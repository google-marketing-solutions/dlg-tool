/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import {afterEach, describe, expect, it, vi} from 'vitest';
import {
  MAX_STORED_ERROR_LENGTH,
  NOT_A_MANAGER_ACCOUNT_ERROR,
  runAdsScriptOrchestrator,
} from '../src/ads/orchestrator';
import {readGlobalConfig, syncDiscoveredAccountsToConfig} from '../src/appscript/config-sheet';
import {
  getOrLockDailyConfig,
  resetTodayPublisherState,
} from '../src/appscript/run-state';
import {
  MAX_ACCOUNT_ATTEMPTS,
  getMccCheckpoint,
} from '../src/shared/sheets/mcc-state';
import {readStagingBatchForMcc} from '../src/shared/sheets/staging-writer';
import {buildMockAccount, setTestDate} from './helpers/factories';
import {
  type MockSheet,
  createMockAdsApp,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

const URL = 'https://docs.google.com/spreadsheets/d/test-sheet/edit';
const DATE = '2026-10-14';
const MCC = {mccId: '111-222-3333', mccName: 'EMEA MCC'};
const BROKEN = '1000000001';
const HEALTHY = '1000000002';

/** GAQL rows for one CAMPAIGN_BUDGET recommendation on campaign 10 of `accountId`. */
function oneRecommendation(query: string, accountId: string): unknown[] {
  const cid = accountId.replace(/\D/g, '');
  if (query.includes('FROM recommendation')) {
    return [
      {
        recommendation: {
          resourceName: `customers/${cid}/recommendations/1`,
          type: 'CAMPAIGN_BUDGET',
          campaign: `customers/${cid}/campaigns/10`,
          campaignBudgetRecommendation: {
            currentBudgetAmountMicros: 50_000_000,
            recommendedBudgetAmountMicros: 80_000_000,
          },
        },
      },
    ];
  }
  if (query.includes("campaign.status = 'ENABLED'") && query.includes('FROM campaign')) {
    return [
      {
        customer: {id: cid},
        campaign: {
          resourceName: `customers/${cid}/campaigns/10`,
          name: `Camp ${cid}`,
          advertisingChannelType: 'SEARCH',
          biddingStrategyType: 'MAXIMIZE_CONVERSIONS',
        },
        campaignBudget: {resourceName: `customers/${cid}/campaignBudgets/1`, amountMicros: 50_000_000, period: 'DAILY'},
      },
    ];
  }
  return [];
}

function setup(
  failRuns: (run: number) => boolean,
  errorMessage: (run: number) => string = (run) => `Simulated GAQL error (run ${run})`
) {
  const spreadsheetApp = createMockSpreadsheetApp(URL);
  const ss = spreadsheetApp.openByUrl(URL);
  const configSheet = ss.insertSheet('Config', 0);
  const childAccounts = [
    buildMockAccount({customerId: '100-000-0001', name: 'Broken Acc', currencyCode: 'USD'}),
    buildMockAccount({customerId: '100-000-0002', name: 'Healthy Acc', currencyCode: 'USD'}),
  ];

  let run = 0;
  let accountsScans = 0;
  const adsApp = createMockAdsApp({
    mccId: MCC.mccId,
    mccName: MCC.mccName,
    childAccounts,
    searchHandler: (query: string, accountId: string) => {
      if (accountId.replace(/\D/g, '') === BROKEN && failRuns(run)) {
        throw new Error(errorMessage(run));
      }
      return oneRecommendation(query, accountId);
    },
  });
  const originalAccounts = adsApp.adsManagerApp.accounts.bind(adsApp.adsManagerApp);
  adsApp.adsManagerApp.accounts = () => {
    accountsScans++;
    return originalAccounts();
  };
  installAppsScriptGlobals({spreadsheetApp, adsApp});

  syncDiscoveredAccountsToConfig(configSheet, MCC, childAccounts);
  configSheet.getRange(22, 1).setValue(true);
  configSheet.getRange(23, 1).setValue(true);
  getOrLockDailyConfig(ss, DATE, () => readGlobalConfig(configSheet));

  const runOnce = () => {
    run++;
    setTestDate(DATE, 9 + run);
    runAdsScriptOrchestrator({spreadsheetUrl: URL});
    return getMccCheckpoint(ss, DATE, MCC.mccId);
  };
  const totalWrites = () =>
    ss.getSheets().reduce((n, s) => n + (s as MockSheet).valueWriteCount, 0);
  return {ss, runOnce, totalWrites, scans: () => accountsScans};
}

describe('Code.gs: failed accounts, idle early exit, non-MCC installs', () => {
  it('caps a failed account\'s stored error message so a long stack trace cannot bloat the checkpoint cell', () => {
    const {runOnce} = setup(() => true, () => 'x'.repeat(5000));
    const cp = runOnce();
    const stored = cp.failedCids?.[BROKEN]?.error ?? '';
    expect(stored.length).toBe(MAX_STORED_ERROR_LENGTH);
    expect(stored.endsWith('…')).toBe(true);
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('retries a transiently failing account on the next run without clearing staging, then completes', () => {
    const {ss, runOnce} = setup((run) => run === 1);

    const cp1 = runOnce();
    expect(cp1.status).toBe('IN_PROGRESS');
    expect(cp1.completedCids).toEqual([HEALTHY]);
    expect(cp1.failedCids?.[BROKEN]).toEqual({error: 'Simulated GAQL error (run 1)', attempts: 1});
    expect(readStagingBatchForMcc(ss, MCC.mccId).recommendations.length).toBe(1);

    const cp2 = runOnce();
    expect(cp2.status).toBe('COMPLETED');
    expect(cp2.completedCids.sort()).toEqual([BROKEN, HEALTHY]);
    expect(cp2.failedCids).toBeUndefined();
    // The healthy account's rec from run 1 survived the retry run
    expect(readStagingBatchForMcc(ss, MCC.mccId).recommendations.length).toBe(2);
  });

  it(`gives up on a permanently failing account after ${MAX_ACCOUNT_ATTEMPTS} runs and completes the MCC`, () => {
    const {ss, runOnce} = setup(() => true);

    for (let attempt = 1; attempt < MAX_ACCOUNT_ATTEMPTS; attempt++) {
      const cp = runOnce();
      expect(cp.status).toBe('IN_PROGRESS');
      expect(cp.failedCids?.[BROKEN].attempts).toBe(attempt);
    }

    const final = runOnce();
    expect(final.status).toBe('COMPLETED');
    expect(final.completedCids.sort()).toEqual([BROKEN, HEALTHY]);
    expect(final.failedCids?.[BROKEN]).toEqual({
      error: `Simulated GAQL error (run ${MAX_ACCOUNT_ATTEMPTS})`,
      attempts: MAX_ACCOUNT_ATTEMPTS,
    });
    expect(readStagingBatchForMcc(ss, MCC.mccId).recommendations.length).toBe(1);
  });

  it('keeps staging when every account failed in the first run', () => {
    const {ss, runOnce} = setup((run) => run <= 2);
    // Make the healthy account fail too in run 1 by unchecking it (and re-locking), then re-check before run 2
    const config = ss.getSheetByName('Config')!;
    config.getRange(23, 1).setValue(false);
    resetTodayPublisherState(ss, DATE);
    getOrLockDailyConfig(ss, DATE, () => readGlobalConfig(config));
    expect(runOnce().completedCids).toEqual([]);
    config.getRange(23, 1).setValue(true);
    resetTodayPublisherState(ss, DATE);
    getOrLockDailyConfig(ss, DATE, () => readGlobalConfig(config));

    // Run 2: progress exists (failed account recorded), so staging is not cleared; healthy account stages its rec
    expect(runOnce().status).toBe('IN_PROGRESS');
    expect(readStagingBatchForMcc(ss, MCC.mccId).recommendations.length).toBe(1);
    // Run 3: broken account recovers; the healthy account's rec from run 2 is still there
    expect(runOnce().status).toBe('COMPLETED');
    expect(readStagingBatchForMcc(ss, MCC.mccId).recommendations.length).toBe(2);
  });

  it('exits idle hourly runs once today is COMPLETED, without scanning accounts or writing', () => {
    const {runOnce, totalWrites, scans} = setup(() => false);
    expect(runOnce().status).toBe('COMPLETED');
    const scansBefore = scans();
    const writesBefore = totalWrites();

    const idle = runOnce();
    expect(idle.status).toBe('COMPLETED');
    expect(idle.completedCids.sort()).toEqual([BROKEN, HEALTHY]);
    expect(scans()).toBe(scansBefore);
    expect(totalWrites()).toBe(writesBefore);
  });

  it('throws a clear error when installed outside a manager (MCC) account', () => {
    const spreadsheetApp = createMockSpreadsheetApp(URL);
    installAppsScriptGlobals({spreadsheetApp, adsApp: createMockAdsApp({mccId: '123-456-7890'})});
    delete (globalThis as Record<string, unknown>).AdsManagerApp;
    setTestDate(DATE, 9);
    expect(() => runAdsScriptOrchestrator({spreadsheetUrl: URL})).toThrow(
      NOT_A_MANAGER_ACCOUNT_ERROR
    );
    expect(spreadsheetApp.openByUrl(URL).getSheetByName('_MCC_1234567890')).toBeNull();
  });
});
