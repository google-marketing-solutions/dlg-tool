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
import {runAdsScriptOrchestrator} from '../src/ads/orchestrator';
import {readGlobalConfig, syncDiscoveredAccountsToConfig} from '../src/appscript/config-sheet';
import {
  getOrLockDailyConfig,
  getRunStateValue,
  withSheetLock,
} from '../src/appscript/run-state';
import {publishFromStagingToData} from '../src/appscript/sheet-ui';
import {readReportRowsFromSheet, writeReportRowsToSheet} from '../src/appscript/sheets/data-writer';
import {createEmptyReportRow} from '../src/shared/schema';
import {
  getMccCheckpoint,
  readMccSheetState,
} from '../src/shared/sheets/mcc-state';
import {buildMockAccount, setTestDate} from './helpers/factories';
import {
  createMockAdsApp,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

describe('Hourly Checkpoint/Resume Engine, Single-Writer Isolation & Multi-MCC Barrier', () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it('writes discovery state exclusively to _MCC_<id> on first trigger (single-writer boundary) and the Sheet-bound script syncs it to Config', () => {
    const url = 'https://docs.google.com/spreadsheets/d/test-sheet/edit';
    const spreadsheetApp = createMockSpreadsheetApp(url);
    let gaqlCalls = 0;

    const adsApp = createMockAdsApp({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      isPreview: false,
      childAccounts: [
        buildMockAccount({ customerId: '100-000-0001', name: 'UK Account', currencyCode: 'GBP' }),
      ],
      searchHandler: () => {
        gaqlCalls++;
        return [];
      },
    });
    installAppsScriptGlobals({ spreadsheetApp, adsApp });
    setTestDate('2026-10-14', 10);

    runAdsScriptOrchestrator({ spreadsheetUrl: url });

    expect(gaqlCalls).toBe(0);

    const ss = spreadsheetApp.openByUrl(url);
    // Code.gs never touches Config or _RunState directly; it writes only to its own _MCC_1112223333 tab
    expect(ss.getSheetByName('Config')).toBeNull();
    expect(ss.getSheetByName('_RunState')).toBeNull();
    const mccPayload = readMccSheetState(ss, '111-222-3333');
    expect(mccPayload?.mcc.mccName).toBe('EMEA MCC');
    expect(mccPayload?.childAccounts?.length).toBe(1);
    expect(mccPayload?.checkpoint).toBeUndefined();
    expect(mccPayload?.statusMessage).toContain('registered');

    // When the Sheet-bound script runs its scheduled trigger (with un-checked or newly synced accounts), it syncs _MCC_1112223333 into Config!
    publishFromStagingToData({ spreadsheet: ss });
    const configSheet = ss.getSheetByName('Config')!;
    expect(configSheet).not.toBeNull();
    expect(configSheet.getRange(11, 2).getValue()).toBe('111-222-3333');
    expect(configSheet.getRange(22, 2).getValue()).toBe('100-000-0001');
  });

  it('checkpoints progress when remaining time < 240s and resumes seamlessly on the next hourly trigger', () => {
    const url = 'https://docs.google.com/spreadsheets/d/test-sheet/edit';
    const spreadsheetApp = createMockSpreadsheetApp(url);
    const ss = spreadsheetApp.openByUrl(url);
    const configSheet = ss.insertSheet('Config', 0);

    const childAccounts = [
      buildMockAccount({ customerId: '100-000-0001', name: 'Acc 1', currencyCode: 'USD' }),
      buildMockAccount({ customerId: '100-000-0002', name: 'Acc 2', currencyCode: 'USD' }),
    ];

    // Pre-register MCC, select both accounts, and lock daily config in _RunState so the next run is a FULL_RUN
    syncDiscoveredAccountsToConfig(
      configSheet,
      { mccId: '111-222-3333', mccName: 'EMEA MCC' },
      childAccounts
    );
    configSheet.getRange(22, 1).setValue(true);
    configSheet.getRange(23, 1).setValue(true);
    getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));
    getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));

    const adsApp = createMockAdsApp({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      remainingTime: 1800,
      childAccounts,
      searchHandler: (query: string, accountId: string) => {
        // After processing the first account, drop remainingTime below 240s to trigger checkpoint!
        if (accountId.replace(/\D/g, '') === '1000000001' && query.includes('DURING LAST_7_DAYS')) {
          adsApp.setRemainingTime(180);
        }
        if (query.includes('FROM recommendation')) {
          return [
            {
              recommendation: {
                resourceName: `customers/${accountId.replace(/\D/g, '')}/recommendations/1`,
                type: 'CAMPAIGN_BUDGET',
                campaign: `customers/${accountId.replace(/\D/g, '')}/campaigns/10`,
                campaignBudgetRecommendation: {
                  currentBudgetAmountMicros: 50_000_000,
                  recommendedBudgetAmountMicros: 80_000_000,
                },
                impact: {
                  baseMetrics: { costMicros: 350_000_000, conversions: 10 },
                  potentialMetrics: { costMicros: 560_000_000, conversions: 16 },
                },
              },
            },
          ];
        }
        if (query.includes("campaign.status = 'ENABLED'") && query.includes('FROM campaign')) {
          return [
            {
              customer: { id: accountId.replace(/\D/g, '') },
              campaign: {
                resourceName: `customers/${accountId.replace(/\D/g, '')}/campaigns/10`,
                name: `Camp for ${accountId}`,
                advertisingChannelType: 'SEARCH',
                biddingStrategyType: 'MAXIMIZE_CONVERSIONS',
              },
              campaignBudget: {
                resourceName: `customers/${accountId.replace(/\D/g, '')}/campaignBudgets/1`,
                name: 'Daily Budget',
                amountMicros: 50_000_000,
                period: 'DAILY',
              },
            },
          ];
        }
        return [];
      },
    });
    installAppsScriptGlobals({ spreadsheetApp, adsApp });

    // Hour 1: Processes Acc 1, then hits remainingTime=180 (<240) before Acc 2 and checkpoints
    setTestDate('2026-10-14', 10);
    runAdsScriptOrchestrator({
      spreadsheetUrl: url,
    });
    const cp1 = getMccCheckpoint(ss, '2026-10-14', '111-222-3333');
    expect(cp1.status).toBe('IN_PROGRESS');
    expect(cp1.completedCids).toEqual(['1000000001']);
    expect(ss.getSheetByName('Data')).toBeNull(); // Data is untouched while partial!

    // Hour 2: Fresh 1800s budget -> resumes at Acc 2 and completes MCC staging in _MCC_1112223333
    setTestDate('2026-10-14', 11);
    adsApp.setRemainingTime(1800);
    runAdsScriptOrchestrator({
      spreadsheetUrl: url,
    });
    const cp2 = getMccCheckpoint(ss, '2026-10-14', '111-222-3333');
    expect(cp2.status).toBe('COMPLETED');
    expect(cp2.completedCids).toEqual(['1000000001', '1000000002']);

    // Verify AdsManagerApp.select was called exactly once per account across the two hourly runs
    expect(adsApp.selectCounts.get('100-000-0001')).toBe(1);
    expect(adsApp.selectCounts.get('100-000-0002')).toBe(1);

    // The Sheet-bound script's scheduled trigger runs the Final Join & Publish into Data
    publishFromStagingToData({
      spreadsheet: ss,
    });
    expect(getRunStateValue(ss, 'PUBLISH_STATUS_2026-10-14')).toBe('PUBLISHED');

    const dataSheet = ss.getSheetByName('Data')!;
    expect(dataSheet).not.toBeNull();
    expect(readReportRowsFromSheet(dataSheet).length).toBe(2);
    expect(dataSheet.getMaxRows()).toBe(4); // 1 header + 1 schema dummy row + 2 recommendation rows
  });

  it('publishes each MCC as soon as it completes, marks pending MCCs as stale without republish loops, and does not resurrect stale rows for 0-rec completed MCCs', () => {
    const url = 'https://docs.google.com/spreadsheets/d/test-sheet/edit';
    const spreadsheetApp = createMockSpreadsheetApp(url);
    const ss = spreadsheetApp.openByUrl(url);
    const configSheet = ss.insertSheet('Config', 0);

    const accA = buildMockAccount({ customerId: '100-000-0001', name: 'EMEA 1', currencyCode: 'USD' });
    const accB = buildMockAccount({ customerId: '200-000-0001', name: 'AMER 1', currencyCode: 'USD' });
    const accC = buildMockAccount({ customerId: '300-000-0001', name: 'APAC 1', currencyCode: 'USD' });

    syncDiscoveredAccountsToConfig(configSheet, { mccId: '111-111-1111', mccName: 'EMEA' }, [accA]);
    syncDiscoveredAccountsToConfig(configSheet, { mccId: '222-222-2222', mccName: 'AMER' }, [accB]);
    syncDiscoveredAccountsToConfig(configSheet, { mccId: '333-333-3333', mccName: 'APAC' }, [accC]);

    // Check EMEA (row 22) and AMER (row 23), leave APAC (row 24) unchecked -> APAC has Selected Accounts = 0 (Inactive / Skipped!)
    configSheet.getRange(22, 1).setValue(true);
    configSheet.getRange(23, 1).setValue(true);
    configSheet.getRange('B5').setValue('EUR');
    getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));

    // Seed yesterday's Data tab with BOTH an old EMEA row (to test no resurrection) and an old AMER row (to test stale fallback)
    const dataSheet = ss.insertSheet('Data');
    const emeaYesterdayRow = createEmptyReportRow();
    emeaYesterdayRow.mccId = '111-111-1111';
    emeaYesterdayRow.mccName = 'EMEA';
    emeaYesterdayRow.accountId = '100-000-0001';
    emeaYesterdayRow.accountName = 'EMEA 1';
    emeaYesterdayRow.campaignId = '888';

    const amerStaleRow = createEmptyReportRow();
    amerStaleRow.mccId = '222-222-2222';
    amerStaleRow.mccName = 'AMER';
    amerStaleRow.accountId = '200-000-0001';
    amerStaleRow.accountName = 'AMER 1';
    amerStaleRow.currencyCode = 'USD';
    amerStaleRow.campaignId = '999';
    amerStaleRow.recommendationCurrentBudgetAmount = 100;
    writeReportRowsToSheet(dataSheet, [emeaYesterdayRow, amerStaleRow]);
    configSheet.getRange('B5').setValue('EUR');
    const runStateSheet = ss.getSheetByName('_RunState') ?? ss.insertSheet('_RunState', 200);
    runStateSheet.formulaEvaluator = (formula: string) => (formula.includes('USDEUR') ? 0.92 : '#N/A');

    // EMEA completes today with 0 recommendations
    const adsAppEmea = createMockAdsApp({
      mccId: '111-111-1111',
      mccName: 'EMEA',
      childAccounts: [accA],
      searchHandler: () => [],
    });
    installAppsScriptGlobals({ spreadsheetApp, adsApp: adsAppEmea });

    // EMEA finishes its run -> stages immediately
    setTestDate('2026-10-14', 14);
    runAdsScriptOrchestrator({
      spreadsheetUrl: url,
    });
    expect(getMccCheckpoint(ss, '2026-10-14', '111-111-1111').status).toBe('COMPLETED');

    // Publisher immediately publishes EMEA's fresh status alongside AMER's stale snapshot (no cutoff wait!)
    publishFromStagingToData({
      spreadsheet: ss,
    });
    expect(getRunStateValue(ss, 'PUBLISH_STATUS_2026-10-14')).toBe('PUBLISHED_WITH_STALE_WARNING');
    // No resurrection: EMEA completed today with 0 recs -> contributes 0 rows (yesterday's EMEA row 888 is NOT resurrected!),
    // while pending AMER keeps its 1 fallback row 999 marked stale!
    const publishedRows = readReportRowsFromSheet(dataSheet);
    expect(publishedRows.length).toBe(1);
    expect(publishedRows[0].mccId).toBe('222-222-2222');
    expect(publishedRows[0].mccSyncStatus).toContain('WARNING: Stale snapshot');
    expect(publishedRows[0].currencyCodeGlobal).toBe('EUR');
    expect(publishedRows[0].recommendationCurrentBudgetAmountGlobal).toBeCloseTo(92);
    expect(String(configSheet.getRange('B2').getValue())).toContain('WARNING');

    // Idle tick is a no-op: next 15-minute trigger with no new MCC completions must be a NO_OP (no infinite republish loop!)
    const writesBeforeIdle = dataSheet.valueWriteCount;
    publishFromStagingToData({
      spreadsheet: ss,
    });
    expect(dataSheet.valueWriteCount).toBe(writesBeforeIdle);

    // When lagging AMER finishes later today (with 0 recs), the signature changes and the Sheet-bound script publishes fresh data!
    setTestDate('2026-10-14', 15);
    const adsAppAmer = createMockAdsApp({
      mccId: '222-222-2222',
      mccName: 'AMER',
      childAccounts: [accB],
      searchHandler: () => [],
    });
    installAppsScriptGlobals({ spreadsheetApp, adsApp: adsAppAmer });
    runAdsScriptOrchestrator({
      spreadsheetUrl: url,
    });

    publishFromStagingToData({
      spreadsheet: ss,
    });
    expect(getRunStateValue(ss, 'PUBLISH_STATUS_2026-10-14')).toBe('PUBLISHED');
    expect(readReportRowsFromSheet(dataSheet).length).toBe(0);
    expect(String(configSheet.getRange('B2').getValue())).toContain('🟢 PUBLISHED');
  });

  it('isolates per-account errors with try/catch so a failing child account never wedges the MCC', () => {
    const url = 'https://docs.google.com/spreadsheets/d/test-sheet/edit';
    const spreadsheetApp = createMockSpreadsheetApp(url);
    const ss = spreadsheetApp.openByUrl(url);
    const configSheet = ss.insertSheet('Config', 0);

    const childAccounts = [
      buildMockAccount({ customerId: '100-000-0001', name: 'Broken Acc', currencyCode: 'USD' }),
      buildMockAccount({ customerId: '100-000-0002', name: 'Healthy Acc', currencyCode: 'USD' }),
    ];
    syncDiscoveredAccountsToConfig(
      configSheet,
      { mccId: '111-222-3333', mccName: 'EMEA MCC' },
      childAccounts
    );
    configSheet.getRange(22, 1).setValue(true);
    configSheet.getRange(23, 1).setValue(true);
    getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));

    const adsApp = createMockAdsApp({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      childAccounts,
      searchHandler: (_query: string, accountId: string) => {
        if (accountId.replace(/\D/g, '') === '1000000001') {
          throw new Error('Simulated transient GAQL quota error');
        }
        return [];
      },
    });
    installAppsScriptGlobals({ spreadsheetApp, adsApp });
    setTestDate('2026-10-14', 10);

    runAdsScriptOrchestrator({
      spreadsheetUrl: url,
    });

    // The healthy account completes; the failing one is retried on the next hourly run (not completed)
    const cp = getMccCheckpoint(ss, '2026-10-14', '111-222-3333');
    expect(cp.status).toBe('IN_PROGRESS');
    expect(cp.completedCids).toEqual(['1000000002']);
    expect(cp.failedCids?.['1000000001']).toEqual({
      error: 'Simulated transient GAQL quota error',
      attempts: 1,
    });
  });

  it('runs the Code.gs orchestrator (discovery and full run) without any LockService', () => {
    const url = 'https://docs.google.com/spreadsheets/d/test-sheet/edit';
    const spreadsheetApp = createMockSpreadsheetApp(url);
    const ss = spreadsheetApp.openByUrl(url);
    const childAccounts = [
      buildMockAccount({ customerId: '100-000-0001', name: 'Acc 1', currencyCode: 'USD' }),
    ];
    const adsApp = createMockAdsApp({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      childAccounts,
      searchHandler: () => [],
    });
    // Google Ads scripts have no document lock, so Code.gs must never call withSheetLock
    const {lockService} = installAppsScriptGlobals({ spreadsheetApp, adsApp });

    setTestDate('2026-10-14', 9);
    installAppsScriptGlobals({ spreadsheetApp, adsApp, lockService: null });
    runAdsScriptOrchestrator({ spreadsheetUrl: url });
    expect(readMccSheetState(ss, '111-222-3333')?.checkpoint).toBeUndefined();

    // Sheet-bound script locks today's config in _RunState (using its document lock)
    installAppsScriptGlobals({ spreadsheetApp, adsApp, lockService });
    const configSheet = ss.insertSheet('Config', 0);
    syncDiscoveredAccountsToConfig(configSheet, { mccId: '111-222-3333', mccName: 'EMEA MCC' }, childAccounts);
    configSheet.getRange(22, 1).setValue(true);
    getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));

    // Code.gs runs with lockService: null and reads the frozen config from _RunState
    installAppsScriptGlobals({ spreadsheetApp, adsApp, lockService: null });
    setTestDate('2026-10-14', 10);
    runAdsScriptOrchestrator({ spreadsheetUrl: url });
    const cp = getMccCheckpoint(ss, '2026-10-14', '111-222-3333');
    expect(cp.status).toBe('COMPLETED');
    expect(cp.completedCids).toEqual(['1000000001']);
  });

  it('acquires and releases the document lock around withSheetLock', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    const {lockService} = installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();

    let heldDuringExecution = false;
    const result = withSheetLock(ss, 'SHEET_TEST', () => {
      heldDuringExecution = lockService!.getDocumentLock().hasLock();
      return 42;
    });

    expect(result).toBe(42);
    expect(heldDuringExecution).toBe(true);
    expect(lockService!.stats).toEqual({ acquireCount: 1, releaseCount: 1, held: false });
    // No lock cell is written to _RunState anymore
    expect(ss.getSheetByName('_RunState')).toBeNull();
  });

  it('releases the document lock when the locked function throws', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    const {lockService} = installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();

    expect(() =>
      withSheetLock(ss, 'SHEET_TEST', () => {
        throw new Error('boom');
      })
    ).toThrow('boom');
    expect(lockService!.stats).toEqual({ acquireCount: 1, releaseCount: 1, held: false });

    // The lock is usable again afterwards
    withSheetLock(ss, 'SHEET_TEST', () => undefined);
    expect(lockService!.stats.acquireCount).toBe(2);
  });

  it('does not re-acquire the document lock for nested (re-entrant) calls', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    const {lockService} = installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();

    withSheetLock(ss, 'OUTER', () => {
      withSheetLock(ss, 'INNER', () => {
        expect(lockService!.getDocumentLock().hasLock()).toBe(true);
      });
      // The inner call must not release the outer lock
      expect(lockService!.getDocumentLock().hasLock()).toBe(true);
    });
    expect(lockService!.stats).toEqual({ acquireCount: 1, releaseCount: 1, held: false });
  });

  it('throws a clear error when no LockService document lock is available', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp, lockService: null });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    let ran = false;

    expect(() =>
      withSheetLock(ss, 'SHEET_TEST', () => {
        ran = true;
      })
    ).toThrow(/must run in the Sheet-bound script \(LockService document lock unavailable\)/);
    expect(ran).toBe(false);
  });

  it('uses injected timeZone or falls back to UTC when spreadsheet.getSpreadsheetTimeZone() returns null or empty string', () => {
    const url = 'https://docs.google.com/spreadsheets/d/null-tz-sheet/edit';
    const spreadsheetApp = createMockSpreadsheetApp(url);
    const ss = spreadsheetApp.openByUrl(url);
    ss.setSpreadsheetTimeZone(null);

    const adsApp = createMockAdsApp({
      mccId: '111-222-3333',
      mccName: 'Null TZ MCC',
      isPreview: false,
      childAccounts: [
        buildMockAccount({ customerId: '100-000-0001', name: 'UK Account', currencyCode: 'GBP' }),
      ],
    });
    installAppsScriptGlobals({ spreadsheetApp, adsApp });
    setTestDate('2026-10-14', 10);

    let capturedTz: string | undefined;
    const origFormatDate = Utilities.formatDate;
    Utilities.formatDate = (date: Date, tz: string, format: string) => {
      capturedTz = tz;
      return origFormatDate(date, tz, format);
    };

    expect(() =>
      runAdsScriptOrchestrator({ spreadsheetUrl: url, timeZone: 'Europe/London' })
    ).not.toThrow();
    expect(capturedTz).toBe('Europe/London');

    expect(() => runAdsScriptOrchestrator({ spreadsheetUrl: url })).not.toThrow();
    expect(capturedTz).toBe('UTC');

    expect(() => publishFromStagingToData({ spreadsheet: ss })).not.toThrow();
    expect(capturedTz).toBe('UTC');

    ss.setSpreadsheetTimeZone('');
    expect(() => runAdsScriptOrchestrator({ spreadsheetUrl: url })).not.toThrow();
    expect(capturedTz).toBe('UTC');
    expect(() => publishFromStagingToData({ spreadsheet: ss })).not.toThrow();
    expect(capturedTz).toBe('UTC');
  });
});
