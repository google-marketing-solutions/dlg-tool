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
import {getRunStateValue} from '../src/appscript/run-state';
import {
  applyConfigChangesNow,
  openDataStudioCreator,
  publishFromStagingToData,
  seedSyntheticDemoData,
} from '../src/appscript/sheet-ui';
import {readReportRowsFromSheet, writeReportRowsToSheet} from '../src/appscript/sheets/data-writer';
import {
  type MccCheckpointState,
  clearMccSheetCheckpoint,
  getMccCheckpoint,
  writeMccSheetState,
} from '../src/shared/sheets/mcc-state';
import {setTestDate} from './helpers/factories';
import {
  type MockSheet,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

const D1 = '2026-10-26';
const D2 = '2026-10-27';
const MCC_A = {mccId: '999-888-7777', mccName: '[Demo] Global Retail MCC', cids: ['1012023030', '2023034040', '3034045050']};
const MCC_B = {
  mccId: '888-777-6666',
  mccName: '[Demo] DACH & Nordics Aviation MCC',
  cids: ['4045056060', '5056067070'],
};

afterEach(() => {
  vi.useRealTimers();
});

function seed() {
  const spreadsheetApp = createMockSpreadsheetApp();
  installAppsScriptGlobals({spreadsheetApp});
  const ss = spreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.insertSheet('Config', 0);
  setTestDate(D1);
  seedSyntheticDemoData({spreadsheet: ss});
  const dataRows = () => readReportRowsFromSheet(ss.getSheetByName('Data'));
  const totalWrites = () => ss.getSheets().reduce((n, s) => n + (s as MockSheet).valueWriteCount, 0);
  return {spreadsheetApp, ss, configSheet, dataRows, totalWrites};
}

function writeCheckpoint(
  ss: ReturnType<typeof seed>['ss'],
  mcc: typeof MCC_A,
  dateStr: string,
  extra: Partial<MccCheckpointState> = {}
) {
  writeMccSheetState(ss, {
    mcc: {mccId: mcc.mccId, mccName: mcc.mccName},
    checkpoint: {mccId: mcc.mccId.replace(/-/g, ''), dateStr, status: 'COMPLETED', completedCids: mcc.cids, ...extra},
    lastCompletedDate: dateStr,
  });
}

describe('Publisher: failed accounts, whole-publish lock, single FX pass', () => {
  it("keeps a given-up account's previous rows with a warning and shows the failed count", () => {
    const {ss, configSheet, dataRows, totalWrites} = seed();
    expect(dataRows().length).toBe(13);
    const failedRowsBefore = dataRows().filter((r) => r.accountId === '404-505-6060').length;
    expect(failedRowsBefore).toBe(2);

    // Next day: both MCCs COMPLETED, but MCC B gave up on 404-505-6060 after 3 attempts
    setTestDate(D2);
    writeCheckpoint(ss, MCC_A, D2);
    writeCheckpoint(ss, MCC_B, D2, {failedCids: {'4045056060': {error: 'Quota exceeded', attempts: 3}}});

    publishFromStagingToData({spreadsheet: ss});
    expect(getRunStateValue(ss, `PUBLISH_STATUS_${D2}`)).toBe('PUBLISHED_WITH_STALE_WARNING');
    const rows = dataRows();
    expect(rows.length).toBe(13);
    const failedRows = rows.filter((r) => r.accountId === '404-505-6060');
    expect(failedRows.length).toBe(2);
    for (const r of failedRows) {
      expect(r.mccSyncStatus).toBe(
        `WARNING: Account 404-505-6060 failed (Quota exceeded); showing snapshot from ${D1}`
      );
      expect(r.mccLastUpdatedDate).toBe(D1);
    }
    for (const r of rows.filter((row) => row.accountId !== '404-505-6060')) {
      expect(r.mccSyncStatus).toBe('OK');
      expect(r.mccLastUpdatedDate).toBe(D2);
    }
    expect(String(configSheet.getRange('B2').getValue())).toContain(
      "1 account(s) failed in today's run (showing their previous snapshot): 404-505-6060"
    );

    // failedCids are part of the signature: the next idle tick is a NO_OP (zero writes)
    const writesBeforeIdle = totalWrites();
    publishFromStagingToData({spreadsheet: ss});
    expect(totalWrites()).toBe(writesBeforeIdle);
  });

  it('shows nothing for a given-up account that has no previous rows', () => {
    const {ss, dataRows} = seed();
    // Data has no previous rows for 505-606-7070
    const dataSheet = ss.getSheetByName('Data')!;
    writeReportRowsToSheet(dataSheet, dataRows().filter((r) => r.accountId !== '505-606-7070'));
    expect(dataRows().length).toBe(12);

    setTestDate(D2);
    writeCheckpoint(ss, MCC_A, D2);
    writeCheckpoint(ss, MCC_B, D2, {failedCids: {'5056067070': {error: 'boom', attempts: 3}}});
    publishFromStagingToData({spreadsheet: ss});
    expect(dataRows().filter((r) => r.accountId === '505-606-7070')).toEqual([]);
    expect(dataRows().length).toBe(12);
  });

  it('does zero sheet writes on an idle tick', () => {
    const {ss, totalWrites} = seed();
    const before = totalWrites();
    publishFromStagingToData({spreadsheet: ss});
    expect(totalWrites()).toBe(before);
  });

  it('resolves FX once across all MCCs and caches the unified rate map in _RunState', () => {
    const {ss} = seed();
    const runState = ss.getSheetByName('_RunState') as MockSheet;
    const evaluatedFormulas: string[] = [];
    runState.formulaEvaluator = (formula: string) => {
      evaluatedFormulas.push(formula);
      return '#N/A';
    };

    publishFromStagingToData({spreadsheet: ss, force: true});
    // One GOOGLEFINANCE evaluation per non-target currency (CHF, EUR, GBP -> USD)
    expect(evaluatedFormulas.length).toBe(3);
    const target = JSON.parse(getRunStateValue(ss, `DAILY_CONFIG_LOCK_${D1}`)!).targetCurrency;
    const cached = JSON.parse(getRunStateValue(ss, `FX_RATES_${D1}_${target}`)!);
    expect(Object.keys(cached).sort()).toEqual(['CHF', 'EUR', 'GBP', 'USD']);
  });
});

describe('Apply Config Changes Now: only reset what a change invalidates', () => {
  it('keeps checkpoints for a Target Currency change, republishes fresh, and stays fresh on the next tick', () => {
    const {ss, configSheet, dataRows, totalWrites} = seed();
    configSheet.getRange('B5').setValue('EUR');
    applyConfigChangesNow();

    expect(getMccCheckpoint(ss, D1, MCC_A.mccId).status).toBe('COMPLETED');
    expect(getMccCheckpoint(ss, D1, MCC_B.mccId).status).toBe('COMPLETED');
    expect(JSON.parse(getRunStateValue(ss, `DAILY_CONFIG_LOCK_${D1}`)!).targetCurrency).toBe('EUR');
    expect(dataRows().length).toBe(13);
    expect(dataRows().every((r) => r.mccSyncStatus === 'OK' && r.currencyCodeGlobal === 'EUR')).toBe(true);

    const writesBeforeIdle = totalWrites();
    publishFromStagingToData({spreadsheet: ss});
    expect(totalWrites()).toBe(writesBeforeIdle);
    expect(dataRows().every((r) => r.mccSyncStatus === 'OK')).toBe(true);
  });

  it('ignores mid-day Config edits (currency, Shopping toggle, and account checkboxes) until Apply Config Changes Now is called', () => {
    const {ss, configSheet, dataRows, totalWrites} = seed();
    expect(dataRows().length).toBe(13);

    // User edits Config mid-day WITHOUT clicking "Apply Config Changes Now":
    // - Changes Target Currency to EUR (B5)
    // - Enables Exclude Shopping & Feed-PMax (B6)
    // - Unchecks all child accounts (rows 22..26)
    configSheet.getRange('B5').setValue('EUR');
    configSheet.getRange('B6').setValue(true);
    for (let r = 22; r <= 26; r++) {
      configSheet.getRange(r, 1).setValue(false);
    }

    // Next 15-minute scheduled publish reads DAILY_CONFIG_LOCK_<D1> from _RunState -> zero writes, Data untouched!
    const writesBefore = totalWrites();
    publishFromStagingToData({spreadsheet: ss});
    expect(totalWrites()).toBe(writesBefore);
    expect(dataRows().length).toBe(13);
    expect(dataRows()[0].currencyCodeGlobal).toBe('USD');
  });

  it('resets all active MCCs when Exclude Shopping & Feed-PMax changes', () => {
    const {ss, configSheet, dataRows} = seed();
    configSheet.getRange('B6').setValue(true);
    applyConfigChangesNow();

    expect(getMccCheckpoint(ss, D1, MCC_A.mccId).status).toBe('IN_PROGRESS');
    expect(getMccCheckpoint(ss, D1, MCC_B.mccId).status).toBe('IN_PROGRESS');
    expect(dataRows().every((r) => r.mccSyncStatus === `WARNING: Stale snapshot from ${D1}`)).toBe(true);
    expect(String(configSheet.getRange('B2').getValue())).toContain('stale snapshot');
  });

  it('resets only the MCC whose newly checked account is missing from its completed accounts', () => {
    const {ss} = seed();
    // MCC B's run today did not include 404-505-6060 (it was unchecked then); it is checked now
    writeCheckpoint(ss, MCC_B, D1, {completedCids: ['5056067070']});
    applyConfigChangesNow();

    expect(getMccCheckpoint(ss, D1, MCC_A.mccId).status).toBe('COMPLETED');
    expect(getMccCheckpoint(ss, D1, MCC_B.mccId).status).toBe('IN_PROGRESS');
  });

  it('never marks a reset MCC fresh, even for a forced publish', () => {
    const {ss, dataRows} = seed();
    clearMccSheetCheckpoint(ss, MCC_A.mccId);
    publishFromStagingToData({spreadsheet: ss, force: true});
    expect(getRunStateValue(ss, `PUBLISH_STATUS_${D1}`)).toBe('PUBLISHED_WITH_STALE_WARNING');
    for (const r of dataRows()) {
      expect(r.mccSyncStatus).toBe(r.mccId === MCC_A.mccId ? `WARNING: Stale snapshot from ${D1}` : 'OK');
    }
  });
});

describe('Publisher: stale _Stage_* fallback only without progress today', () => {
  function seedWithoutMccBDataRows() {
    const seeded = seed();
    const dataSheet = seeded.ss.getSheetByName('Data')!;
    writeReportRowsToSheet(dataSheet, seeded.dataRows().filter((r) => r.mccId !== MCC_B.mccId));
    expect(seeded.dataRows().some((r) => r.mccId === MCC_B.mccId)).toBe(false);
    return seeded;
  }

  it('publishes nothing for an MCC that is mid-run today and has no Data rows', () => {
    const {ss, configSheet, dataRows} = seedWithoutMccBDataRows();
    setTestDate(D2);
    writeCheckpoint(ss, MCC_A, D2);
    // MCC B flushed its first account today, so _Stage_* holds a partial batch
    writeMccSheetState(ss, {
      mcc: {mccId: MCC_B.mccId, mccName: MCC_B.mccName},
      checkpoint: {mccId: '8887776666', dateStr: D2, status: 'IN_PROGRESS', completedCids: [MCC_B.cids[0]]},
      lastCompletedDate: D1,
    });

    publishFromStagingToData({spreadsheet: ss});
    expect(getRunStateValue(ss, `PUBLISH_STATUS_${D2}`)).toBe('PUBLISHED_WITH_STALE_WARNING');
    expect(dataRows().filter((r) => r.mccId === MCC_B.mccId)).toEqual([]);
    expect(dataRows().filter((r) => r.mccId === MCC_A.mccId).length).toBe(10);
    expect(String(configSheet.getRange('B2').getValue())).toContain(`pending MCC(s): ${MCC_B.mccId}`);
  });

  it('still falls back to _Stage_* when the MCC has made no progress today', () => {
    const {ss, dataRows} = seedWithoutMccBDataRows();
    setTestDate(D2);
    writeCheckpoint(ss, MCC_A, D2);
    // MCC B has no checkpoint for today: _Stage_* still holds its completed run from D1

    publishFromStagingToData({spreadsheet: ss});
    const rowsB = dataRows().filter((r) => r.mccId === MCC_B.mccId);
    expect(rowsB.length).toBe(3);
    expect(rowsB.every((r) => r.mccSyncStatus === `WARNING: Stale snapshot from ${D1}`)).toBe(true);
  });
});

describe('Create New Data Studio Dashboard', () => {
  it('creates and initializes the Data tab before building the URL with its real gid', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({spreadsheetApp});
    const ss = spreadsheetApp.getActiveSpreadsheet();
    expect(ss.getSheetByName('Data')).toBeNull();

    openDataStudioCreator();

    const dataSheet = ss.getSheetByName('Data')!;
    expect(dataSheet.getLastRow()).toBe(2);
    expect(dataSheet.getRange(1, 1).getValue()).toBe('recommendationNumber');
    expect(String(dataSheet.getRange(2, 8).getValue())).toBe('!!! SCHEMA DEFINITION (FILTER ME OUT)');
    const html = spreadsheetApp.modals[spreadsheetApp.modals.length - 1].html;
    expect(html).toContain(`worksheetId=${dataSheet.getSheetId()}`);
  });
});

describe('Publisher: accounts discovered under more than one MCC', () => {
  it('warns in the status banner and stops warning once the duplicate is gone', () => {
    const {ss, configSheet} = seed();
    // A sub-manager account that also has the script installed re-discovers 101-202-3030 from MCC A
    const subMcc = {mccId: '777-666-5555', mccName: 'Sub-manager'};
    writeMccSheetState(ss, {
      mcc: subMcc,
      childAccounts: [{customerId: '101-202-3030', name: '[Demo] US Search & AI Max', currencyCode: 'USD'}],
    });

    publishFromStagingToData({spreadsheet: ss});
    const banner = String(configSheet.getRange('B2').getValue());
    expect(banner).toContain('1 account(s) are under more than one MCC');
    expect(banner).toContain('101-202-3030 (MCCs 777-666-5555, 999-888-7777)');

    writeMccSheetState(ss, {mcc: subMcc, childAccounts: []});
    publishFromStagingToData({spreadsheet: ss});
    expect(String(configSheet.getRange('B2').getValue())).not.toContain('more than one MCC');
  });
});

describe('Publisher: outdated Google Ads Script copies', () => {
  it("warns when an MCC's _MCC_ tab was written by a different script version", () => {
    const {ss, configSheet} = seed();
    const mccTab = ss.getSheetByName('_MCC_9998887777')!;
    const keys = mccTab.getRange(1, 1, mccTab.getLastRow(), 1).getValues().map((r) => r[0]);
    mccTab.getRange(keys.indexOf('SCRIPT_VERSION') + 1, 2).setValue('1.9.0');

    publishFromStagingToData({spreadsheet: ss});
    expect(String(configSheet.getRange('B2').getValue())).toContain(
      'Google Ads Script out of date in MCC(s) 999-888-7777 (v1.9.0)'
    );
  });
});
