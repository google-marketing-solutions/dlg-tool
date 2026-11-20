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

import {describe, expect, it} from 'vitest';
import {
  CID_TABLE_HEADERS,
  MCC_TABLE_HEADERS,
  assertConfigLayout,
  getMccSelectionSummary,
  getSelectedCidsForMcc,
  readGlobalConfig,
  syncAllMccSheetsToConfig,
  syncDiscoveredAccountsToConfig,
} from '../src/appscript/config-sheet';
import {
  getOrLockDailyConfig,
  readRunStateMap,
  resetTodayRunState,
  setRunStateValue,
  setRunStateValues,
  withSheetLock,
} from '../src/appscript/run-state';
import {
  getMccCheckpoint,
  readMccSheetState,
  writeMccSheetState,
} from '../src/shared/sheets/mcc-state';
import {buildMockAccount} from './helpers/factories';
import {createMockSpreadsheetApp, installAppsScriptGlobals} from './mocks/google-apps-script.mock';

describe('Config Tab Manager, Auto-Registration, Checkbox CID Filter & Daily Lock', () => {
  it('auto-registers a new MCC, appends child CIDs without duplicates, defaults Target Currency to first registered currency, and updates dropdown', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    const mccA = { mccId: '111-222-3333', mccName: 'EMEA Retail MCC' };
    const accountsPass1 = [
      buildMockAccount({ customerId: '123-456-7890', name: 'UK Search', currencyCode: 'GBP' }),
      buildMockAccount({ customerId: '234-567-8901', name: 'DE Search', currencyCode: 'EUR' }),
    ];

    const firstSync = syncDiscoveredAccountsToConfig(configSheet, mccA, accountsPass1);
    expect(firstSync.isNewlyRegisteredMcc).toBe(true);
    expect(firstSync.newlyAddedCidsCount).toBe(2);
    // Newly discovered MCCs and CIDs default to unselected (false)
    expect(configSheet.getRange(11, 1).getValue()).toBe(false);
    expect(configSheet.getRange(22, 1).getValue()).toBe(false);
    expect(configSheet.getRange(23, 1).getValue()).toBe(false);

    // Target Currency (B5) defaults to the first currency registered by an MCC ('GBP'), and dropdown has no 'NONE (Native)'
    expect(configSheet.getRange('B5').getValue()).toBe('GBP');
    expect(readGlobalConfig(configSheet).targetCurrency).toBe('GBP');
    const validation = configSheet.getRange('B5').getDataValidation();
    expect(validation?.values).toEqual(['EUR', 'GBP']);

    // User changes selection to 'EUR' -> preserved on subsequent syncs
    configSheet.getRange('B5').setValue('EUR');

    // Second trigger on the same MCC where 234-567-8901 was unlinked/removed, 123-456-7890 remains,
    // and 345-678-9012 is newly added (with Select All in MCC checked -> auto-checked!)
    configSheet.getRange(11, 1).setValue(true);
    const accountsPass2 = [
      buildMockAccount({ customerId: '123-456-7890', name: 'UK Search', currencyCode: 'GBP' }),
      buildMockAccount({ customerId: '345-678-9012', name: 'US Search', currencyCode: 'USD' }),
    ];
    const secondSync = syncDiscoveredAccountsToConfig(configSheet, mccA, accountsPass2);
    expect(secondSync.isNewlyRegisteredMcc).toBe(false);
    expect(secondSync.newlyAddedCidsCount).toBe(1);
    expect(secondSync.removedCidsCount).toBe(1);
    // Row 22 is 123-456-7890, Row 23 is now 345-678-9012 (auto-checked because MCC row 11 was true!)
    expect(configSheet.getRange(22, 2).getValue()).toBe('123-456-7890');
    expect(configSheet.getRange(23, 2).getValue()).toBe('345-678-9012');
    expect(configSheet.getRange(23, 1).getValue()).toBe(true);
    expect(configSheet.getRange('B5').getValue()).toBe('EUR');
    expect(configSheet.getRange('B5').getDataValidation()?.values).toEqual([
      'EUR',
      'GBP',
      'USD',
    ]);
  });

  it('filters checked CIDs strictly for the current MCC and marks MCCs with 0 selected accounts as Inactive', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    const mccA = { mccId: '111-222-3333', mccName: 'EMEA MCC' };
    const mccB = { mccId: '444-555-6666', mccName: 'AMER MCC' };

    syncDiscoveredAccountsToConfig(configSheet, mccA, [
      buildMockAccount({ customerId: '100-000-0001', name: 'EMEA 1', currencyCode: 'EUR' }),
      buildMockAccount({ customerId: '100-000-0002', name: 'EMEA 2', currencyCode: 'EUR' }),
    ]);
    syncDiscoveredAccountsToConfig(configSheet, mccB, [
      buildMockAccount({ customerId: '200-000-0001', name: 'AMER 1', currencyCode: 'USD' }),
    ]);

    // By default all newly discovered accounts start unselected (false); user checks only EMEA 1 (row 22)
    configSheet.getRange(22, 1).setValue(true);

    // Even if EMEA MCC's "Select All in MCC" checkbox (row 11, col 1) is unchecked, EMEA 1 is still checked!
    configSheet.getRange(11, 1).setValue(false);
    configSheet.getRange(12, 1).setValue(false);

    const selectedForA = getSelectedCidsForMcc(configSheet, mccA.mccId);
    const selectedForB = getSelectedCidsForMcc(configSheet, mccB.mccId);

    expect(selectedForA).toEqual(['1000000001']);
    expect(selectedForB).toEqual([]);

    const summaries = getMccSelectionSummary(configSheet);
    const summaryA = summaries.find((s) => s.mccId === '1112223333');
    const summaryB = summaries.find((s) => s.mccId === '4445556666');

    expect(summaryA?.discoveredCount).toBe(2);
    expect(summaryA?.selectedCount).toBe(1);
    expect(summaryA?.isActive).toBe(true);

    // AMER MCC has 0 selected accounts -> marked Inactive / Skipped!
    expect(summaryB?.discoveredCount).toBe(1);
    expect(summaryB?.selectedCount).toBe(0);
    expect(summaryB?.isActive).toBe(false);
  });

  it('reads Global Settings (Target Currency and Exclude Shopping & Feed PMax) from B5:B6', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    configSheet.getRange('B5').setValue('eur');
    configSheet.getRange('B6').setValue(true);

    expect(readGlobalConfig(configSheet)).toEqual({
      targetCurrency: 'EUR',
      excludeShoppingAndPmax: true,
      selectedCidsByMcc: {},
    });
  });

  it('keeps exactly one empty accounts data row when every account of the only MCC is unlinked', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);
    const mcc = { mccId: '111-222-3333', mccName: 'EMEA' };

    syncDiscoveredAccountsToConfig(configSheet, mcc, [
      buildMockAccount({ customerId: '100-000-0001', name: 'A', currencyCode: 'EUR' }),
      buildMockAccount({ customerId: '100-000-0002', name: 'B', currencyCode: 'EUR' }),
      buildMockAccount({ customerId: '100-000-0003', name: 'C', currencyCode: 'EUR' }),
    ]);
    const maxRowsWithAccounts = configSheet.getMaxRows();

    const res = syncDiscoveredAccountsToConfig(configSheet, mcc, []);
    expect(res.removedCidsCount).toBe(3);
    // Rows 23-24 are deleted; row 22 stays (blank) so later inserted rows copy a data row, not the header
    expect(configSheet.getMaxRows()).toBe(maxRowsWithAccounts - 2);
    expect(configSheet.getRange(22, 2).getValue()).toBe('');
  });

  it('locks daily GlobalConfig in _RunState until tomorrow or until resetTodayRunState is called', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    configSheet.getRange('B5').setValue('USD');
    configSheet.getRange('B6').setValue(false);

    const locked1 = getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));
    expect(locked1.targetCurrency).toBe('USD');
    expect(locked1.excludeShoppingAndPmax).toBe(false);

    // User edits Config mid-day during an hourly batch cycle
    configSheet.getRange('B5').setValue('GBP');
    configSheet.getRange('B6').setValue(true);

    const locked2 = getOrLockDailyConfig(ss, '2026-10-14', () => readGlobalConfig(configSheet));
    expect(locked2.targetCurrency).toBe('USD');
    expect(locked2.excludeShoppingAndPmax).toBe(false);

    // Clicking "Apply Config Changes Now" resets today's lock
    resetTodayRunState(ss, '2026-10-14');
    const lockedAfterReset = getOrLockDailyConfig(ss, '2026-10-14', () =>
      readGlobalConfig(configSheet)
    );
    expect(lockedAfterReset.targetCurrency).toBe('GBP');
    expect(lockedAfterReset.excludeShoppingAndPmax).toBe(true);

    // Populate additional day-1 state (FX_RATES plus persistent multi-day keys)
    setRunStateValue(ss, 'FX_RATES_2026-10-14_GBP', '{"USD":0.79}', '', '2026-10-14');
    setRunStateValue(ss, 'FX_LAST_GOOD_USD_GBP', '0.79', '', '2026-10-14');

    // Next day (2026-10-15): locking daily config automatically prunes 2026-10-14 date-scoped keys
    getOrLockDailyConfig(ss, '2026-10-15', () => readGlobalConfig(configSheet));
    const stateMap = readRunStateMap(ss);

    expect(stateMap.has('DAILY_CONFIG_LOCK_2026-10-14')).toBe(false);
    expect(stateMap.has('FX_RATES_2026-10-14_GBP')).toBe(false);
    expect(stateMap.has('DAILY_CONFIG_LOCK_2026-10-15')).toBe(true);
    // Persistent cross-day keys remain intact
    expect(stateMap.get('FX_LAST_GOOD_USD_GBP')).toBe('0.79');
  });

  it('uses open-ended formulas and skips redundant setDataValidation calls', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    const mcc = { mccId: '111-222-3333', mccName: 'EMEA' };
    const accounts = [
      buildMockAccount({ customerId: '001-234-5678', name: 'Leading Zero CID', currencyCode: 'EUR' }),
    ];

    syncDiscoveredAccountsToConfig(configSheet, mcc, accounts);
    const callsAfterFirst = configSheet.setDataValidationCallCount;

    // Re-syncing identical currencies must not call setDataValidation again on B5
    syncDiscoveredAccountsToConfig(configSheet, mcc, accounts);
    expect(configSheet.setDataValidationCallCount).toBe(callsAfterFirst);

    // Open-ended COUNTIF($D$22:$D, ...) formula in row 11
    expect(String(configSheet.getRange(11, 4).getValue())).toContain('$D$22:$D');
    expect(String(configSheet.getRange(11, 5).getValue())).toContain('$A$22:$A');
  });
});

describe('Config & _RunState', () => {
  function setup() {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({spreadsheetApp});
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);
    return {ss, configSheet};
  }

  it('accepts the Master Template header rows and rejects a shifted Config layout before writing', () => {
    const {ss, configSheet} = setup();
    configSheet.getRange(10, 1, 1, 6).setValues([[...MCC_TABLE_HEADERS]]);
    configSheet.getRange(21, 1, 1, 6).setValues([[...CID_TABLE_HEADERS]]);
    writeMccSheetState(ss, {
      mcc: {mccId: '111-222-3333', mccName: 'EMEA'},
      childAccounts: [{customerId: '100-000-0001', name: 'UK', currencyCode: 'GBP'}],
    });
    expect(() => assertConfigLayout(configSheet)).not.toThrow();
    syncAllMccSheetsToConfig(ss, configSheet);
    expect(configSheet.getRange(22, 2).getValue()).toBe('100-000-0001');

    // A row inserted above the Child Accounts table moves its header from row 21 to row 22
    configSheet.getRange(21, 1, 1, 6).setValues([['', '', '', '', '', '']]);
    const writesBefore = configSheet.valueWriteCount;
    expect(() => syncAllMccSheetsToConfig(ss, configSheet)).toThrow(
      'Config layout changed: expected "Include" in A21, found ""'
    );
    expect(configSheet.valueWriteCount).toBe(writesBefore);
  });

  it('stores formula-like account and MCC names as inert text and does not rewrite them on the next sync', () => {
    const {configSheet} = setup();
    const evilName = '=HYPERLINK("https://evil.example","x")';
    const sync = () =>
      syncDiscoveredAccountsToConfig(configSheet, {mccId: '111-222-3333', mccName: '+MCC'}, [
        buildMockAccount({customerId: '100-000-0001', name: evilName, currencyCode: 'USD'}),
      ]);
    sync();

    // MCC table row 11: B (ID) and C (name) are text; the name is escaped
    expect(configSheet.getRange(11, 3).getValue()).toBe("'+MCC");
    expect(configSheet.getCellFormat(11, 3)).toBe('@');
    // Child Accounts row 22: name (C) and parent MCC name (E) are escaped; B:F are text
    expect(configSheet.getRange(22, 3).getValue()).toBe(`'${evilName}`);
    expect(configSheet.getRange(22, 5).getValue()).toBe("'+MCC");
    for (let col = 2; col <= 6; col++) {
      expect(configSheet.getCellFormat(22, col)).toBe('@');
    }
    // Readers see the original names
    expect(getMccSelectionSummary(configSheet)[0].mccName).toBe('+MCC');

    // The next sync compares unescaped values, so nothing looks renamed and nothing is written
    const writesBefore = configSheet.valueWriteCount;
    sync();
    expect(configSheet.valueWriteCount).toBe(writesBefore);
  });

  it('escapes renamed account names written in place', () => {
    const {configSheet} = setup();
    const mcc = {mccId: '111-222-3333', mccName: 'EMEA'};
    syncDiscoveredAccountsToConfig(configSheet, mcc, [
      buildMockAccount({customerId: '100-000-0001', name: 'Safe', currencyCode: 'USD'}),
    ]);
    syncDiscoveredAccountsToConfig(configSheet, mcc, [
      buildMockAccount({customerId: '100-000-0001', name: '@Renamed', currencyCode: 'USD'}),
    ]);
    expect(configSheet.getRange(22, 3).getValue()).toBe("'@Renamed");
    expect(configSheet.getCellFormat(22, 3)).toBe('@');
  });

  it('falls back to the first registered currency when Target Currency is not a 3-letter code', () => {
    const {configSheet} = setup();
    syncDiscoveredAccountsToConfig(configSheet, {mccId: '111-222-3333', mccName: 'EMEA'}, [
      buildMockAccount({customerId: '100-000-0001', name: 'UK', currencyCode: 'GBP'}),
    ]);
    configSheet.getRange('B5').setValue('USD")&IMPORTDATA("https://evil.example")&("');
    expect(readGlobalConfig(configSheet).targetCurrency).toBe('GBP');
  });

  it('never adds currencies from discovered accounts without a valid customer ID', () => {
    const {configSheet} = setup();
    syncDiscoveredAccountsToConfig(configSheet, {mccId: '111-222-3333', mccName: 'EMEA'}, [
      buildMockAccount({customerId: '123-456-7890', name: 'UK', currencyCode: 'GBP'}),
      buildMockAccount({customerId: '', name: 'Broken', currencyCode: 'JPY'}),
      buildMockAccount({customerId: 'MCC 1', name: 'Label', currencyCode: 'CHF'}),
    ]);
    expect(configSheet.getRange('B5').getDataValidation()?.values).toEqual(['GBP']);
    expect(configSheet.getRange(23, 2).getValue()).toBe('');
  });

  it('updates renamed account and MCC names in place, keeps checkboxes, and writes nothing when unchanged', () => {
    const {configSheet} = setup();
    const mcc = {mccId: '111-222-3333', mccName: 'EMEA'};
    const accounts = [
      buildMockAccount({customerId: '123-456-7890', name: 'UK Search', currencyCode: 'GBP'}),
      buildMockAccount({customerId: '234-567-8901', name: 'DE Search', currencyCode: 'EUR'}),
    ];
    syncDiscoveredAccountsToConfig(configSheet, mcc, accounts);
    configSheet.getRange(11, 1).setValue(true);
    configSheet.getRange(23, 1).setValue(true);

    // Idle re-sync: zero writes
    const writesBefore = configSheet.valueWriteCount;
    const validationsBefore = configSheet.setDataValidationCallCount;
    syncDiscoveredAccountsToConfig(configSheet, mcc, accounts);
    expect(configSheet.valueWriteCount).toBe(writesBefore);
    expect(configSheet.setDataValidationCallCount).toBe(validationsBefore);

    // Account and MCC renamed in Google Ads
    const renamedMcc = {mccId: '111-222-3333', mccName: 'EMEA (Renamed)'};
    syncDiscoveredAccountsToConfig(configSheet, renamedMcc, [
      accounts[0],
      buildMockAccount({customerId: '234-567-8901', name: 'DE Search (New)', currencyCode: 'EUR'}),
    ]);
    expect(configSheet.getRange(11, 3).getValue()).toBe('EMEA (Renamed)');
    expect(configSheet.getRange(22, 3).getValue()).toBe('UK Search');
    expect(configSheet.getRange(23, 3).getValue()).toBe('DE Search (New)');
    expect(configSheet.getRange(22, 5).getValue()).toBe('EMEA (Renamed)');
    expect(configSheet.getRange(23, 5).getValue()).toBe('EMEA (Renamed)');
    // User checkbox state is untouched
    expect(configSheet.getRange(11, 1).getValue()).toBe(true);
    expect(configSheet.getRange(22, 1).getValue()).toBe(false);
    expect(configSheet.getRange(23, 1).getValue()).toBe(true);
  });

  it('drops currencies of unlinked accounts from the dropdown but keeps the selected Target Currency', () => {
    const {configSheet} = setup();
    const mcc = {mccId: '111-222-3333', mccName: 'EMEA'};
    syncDiscoveredAccountsToConfig(configSheet, mcc, [
      buildMockAccount({customerId: '123-456-7890', name: 'UK', currencyCode: 'GBP'}),
      buildMockAccount({customerId: '234-567-8901', name: 'DE', currencyCode: 'EUR'}),
      buildMockAccount({customerId: '345-678-9012', name: 'JP', currencyCode: 'JPY'}),
    ]);
    expect(configSheet.getRange('B5').getDataValidation()?.values).toEqual(['EUR', 'GBP', 'JPY']);
    configSheet.getRange('B5').setValue('EUR');

    // DE (EUR, the selected currency) and JP (JPY) are unlinked
    syncDiscoveredAccountsToConfig(configSheet, mcc, [
      buildMockAccount({customerId: '123-456-7890', name: 'UK', currencyCode: 'GBP'}),
    ]);
    expect(configSheet.getRange('B5').getValue()).toBe('EUR');
    expect(configSheet.getRange('B5').getDataValidation()?.values).toEqual(['EUR', 'GBP']);
  });

  it('keeps child accounts in writeMccSheetState when omitted and clears them for an explicit []', () => {
    const {ss} = setup();
    const mcc = {mccId: '111-222-3333', mccName: 'EMEA'};
    writeMccSheetState(ss, {
      mcc,
      childAccounts: [{customerId: '123-456-7890', name: 'UK', currencyCode: 'GBP'}],
    });
    writeMccSheetState(ss, {mcc, statusMessage: 'still here'});
    expect(readMccSheetState(ss, mcc.mccId)?.childAccounts?.length).toBe(1);

    writeMccSheetState(ss, {mcc, childAccounts: []});
    expect(readMccSheetState(ss, mcc.mccId)?.childAccounts).toEqual([]);
    expect(readMccSheetState(ss, mcc.mccId)?.statusMessage).toBe('still here');
  });

  it('reads only _MCC_<id> in getMccCheckpoint and ignores legacy _RunState checkpoint rows', () => {
    const {ss} = setup();
    setRunStateValues(
      ss,
      [
        {
          key: 'MCC_CHECKPOINT_2026-10-14_1112223333',
          value: JSON.stringify({mccId: '1112223333', dateStr: '2026-10-14', status: 'COMPLETED', completedCids: []}),
        },
        {key: 'MCC_LAST_COMPLETED_DATE_1112223333', value: '2026-10-14'},
      ],
      '2026-10-14'
    );
    const legacy = getMccCheckpoint(ss, '2026-10-14', '111-222-3333');
    expect(legacy.status).toBe('IN_PROGRESS');
    expect(legacy.lastCompletedDate).toBeUndefined();

    writeMccSheetState(ss, {
      mcc: {mccId: '111-222-3333', mccName: 'EMEA'},
      checkpoint: {mccId: '1112223333', dateStr: '2026-10-14', status: 'COMPLETED', completedCids: ['1']},
    });
    const fromMccSheet = getMccCheckpoint(ss, '2026-10-14', '111-222-3333');
    expect(fromMccSheet.status).toBe('COMPLETED');
    expect(fromMccSheet.lastCompletedDate).toBe('2026-10-14');
  });

  it('upserts several keys in a single _RunState write via setRunStateValues', () => {
    const {ss} = setup();
    setRunStateValue(ss, 'A', '1');
    const runStateSheet = ss.getSheetByName('_RunState')!;
    const writesBefore = runStateSheet.valueWriteCount;
    setRunStateValues(ss, [
      {key: 'A', value: '2'},
      {key: 'B', value: '3', notes: 'note'},
    ]);
    expect(runStateSheet.valueWriteCount - writesBefore).toBe(1);
    const map = readRunStateMap(ss);
    expect(map.get('A')).toBe('2');
    expect(map.get('B')).toBe('3');

    setRunStateValues(ss, []);
    expect(runStateSheet.valueWriteCount - writesBefore).toBe(1);
  });

  it('flushes pending writes in withSheetLock before releasing the document lock, even when fn throws', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    const {lockService} = installAppsScriptGlobals({spreadsheetApp});
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const calls: string[] = [];
    spreadsheetApp.flush = () => {
      calls.push(lockService!.getDocumentLock().hasLock() ? 'flush(locked)' : 'flush(unlocked)');
    };
    const lock = lockService!.getDocumentLock();
    const originalRelease = lock.releaseLock.bind(lock);
    lock.releaseLock = () => {
      calls.push('release');
      originalRelease();
    };

    withSheetLock(ss, 'TEST', () => calls.push('fn'));
    expect(calls).toEqual(['fn', 'flush(locked)', 'release']);

    calls.length = 0;
    expect(() =>
      withSheetLock(ss, 'TEST', () => {
        throw new Error('boom');
      })
    ).toThrow('boom');
    expect(calls).toEqual(['flush(locked)', 'release']);

    // Release still happens when flush itself throws
    calls.length = 0;
    spreadsheetApp.flush = () => {
      throw new Error('flush failed');
    };
    expect(() => withSheetLock(ss, 'TEST', () => undefined)).toThrow('flush failed');
    expect(calls).toEqual(['release']);
    expect(lock.hasLock()).toBe(false);
  });
});
