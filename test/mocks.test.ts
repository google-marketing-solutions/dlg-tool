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

import {afterEach, describe, expect, it} from 'vitest';
import {
  MockSheet,
  createMockAdsApp,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

const g = globalThis as Record<string, unknown>;

describe('Apps Script mocks: row/column deletion', () => {
  function seedSheet(): MockSheet {
    const sheet = new MockSheet('S', 1);
    sheet.getRange(1, 1, 3, 3).setValues([
      ['a1', 'b1', 'c1'],
      ['a2', 'b2', 'c2'],
      ['a3', 'b3', 'c3'],
    ]);
    sheet.getRange(3, 1).setNumberFormat('@');
    sheet.getRange(3, 3).setNumberFormat('@');
    sheet.getRange(3, 2).insertCheckboxes();
    sheet.getRange(1, 3).insertCheckboxes();
    return sheet;
  }

  it('deleteRows shifts values, formats, and validations up together', () => {
    const sheet = seedSheet();
    sheet.deleteRows(1, 2);

    expect(sheet.getRange(1, 1, 1, 3).getValues()).toEqual([['a3', 'b3', 'c3']]);
    expect(sheet.getCellFormat(1, 1)).toBe('@');
    expect(sheet.getCellFormat(1, 3)).toBe('@');
    expect(sheet.getCellValidation(1, 2)?.type).toBe('BOOLEAN');
    // Nothing is left behind at the old row, and the deleted row's validation is gone
    expect(sheet.getCellFormat(3, 1)).toBeUndefined();
    expect(sheet.getCellValidation(3, 2)).toBeNull();
    expect(sheet.getCellValidation(1, 3)).toBeNull();
  });

  it('deleteColumns shifts values, formats, and validations right of the range left', () => {
    const sheet = seedSheet();
    sheet.deleteColumns(1, 1);

    expect(sheet.getMaxColumns()).toBe(25);
    expect(sheet.getRange(1, 1, 3, 2).getValues()).toEqual([
      ['b1', 'c1'],
      ['b2', 'c2'],
      ['b3', 'c3'],
    ]);
    expect(sheet.getLastColumn()).toBe(2);
    expect(sheet.getCellFormat(3, 2)).toBe('@');
    expect(sheet.getCellFormat(3, 1)).toBeUndefined();
    expect(sheet.getCellValidation(3, 1)?.type).toBe('BOOLEAN');
    expect(sheet.getCellValidation(1, 2)?.type).toBe('BOOLEAN');
    expect(sheet.getCellValidation(1, 3)).toBeNull();
    // getLastColumn() stays within the grid, so a full-width read never goes out of bounds
    expect(() => sheet.getRange(1, 1, 3, sheet.getLastColumn()).getValues()).not.toThrow();
  });

  it('deleteColumns in the middle keeps columns to the left in place', () => {
    const sheet = seedSheet();
    sheet.deleteColumns(2, 1);
    expect(sheet.getRange(1, 1, 3, 2).getValues()).toEqual([
      ['a1', 'c1'],
      ['a2', 'c2'],
      ['a3', 'c3'],
    ]);
    expect(sheet.getCellFormat(3, 1)).toBe('@');
    expect(sheet.getCellFormat(3, 2)).toBe('@');
    expect(sheet.getCellValidation(3, 2)).toBeNull();
    expect(sheet.getCellValidation(1, 2)?.type).toBe('BOOLEAN');
  });

  it('clear() removes values, formats, and validations', () => {
    const sheet = seedSheet();
    sheet.clear();
    expect(sheet.getLastRow()).toBe(0);
    expect(sheet.getCellFormat(3, 1)).toBeUndefined();
    expect(sheet.getCellValidation(3, 2)).toBeNull();
  });
});

describe('Apps Script mocks: insertSheet', () => {
  it('throws on a duplicate sheet name, like real Sheets', () => {
    const ss = createMockSpreadsheetApp().getActiveSpreadsheet();
    ss.insertSheet('Config');
    expect(() => ss.insertSheet('Config')).toThrow(/already exists/);
    expect(ss.getSheets().length).toBe(1);
  });

  it('never gives an auto-generated sheet ID that collides with an explicit ID', () => {
    const ss = createMockSpreadsheetApp().getActiveSpreadsheet();
    const explicit100 = ss.insertSheet('Data', 100);
    const explicit101 = ss.insertSheet('_RunState', 101);
    const auto1 = ss.insertSheet('Auto 1');
    const auto2 = ss.insertSheet('Auto 2');
    const ids = [explicit100, explicit101, auto1, auto2].map((s) => s.getSheetId());
    expect(explicit100.getSheetId()).toBe(100);
    expect(explicit101.getSheetId()).toBe(101);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('rejects an explicit sheet ID that is already in use', () => {
    const ss = createMockSpreadsheetApp().getActiveSpreadsheet();
    const auto = ss.insertSheet('Auto');
    expect(() => ss.insertSheet('Other', auto.getSheetId())).toThrow(/already in use/);
  });
});

describe('Apps Script mocks: installAppsScriptGlobals', () => {
  afterEach(() => {
    installAppsScriptGlobals({});
  });

  it('removes services that are not passed, so globals never leak between tests', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    const adsApp = createMockAdsApp({mccId: '111-222-3333', mccName: 'MCC'});
    installAppsScriptGlobals({spreadsheetApp, adsApp});
    expect(g.SpreadsheetApp).toBe(spreadsheetApp);
    expect(g.AdsApp).toBe(adsApp);
    expect(g.AdsManagerApp).toBe(adsApp.adsManagerApp);

    installAppsScriptGlobals({spreadsheetApp});
    expect(g.SpreadsheetApp).toBe(spreadsheetApp);
    expect('AdsApp' in g).toBe(false);
    expect('AdsManagerApp' in g).toBe(false);

    installAppsScriptGlobals({adsApp});
    expect('SpreadsheetApp' in g).toBe(false);
    expect(g.AdsApp).toBe(adsApp);

    g.ScriptApp = {};
    g.__EMBEDDED_ADS_SCRIPT__ = 'test';
    installAppsScriptGlobals({});
    expect('SpreadsheetApp' in g).toBe(false);
    expect('AdsApp' in g).toBe(false);
    expect('AdsManagerApp' in g).toBe(false);
    expect('ScriptApp' in g).toBe(false);
    expect('__EMBEDDED_ADS_SCRIPT__' in g).toBe(false);
  });
});
