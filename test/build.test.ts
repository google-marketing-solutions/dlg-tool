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

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {describe, expect, it} from 'vitest';
import {buildAll, buildSheetManifest} from '../scripts/build.mjs';
import {
  createMockAdsApp,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

describe('Build Pipeline & Apps Script Mocks', () => {
  it('bundles dist/ads/Code.gs and dist/sheet/Code.gs with required Apps Script entrypoints', async () => {
    const tmpOutDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dlg-build-test-'));
    try {
      await buildAll(tmpOutDir);

      const codeGsPath = path.join(tmpOutDir, 'ads', 'Code.gs');
      const sheetCodePath = path.join(tmpOutDir, 'sheet', 'Code.gs');
      const manifestPath = path.join(tmpOutDir, 'sheet', 'appsscript.json');

      expect(fs.existsSync(codeGsPath)).toBe(true);
      expect(fs.existsSync(sheetCodePath)).toBe(true);
      expect(fs.existsSync(manifestPath)).toBe(true);

      const codeGs = fs.readFileSync(codeGsPath, 'utf8');
      const sheetCode = fs.readFileSync(sheetCodePath, 'utf8');

      expect(codeGs.startsWith('/**\n * Copyright 2026 Google LLC')).toBe(true);
      expect(sheetCode.startsWith('/**\n * Copyright 2026 Google LLC')).toBe(true);
      expect(codeGs).toContain('var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";');
      expect(codeGs).toContain('var SPREADSHEET_TIMEZONE = "YOUR_SPREADSHEET_TIMEZONE_HERE";');
      expect(codeGs).toContain('function main()');
      expect(codeGs).not.toContain('withSheetLock');
      expect(codeGs).not.toContain('syncDiscoveredAccountsToConfig');
      expect(codeGs).not.toContain('publishFromStagingToData');
      expect(sheetCode).toContain('@OnlyCurrentDoc');
      expect(sheetCode).toContain('function onOpen()');
      expect(sheetCode).toContain('function onEdit(e)');
      expect(sheetCode).toContain('EMBEDDED_ADS_SCRIPT_TEMPLATE');

      const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
      expect(manifest).toEqual(buildSheetManifest());
    } finally {
      fs.rmSync(tmpOutDir, {recursive: true, force: true});
    }
  });

  it('pins exactly the OAuth scopes the Sheet-bound script needs', () => {
    expect(buildSheetManifest().oauthScopes).toEqual([
      'https://www.googleapis.com/auth/spreadsheets.currentonly',
      'https://www.googleapis.com/auth/script.container.ui',
      'https://www.googleapis.com/auth/script.scriptapp',
    ]);
  });

  it('provides functional in-memory mocks for SpreadsheetApp and AdsApp', () => {
    const spreadsheetApp = createMockSpreadsheetApp('https://docs.google.com/spreadsheets/d/test-id/edit');
    const adsApp = createMockAdsApp({
      mccId: '111-222-3333',
      mccName: 'Test MCC',
      remainingTime: 1800,
      isPreview: false,
    });

    installAppsScriptGlobals({ spreadsheetApp, adsApp });

    const ss = spreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.getRange(1, 1, 1, 2).setValues([['Target Currency', 'USD']]);

    expect(sheet.getRange(1, 1, 1, 2).getValues()).toEqual([['Target Currency', 'USD']]);
    expect(adsApp.currentAccount().getCustomerId()).toBe('111-222-3333');
    expect(adsApp.getExecutionInfo().getRemainingTime()).toBe(1800);
  });
});
