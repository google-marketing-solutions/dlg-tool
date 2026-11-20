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
import {syncDiscoveredAccountsToConfig} from '../src/appscript/config-sheet';
import {getOrLockDailyConfig, getRunStateValue} from '../src/appscript/run-state';
import {
  DEMO_TOOLS_PROPERTY,
  PUBLISHER_HEARTBEAT_PROPERTY,
  PUBLISHER_HEARTBEAT_STALE_MS,
  applyConfigChangesNow,
  buildDataStudioCreationUrlForActiveSheet,
  ensureAutoPublishTriggerInstalled,
  generateAdsScriptForCurrentSheet,
  onEdit,
  onOpen,
  openSavedDashboard,
  publishFromStagingToData,
  runScheduledPublish,
  seedSyntheticDemoData,
  seedSyntheticDemoDataFromMenu,
  showCopyAdsScriptModal,
} from '../src/appscript/sheet-ui';
import {readReportRowsFromSheet} from '../src/appscript/sheets/data-writer';
import {SCHEMA_COLUMN_NAMES, SCHEMA_DUMMY_ACCOUNT_NAME} from '../src/shared/schema';
import {
  clearMccSheetCheckpoint,
  getMccCheckpoint,
  writeMccSheetState,
} from '../src/shared/sheets/mcc-state';
import {buildMockAccount, setTestDate} from './helpers/factories';
import {createMockSpreadsheetApp, installAppsScriptGlobals} from './mocks/google-apps-script.mock';

describe('Master Google Sheet Bound Script, 4 Setup Actions & onEdit Checkbox Sync', () => {
  afterEach(() => {
    vi.useRealTimers();
    installAppsScriptGlobals({});
  });

  it('injects the current Google Sheet URL and timezone into the embedded Code.gs template and throws on missing placeholders', () => {
    const template = [
      '// Header',
      'var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";',
      'var SPREADSHEET_TIMEZONE = "YOUR_SPREADSHEET_TIMEZONE_HERE";',
      'function main() {}',
    ].join('\n');
    const injected = generateAdsScriptForCurrentSheet(
      'https://docs.google.com/spreadsheets/d/my-sheet-123/edit',
      'Europe/London',
      template
    );

    expect(injected).toContain(
      'var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/my-sheet-123/edit";'
    );
    expect(injected).toContain('var SPREADSHEET_TIMEZONE = "Europe/London";');
    expect(injected).not.toContain('YOUR_SPREADSHEET_URL_HERE');
    expect(injected).not.toContain('YOUR_SPREADSHEET_TIMEZONE_HERE');

    expect(() =>
      generateAdsScriptForCurrentSheet(
        'https://docs.google.com/spreadsheets/d/123',
        'UTC',
        '// broken'
      )
    ).toThrow(/SPREADSHEET_URL placeholder/);
    expect(() =>
      generateAdsScriptForCurrentSheet(
        'https://docs.google.com/spreadsheets/d/123',
        'UTC',
        'var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";'
      )
    ).toThrow(/SPREADSHEET_TIMEZONE placeholder/);
  });

  it('builds Data Studio creation URL with c.reportId and refreshFields=false', () => {
    const spreadsheetApp = createMockSpreadsheetApp(
      'https://docs.google.com/spreadsheets/d/my-sheet-123/edit'
    );
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Data', 100);

    const lsUrl = buildDataStudioCreationUrlForActiveSheet(ss);
    expect(lsUrl).toContain('c.reportId=8efcb962-4ce8-4df1-bca0-bd354d946b2d');
    expect(lsUrl).not.toContain('r.reportId=');
    expect(lsUrl).toContain('ds.ds0.spreadsheetId=mock-spreadsheet-id');
    expect(lsUrl).toContain('ds.ds0.worksheetId=100');
    // Field types come from the template, not from (possibly empty) sheet columns
    expect(lsUrl).toContain('ds.ds0.refreshFields=false');
  });

  it('synchronizes Global, Per-MCC, and Individual CID checkboxes via onEdit(e) including multi-row edits', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);
    configSheet.getRange('A9').setValue(true);

    syncDiscoveredAccountsToConfig(
      configSheet,
      { mccId: '111-222-3333', mccName: 'EMEA MCC' },
      [
        buildMockAccount({ customerId: '100-000-0001', name: 'EMEA 1', currencyCode: 'EUR' }),
        buildMockAccount({ customerId: '100-000-0002', name: 'EMEA 2', currencyCode: 'EUR' }),
      ]
    );
    syncDiscoveredAccountsToConfig(
      configSheet,
      { mccId: '444-555-6666', mccName: 'AMER MCC' },
      [buildMockAccount({ customerId: '200-000-0001', name: 'AMER 1', currencyCode: 'USD' })]
    );

    // Initially all newly discovered MCCs (rows 11, 12) and CIDs (rows 22..24) default to false
    expect(configSheet.getRange(11, 1).getValue()).toBe(false);
    expect(configSheet.getRange(12, 1).getValue()).toBe(false);
    expect(configSheet.getRange(22, 1).getValue()).toBe(false);
    expect(configSheet.getRange(23, 1).getValue()).toBe(false);
    expect(configSheet.getRange(24, 1).getValue()).toBe(false);

    // 1. User checks AMER MCC's "Select All in MCC" (Row 12, Col 1)
    configSheet.getRange(12, 1).setValue(true);
    onEdit({ range: configSheet.getRange(12, 1) });

    // Both EMEA CIDs (rows 22 & 23) stay unchecked, while AMER CID (row 24) becomes checked!
    expect(configSheet.getRange(22, 1).getValue()).toBe(false);
    expect(configSheet.getRange(23, 1).getValue()).toBe(false);
    expect(configSheet.getRange(24, 1).getValue()).toBe(true);
    expect(configSheet.getRange('A9').getValue()).toBe(false);

    // 2. Multi-row range edit across rows 22..23 (checking both EMEA 1 & EMEA 2 at once)
    configSheet.getRange(22, 1, 2, 1).setValues([[true], [true]]);
    onEdit({ range: configSheet.getRange(22, 1, 2, 1) });
    expect(configSheet.getRange(11, 1).getValue()).toBe(true);
    expect(configSheet.getRange('A9').getValue()).toBe(true);

    // 3. User unchecks Global Select All (Row 9, Col 1) -> All MCCs and all CIDs become false
    configSheet.getRange(9, 1).setValue(false);
    onEdit({ range: configSheet.getRange(9, 1) });
    expect(configSheet.getRange(11, 1).getValue()).toBe(false);
    expect(configSheet.getRange(12, 1).getValue()).toBe(false);
    expect(configSheet.getRange(22, 1).getValue()).toBe(false);
    expect(configSheet.getRange(23, 1).getValue()).toBe(false);
    expect(configSheet.getRange(24, 1).getValue()).toBe(false);
  });

  it('calls 3-arg ui.alert on missing/invalid Saved Dashboard URL and opens a modal when valid', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Config', 0);
    const setupSheet = ss.insertSheet('Setup', 0);

    // Empty Setup!D13 -> triggers 3-argument ui.alert without throwing
    openSavedDashboard();
    expect(spreadsheetApp.alerts.length).toBe(1);
    expect(spreadsheetApp.alerts[0].buttons).toBe('OK');
    expect(spreadsheetApp.alerts[0].prompt).toContain('cell D13 on the Setup tab');

    // Valid https:// URL in Setup!D13 -> opens modal with escaped href
    setupSheet.getRange('D13').setValue('https://datastudio.google.com/reporting/abc');
    openSavedDashboard();
    expect(spreadsheetApp.modals.length).toBe(1);
    expect(spreadsheetApp.modals[0].html).toContain('https://datastudio.google.com/reporting/abc');
  });

  it('rejects non-Data Studio and script-injecting dashboard URLs and keeps URL text out of <script>', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Config', 0);
    const setupSheet = ss.insertSheet('Setup', 0);

    for (const bad of [
      'https://example.com/</script><script>google.script.run.seedSyntheticDemoData()</script>',
      'https://lookerstudio.google.com/reporting/x</script><script>alert(1)</script>',
      'https://evil.example/reporting/abc',
      'javascript:alert(1)',
    ]) {
      setupSheet.getRange('D13').setValue(bad);
      openSavedDashboard();
    }
    expect(spreadsheetApp.modals.length).toBe(0);
    expect(spreadsheetApp.alerts.length).toBe(4);

    setupSheet.getRange('D13').setValue('https://lookerstudio.google.com/reporting/abc/page/xyz');
    openSavedDashboard();
    expect(spreadsheetApp.modals.length).toBe(1);
    const html = spreadsheetApp.modals[0].html;
    const script = html.slice(html.indexOf('<script>'), html.indexOf('</script>'));
    expect(script).not.toContain('lookerstudio');
    expect(html).toContain('href="https://lookerstudio.google.com/reporting/abc/page/xyz"');
  });

  it('installs the 15-minute auto-publish trigger and renders the Copy Ads Script modal with copyBtn above the code box', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Config', 0);

    let installedMinutes = 0;
    const g = globalThis as Record<string, unknown>;
    g.ScriptApp = {
      getProjectTriggers: () => [],
      newTrigger: (fnName: string) => ({
        timeBased: () => ({
          everyMinutes: (mins: number) => ({
            create: () => {
              if (fnName === 'runScheduledPublish') installedMinutes = mins;
            },
          }),
        }),
      }),
    };
    g.__EMBEDDED_ADS_SCRIPT__ = [
      'var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";',
      'var SPREADSHEET_TIMEZONE = "YOUR_SPREADSHEET_TIMEZONE_HERE";',
    ].join('\n');

    ensureAutoPublishTriggerInstalled();
    expect(installedMinutes).toBe(15);

    showCopyAdsScriptModal();
    expect(spreadsheetApp.modals.length).toBe(1);
    expect(spreadsheetApp.modals[0].html).toContain('Frequency = Hourly');
    expect(spreadsheetApp.modals[0].html).toContain('15-minute background publisher trigger');
    expect(spreadsheetApp.modals[0].html).toContain('America/New_York');
    // The Copy button sits above the code box so it is visible without scrolling
    expect(spreadsheetApp.modals[0].html.indexOf('id="copyBtn"')).toBeLessThan(
      spreadsheetApp.modals[0].html.indexOf('<textarea')
    );
  });

  it('clears today\'s config lock when applyConfigChangesNow is invoked', () => {
    const todayStr = '2026-10-14';
    setTestDate(todayStr);

    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    getOrLockDailyConfig(ss, todayStr, () => ({
      targetCurrency: 'EUR',
      excludeShoppingAndPmax: false,
    }));

    applyConfigChangesNow();

    const relocked = getOrLockDailyConfig(ss, todayStr, () => ({
      targetCurrency: 'GBP',
      excludeShoppingAndPmax: true,
    }));
    expect(relocked.targetCurrency).toBe('GBP');
    expect(String(configSheet.getRange('B2').getValue())).toContain('Config changes applied');
  });

  it('seeds 2 synthetic Demo MCCs and 13 recommendations into Config, _Stage_*, and Data for end-to-end Multi-MCC Dashboard testing', () => {
    setTestDate('2026-10-19');

    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);

    seedSyntheticDemoData({ spreadsheet: ss });
    const dataSheet = ss.getSheetByName('Data')!;
    expect(readReportRowsFromSheet(dataSheet).length).toBe(13); // 10 recs from 999-888-7777 + 3 recs from 888-777-6666
    expect(configSheet.getRange(11, 2).getValue()).toBe('999-888-7777');
    expect(configSheet.getRange(12, 2).getValue()).toBe('888-777-6666');
    expect(configSheet.getRange(22, 1).getValue()).toBe(true);
    expect(configSheet.getRange(23, 1).getValue()).toBe(true);
    expect(configSheet.getRange(24, 1).getValue()).toBe(true);
    expect(configSheet.getRange(25, 1).getValue()).toBe(true);
    expect(configSheet.getRange(26, 1).getValue()).toBe(true);

    // Check Exclude Shopping & Feed PMax (B6 = true) and click Apply Config Changes Now -> filters out the 1 SHOPPING rec!
    configSheet.getRange('B6').setValue(true);
    applyConfigChangesNow();
    expect(dataSheet.getMaxRows()).toBe(14); // 1 header + 1 schema dummy row + 12 remaining recommendation rows
  });

  it('seeds budget recs on portfolio campaigns and a campaign with both a budget and a bid rec as separate rows', () => {
    setTestDate('2026-10-19');

    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Config', 0);

    seedSyntheticDemoData({ spreadsheet: ss });
    const rows = readReportRowsFromSheet(ss.getSheetByName('Data')!);
    const summarize = (accountId: string) =>
      rows
        .filter((r) => r.accountId === accountId)
        .map((r) => `${r.recommendationScope}|${r.recommendationType}|${r.campaignName}`)
        .sort();

    // UK portfolio: one PORTFOLIO bid row + one CAMPAIGN budget row per member campaign
    expect(summarize('303-404-5050')).toEqual([
      'CAMPAIGN|CAMPAIGN_BUDGET|UK - Search - London & South East',
      'CAMPAIGN|CAMPAIGN_BUDGET|UK - Search - Midlands & North',
      'PORTFOLIO|LOWER_TARGET_ROAS|MCC Global ROAS Portfolio',
      'PORTFOLIO|RAISE_TARGET_CPA|UK - Target CPA Growth Portfolio',
    ]);

    // US Brand Search: a budget rec (Budgets page) and a bid rec (Bids page) on the same campaign
    const brandRows = rows.filter((r) => r.campaignName === 'US - Brand Search - AI Max');
    expect(brandRows.map((r) => r.recommendationType).sort()).toEqual([
      'CAMPAIGN_BUDGET',
      'RAISE_TARGET_CPA',
    ]);
    expect(brandRows.every((r) => r.recommendationScope === 'CAMPAIGN')).toBe(true);
  });

  it('keeps Data timestamps as valid date strings across repeated stale publishes', () => {
    setTestDate('2026-10-26');

    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Config', 0);

    seedSyntheticDemoData({ spreadsheet: ss });
    // Clearing today's _MCC_ checkpoints (as a GAQL-affecting Config change would) makes the next
    // scheduled (non-forced) publish treat both demo MCCs as stale and re-read the existing Data rows.
    clearMccSheetCheckpoint(ss, '999-888-7777');
    clearMccSheetCheckpoint(ss, '888-777-6666');

    const dataSheet = ss.getSheetByName('Data')!;
    const tsCol = SCHEMA_COLUMN_NAMES.indexOf('timestamp') + 1;
    for (const dateStr of ['2026-10-26', '2026-10-27']) {
      setTestDate(dateStr);
      publishFromStagingToData({ spreadsheet: ss });
      expect(getRunStateValue(ss, `PUBLISH_STATUS_${dateStr}`)).toBe('PUBLISHED_WITH_STALE_WARNING');

      const rows = readReportRowsFromSheet(dataSheet);
      expect(rows.length).toBe(13);
      for (const row of rows) {
        expect(row.mccSyncStatus).toBe('WARNING: Stale snapshot from 2026-10-26');
        expect(row.timestamp).toBe('2026-10-26 00:00:00');
      }
      for (let r = 2; r <= dataSheet.getLastRow(); r++) {
        expect(typeof dataSheet.getRange(r, tsCol).getValue()).not.toBe('number');
      }
    }
  });

  it('keeps exactly one uncounted schema dummy row across publish, idle, and stale republishes', () => {
    setTestDate('2026-10-26');

    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.insertSheet('Config', 0);
    const accountNameCol = SCHEMA_COLUMN_NAMES.indexOf('accountName') + 1;

    const dataSheet = (): NonNullable<ReturnType<typeof ss.getSheetByName>> =>
      ss.getSheetByName('Data')!;
    const expectSingleDummyRow = (realRowCount: number) => {
      const sheet = dataSheet();
      expect(sheet.getMaxRows()).toBe(realRowCount + 2); // header + dummy + real rows
      expect(sheet.getLastRow()).toBe(realRowCount + 2);
      const names = sheet
        .getRange(1, accountNameCol, sheet.getLastRow(), 1)
        .getValues()
        .map((r) => r[0]);
      expect(names.filter((n) => n === SCHEMA_DUMMY_ACCOUNT_NAME)).toHaveLength(1);
      expect(names[1]).toBe(SCHEMA_DUMMY_ACCOUNT_NAME);
      expect(readReportRowsFromSheet(sheet)).toHaveLength(realRowCount);
    };

    // 1. Fresh publish
    seedSyntheticDemoData({ spreadsheet: ss });
    // The status banner counts real recommendations only
    expect(String(configSheet.getRange('B2').getValue())).toContain('13 recommendations');
    expectSingleDummyRow(13);

    // 2. Idle tick: nothing changed, nothing written
    setTestDate('2026-10-26', 11);
    const writesBeforeIdle = dataSheet().valueWriteCount;
    publishFromStagingToData({ spreadsheet: ss });
    expect(dataSheet().valueWriteCount).toBe(writesBeforeIdle);
    expectSingleDummyRow(13);

    // 3. Stale republishes re-read Data (including its dummy row) and write it back
    clearMccSheetCheckpoint(ss, '999-888-7777');
    clearMccSheetCheckpoint(ss, '888-777-6666');
    for (const dateStr of ['2026-10-26', '2026-10-27', '2026-10-28']) {
      setTestDate(dateStr);
      publishFromStagingToData({ spreadsheet: ss });
      expect(getRunStateValue(ss, `PUBLISH_STATUS_${dateStr}`)).toBe('PUBLISHED_WITH_STALE_WARNING');
      expectSingleDummyRow(13);
    }
  });
});

describe('Demo tools flag and publisher heartbeat', () => {
  afterEach(() => {
    vi.useRealTimers();
    installAppsScriptGlobals({});
  });

  function setup() {
    const spreadsheetApp = createMockSpreadsheetApp();
    const {propertiesService} = installAppsScriptGlobals({spreadsheetApp});
    const ss = spreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('Config', 0);
    return {spreadsheetApp, propertiesService, ss};
  }

  it('shows the demo menu item only when DLG_ENABLE_DEMO_TOOLS is true', () => {
    const {spreadsheetApp, propertiesService} = setup();
    onOpen();
    expect(spreadsheetApp.menuItems.map((m) => m.fnName)).not.toContain('seedSyntheticDemoDataFromMenu');

    propertiesService.scriptProperties.setProperty(DEMO_TOOLS_PROPERTY, 'true');
    onOpen();
    expect(spreadsheetApp.menuItems.map((m) => m.fnName)).toContain('seedSyntheticDemoDataFromMenu');
    expect(spreadsheetApp.menuItems.map((m) => m.fnName)).not.toContain('seedSyntheticDemoData');
  });

  it('refuses to seed when the flag is off and asks for confirmation when it is on', () => {
    setTestDate('2026-10-19');
    const {spreadsheetApp, propertiesService, ss} = setup();
    expect(() => seedSyntheticDemoDataFromMenu()).toThrow(/Demo tools are disabled/);
    expect(ss.getSheetByName('Data')).toBeNull();

    propertiesService.scriptProperties.setProperty(DEMO_TOOLS_PROPERTY, 'true');
    spreadsheetApp.ui.alertResponse = 'NO';
    seedSyntheticDemoDataFromMenu();
    expect(spreadsheetApp.alerts.at(-1)?.buttons).toBe('YES_NO');
    expect(ss.getSheetByName('Data')).toBeNull();

    spreadsheetApp.ui.alertResponse = 'YES';
    seedSyntheticDemoDataFromMenu();
    expect(readReportRowsFromSheet(ss.getSheetByName('Data')!).length).toBe(13);
  });

  it('does not reset the checkpoint of a real MCC when seeding demo data', () => {
    setTestDate('2026-10-19');
    const {ss} = setup();
    const realMcc = {mccId: '111-222-3333', mccName: 'Real MCC'};
    writeMccSheetState(ss, {
      mcc: realMcc,
      checkpoint: {mccId: '1112223333', dateStr: '2026-10-19', status: 'COMPLETED', completedCids: ['1000000001']},
    });
    seedSyntheticDemoData({spreadsheet: ss});
    expect(getMccCheckpoint(ss, '2026-10-19', realMcc.mccId).status).toBe('COMPLETED');
  });

  it("does not install a second publisher trigger while another editor's heartbeat is fresh", () => {
    setTestDate('2026-10-19', 10);
    const {propertiesService} = setup();
    let created = 0;
    (globalThis as Record<string, unknown>).ScriptApp = {
      getProjectTriggers: () => [],
      newTrigger: () => ({timeBased: () => ({everyMinutes: () => ({create: () => created++})})}),
    };

    // Another editor's trigger ran 10 minutes ago
    propertiesService.documentProperties.setProperty(
      PUBLISHER_HEARTBEAT_PROPERTY,
      String(Date.now() - 10 * 60 * 1000)
    );
    ensureAutoPublishTriggerInstalled();
    expect(created).toBe(0);

    // ...and is gone: the heartbeat is older than 45 minutes
    propertiesService.documentProperties.setProperty(
      PUBLISHER_HEARTBEAT_PROPERTY,
      String(Date.now() - PUBLISHER_HEARTBEAT_STALE_MS - 1)
    );
    ensureAutoPublishTriggerInstalled();
    expect(created).toBe(1);
    // The new trigger claims the heartbeat immediately
    expect(Number(propertiesService.documentProperties.getProperty(PUBLISHER_HEARTBEAT_PROPERTY))).toBe(Date.now());
  });

  it('records the heartbeat on every scheduled publish', () => {
    setTestDate('2026-10-19', 10);
    const {propertiesService, ss} = setup();
    runScheduledPublish({spreadsheet: ss});
    expect(Number(propertiesService.documentProperties.getProperty(PUBLISHER_HEARTBEAT_PROPERTY))).toBe(Date.now());
  });
});
