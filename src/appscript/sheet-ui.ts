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
 *
 * @fileoverview Entrypoint and UI controller for the Google Sheet-bound script (`dist/sheet/Code.gs`).
 * Provides:
 * - `onOpen()`: Builds the "🚀 DLG Dashboard v2" menu with the setup and operational actions,
 * - `onEdit(e)`: Synchronizes the 3-tier checkbox hierarchy (`Global Select All A9` <-> `Per-MCC A11:A18` <-> `Individual CID A22+`)
 *   using batched `setValues()` calls and supporting multi-row range edits,
 * - `showCopyAdsScriptModal()`: Injects the active Sheet URL into the embedded `Code.gs` template with resilient clipboard fallback, and
 * - `runScheduledPublish()` / `applyConfigChangesNow()`: Syncs `_MCC_*` discoveries, resets today's publisher state (and only
 *   the MCC checkpoints a change really invalidates), and triggers `publishFromStagingToData`.
 */

import {
  clearMccSheetCheckpoint,
  getMccCheckpoint,
} from '../shared/sheets/mcc-state';
import {
  type SpreadsheetLike,
  getOrCreateSheet,
  parseSheetBoolean,
} from '../shared/sheets/sheet-utils';
import {normalizeCid} from '../shared/utils/cid';
import {resolveSpreadsheetDate} from '../shared/utils/date';
import {
  CID_COL,
  CID_TABLE_START_ROW,
  CONFIG_SHEET_NAME,
  GLOBAL_SELECT_ALL_CELL,
  GLOBAL_SELECT_ALL_ROW,
  MCC_COL,
  MCC_TABLE_END_ROW,
  MCC_TABLE_START_ROW,
  SAVED_DASHBOARD_URL_CELL,
  SETUP_SHEET_NAME,
  STATUS_BANNER_CELL,
  getMccSelectionSummary,
  getSelectedCidsForMcc,
  readGlobalConfig,
  sheetCol,
  syncAllMccSheetsToConfig,
} from './config-sheet';
import {seedSyntheticDemoData} from './demo-seed';
import {
  type PublishOptions,
  publishFromStagingToData,
} from './publisher';
import {
  readLockedDailyConfig,
  resetTodayPublisherState,
  withSheetLock,
} from './run-state';
import {DATA_SHEET_NAME, writeReportRowsToSheet} from './sheets/data-writer';

export {
  publishFromStagingToData,
  seedSyntheticDemoData,
  type PublishOptions,
};

declare const __EMBEDDED_ADS_SCRIPT__: string;

/** Embedded `Code.gs` template injected at build time by `scripts/build.mjs`. */
export const EMBEDDED_ADS_SCRIPT_TEMPLATE =
  typeof __EMBEDDED_ADS_SCRIPT__ !== 'undefined' ? __EMBEDDED_ADS_SCRIPT__ : '';

const SPREADSHEET_URL_PLACEHOLDER = 'var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";';
const SPREADSHEET_TIMEZONE_PLACEHOLDER =
  'var SPREADSHEET_TIMEZONE = "YOUR_SPREADSHEET_TIMEZONE_HERE";';
const DATA_STUDIO_TEMPLATE_REPORT_ID = '8efcb962-4ce8-4df1-bca0-bd354d946b2d';

/**
 * Script Property that enables the demo/testing tools (the "Seed Synthetic Demo Data" menu item). Set it to
 * `true` in Apps Script → Project Settings → Script Properties; it is off by default.
 */
export const DEMO_TOOLS_PROPERTY = 'DLG_ENABLE_DEMO_TOOLS';
/** Document Property holding the epoch ms of the last scheduled publish run (by any editor's trigger). */
export const PUBLISHER_HEARTBEAT_PROPERTY = 'DLG_PUBLISHER_HEARTBEAT_MS';
/** A heartbeat older than this (3 missed 15-minute ticks) means no editor's publisher trigger is running. */
export const PUBLISHER_HEARTBEAT_STALE_MS = 45 * 60 * 1000;

/**
 * Injects the active spreadsheet's URL and timezone into the embedded `Code.gs` template.
 * Throws if the expected `SPREADSHEET_URL` or `SPREADSHEET_TIMEZONE` placeholder is missing from `template`.
 */
export function generateAdsScriptForCurrentSheet(
  spreadsheetUrl: string,
  timeZone = 'UTC',
  template = EMBEDDED_ADS_SCRIPT_TEMPLATE ||
    ((globalThis as Record<string, unknown>).__EMBEDDED_ADS_SCRIPT__ as string) ||
    ''
): string {
  if (!template.includes(SPREADSHEET_URL_PLACEHOLDER)) {
    throw new Error(
      'Embedded Ads Script template is missing the SPREADSHEET_URL placeholder.'
    );
  }
  if (!template.includes(SPREADSHEET_TIMEZONE_PLACEHOLDER)) {
    throw new Error(
      'Embedded Ads Script template is missing the SPREADSHEET_TIMEZONE placeholder.'
    );
  }
  return template
    .replace(
      SPREADSHEET_URL_PLACEHOLDER,
      `var SPREADSHEET_URL = ${JSON.stringify(spreadsheetUrl)};`
    )
    .replace(
      SPREADSHEET_TIMEZONE_PLACEHOLDER,
      `var SPREADSHEET_TIMEZONE = ${JSON.stringify(timeZone || 'UTC')};`
    );
}

/**
 * Builds the Data Studio Linking API creation URL bound to the active spreadsheet and its `Data` tab.
 */
export function buildDataStudioCreationUrlForActiveSheet(spreadsheet: {
  getId(): string;
  getSheetByName(name: string): { getSheetId(): number } | null;
}): string {
  const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  const gid = dataSheet?.getSheetId() ?? 100;
  const params = [
    'reporting/create?',
    'c.mode=edit',
    `&c.reportId=${DATA_STUDIO_TEMPLATE_REPORT_ID}`,
    '&r.reportName=DLG_Dashboard_v2',
    '&ds.ds0.connector=googleSheets',
    `&ds.ds0.spreadsheetId=${encodeURIComponent(spreadsheet.getId())}`,
    `&ds.ds0.worksheetId=${encodeURIComponent(String(gid))}`,
    '&ds.ds0.includeHiddenCells=true',
    '&ds.ds0.includeFilteredCells=true',
    '&ds.ds0.refreshFields=false',
  ].join('');
  return `https://datastudio.google.com/${params}`;
}

/**
 * Adds the top-level "🚀 DLG Dashboard v2" menu with the setup and sync actions when the Google Sheet opens.
 */
export function onOpen(): void {
  if (typeof SpreadsheetApp === 'undefined' || typeof SpreadsheetApp.getUi !== 'function') {
    return;
  }
  const ui = SpreadsheetApp.getUi();
  if (!ui) return;

  const menu = ui
    .createMenu('🚀 DLG Dashboard v2')
    .addItem('📋 Copy Google Ads Script', 'showCopyAdsScriptModal')
    .addItem('📊 Create New Data Studio Dashboard', 'openDataStudioCreator')
    .addItem('🚀 Open Saved Dashboard', 'openSavedDashboard')
    .addSeparator()
    .addItem("🔄 Apply Config Changes Now (Reset Today's Lock)", 'applyConfigChangesNow');
  if (isDemoToolsEnabled()) {
    menu
      .addSeparator()
      .addItem('🧪 Seed Synthetic Demo Data (For Testing)', 'seedSyntheticDemoDataFromMenu');
  }
  menu.addToUi();
}

/** True when the `DLG_ENABLE_DEMO_TOOLS` Script Property is `true`. */
export function isDemoToolsEnabled(): boolean {
  return (
    typeof PropertiesService !== 'undefined' &&
    PropertiesService.getScriptProperties().getProperty(DEMO_TOOLS_PROPERTY) === 'true'
  );
}

/**
 * Menu entry for seeding demo data. Checks the demo-tools flag on the server (hiding the menu item alone does
 * not stop the function from being run) and asks for confirmation first.
 */
export function seedSyntheticDemoDataFromMenu(): void {
  if (!isDemoToolsEnabled()) {
    throw new Error(
      `Demo tools are disabled. Set the Script Property ${DEMO_TOOLS_PROPERTY}=true to enable them.`
    );
  }
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Seed demo data?',
    'Adds 2 demo MCCs (using 2 of the 8 MCC slots) and sample recommendations to Config and Data. Real MCCs are not reset.',
    ui.ButtonSet.YES_NO
  );
  if (answer === ui.Button.YES) {
    seedSyntheticDemoData();
  }
}

/**
 * Apps Script simple `onEdit(e)` trigger event type, extending `@types/google-apps-script`'s
 * `GoogleAppsScript.Events.SheetsOnEdit` with structural range/sheet compatibility for unit tests.
 */
export type SheetsOnEditEvent = Partial<Omit<GoogleAppsScript.Events.SheetsOnEdit, 'range'>> & {
  range?: Pick<GoogleAppsScript.Spreadsheet.Range, 'getRow' | 'getColumn' | 'getValue'> & {
    columnStart?: number;
    columnEnd?: number;
    rowStart?: number;
    rowEnd?: number;
    getNumRows?(): number;
    getNumColumns?(): number;
    getSheet(): Pick<GoogleAppsScript.Spreadsheet.Sheet, 'getName' | 'getLastRow'> & {
      getRange(
        rowOrA1: number | string,
        col?: number,
        numRows?: number,
        numCols?: number
      ): {
        getValue(): unknown;
        setValue(val: unknown): unknown;
        getValues(): unknown[][];
        setValues(vals: unknown[][]): unknown;
      };
    };
  };
};

/**
 * Simple trigger that synchronizes the Global Select All checkbox (`A9`),
 * per-MCC `Select All in MCC` checkboxes (`A11:A18`), and individual CID `Include` checkboxes (`A22+`).
 * Optimized for minimum latency:
 * - Fast-exits in 0 RPCs using in-memory `e.range.columnStart` / `rowStart` / `rowEnd` properties when available,
 * - Reads the entire checkbox & ID block (`A9:D<lastRow>`) in a single `getValues()` RPC, and
 * - Only writes cells whose boolean state actually changed (0 write RPCs on intermediate individual CID clicks).
 */
export function onEdit(e?: SheetsOnEditEvent): void {
  const range = e?.range;
  if (!range) return;

  // 1. Zero-RPC fast-exit check on column & row bounds using event properties when present
  const colStart = range.columnStart ?? range.getColumn();
  if (colStart !== 1) return;

  const startRow = range.rowStart ?? range.getRow();
  const endRow = range.rowEnd ?? startRow + (range.getNumRows?.() ?? 1) - 1;
  if (endRow < GLOBAL_SELECT_ALL_ROW || (startRow > MCC_TABLE_END_ROW && endRow < CID_TABLE_START_ROW)) {
    return;
  }

  const sheet = range.getSheet();
  if (sheet.getName() !== CONFIG_SHEET_NAME) {
    return;
  }

  // 2. Single batch read of A9:D<lastRow> (1 read RPC total for all 3 checkbox tiers!)
  const lastRow = Math.max(sheet.getLastRow(), MCC_TABLE_END_ROW);
  const blockRowsCount = lastRow - GLOBAL_SELECT_ALL_ROW + 1;
  const blockWidth = sheetCol(CID_COL.PARENT_MCC_ID);
  const block = sheet.getRange(GLOBAL_SELECT_ALL_ROW, 1, blockRowsCount, blockWidth).getValues();
  const getBlockRow = (sheetRow: number): unknown[] => block[sheetRow - GLOBAL_SELECT_ALL_ROW] ?? new Array(blockWidth).fill('');

  const mccCount = MCC_TABLE_END_ROW - MCC_TABLE_START_ROW + 1;
  const hasCids = lastRow >= CID_TABLE_START_ROW;
  const cidCount = hasCids ? lastRow - CID_TABLE_START_ROW + 1 : 0;

  // Case 1: Global "Select / Deselect All Accounts" checkbox (A9)
  if (startRow <= GLOBAL_SELECT_ALL_ROW && endRow >= GLOBAL_SELECT_ALL_ROW) {
    const isChecked = e?.value !== undefined ? parseSheetBoolean(e.value) : parseSheetBoolean(getBlockRow(GLOBAL_SELECT_ALL_ROW)[0]);

    const nextMccColA: boolean[][] = [];
    let mccChanged = false;
    for (let r = MCC_TABLE_START_ROW; r <= MCC_TABLE_END_ROW; r++) {
      const row = getBlockRow(r);
      const hasMcc = Boolean(normalizeCid(row[MCC_COL.MCC_ID] as string | number));
      const targetVal = hasMcc ? isChecked : false;
      if (parseSheetBoolean(row[MCC_COL.SELECT_ALL]) !== targetVal || (row[MCC_COL.SELECT_ALL] !== false && !hasMcc)) {
        mccChanged = true;
      }
      nextMccColA.push([targetVal]);
    }
    if (mccChanged) {
      sheet.getRange(MCC_TABLE_START_ROW, 1, mccCount, 1).setValues(nextMccColA);
    }

    if (hasCids) {
      const nextCidColA: boolean[][] = [];
      let cidChanged = false;
      for (let r = CID_TABLE_START_ROW; r <= lastRow; r++) {
        const row = getBlockRow(r);
        const hasCid = Boolean(normalizeCid(row[CID_COL.CID] as string | number));
        const targetVal = hasCid ? isChecked : false;
        if (parseSheetBoolean(row[CID_COL.INCLUDE]) !== targetVal || (row[CID_COL.INCLUDE] !== false && !hasCid)) {
          cidChanged = true;
        }
        nextCidColA.push([targetVal]);
      }
      if (cidChanged) {
        sheet.getRange(CID_TABLE_START_ROW, 1, cidCount, 1).setValues(nextCidColA);
      }
    }
    return;
  }

  // Case 2: Per-MCC "Select All in MCC" checkbox (Rows 11–18)
  if (startRow >= MCC_TABLE_START_ROW && startRow <= MCC_TABLE_END_ROW) {
    const mccEnd = Math.min(endRow, MCC_TABLE_END_ROW);
    const toggledMccStates = new Map<string, boolean>();
    for (let r = startRow; r <= mccEnd; r++) {
      const row = getBlockRow(r);
      const mccId = normalizeCid(row[MCC_COL.MCC_ID] as string | number);
      if (mccId) {
        toggledMccStates.set(mccId, parseSheetBoolean(row[MCC_COL.SELECT_ALL]));
      }
    }

    if (hasCids && toggledMccStates.size > 0) {
      const nextCidColA: unknown[][] = [];
      let cidChanged = false;
      let totalValidCids = 0;
      let checkedValidCids = 0;

      for (let r = CID_TABLE_START_ROW; r <= lastRow; r++) {
        const row = getBlockRow(r);
        const rowCid = normalizeCid(row[CID_COL.CID] as string | number);
        if (!rowCid) {
          if (row[CID_COL.INCLUDE] !== false) cidChanged = true;
          nextCidColA.push([false]);
          continue;
        }
        totalValidCids++;
        const rowMcc = normalizeCid(row[CID_COL.PARENT_MCC_ID] as string | number);
        const currentChecked = parseSheetBoolean(row[CID_COL.INCLUDE]);
        const targetChecked = toggledMccStates.has(rowMcc)
          ? toggledMccStates.get(rowMcc)!
          : currentChecked;
        if (currentChecked !== targetChecked) {
          cidChanged = true;
        }
        if (targetChecked) {
          checkedValidCids++;
        }
        nextCidColA.push([targetChecked]);
      }

      if (cidChanged) {
        sheet.getRange(CID_TABLE_START_ROW, 1, cidCount, 1).setValues(nextCidColA);
      }

      if (totalValidCids > 0) {
        const nextGlobal = checkedValidCids === totalValidCids;
        if (parseSheetBoolean(getBlockRow(GLOBAL_SELECT_ALL_ROW)[0]) !== nextGlobal) {
          sheet.getRange(GLOBAL_SELECT_ALL_CELL).setValue(nextGlobal);
        }
      }
    }
    return;
  }

  // Case 3: Individual CID "Include" checkbox (Row 22+)
  if (endRow >= CID_TABLE_START_ROW && hasCids) {
    const cidStart = Math.max(startRow, CID_TABLE_START_ROW);
    const cidEnd = Math.min(endRow, lastRow);
    const affectedMccs = new Set<string>();
    for (let r = cidStart; r <= cidEnd; r++) {
      const parentMccId = normalizeCid(getBlockRow(r)[3] as string | number);
      if (parentMccId) {
        affectedMccs.add(parentMccId);
      }
    }

    // Tally CID counts per MCC and globally in memory from the already-read block
    const mccTotals = new Map<string, { total: number; checked: number }>();
    let totalValidCids = 0;
    let checkedValidCids = 0;

    for (let r = CID_TABLE_START_ROW; r <= lastRow; r++) {
      const row = getBlockRow(r);
      const cid = normalizeCid(row[CID_COL.CID] as string | number);
      if (!cid) continue;
      totalValidCids++;
      const isRowChecked = parseSheetBoolean(row[CID_COL.INCLUDE]);
      if (isRowChecked) checkedValidCids++;

      const mcc = normalizeCid(row[CID_COL.PARENT_MCC_ID] as string | number);
      if (mcc && affectedMccs.has(mcc)) {
        const entry = mccTotals.get(mcc) ?? { total: 0, checked: 0 };
        entry.total++;
        if (isRowChecked) entry.checked++;
        mccTotals.set(mcc, entry);
      }
    }

    // Only write to an MCC row if its "Select All in MCC" boolean state actually changed!
    for (let r = MCC_TABLE_START_ROW; r <= MCC_TABLE_END_ROW; r++) {
      const mccRow = getBlockRow(r);
      const mccId = normalizeCid(mccRow[MCC_COL.MCC_ID] as string | number);
      if (!mccId || !affectedMccs.has(mccId)) continue;
      const stats = mccTotals.get(mccId);
      const allChecked = Boolean(stats && stats.total > 0 && stats.checked === stats.total);
      if (parseSheetBoolean(mccRow[MCC_COL.SELECT_ALL]) !== allChecked) {
        sheet.getRange(r, 1).setValue(allChecked);
      }
    }

    // Only write to Global Select All (A9) if its boolean state actually changed!
    if (totalValidCids > 0) {
      const nextGlobal = checkedValidCids === totalValidCids;
      if (parseSheetBoolean(getBlockRow(GLOBAL_SELECT_ALL_ROW)[0]) !== nextGlobal) {
        sheet.getRange(GLOBAL_SELECT_ALL_CELL).setValue(nextGlobal);
      }
    }
  }
}

/**
 * Ensures a 15-minute time-driven `runScheduledPublish` trigger is running for this Sheet.
 * `ScriptApp.getProjectTriggers()` only returns the current user's triggers, so another editor's trigger is
 * detected through the publisher heartbeat instead: a new trigger is installed only when the current user has
 * none and no publisher has run in the last `PUBLISHER_HEARTBEAT_STALE_MS`.
 */
export function ensureAutoPublishTriggerInstalled(): void {
  if (
    typeof ScriptApp === 'undefined' ||
    typeof ScriptApp.getProjectTriggers !== 'function' ||
    typeof ScriptApp.newTrigger !== 'function'
  ) {
    return;
  }

  const existing = ScriptApp.getProjectTriggers();
  if (existing.some((t) => t.getHandlerFunction() === 'runScheduledPublish')) {
    return;
  }
  const lastHeartbeat = readPublisherHeartbeat();
  if (lastHeartbeat !== null && Date.now() - lastHeartbeat < PUBLISHER_HEARTBEAT_STALE_MS) {
    return; // another editor's publisher trigger is running
  }
  ScriptApp.newTrigger('runScheduledPublish').timeBased().everyMinutes(15).create();
  recordPublisherHeartbeat(); // claim it right away so a second editor does not install another one
}

/**
 * Entrypoint for the 15-minute time-driven background publisher trigger.
 */
export function runScheduledPublish(options?: PublishOptions): void {
  recordPublisherHeartbeat();
  publishFromStagingToData(options);
}

function readPublisherHeartbeat(): number | null {
  if (typeof PropertiesService === 'undefined') return null;
  const raw = PropertiesService.getDocumentProperties().getProperty(PUBLISHER_HEARTBEAT_PROPERTY);
  const ms = Number(raw);
  return raw && Number.isFinite(ms) ? ms : null;
}

/** Records "a publisher ran now" in a Document Property (not a sheet cell, so idle ticks stay write-free). */
function recordPublisherHeartbeat(): void {
  if (typeof PropertiesService === 'undefined') return;
  PropertiesService.getDocumentProperties().setProperty(
    PUBLISHER_HEARTBEAT_PROPERTY,
    String(Date.now())
  );
}

/**
 * Displays a modal dialog containing the pre-configured `Code.gs` script with the current Sheet URL injected.
 */
export function showCopyAdsScriptModal(): void {
  ensureAutoPublishTriggerInstalled();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptCode = generateAdsScriptForCurrentSheet(
    ss.getUrl(),
    ss.getSpreadsheetTimeZone() || 'UTC'
  );
  const escaped = escapeHtml(scriptCode);

  const html = `
    <div style="font-family: Arial, sans-serif; padding: 4px 12px 12px; box-sizing: border-box;">
      <p style="margin-top:0; margin-bottom:6px;"><b>Setup Steps (repeat for each MCC):</b></p>
      <ol style="margin-top:0; margin-bottom:12px; padding-left:20px; font-size:12px; line-height:1.45;">
        <li>Click <b>📋 Copy to Clipboard</b> below.</li>
        <li>In Google Ads (<b>Tools &amp; settings → Bulk actions → Scripts</b>), create a new script in your MCC, paste the code, click <b>Save</b>, then <b>Authorize</b> and <b>Preview</b> once to register its accounts on the <b>Config</b> tab.</li>
        <li>Set the script <b>Frequency = Hourly</b> (required for automatic daily updates and 30-min timeout checkpoint recovery). <i>This Sheet's 15-minute background publisher trigger is installed automatically.</i></li>
        <li>Return to the <b>Config</b> tab in this Sheet and check the accounts you want to include.</li>
      </ol>
      <div style="margin-bottom:8px;">
        <button id="copyBtn" onclick="copyAdsScript(this)" style="padding:8px 16px; background:#1a73e8; color:#fff; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">📋 Copy to Clipboard</button>
      </div>
      <textarea id="code" readonly onclick="this.select();" style="width:100%; height:200px; box-sizing:border-box; font-family:monospace; font-size:11px;">${escaped}</textarea>
      <script>
        (function() {
          var el = document.getElementById('code');
          if (el) { el.focus(); el.select(); }
        })();
        function copyAdsScript(btn) {
          var el = document.getElementById('code');
          el.focus();
          el.select();
          function fallbackCopy() {
            try {
              if (document.execCommand && document.execCommand('copy')) {
                btn.innerText = '✅ Copied!';
                return;
              }
            } catch (e) {}
            btn.innerText = '⚠️ Select & press Ctrl+C / Cmd+C';
          }
          if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(el.value).then(function() {
              btn.innerText = '✅ Copied!';
            }).catch(fallbackCopy);
          } else {
            fallbackCopy();
          }
        }
      </script>
    </div>
  `;
  if (typeof HtmlService !== 'undefined' && typeof HtmlService.createHtmlOutput === 'function') {
    const output = HtmlService.createHtmlOutput(html).setWidth(680).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(output, '📋 Copy DLG v2 Google Ads Script');
  }
}

/**
 * Opens a modal dialog launching the Data Studio Linking API URL to create a new report copy.
 */
export function openDataStudioCreator(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDataSheetInitialized(ss);
  const url = buildDataStudioCreationUrlForActiveSheet(ss);
  openUrlModal(url, '📊 Create New Data Studio Dashboard');
}

/**
 * Makes sure the `Data` tab exists and, when empty, holds the header and schema dummy row, so Data Studio
 * binds to its real gid with all typed columns even before the first publish.
 */
export function ensureDataSheetInitialized(ss: SpreadsheetLike): void {
  withSheetLock(ss, 'SHEET_INIT_DATA', () => {
    const dataSheet = getOrCreateSheet(ss, DATA_SHEET_NAME);
    if (dataSheet.getLastRow() === 0) {
      writeReportRowsToSheet(dataSheet, [], ss);
    }
  });
}

/**
 * Opens the user's saved Data Studio Dashboard URL from cell `Setup!D13`, or alerts if unset.
 */
export function openSavedDashboard(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SETUP_SHEET_NAME);
  const savedUrl = String(setupSheet?.getRange(SAVED_DASHBOARD_URL_CELL)?.getValue() ?? '').trim();

  if (!isAllowedDashboardUrl(savedUrl)) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'No Saved Dashboard URL Found',
      `Paste your saved Data Studio report URL (starting with https://lookerstudio.google.com/ or https://datastudio.google.com/) into cell ${SAVED_DASHBOARD_URL_CELL} on the ${SETUP_SHEET_NAME} tab first.`,
      ui.ButtonSet.OK
    );
    return;
  }
  openUrlModal(savedUrl, '🚀 Open Saved Data Studio Dashboard');
}

/**
 * Applies mid-day `Config` edits immediately:
 * - Publisher-side changes only (Target Currency, unchecked accounts): clears today's config
 *   lock, FX cache and publish status/signature, keeps every `_MCC_` checkpoint, and republishes at once (no Google
 *   Ads re-runs, no flip to "stale" on the next trigger).
 * - `Exclude Shopping & Feed-PMax` changed: resets the checkpoints of all active MCCs, because Shopping and
 *   feed-connected PMax campaigns are filtered when the Google Ads Script queries the API.
 * - Newly checked accounts missing from an MCC's completed accounts: resets only that MCC's checkpoint.
 * Reset MCCs show as pending/stale until their next hourly Google Ads Script run.
 */
export function applyConfigChangesNow(): void {
  ensureAutoPublishTriggerInstalled();
  const ss: SpreadsheetLike = SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = resolveSpreadsheetDate(ss);

  withSheetLock(ss, 'SHEET_RESET', () => {
    const configSheet = getOrCreateSheet(ss, CONFIG_SHEET_NAME, { index: 0 });
    syncAllMccSheetsToConfig(ss, configSheet);

    const lockedConfig = readLockedDailyConfig(ss, dateStr);
    const liveConfig = readGlobalConfig(configSheet);
    const shoppingFilterChanged =
      lockedConfig !== null &&
      Boolean(lockedConfig.excludeShoppingAndPmax) !== Boolean(liveConfig.excludeShoppingAndPmax);

    const resetMccIds: string[] = [];
    for (const mcc of getMccSelectionSummary(configSheet).filter((m) => m.isActive)) {
      const checkpoint = getMccCheckpoint(ss, dateStr, mcc.mccId);
      const hasProgressToday =
        checkpoint.dateStr === dateStr &&
        (checkpoint.status === 'COMPLETED' || checkpoint.completedCids.length > 0);
      if (!hasProgressToday) {
        continue; // nothing collected today yet: the next run uses the new Config anyway
      }
      // An unfinished run picks up newly checked accounts by itself; a COMPLETED one must be reset
      const completed = new Set(checkpoint.completedCids);
      const hasNewlyCheckedCids =
        checkpoint.status === 'COMPLETED' &&
        getSelectedCidsForMcc(configSheet, mcc.mccId).some((cid) => !completed.has(cid));
      if (shoppingFilterChanged || hasNewlyCheckedCids) {
        resetMccIds.push(mcc.mccId);
      }
    }

    resetTodayPublisherState(ss, dateStr);
    for (const mccId of resetMccIds) {
      clearMccSheetCheckpoint(ss, mccId);
    }

    const resetNote =
      resetMccIds.length > 0
        ? ` Re-collecting ${resetMccIds.length} MCC(s) on their next hourly Google Ads Script run${shoppingFilterChanged ? ' (Exclude Shopping & Feed-PMax changed)' : ' (newly checked accounts)'}.`
        : ' Republishing from the data already collected today.';
    configSheet
      .getRange(STATUS_BANNER_CELL)
      .setValue(`🔄 Config changes applied (${dateStr}).${resetNote}`);
  });

  // Immediately re-publish from staged data
  publishFromStagingToData({
    spreadsheet: ss,
    force: true,
  });
}

/**
 * Data Studio report links only: an https URL on `lookerstudio.google.com` or `datastudio.google.com` with no
 * whitespace, quotes, angle brackets, or backslashes. (Apps Script's V8 runtime has no WHATWG `URL` class.)
 */
const DASHBOARD_URL_RE = /^https:\/\/(lookerstudio|datastudio)\.google\.com\/[^\s"'<>\\]*$/i;

function isAllowedDashboardUrl(url: string): boolean {
  return DASHBOARD_URL_RE.test(url);
}

function escapeHtml(raw: string): string {
  return raw
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function openUrlModal(url: string, title: string): void {
  if (!isAllowedDashboardUrl(url)) {
    throw new Error('Only Data Studio report links (lookerstudio.google.com / datastudio.google.com) can be opened.');
  }
  const safeHref = escapeHtml(url);
  const safeTitle = escapeHtml(title);
  // The inline script reads the URL from the escaped anchor, so no URL text is ever placed inside <script>
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 16px;">
      <p>Click below if the tab did not open automatically:</p>
      <p><a id="target" href="${safeHref}" target="_blank" rel="noopener" style="font-weight:bold; color:#1a73e8;">${safeTitle}</a></p>
      <script>window.open(document.getElementById('target').href, '_blank');</script>
    </div>
  `;
  if (typeof HtmlService !== 'undefined' && typeof HtmlService.createHtmlOutput === 'function') {
    const output = HtmlService.createHtmlOutput(html).setWidth(480).setHeight(180);
    SpreadsheetApp.getUi().showModalDialog(output, title);
  }
}
