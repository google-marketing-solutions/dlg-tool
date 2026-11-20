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
 * @fileoverview Manager for the visible `Setup` and `Config` tabs in the Master Template Google Sheet.
 * `Setup` layout: step-by-step setup guide and the saved Data Studio dashboard URL `D13`.
 * `Config` layout: status banner `B2`, Target Currency `B5`, Exclude Shopping & Feed-PMax `B6`,
 * Global Select All `A9`, Discovered MCCs `Rows 11–18`, and Child Accounts `Row 22+`.
 * Handles:
 * - Reading global settings (`readGlobalConfig`: Target Currency, Exclude Shopping & Feed-PMax),
 * - Auto-registering MCCs, appending discovered child accounts, and updating renamed account/MCC
 *   names in place (`syncDiscoveredAccountsToConfig` and `syncAllMccSheetsToConfig`),
 * - Populating the `Target Currency` dropdown (`B5`) from the currencies of accounts in `Config`, and
 * - Filtering checked CIDs per MCC and marking MCCs with `0` selected accounts as Inactive/Skipped.
 * Bundled only into `dist/sheet/Code.gs` (sole reader/writer of `Config`, which locks daily state into `_RunState`).
 */

import {MCC_SHEET_PREFIX, readMccSheetState} from '../shared/sheets/mcc-state';
import {
  type SheetLike,
  type SpreadsheetLike,
  escapeFormulaString,
  parseSheetBoolean,
  unescapeFormulaString,
} from '../shared/sheets/sheet-utils';
import type {ChildAccountInput, GlobalConfig, MccIdentity} from '../shared/types';
import {DLG_VERSION} from '../shared/version';
import {formatCidHyphenated, normalizeCid} from '../shared/utils/cid';
import {normalizeCurrencyCode} from '../shared/utils/currency';

export type {ChildAccountInput, MccIdentity};

/** Name of the visible user-facing setup guide tab. */
export const SETUP_SHEET_NAME = 'Setup';
/** A1 address of the saved Data Studio dashboard URL cell on the `Setup` tab. */
export const SAVED_DASHBOARD_URL_CELL = 'D13';

/** Name of the visible user-facing configuration tab (looked up by name, not gid). */
export const CONFIG_SHEET_NAME = 'Config';
/** A1 address of the live sync status banner cell on the `Config` tab. */
export const STATUS_BANNER_CELL = 'B2';
/** A1 address of the Target Currency dropdown cell on the `Config` tab. */
export const TARGET_CURRENCY_CELL = 'B5';
/** A1 address of the Exclude Shopping & Feed-PMax toggle cell on the `Config` tab. */
export const EXCLUDE_SHOPPING_PMAX_CELL = 'B6';
/** A1 range of the Global Settings cells (`Target Currency` and `Exclude Shopping & Feed PMax`) on the `Config` tab. */
export const GLOBAL_SETTINGS_RANGE = 'B5:B6';
/** 1-based row of the Global Select/Deselect All Accounts checkbox on the `Config` tab. */
export const GLOBAL_SELECT_ALL_ROW = 9;
/** A1 address of the Global Select/Deselect All Accounts checkbox on the `Config` tab. */
export const GLOBAL_SELECT_ALL_CELL = `A${GLOBAL_SELECT_ALL_ROW}`;

/** First 1-based row of the Discovered MCCs table (`Rows 11–18`) on the `Config` tab. */
export const MCC_TABLE_START_ROW = 11;
/** Last 1-based row of the Discovered MCCs table (`Rows 11–18`) on the `Config` tab. */
export const MCC_TABLE_END_ROW = 18;
/** Maximum number of MCCs supported per Google Sheet. */
export const MAX_SUPPORTED_MCCS = MCC_TABLE_END_ROW - MCC_TABLE_START_ROW + 1;
/** First 1-based data row of the Generated Child Accounts (CIDs) table on the `Config` tab. */
export const CID_TABLE_START_ROW = 22;

/** Width (columns A:F) of both the Discovered MCCs table and the Child Accounts table. */
export const CONFIG_TABLE_WIDTH = 6;
/** 0-based column offsets inside the Discovered MCCs table (columns A:F, rows 11–18). */
export const MCC_COL = {
  SELECT_ALL: 0,
  MCC_ID: 1,
  NAME: 2,
  DISCOVERED_COUNT: 3,
  SELECTED_COUNT: 4,
  STATUS: 5,
} as const;
/** 0-based column offsets inside the Child Accounts table (columns A:F, from row `CID_TABLE_START_ROW`). */
export const CID_COL = {
  INCLUDE: 0,
  CID: 1,
  NAME: 2,
  PARENT_MCC_ID: 3,
  PARENT_MCC_NAME: 4,
  CURRENCY: 5,
} as const;

/** Header row (row 10) of the Discovered MCCs table in the Master Template, by `MCC_COL` offset. */
export const MCC_TABLE_HEADERS: readonly string[] = [
  'Select All in MCC',
  'MCC ID',
  'MCC Name',
  'Discovered Accounts',
  'Selected Accounts',
  'Sync Status',
];
/** Header row (row 21) of the Child Accounts table in the Master Template, by `CID_COL` offset. */
export const CID_TABLE_HEADERS: readonly string[] = [
  'Include',
  'Account CID',
  'Account Name',
  'Parent MCC ID',
  'Parent MCC Name',
  'Currency',
];

/** Converts a 0-based table column offset into a 1-based Sheets column number. */
export function sheetCol(offset: number): number {
  return offset + 1;
}

/** Converts a 0-based table column offset (A–Z only) into its column letter. */
export function sheetColLetter(offset: number): string {
  return String.fromCharCode('A'.charCodeAt(0) + offset);
}

/** Summary of discovered and selected child accounts for a single MCC row in `Config`. */
export interface MccSelectionSummary {
  mccId: string;
  mccName: string;
  selectAllChecked: boolean;
  discoveredCount: number;
  selectedCount: number;
  isActive: boolean;
}

/**
 * Throws when the `Config` table header rows (rows 10 and 21) do not match the Master Template, for example
 * after a row or column was inserted above or inside the tables, so the scripts never read or write the wrong
 * cells. A completely blank `Config` tab (both header rows empty, i.e. not created from the template) passes.
 */
export function assertConfigLayout(configSheet: SheetLike): void {
  const checks = [
    { row: MCC_TABLE_START_ROW - 1, expected: MCC_TABLE_HEADERS },
    { row: CID_TABLE_START_ROW - 1, expected: CID_TABLE_HEADERS },
  ];
  const actualRows = checks.map((c) =>
    configSheet
      .getRange(c.row, 1, 1, CONFIG_TABLE_WIDTH)
      .getValues()[0]
      .map((v) => String(v ?? '').trim())
  );
  if (actualRows.every((cells) => cells.every((v) => v === ''))) {
    return;
  }
  checks.forEach((check, i) => {
    check.expected.forEach((header, col) => {
      const actual = actualRows[i][col] ?? '';
      if (actual !== header) {
        throw new Error(
          `Config layout changed: expected "${header}" in ${sheetColLetter(col)}${check.row}, found "${actual}". ` +
            'Restore the Master Template layout of the Config tab (do not insert or delete rows or columns above or inside its tables).'
        );
      }
    });
  });
}

/** One row of the Child Accounts table, before formula-escaping. */
interface CidRow {
  include: boolean;
  cid: string;
  name: string;
  parentMccId: string;
  parentMccName: string;
  currency: string;
}

interface CidTableDiff {
  existingRowsCount: number;
  retainedRows: unknown[][];
  rowsToAppend: unknown[][];
  renamedCells: Array<{ row: number; col: number; value: string }>;
  removedCidsCount: number;
  nextWriteRow: number;
}

/**
 * Reads the global configuration settings and per-MCC selected account CIDs from the `Config` tab.
 * If `Target Currency` is not a 3-letter currency code (empty, `NONE`, or anything else), defaults to the
 * first currency registered in the Child Accounts table.
 */
export function readGlobalConfig(configSheet: SheetLike): GlobalConfig {
  const values = configSheet.getRange(GLOBAL_SETTINGS_RANGE).getValues();
  const targetCurrency =
    normalizeCurrencyCode(values[0]?.[0]) ?? getFirstRegisteredCurrency(configSheet);

  const excludeRaw = values[1]?.[0];
  const excludeShoppingAndPmax = parseSheetBoolean(excludeRaw);

  const selectedCidsByMcc: Record<string, string[]> = {};
  for (const summary of getMccSelectionSummary(configSheet)) {
    selectedCidsByMcc[summary.mccId] = getSelectedCidsForMcc(configSheet, summary.mccId);
  }

  return {
    targetCurrency,
    excludeShoppingAndPmax,
    selectedCidsByMcc,
  };
}

/**
 * Returns the first currency code registered in the Child Accounts table, falling back to `'USD'`
 * if no child accounts have been registered yet.
 */
export function getFirstRegisteredCurrency(configSheet: SheetLike): string {
  const lastRow = configSheet.getLastRow();
  if (lastRow >= CID_TABLE_START_ROW) {
    const cidCount = lastRow - CID_TABLE_START_ROW + 1;
    const currRows = configSheet
      .getRange(CID_TABLE_START_ROW, sheetCol(CID_COL.CURRENCY), cidCount, 1)
      .getValues();
    for (const r of currRows) {
      const curr = normalizeCurrencyCode(r[0]);
      if (curr) {
        return curr;
      }
    }
  }
  return 'USD';
}

/**
 * Syncs the MCC and its child CIDs into the config table.
 * - Registers the current MCC in the Discovered MCCs Table if not yet present.
 * - Appends newly discovered child accounts to the Generated CIDs Table.
 * - Updates renamed account names and MCC names.
 * - Updates the `Target Currency` dropdown validation options.
 */
export function syncDiscoveredAccountsToConfig(
  configSheet: SheetLike,
  mcc: MccIdentity,
  childAccounts: ChildAccountInput[]
): { isNewlyRegisteredMcc: boolean; newlyAddedCidsCount: number; removedCidsCount: number } {
  // 1. Register or update the MCC row
  const { isNewlyRegisteredMcc, defaultIncludeForNewCids } = syncMccTableRow(configSheet, mcc);

  // 2. Reconcile existing CID rows against live childAccounts
  const diff = reconcileCidRows(configSheet, mcc, childAccounts, defaultIncludeForNewCids);

  // 3. Apply CID table changes (rewrite if accounts were unlinked, otherwise rename + append)
  applyCidTableChanges(configSheet, diff);

  // 4. Refresh Target Currency dropdown (B5) from all active CID rows
  updateTargetCurrencyDropdown(configSheet, [...diff.retainedRows, ...diff.rowsToAppend]);

  return {
    isNewlyRegisteredMcc,
    newlyAddedCidsCount: diff.rowsToAppend.length,
    removedCidsCount: diff.removedCidsCount,
  };
}

/** A child account discovered under more than one registered MCC (normalized 10-digit IDs). */
export interface DuplicateAccount {
  cid: string;
  mccIds: string[];
}

/** An MCC whose `_MCC_<id>` tab was written by a different Google Ads Script version than this Sheet's. */
export interface OutdatedScript {
  mccId: string;
  /** `null` when the tab was written before script versions were recorded. */
  version: string | null;
}

/** Result of `syncAllMccSheetsToConfig`. */
export interface MccSyncResult {
  newlyRegisteredMccs: number;
  newlyAddedCids: number;
  latestStatusMessages: string[];
  /**
   * Accounts discovered under more than one MCC (the script was installed in a sub-manager account as well
   * as its parent). Only the first MCC synced owns such an account in `Config`.
   */
  duplicateAccounts: DuplicateAccount[];
  /** MCCs whose Google Ads Script copy does not match this Sheet's version (`DLG_VERSION`). */
  outdatedScripts: OutdatedScript[];
}

/**
 * Scans all `_MCC_<normMccId>` sheets in the workbook and syncs their discovered MCCs and child
 * accounts into the `Config` tab.
 */
export function syncAllMccSheetsToConfig(
  spreadsheet: SpreadsheetLike,
  configSheet: SheetLike
): MccSyncResult {
  assertConfigLayout(configSheet);
  const sheets = spreadsheet.getSheets();
  let newlyRegisteredMccs = 0;
  let newlyAddedCids = 0;
  const latestStatusMessages: string[] = [];
  const mccIdsByCid = new Map<string, string[]>();
  const outdatedScripts: OutdatedScript[] = [];

  for (const s of sheets) {
    const name = s.getName();
    if (!name.startsWith(MCC_SHEET_PREFIX)) continue;
    const normMcc = normalizeCid(name.slice(MCC_SHEET_PREFIX.length));
    if (!normMcc) continue;

    const payload = readMccSheetState(spreadsheet, normMcc);
    if (!payload) continue;

    const syncRes = syncDiscoveredAccountsToConfig(
      configSheet,
      payload.mcc,
      payload.childAccounts ?? []
    );
    if (syncRes.isNewlyRegisteredMcc) newlyRegisteredMccs++;
    newlyAddedCids += syncRes.newlyAddedCidsCount;
    if (payload.statusMessage) {
      latestStatusMessages.push(payload.statusMessage);
    }
    if (payload.scriptVersion !== DLG_VERSION) {
      outdatedScripts.push({ mccId: normMcc, version: payload.scriptVersion ?? null });
    }
    for (const acc of payload.childAccounts ?? []) {
      const cid = normalizeCid(acc.customerId);
      if (!cid) continue;
      const mccIds = mccIdsByCid.get(cid) ?? [];
      if (!mccIds.includes(normMcc)) mccIds.push(normMcc);
      mccIdsByCid.set(cid, mccIds);
    }
  }

  const duplicateAccounts = Array.from(mccIdsByCid.entries())
    .filter(([, mccIds]) => mccIds.length > 1)
    .map(([cid, mccIds]) => ({ cid, mccIds: [...mccIds].sort() }))
    .sort((x, y) => x.cid.localeCompare(y.cid));

  outdatedScripts.sort((x, y) => x.mccId.localeCompare(y.mccId));
  return { newlyRegisteredMccs, newlyAddedCids, latestStatusMessages, duplicateAccounts, outdatedScripts };
}

/**
 * Registers `mcc` in the Discovered MCCs Table if not yet present, or updates its display name in
 * place if renamed. Returns whether it was newly registered and whether its "Select All in MCC"
 * checkbox is checked.
 */
function syncMccTableRow(
  configSheet: SheetLike,
  mcc: MccIdentity
): { isNewlyRegisteredMcc: boolean; defaultIncludeForNewCids: boolean } {
  const normalizedMccId = normalizeCid(mcc.mccId);
  const formattedMccId = formatCidHyphenated(mcc.mccId);
  const mccTableValues = configSheet
    .getRange(MCC_TABLE_START_ROW, 1, MAX_SUPPORTED_MCCS, CONFIG_TABLE_WIDTH)
    .getValues();

  const registeredMccIds = mccTableValues
    .map((row) => normalizeCid(row[MCC_COL.MCC_ID]))
    .filter(Boolean);
  const mccRowIndex = registeredMccIds.indexOf(normalizedMccId);

  if (mccRowIndex === -1) {
    if (registeredMccIds.length >= MAX_SUPPORTED_MCCS) {
      throw new Error(
        `Maximum of ${MAX_SUPPORTED_MCCS} MCCs already registered in Config (Rows ${MCC_TABLE_START_ROW}-${MCC_TABLE_END_ROW}).`
      );
    }
    const sheetRow = MCC_TABLE_START_ROW + registeredMccIds.length;
    const mccIdCell = `${sheetColLetter(MCC_COL.MCC_ID)}${sheetRow}`;
    const selectedCountCell = `${sheetColLetter(MCC_COL.SELECTED_COUNT)}${sheetRow}`;
    const parentCol = sheetColLetter(CID_COL.PARENT_MCC_ID);
    const includeCol = sheetColLetter(CID_COL.INCLUDE);
    const parentRange = `$${parentCol}$${CID_TABLE_START_ROW}:$${parentCol}`;
    const includeRange = `$${includeCol}$${CID_TABLE_START_ROW}:$${includeCol}`;

    const row: unknown[] = new Array(CONFIG_TABLE_WIDTH);
    row[MCC_COL.SELECT_ALL] = false;
    row[MCC_COL.MCC_ID] = formattedMccId;
    row[MCC_COL.NAME] = escapeFormulaString(mcc.mccName);
    row[MCC_COL.DISCOVERED_COUNT] = `=IF(${mccIdCell}="", "", COUNTIF(${parentRange}, ${mccIdCell}))`;
    row[MCC_COL.SELECTED_COUNT] =
      `=IF(${mccIdCell}="", "", COUNTIFS(${parentRange}, ${mccIdCell}, ${includeRange}, TRUE))`;
    row[MCC_COL.STATUS] =
      `=IF(${mccIdCell}="", "", IF(${selectedCountCell}=0, "⏸️ INACTIVE (0 Selected — Skipped)", "🟢 ACTIVE (" & ${selectedCountCell} & " Selected)"))`;

    // MCC ID and name are text; the count/status columns are our own formulas and must stay formulas
    configSheet.getRange(sheetRow, sheetCol(MCC_COL.MCC_ID), 1, 2).setNumberFormat('@');
    configSheet.getRange(sheetRow, 1, 1, CONFIG_TABLE_WIDTH).setValues([row]);
    configSheet.getRange(sheetRow, sheetCol(MCC_COL.SELECT_ALL), 1, 1).insertCheckboxes();
    return { isNewlyRegisteredMcc: true, defaultIncludeForNewCids: false };
  }

  const mccRow = mccTableValues[mccRowIndex];
  const defaultIncludeForNewCids = parseSheetBoolean(mccRow[MCC_COL.SELECT_ALL]);
  if (mcc.mccName && readTextCell(mccRow[MCC_COL.NAME]) !== mcc.mccName) {
    writeTextCell(configSheet, MCC_TABLE_START_ROW + mccRowIndex, sheetCol(MCC_COL.NAME), mcc.mccName);
  }
  return { isNewlyRegisteredMcc: false, defaultIncludeForNewCids };
}

/**
 * Compares existing CID rows against `childAccounts` for `mcc`, computing retained rows,
 * in-place name updates, unlinked account removals, and newly discovered rows to append.
 */
function reconcileCidRows(
  configSheet: SheetLike,
  mcc: MccIdentity,
  childAccounts: ChildAccountInput[],
  defaultIncludeForNewCids: boolean
): CidTableDiff {
  const normalizedMccId = normalizeCid(mcc.mccId);
  const formattedMccId = formatCidHyphenated(mcc.mccId);

  const liveAccountsByCid = new Map<string, ChildAccountInput>();
  for (const acc of childAccounts) {
    const normCid = normalizeCid(acc.customerId);
    if (normCid && !liveAccountsByCid.has(normCid)) {
      liveAccountsByCid.set(normCid, acc);
    }
  }

  const lastRow = Math.max(configSheet.getLastRow(), CID_TABLE_START_ROW - 1);
  const existingRowsCount = Math.max(0, lastRow - CID_TABLE_START_ROW + 1);
  const existingRows =
    existingRowsCount > 0
      ? configSheet.getRange(CID_TABLE_START_ROW, 1, existingRowsCount, CONFIG_TABLE_WIDTH).getValues()
      : [];

  const existingCids = new Set<string>();
  const retainedRows: unknown[][] = [];
  const renamedCells: Array<{ row: number; col: number; value: string }> = [];
  let removedCidsCount = 0;
  let nextWriteRow = CID_TABLE_START_ROW;

  for (let i = 0; i < existingRows.length; i++) {
    const row = existingRows[i];
    const cid = normalizeCid(row[CID_COL.CID]);
    const parentMcc = normalizeCid(row[CID_COL.PARENT_MCC_ID]);
    if (!cid) {
      continue;
    }
    if (parentMcc === normalizedMccId && !liveAccountsByCid.has(cid)) {
      removedCidsCount++;
      continue;
    }
    existingCids.add(cid);

    const liveName = liveAccountsByCid.get(cid)?.name;
    if (liveName && readTextCell(row[CID_COL.NAME]) !== liveName) {
      row[CID_COL.NAME] = liveName;
      renamedCells.push({ row: CID_TABLE_START_ROW + i, col: sheetCol(CID_COL.NAME), value: liveName });
    }
    if (
      parentMcc === normalizedMccId &&
      mcc.mccName &&
      readTextCell(row[CID_COL.PARENT_MCC_NAME]) !== mcc.mccName
    ) {
      row[CID_COL.PARENT_MCC_NAME] = mcc.mccName;
      renamedCells.push({
        row: CID_TABLE_START_ROW + i,
        col: sheetCol(CID_COL.PARENT_MCC_NAME),
        value: mcc.mccName,
      });
    }
    // Names are re-escaped on write (escapeFormulaString is idempotent), so rewrites stay inert
    row[CID_COL.NAME] = escapeFormulaString(String(row[CID_COL.NAME] ?? ''));
    row[CID_COL.PARENT_MCC_NAME] = escapeFormulaString(String(row[CID_COL.PARENT_MCC_NAME] ?? ''));
    retainedRows.push(row);
    nextWriteRow = CID_TABLE_START_ROW + i + 1;
  }

  const rowsToAppend: unknown[][] = [];
  for (const [normCid, acc] of liveAccountsByCid) {
    if (existingCids.has(normCid)) {
      continue;
    }
    rowsToAppend.push(
      toCidRowValues({
        include: defaultIncludeForNewCids,
        cid: formatCidHyphenated(normCid),
        name: acc.name,
        parentMccId: formattedMccId,
        parentMccName: mcc.mccName,
        currency: String(acc.currencyCode ?? '').trim().toUpperCase(),
      })
    );
  }

  return {
    existingRowsCount: existingRows.length,
    retainedRows,
    rowsToAppend,
    renamedCells,
    removedCidsCount,
    nextWriteRow,
  };
}

/**
 * Writes the reconciled CID table changes to `configSheet`:
 * - When accounts were unlinked (`removedCidsCount > 0`), rewrites the table in place and trims excess rows.
 * - Otherwise, updates renamed cells in place and appends any newly discovered rows at the bottom.
 */
function applyCidTableChanges(configSheet: SheetLike, diff: CidTableDiff): void {
  if (diff.removedCidsCount > 0) {
    rewriteCidTable(configSheet, [...diff.retainedRows, ...diff.rowsToAppend], diff.existingRowsCount);
    return;
  }

  for (const cell of diff.renamedCells) {
    writeTextCell(configSheet, cell.row, cell.col, cell.value);
  }
  if (diff.rowsToAppend.length > 0) {
    writeFormattedCidRows(configSheet, diff.nextWriteRow, diff.rowsToAppend);
  }
}

function rewriteCidTable(
  configSheet: SheetLike,
  combinedRows: unknown[][],
  existingRowsCount: number
): void {
  const totalToWrite = Math.max(existingRowsCount, combinedRows.length);
  if (totalToWrite === 0) {
    return;
  }
  ensureSheetMaxRows(configSheet, CID_TABLE_START_ROW + totalToWrite - 1);

  const paddedRows = [...combinedRows];
  while (paddedRows.length < totalToWrite) {
    paddedRows.push(new Array(CONFIG_TABLE_WIDTH).fill(''));
  }
  if (combinedRows.length > 0) {
    formatCidTextColumns(configSheet, CID_TABLE_START_ROW, combinedRows.length);
  }
  configSheet.getRange(CID_TABLE_START_ROW, 1, totalToWrite, CONFIG_TABLE_WIDTH).setValues(paddedRows);
  if (combinedRows.length > 0) {
    configSheet
      .getRange(CID_TABLE_START_ROW, sheetCol(CID_COL.INCLUDE), combinedRows.length, 1)
      .insertCheckboxes();
  }
  // Keep at least 1 data row so future appended rows copy data-row formatting rather than the header row
  const keptRows = Math.max(combinedRows.length, 1);
  if (existingRowsCount > keptRows) {
    configSheet.deleteRows(CID_TABLE_START_ROW + keptRows, existingRowsCount - keptRows);
  }
}

function writeFormattedCidRows(
  configSheet: SheetLike,
  startRow: number,
  rows: unknown[][]
): void {
  ensureSheetMaxRows(configSheet, startRow + rows.length - 1);
  formatCidTextColumns(configSheet, startRow, rows.length);
  configSheet.getRange(startRow, 1, rows.length, CONFIG_TABLE_WIDTH).setValues(rows);
  configSheet.getRange(startRow, sheetCol(CID_COL.INCLUDE), rows.length, 1).insertCheckboxes();
}

/** Formats every Child Accounts column except the Include checkbox (B:F) as plain text. */
function formatCidTextColumns(configSheet: SheetLike, startRow: number, numRows: number): void {
  configSheet
    .getRange(startRow, sheetCol(CID_COL.CID), numRows, CONFIG_TABLE_WIDTH - 1)
    .setNumberFormat('@');
}

/** Builds a Child Accounts row; account and MCC names come from Google Ads and are formula-escaped. */
function toCidRowValues(r: CidRow): unknown[] {
  const row: unknown[] = new Array(CONFIG_TABLE_WIDTH);
  row[CID_COL.INCLUDE] = r.include;
  row[CID_COL.CID] = r.cid;
  row[CID_COL.NAME] = escapeFormulaString(r.name);
  row[CID_COL.PARENT_MCC_ID] = r.parentMccId;
  row[CID_COL.PARENT_MCC_NAME] = escapeFormulaString(r.parentMccName);
  row[CID_COL.CURRENCY] = r.currency;
  return row;
}

/** Writes an external name into a single cell as inert text (text format + formula escape). */
function writeTextCell(configSheet: SheetLike, row: number, col: number, value: string): void {
  const cell = configSheet.getRange(row, col);
  cell.setNumberFormat('@');
  cell.setValue(escapeFormulaString(value));
}

/** Reads a text cell written by `writeTextCell` / `toCidRowValues` back to its original value. */
function readTextCell(value: unknown): string {
  return unescapeFormulaString(String(value ?? ''));
}

function ensureSheetMaxRows(configSheet: SheetLike, neededMaxRow: number): void {
  if (neededMaxRow > configSheet.getMaxRows()) {
    configSheet.insertRowsAfter(configSheet.getMaxRows(), neededMaxRow - configSheet.getMaxRows());
  }
}

function updateTargetCurrencyDropdown(
  configSheet: SheetLike,
  activeCidRows: unknown[][]
): void {
  const currencies = new Set<string>();
  let firstRegisteredCurrency = '';
  for (const row of activeCidRows) {
    const curr = normalizeCurrencyCode(row[CID_COL.CURRENCY]);
    if (curr) {
      currencies.add(curr);
      if (!firstRegisteredCurrency) {
        firstRegisteredCurrency = curr;
      }
    }
  }

  const cell = configSheet.getRange(TARGET_CURRENCY_CELL);
  if (!normalizeCurrencyCode(cell.getValue()) && firstRegisteredCurrency) {
    cell.setValue(firstRegisteredCurrency);
  }

  const existingValidation = cell.getDataValidation();
  const rawExistingList =
    (existingValidation?.getCriteriaValues()?.[0] as string[] | undefined) ?? [];
  const previousOptions = Array.isArray(rawExistingList) ? rawExistingList.map(String) : [];

  // Options = currencies of accounts currently in Config plus the selected Target Currency, so currencies
  // of unlinked accounts (or template placeholders such as NONE) drop out.
  const selected = normalizeCurrencyCode(cell.getValue());
  if (selected) {
    currencies.add(selected);
  }
  const options = Array.from(currencies).sort();
  if (options.length === 0) {
    return;
  }

  // Skip setDataValidation RPC if the option list is already identical
  if (
    previousOptions.length === options.length &&
    previousOptions.every((val, idx) => val === options[idx])
  ) {
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
}

/**
 * Returns normalized 10-digit CIDs where `Include (Col A) === true` AND `Parent MCC ID (Col D) === mccId`.
 */
export function getSelectedCidsForMcc(
  configSheet: SheetLike,
  mccId: string
): string[] {
  const targetMcc = normalizeCid(mccId);
  const lastRow = configSheet.getLastRow();
  if (lastRow < CID_TABLE_START_ROW) {
    return [];
  }
  const rows = configSheet
    .getRange(CID_TABLE_START_ROW, 1, lastRow - CID_TABLE_START_ROW + 1, CONFIG_TABLE_WIDTH)
    .getValues();

  const selected: string[] = [];
  for (const row of rows) {
    const include = parseSheetBoolean(row[CID_COL.INCLUDE]);
    const cid = normalizeCid(row[CID_COL.CID]);
    const parentMcc = normalizeCid(row[CID_COL.PARENT_MCC_ID]);

    if (!include || !cid || parentMcc !== targetMcc) {
      continue;
    }
    selected.push(cid);
  }
  return selected;
}

/**
 * Summarizes each registered MCC in the Discovered MCCs Table, computing `discoveredCount`,
 * `selectedCount`, and `isActive` (`selectedCount > 0`).
 * An MCC with `selectedCount === 0` is marked Inactive and skipped by the Multi-MCC barrier.
 */
export function getMccSelectionSummary(configSheet: SheetLike): MccSelectionSummary[] {
  const mccRows = configSheet
    .getRange(MCC_TABLE_START_ROW, 1, MAX_SUPPORTED_MCCS, CONFIG_TABLE_WIDTH)
    .getValues();

  const lastRow = configSheet.getLastRow();
  const cidRows =
    lastRow >= CID_TABLE_START_ROW
      ? configSheet
          .getRange(CID_TABLE_START_ROW, 1, lastRow - CID_TABLE_START_ROW + 1, CONFIG_TABLE_WIDTH)
          .getValues()
      : [];

  const summaries: MccSelectionSummary[] = [];
  for (const mccRow of mccRows) {
    const mccId = normalizeCid(mccRow[MCC_COL.MCC_ID]);
    if (!mccId) continue;

    const selectAllChecked = parseSheetBoolean(mccRow[MCC_COL.SELECT_ALL]);
    const mccName = readTextCell(mccRow[MCC_COL.NAME]).trim();

    let discoveredCount = 0;
    let selectedCount = 0;
    for (const cidRow of cidRows) {
      const cid = normalizeCid(cidRow[CID_COL.CID]);
      const parentMcc = normalizeCid(cidRow[CID_COL.PARENT_MCC_ID]);
      if (!cid || parentMcc !== mccId) continue;

      discoveredCount++;
      const included = parseSheetBoolean(cidRow[CID_COL.INCLUDE]);
      if (included) {
        selectedCount++;
      }
    }

    summaries.push({
      mccId,
      mccName,
      selectAllChecked,
      discoveredCount,
      selectedCount,
      isActive: selectedCount > 0,
    });
  }

  return summaries;
}
