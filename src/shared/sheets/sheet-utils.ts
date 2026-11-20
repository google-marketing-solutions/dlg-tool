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
 * @fileoverview Shared Google Sheets interfaces (`SheetLike`, `SpreadsheetLike`), low-level grid/tab
 * manipulation helpers (`getOrCreateSheet`, `trimSheetGrid`, `writeMatrixInChunks`, formula escaping),
 * and read-only `_RunState` lookups (`readRunStateMap`, `getRunStateValue`, `readLockedDailyConfig`).
 */

import {parseLockedDailyConfig} from '../contract';
import type {GlobalConfig} from '../types';

/** Name of the hidden operational state sheet (`_RunState`). */
export const RUN_STATE_SHEET_NAME = '_RunState';
/** Maximum number of rows written per `setValues` / `setNumberFormats` chunk. */
export const SHEET_WRITE_CHUNK_ROWS = 1000;

/** Structural subset of `GoogleAppsScript.Spreadsheet.Sheet` used by DLG v2 sheet readers and writers. */
export interface SheetLike {
  getName(): string;
  getSheetId(): number;
  hideSheet(): unknown;
  getLastRow(): number;
  getMaxRows(): number;
  getMaxColumns(): number;
  insertRowsAfter(afterPosition: number, howMany: number): unknown;
  insertColumnsAfter(afterPosition: number, howMany: number): unknown;
  deleteRows(rowPosition: number, howMany: number): unknown;
  deleteColumns(colPosition: number, howMany: number): unknown;
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
    setDataValidation(rule: GoogleAppsScript.Spreadsheet.DataValidation | null): unknown;
    getDataValidation(): Pick<GoogleAppsScript.Spreadsheet.DataValidation, 'getCriteriaValues'> | null;
    insertCheckboxes(): unknown;
    setNumberFormat(format: string): unknown;
    setNumberFormats(formats: string[][]): unknown;
    setFormula(formula: string): unknown;
  };
}

/** Structural interface for a Google Sheets `Spreadsheet` workbook. */
export interface SpreadsheetLike {
  getId?(): string;
  getUrl?(): string;
  getSheetByName(name: string): SheetLike | null;
  insertSheet(name: string, index?: number): SheetLike;
  getSheets(): SheetLike[];
  getSpreadsheetTimeZone(): string | null;
}

/**
 * Retrieves or creates a sheet by name.
 */
export function getOrCreateSheet<T extends { hideSheet(): unknown } = SheetLike>(
  spreadsheet: {
    getSheetByName(name: string): T | null;
    insertSheet(name: string, index?: number): T;
  },
  name: string,
  options: { index?: number; hidden?: boolean } = {}
): T {
  let sheet = spreadsheet.getSheetByName(name);
  if (sheet) {
    return sheet;
  }
  sheet =
    options.index !== undefined
      ? spreadsheet.insertSheet(name, options.index)
      : spreadsheet.insertSheet(name);
  if (options.hidden) {
    sheet.hideSheet();
  }
  return sheet;
}

/**
 * Reads all key-value pairs from `_RunState` in a single batch `getValues()` call (purely read-only;
 * returns an empty map when `_RunState` does not exist yet).
 */
export function readRunStateMap(spreadsheet: SpreadsheetLike): Map<string, string> {
  const map = new Map<string, string>();
  const sheet = spreadsheet.getSheetByName(RUN_STATE_SHEET_NAME);
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const row of rows) {
    const k = String(row[0] ?? '').trim();
    if (k) {
      map.set(k, String(row[1] ?? ''));
    }
  }
  return map;
}

/**
 * Looks up a single key in `_RunState`, optionally reusing a pre-read `stateMap`.
 */
export function getRunStateValue(
  spreadsheet: SpreadsheetLike,
  key: string,
  stateMap?: ReadonlyMap<string, string>
): string | null {
  const map = stateMap ?? readRunStateMap(spreadsheet);
  return map.get(key) ?? null;
}

/**
 * Reads today's frozen `GlobalConfig` from `DAILY_CONFIG_LOCK_<dateStr>` in `_RunState`.
 * Returns `null` when the Sheet App Script has not yet written today's lock, or when the lock is corrupted or
 * written with an unsupported `schemaVersion` (see `contract.ts`).
 */
export function readLockedDailyConfig(
  spreadsheet: SpreadsheetLike,
  dateStr: string,
  stateMap?: ReadonlyMap<string, string>
): GlobalConfig | null {
  const lockedJson = getRunStateValue(spreadsheet, `DAILY_CONFIG_LOCK_${dateStr}`, stateMap);
  return lockedJson ? parseLockedDailyConfig(lockedJson) : null;
}

/** Leading characters that make a spreadsheet treat a value as a formula (OWASP CSV-injection set). */
const FORMULA_PREFIX_RE = /^[=+\-@\t\r]/;
/** A value escaped by `escapeFormulaString`. */
const ESCAPED_FORMULA_PREFIX_RE = /^'[=+\-@\t\r]/;

/**
 * Prefixes strings starting with `=`, `+`, `-`, `@`, a tab, or a carriage return with a leading
 * apostrophe (`'`) so Google Sheets stores them as literal text rather than evaluating them as formulas.
 */
export function escapeFormulaString(val: string): string {
  return FORMULA_PREFIX_RE.test(val) ? `'${val}` : val;
}

/**
 * Removes the leading apostrophe added by `escapeFormulaString` when reading values back.
 */
export function unescapeFormulaString(val: string): string {
  return ESCAPED_FORMULA_PREFIX_RE.test(val) ? val.slice(1) : val;
}

/** Sheets returns checkboxes as booleans, but pasted or legacy cells can hold `"TRUE"` / `"true"`. */
export function parseSheetBoolean(value: unknown): boolean {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

/**
 * Builds a `setValues` matrix from `items`, one row per item and one cell per column: `null` / `undefined`
 * become `''`, strings are formula-escaped (`escapeFormulaString`), and other values are written as-is.
 */
export function toSheetValueMatrix<T extends object>(
  columns: readonly string[],
  items: readonly T[]
): unknown[][] {
  return items.map((item) => {
    const rec = item as Record<string, unknown>;
    return columns.map((k) => {
      const v = rec[k];
      if (v === null || v === undefined) return '';
      return typeof v === 'string' ? escapeFormulaString(v) : v;
    });
  });
}

/**
 * Writes a 2D cell matrix to `sheet` in batches of `chunkRows` rows per `setValues` call.
 */
export function writeMatrixInChunks(
  sheet: SheetLike,
  startRow: number,
  startCol: number,
  matrix: ReadonlyArray<readonly unknown[]>,
  chunkRows: number
): void {
  if (matrix.length === 0) return;
  const numCols = matrix[0].length;
  for (let offset = 0; offset < matrix.length; offset += chunkRows) {
    const slice = matrix.slice(offset, offset + chunkRows).map((r) => [...r]);
    sheet.getRange(startRow + offset, startCol, slice.length, numCols).setValues(slice);
  }
}

/**
 * Deletes trailing unused rows and columns beyond `targetRows` x `targetCols`.
 */
export function trimSheetGrid(sheet: SheetLike, targetRows: number, targetCols: number): void {
  const updatedMaxRows = sheet.getMaxRows();
  if (updatedMaxRows > targetRows) {
    sheet.deleteRows(targetRows + 1, updatedMaxRows - targetRows);
  }
  const updatedMaxCols = sheet.getMaxColumns();
  if (updatedMaxCols > targetCols) {
    sheet.deleteColumns(targetCols + 1, updatedMaxCols - targetCols);
  }
}
