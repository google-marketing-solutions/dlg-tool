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
 * @fileoverview Sheet writer and reader for the 82-column `Data` tab (`writeReportRowsToSheet`).
 * Expands sheet grids before writing to prevent out-of-bounds errors, always writes the schema dummy row
 * as Row 2 so Data Studio infers every column's type, applies cell number formats BEFORE
 * writing values (preventing leading-zero/Date coercion), overwrites rows in place without `sheet.clear()`,
 * and trims trailing unused rows and columns.
 * Bundled only into `dist/sheet/Code.gs`.
 */

import {
  REPORT_SCHEMA_TEMPLATE,
  SCHEMA_COLUMN_NAMES,
  createDummyFormattingRow,
  isDummyFormattingRow,
} from '../../shared/schema';
import {
  SHEET_WRITE_CHUNK_ROWS,
  type SheetLike,
  type SpreadsheetLike,
  parseSheetBoolean,
  toSheetValueMatrix,
  trimSheetGrid,
  unescapeFormulaString,
  writeMatrixInChunks,
} from '../../shared/sheets/sheet-utils';
import type {ReportColumnKey, ReportRow} from '../../shared/types';
import {withSheetLock} from '../run-state';

/** Name of the published 82-column `Data` tab read by Data Studio. */
export const DATA_SHEET_NAME = 'Data';

/**
 * Writes an array of `ReportRow` objects to the `Data` sheet (runs in `dist/sheet/Code.gs`):
 * 1. Expands grid rows/columns before writing to prevent out-of-bounds errors.
 * 2. Always writes the schema dummy row as Row 2 (directly under the header), followed by the real rows.
 *    The Sheets connector infers field types from column values, so a column that is blank in every real
 *    row (e.g. no `RAISE_TARGET_CPA` rows today) would otherwise lose its numeric/date/boolean type.
 *    Any dummy row in `rows` is dropped first, so republishing rows never stacks extra dummy rows.
 * 3. Applies all cell number formats (`@`, `0.00`, `0.0%`, etc.) in 1,000-row chunks BEFORE `setValues`, so Sheets never
 *    turns IDs into numbers or dates.
 * 4. Overwrites rows in place without calling `sheet.clear()` (which would also wipe formats and validation)
 *    and trims trailing unused rows/columns
 *    (the sheet keeps exactly header + dummy + real rows).
 */
export function writeReportRowsToSheet(
  sheet: SheetLike,
  rows: readonly ReportRow[],
  workbook?: SpreadsheetLike
): void {
  const runWrite = () => {
    const effectiveRows = [
      createDummyFormattingRow(),
      ...rows.filter((r) => !isDummyFormattingRow(r)),
    ];
    const numCols = SCHEMA_COLUMN_NAMES.length;
    const totalSheetRows = effectiveRows.length + 1; // 1 header row + data rows

    const currentMaxCols = sheet.getMaxColumns();
    if (currentMaxCols < numCols) {
      sheet.insertColumnsAfter(currentMaxCols, numCols - currentMaxCols);
    }
    const currentMaxRows = sheet.getMaxRows();
    if (currentMaxRows < totalSheetRows) {
      sheet.insertRowsAfter(currentMaxRows, totalSheetRows - currentMaxRows);
    }

    const singleRowFormats = SCHEMA_COLUMN_NAMES.map(
      (colKey) => REPORT_SCHEMA_TEMPLATE[colKey].format
    );
    writeFormatsInChunks(
      sheet,
      2,
      1,
      effectiveRows.length,
      singleRowFormats,
      SHEET_WRITE_CHUNK_ROWS
    );

    const valuesMatrix: unknown[][] = [
      SCHEMA_COLUMN_NAMES,
      ...toSheetValueMatrix(SCHEMA_COLUMN_NAMES, effectiveRows),
    ];

    writeMatrixInChunks(sheet, 1, 1, valuesMatrix, SHEET_WRITE_CHUNK_ROWS);
    trimSheetGrid(sheet, totalSheetRows, numCols);
  };

  if (workbook) {
    withSheetLock(workbook, 'DATA_WRITE', runWrite);
  } else {
    runWrite();
  }
}

/**
 * Reads and coerces `ReportRow` objects from the `Data` sheet, automatically skipping the dummy schema-preservation row.
 *
 * Reads every physical column (`getMaxColumns()`) and matches schema keys to columns by Row 1 header
 * name only. A schema key whose header is absent gets `REPORT_SCHEMA_TEMPLATE[key].defaultValue`; values
 * are never taken from a column at the same position, so older layouts with retired or reordered
 * columns cannot leak neighboring values into the wrong field. If a header name appears
 * more than once, the first (left-most) occurrence wins.
 *
 * Date-formatted columns (`timestamp`, `mccLastUpdatedDate`) come back as `YYYY-MM-DD HH:mm:ss` /
 * `YYYY-MM-DD` strings, the same shape the fresh publish path writes. Google Sheets returns these cells
 * as `Date` objects, which are formatted in `timeZone` (pass the spreadsheet time zone, because Sheets
 * builds the `Date` from the cell's wall-clock time in that zone). Epoch-millisecond numbers left by
 * earlier stale publishes are converted back the same way, so corrupted rows self-heal.
 */
export function readReportRowsFromSheet(
  sheet: SheetLike | null,
  timeZone: string = resolveDefaultTimeZone()
): ReportRow[] {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const numCols = sheet.getMaxColumns();
  if (numCols < 1) return [];
  const allValues = sheet.getRange(1, 1, lastRow, numCols).getValues();
  const headerRow = allValues[0] ?? [];

  const colIndexByKey = new Map<string, number>();
  for (let c = 0; c < headerRow.length; c++) {
    const headerName = String(headerRow[c] ?? '').trim();
    if (headerName && !colIndexByKey.has(headerName)) {
      colIndexByKey.set(headerName, c);
    }
  }

  const out: ReportRow[] = [];
  for (let r = 1; r < allValues.length; r++) {
    const rawRow = allValues[r];
    const obj: Record<string, unknown> = {};

    for (const key of SCHEMA_COLUMN_NAMES) {
      const colIdx = colIndexByKey.get(key);
      const cellVal = colIdx === undefined ? undefined : rawRow[colIdx];
      obj[key] = coerceReportCell(key, cellVal, timeZone);
    }

    const typed = obj as unknown as ReportRow;
    if (isDummyFormattingRow(typed)) {
      continue;
    }
    out.push(typed);
  }

  return out;
}

function coerceReportCell(key: ReportColumnKey, cellVal: unknown, timeZone: string): unknown {
  const def = REPORT_SCHEMA_TEMPLATE[key];
  if (cellVal === '' || cellVal === null || cellVal === undefined) {
    return def.defaultValue;
  }
  if (typeof def.defaultValue === 'boolean') {
    return parseSheetBoolean(cellVal);
  }
  if (isDateFormat(def.format)) {
    return coerceDateCell(cellVal, def.format, def.defaultValue, timeZone);
  }
  if (def.format === '@') {
    if (cellVal instanceof Date) {
      return cellVal.toISOString().replace('T', ' ').slice(0, 19);
    }
    return unescapeFormulaString(String(cellVal));
  }
  const num = Number(cellVal);
  return Number.isFinite(num) ? num : def.defaultValue;
}

/**
 * Numbers above this are treated as epoch milliseconds (1e11 ms is March 1973). Real Sheets date
 * serials are far smaller (about 46,000 for 2026).
 */
const MIN_EPOCH_MS = 1e11;

/** Sheets date/time number formats in `REPORT_SCHEMA_TEMPLATE` all start with `yyyy`. */
function isDateFormat(format: string): boolean {
  return format.startsWith('yyyy');
}

/**
 * Coerces a date-formatted `Data` cell to the string shape written by the publisher:
 * - `Date` (what `getValues()` returns for parsed date cells) -> formatted string in `timeZone`.
 * - Finite number above `MIN_EPOCH_MS` -> treated as epoch ms (written by older stale publishes) and
 *   formatted the same way, so corrupted rows self-heal.
 * - Any other number -> `defaultValue`, never written back as a raw number.
 * - Strings pass through unchanged.
 */
function coerceDateCell(
  cellVal: unknown,
  format: string,
  defaultValue: unknown,
  timeZone: string
): unknown {
  const pattern = format.includes('hh') ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd';
  if (cellVal instanceof Date) {
    return formatDateInTimeZone(cellVal, timeZone, pattern) ?? defaultValue;
  }
  if (typeof cellVal === 'number') {
    if (Number.isFinite(cellVal) && cellVal > MIN_EPOCH_MS) {
      return formatDateInTimeZone(new Date(cellVal), timeZone, pattern) ?? defaultValue;
    }
    return defaultValue;
  }
  return String(cellVal);
}

/**
 * Formats `date` with `Utilities.formatDate` in `timeZone` when running in Apps Script, otherwise
 * falls back to UTC via `toISOString`. Returns `null` for invalid dates.
 */
function formatDateInTimeZone(
  date: Date,
  timeZone: string,
  pattern: 'yyyy-MM-dd HH:mm:ss' | 'yyyy-MM-dd'
): string | null {
  if (Number.isNaN(date.getTime())) return null;
  if (typeof Utilities !== 'undefined' && typeof Utilities.formatDate === 'function') {
    return Utilities.formatDate(date, timeZone || 'UTC', pattern);
  }
  const iso = date.toISOString().replace('T', ' ');
  return pattern === 'yyyy-MM-dd' ? iso.slice(0, 10) : iso.slice(0, 19);
}

/** Script time zone when running in Apps Script, otherwise `UTC`. */
function resolveDefaultTimeZone(): string {
  if (typeof Session !== 'undefined' && typeof Session.getScriptTimeZone === 'function') {
    return Session.getScriptTimeZone() || 'UTC';
  }
  return 'UTC';
}

function writeFormatsInChunks(
  sheet: SheetLike,
  startRow: number,
  startCol: number,
  totalDataRows: number,
  singleRowFormats: readonly string[],
  chunkRows: number
): void {
  const numCols = singleRowFormats.length;
  for (let offset = 0; offset < totalDataRows; offset += chunkRows) {
    const count = Math.min(chunkRows, totalDataRows - offset);
    const formatsSlice: string[][] = Array.from({ length: count }, () => [...singleRowFormats]);
    sheet
      .getRange(startRow + offset, startCol, count, numCols)
      .setNumberFormats(formatsSlice);
  }
}
