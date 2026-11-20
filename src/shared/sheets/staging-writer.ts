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
 * @fileoverview Sheet reader, chunked writer, and lifecycle manager for normalized tabular staging tabs.
 * Writes exclusively to per-MCC isolated staging tabs (`_Stage_Accounts_<mccId>`, `_Stage_Budgets_<mccId>`,
 * `_Stage_Campaigns_<mccId>`, `_Stage_Recs_<mccId>`) so concurrent `Code.gs` MCC executions never contend
 * on shared sheets.
 * Formats cells as `@` before writing, escapes formula-prefix strings (`=`, `+`, `-`, `@`), overwrites in place
 * without `sheet.clear()`, and coerces values back to their schema types on read.
 */

import type {
  StagedAccount,
  StagedBudget,
  StagedCampaign,
  StagedRecommendation,
  StagingBatch,
} from '../types';
import {normalizeCid} from '../utils/cid';
import {
  SHEET_WRITE_CHUNK_ROWS,
  type SheetLike,
  type SpreadsheetLike,
  getOrCreateSheet,
  parseSheetBoolean,
  toSheetValueMatrix,
  trimSheetGrid,
  unescapeFormulaString,
  writeMatrixInChunks,
} from './sheet-utils';

/** Base name for the normalized staged accounts sheet. */
export const STAGE_ACCOUNTS_SHEET = '_Stage_Accounts';
/** Base name for the normalized staged budgets sheet. */
export const STAGE_BUDGETS_SHEET = '_Stage_Budgets';
/** Base name for the normalized staged campaigns sheet. */
export const STAGE_CAMPAIGNS_SHEET = '_Stage_Campaigns';
/** Base name for the normalized staged recommendations sheet. */
export const STAGE_RECS_SHEET = '_Stage_Recs';

/**
 * Returns the 4 hidden per-MCC staging sheet names (`_Stage_*_<10DigitMccId>`) for `mccId`.
 */
export function getMccStagingSheetNames(mccId: string): {
  accounts: string;
  budgets: string;
  campaigns: string;
  recommendations: string;
} {
  const norm = normalizeCid(mccId);
  return {
    accounts: `${STAGE_ACCOUNTS_SHEET}_${norm}`,
    budgets: `${STAGE_BUDGETS_SHEET}_${norm}`,
    campaigns: `${STAGE_CAMPAIGNS_SHEET}_${norm}`,
    recommendations: `${STAGE_RECS_SHEET}_${norm}`,
  };
}

type FieldType = 'STRING' | 'NULLABLE_STRING' | 'NUMBER' | 'BOOLEAN';

const STAGED_ACCOUNT_SCHEMA: Record<keyof StagedAccount, FieldType> = {
  mccId: 'STRING',
  mccName: 'STRING',
  accountId: 'STRING',
  accountName: 'STRING',
  currencyCode: 'STRING',
  timestamp: 'STRING',
};

const STAGED_BUDGET_SCHEMA: Record<keyof StagedBudget, FieldType> = {
  mccId: 'STRING',
  accountId: 'STRING',
  budgetResourceName: 'STRING',
  campaignBudgetName: 'STRING',
  campaignBudgetAmount: 'NUMBER',
  campaignBudgetTotalAmount: 'NUMBER',
  campaignBudgetDeliveryMethod: 'STRING',
  campaignBudgetType: 'STRING',
  campaignBudgetIsShared: 'BOOLEAN',
};

const STAGED_CAMPAIGN_SCHEMA: Record<keyof StagedCampaign, FieldType> = {
  mccId: 'STRING',
  accountId: 'STRING',
  campaignResourceName: 'STRING',
  campaignId: 'STRING',
  campaignName: 'STRING',
  campaignType: 'STRING',
  campaignSubType: 'STRING',
  campaignIsAiMax: 'BOOLEAN',
  campaignAiMaxTextCustomizationEnabled: 'BOOLEAN',
  campaignAiMaxFinalUrlExpansionEnabled: 'BOOLEAN',
  budgetResourceName: 'STRING',
  campaignIsPortfolioBiddingStrategy: 'BOOLEAN',
  campaignPortfolioBiddingStrategyName: 'STRING',
  portfolioStrategyResourceName: 'STRING',
  portfolioOwnerId: 'STRING',
  isLocallyOwnedPortfolio: 'BOOLEAN',
  campaignPortfolioIsManagerOwned: 'BOOLEAN',
  campaignBiddingStrategyType: 'STRING',
  campaignBiddingTargetRoas: 'NUMBER',
  campaignBiddingTargetCpa: 'NUMBER',
  campaignStatsYesterdayCost: 'NUMBER',
  campaignStats7DaysCost: 'NUMBER',
  campaign30DaysCost: 'NUMBER',
  campaign30DaysConversionsValue: 'NUMBER',
  campaign30DaysConversions: 'NUMBER',
  campaign30DaysSearchRankLostImpressionShare: 'NUMBER',
  campaign30DaysSearchBudgetLostImpressionShare: 'NUMBER',
};

const STAGED_REC_SCHEMA: Record<keyof StagedRecommendation, FieldType> = {
  mccId: 'STRING',
  accountId: 'STRING',
  recommendationId: 'STRING',
  recommendationType: 'STRING',
  campaignResourceName: 'STRING',
  recommendationCurrentBudgetAmount: 'NUMBER',
  recommendationNewBudgetAmount: 'NUMBER',
  recommendationNewTargetCpa: 'NUMBER',
  recommendationNewTargetRoas: 'NUMBER',
  recommendationBaseCost: 'NUMBER',
  recommendationPotentialCost: 'NUMBER',
  recommendationBaseClicks: 'NUMBER',
  recommendationPotentialClicks: 'NUMBER',
  recommendationBaseConversions: 'NUMBER',
  recommendationPotentialConversions: 'NUMBER',
  recommendationBaseConversionsValue: 'NUMBER',
  recommendationPotentialConversionsValue: 'NUMBER',
  recommendationBaseImpressions: 'NUMBER',
  recommendationPotentialImpressions: 'NUMBER',
  recommendationBaseVideoViews: 'NUMBER',
  recommendationPotentialVideoViews: 'NUMBER',
  moveBudgetSourceBudgetResourceName: 'STRING',
  moveBudgetSourceCampaigns: 'STRING',
  moveBudgetSourceTargetCpa: 'NUMBER',
  moveBudgetSourceTargetRoas: 'NUMBER',
};

/** Ordered `_Stage_Accounts*` column names. */
export const STAGED_ACCOUNT_KEYS = Object.keys(STAGED_ACCOUNT_SCHEMA) as Array<keyof StagedAccount>;
/** Ordered `_Stage_Budgets*` column names. */
export const STAGED_BUDGET_KEYS = Object.keys(STAGED_BUDGET_SCHEMA) as Array<keyof StagedBudget>;
/** Ordered `_Stage_Campaigns*` column names. */
export const STAGED_CAMPAIGN_KEYS = Object.keys(STAGED_CAMPAIGN_SCHEMA) as Array<keyof StagedCampaign>;
/** Ordered `_Stage_Recs*` column names. */
export const STAGED_REC_KEYS = Object.keys(STAGED_REC_SCHEMA) as Array<keyof StagedRecommendation>;

/**
 * Flushes an in-memory `StagingBatch` into the MCC's 4 normalized hidden staging sheets
 * (`_Stage_Accounts_<normMcc>`, `_Stage_Budgets_<normMcc>`, `_Stage_Campaigns_<normMcc>`, `_Stage_Recs_<normMcc>`).
 */
export function appendStagingBatch(
  workbook: SpreadsheetLike,
  batch: StagingBatch,
  mccId: string
): void {
  const names = getMccStagingSheetNames(mccId);
  upsertStagingTable(
    getOrCreateHiddenSheet(workbook, names.accounts),
    STAGED_ACCOUNT_KEYS,
    STAGED_ACCOUNT_SCHEMA,
    batch.accounts,
    (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}`
  );
  upsertStagingTable(
    getOrCreateHiddenSheet(workbook, names.budgets),
    STAGED_BUDGET_KEYS,
    STAGED_BUDGET_SCHEMA,
    batch.budgets,
    (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.budgetResourceName}`
  );
  upsertStagingTable(
    getOrCreateHiddenSheet(workbook, names.campaigns),
    STAGED_CAMPAIGN_KEYS,
    STAGED_CAMPAIGN_SCHEMA,
    batch.campaigns,
    (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.campaignResourceName}`
  );
  upsertStagingTable(
    getOrCreateHiddenSheet(workbook, names.recommendations),
    STAGED_REC_KEYS,
    STAGED_REC_SCHEMA,
    batch.recommendations,
    (r) =>
      `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.recommendationId}::${r.campaignResourceName}`
  );
}

/**
 * Reads all staged entities for `mccId` from its per-MCC hidden tabs (`_Stage_*_<normMcc>`).
 */
export function readStagingBatchForMcc(
  workbook: SpreadsheetLike,
  mccId: string
): StagingBatch {
  const normMcc = normalizeCid(mccId);
  const perMccNames = getMccStagingSheetNames(normMcc);

  const accounts = readTableRows<StagedAccount>(
    workbook.getSheetByName(perMccNames.accounts),
    STAGED_ACCOUNT_KEYS,
    STAGED_ACCOUNT_SCHEMA
  ).filter((r) => normalizeCid(r.mccId) === normMcc);

  const budgets = readTableRows<StagedBudget>(
    workbook.getSheetByName(perMccNames.budgets),
    STAGED_BUDGET_KEYS,
    STAGED_BUDGET_SCHEMA
  ).filter((r) => normalizeCid(r.mccId) === normMcc);

  const campaigns = readTableRows<StagedCampaign>(
    workbook.getSheetByName(perMccNames.campaigns),
    STAGED_CAMPAIGN_KEYS,
    STAGED_CAMPAIGN_SCHEMA
  ).filter((r) => normalizeCid(r.mccId) === normMcc);

  const recommendations = readTableRows<StagedRecommendation>(
    workbook.getSheetByName(perMccNames.recommendations),
    STAGED_REC_KEYS,
    STAGED_REC_SCHEMA
  ).filter((r) => normalizeCid(r.mccId) === normMcc);

  return { accounts, budgets, campaigns, recommendations };
}

/** Returns an empty `StagingBatch` (no accounts, budgets, campaigns, or recommendations). */
export function createEmptyStagingBatch(): StagingBatch {
  return { accounts: [], budgets: [], campaigns: [], recommendations: [] };
}

/**
 * Clears a single MCC's staged entities (`_Stage_*_<normMcc>`) when starting a fresh daily run for that MCC.
 */
export function clearMccFromStaging(
  workbook: SpreadsheetLike,
  mccId: string
): void {
  const normMcc = normalizeCid(mccId);
  const perMccNames = getMccStagingSheetNames(normMcc);
  const perMccSpecs: Array<[string, readonly string[], Record<string, FieldType>]> = [
    [perMccNames.accounts, STAGED_ACCOUNT_KEYS, STAGED_ACCOUNT_SCHEMA],
    [perMccNames.budgets, STAGED_BUDGET_KEYS, STAGED_BUDGET_SCHEMA],
    [perMccNames.campaigns, STAGED_CAMPAIGN_KEYS, STAGED_CAMPAIGN_SCHEMA],
    [perMccNames.recommendations, STAGED_REC_KEYS, STAGED_REC_SCHEMA],
  ];

  for (const [name, keys, schema] of perMccSpecs) {
    const sheet = workbook.getSheetByName(name);
    if (sheet) {
      writeTableRows(sheet, keys, schema, []);
    }
  }
}

function getOrCreateHiddenSheet(workbook: SpreadsheetLike, name: string): SheetLike {
  return getOrCreateSheet(workbook, name, { hidden: true });
}

/**
 * Merges `newItems` into a staging table by key. Fast path (the normal hourly flush): when the header row
 * already matches `keys` exactly and no new key collides with an existing or another new row, only the new rows
 * are formatted and appended after the last row. Otherwise (first write, header change, or a key update after a
 * mid-flush crash) the whole merged table is rewritten.
 */
function upsertStagingTable<T extends object>(
  sheet: SheetLike,
  keys: ReadonlyArray<keyof T>,
  schema: Record<string, FieldType>,
  newItems: readonly T[],
  keyFn: (item: T) => string
): void {
  const lastRow = sheet.getLastRow();
  if (newItems.length === 0 && lastRow > 0) return;

  const existing = readTableRows<T>(sheet, keys, schema);
  const existingKeys = new Set(existing.map(keyFn));
  const newKeys = new Set<string>();
  let collides = false;
  for (const item of newItems) {
    const k = keyFn(item);
    if (existingKeys.has(k) || newKeys.has(k)) {
      collides = true;
      break;
    }
    newKeys.add(k);
  }

  if (lastRow > 0 && !collides && headerMatchesKeys(sheet, keys as readonly string[])) {
    appendTableRows(sheet, keys as readonly string[], schema, newItems, lastRow + 1);
    return;
  }

  const mergedMap = new Map<string, T>();
  for (const item of existing) {
    mergedMap.set(keyFn(item), item);
  }
  for (const item of newItems) {
    mergedMap.set(keyFn(item), item);
  }

  writeTableRows(sheet, keys as readonly string[], schema, Array.from(mergedMap.values()));
}

/** True when Row 1 holds exactly `keys` (in order) with no extra header cells. */
function headerMatchesKeys(sheet: SheetLike, keys: readonly string[]): boolean {
  const maxCols = sheet.getMaxColumns();
  if (maxCols < keys.length) return false;
  const header = sheet.getRange(1, 1, 1, maxCols).getValues()[0] ?? [];
  for (let c = 0; c < maxCols; c++) {
    const cell = String(header[c] ?? '').trim();
    if (cell !== (c < keys.length ? keys[c] : '')) return false;
  }
  return true;
}

/** Number format applied to a staging column of `fieldType` before values are written. */
function formatForFieldType(fieldType: FieldType | undefined): string {
  switch (fieldType) {
    case 'NUMBER':
      return '0.########';
    case 'BOOLEAN':
      return 'General';
    default:
      // STRING and NULLABLE_STRING are stored as text so Sheets never strips leading zeros or parses dates
      return '@';
  }
}

/** Formats and appends `items` starting at `startRow`, growing the grid as needed. */
function appendTableRows<T extends object>(
  sheet: SheetLike,
  keys: readonly string[],
  schema: Record<string, FieldType>,
  items: readonly T[],
  startRow: number
): void {
  if (items.length === 0) return;
  const numCols = keys.length;
  const lastNeededRow = startRow + items.length - 1;
  if (sheet.getMaxRows() < lastNeededRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), lastNeededRow - sheet.getMaxRows());
  }
  const rowFormats = keys.map((k) => formatForFieldType(schema[k]));
  for (let offset = 0; offset < items.length; offset += SHEET_WRITE_CHUNK_ROWS) {
    const count = Math.min(SHEET_WRITE_CHUNK_ROWS, items.length - offset);
    const fmtMatrix = Array.from({ length: count }, () => [...rowFormats]);
    sheet.getRange(startRow + offset, 1, count, numCols).setNumberFormats(fmtMatrix);
  }
  writeMatrixInChunks(sheet, startRow, 1, toSheetValueMatrix(keys, items), SHEET_WRITE_CHUNK_ROWS);
}

function writeTableRows<T extends object>(
  sheet: SheetLike,
  keys: readonly string[],
  schema: Record<string, FieldType>,
  items: readonly T[]
): void {
  const numCols = keys.length;
  const totalRows = items.length + 1;

  if (sheet.getMaxColumns() < numCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < totalRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  }

  // Format columns BEFORE setValues (text columns as '@' so Sheets never strips leading zeros or parses dates)
  const rowFormats = keys.map((k) => formatForFieldType(schema[k]));
  for (let offset = 0; offset < totalRows; offset += SHEET_WRITE_CHUNK_ROWS) {
    const count = Math.min(SHEET_WRITE_CHUNK_ROWS, totalRows - offset);
    const fmtMatrix = Array.from({ length: count }, () =>
      offset === 0 ? keys.map(() => '@') : [...rowFormats]
    );
    if (offset === 0 && count > 1) {
      for (let r = 1; r < count; r++) {
        fmtMatrix[r] = [...rowFormats];
      }
    }
    sheet.getRange(1 + offset, 1, count, numCols).setNumberFormats(fmtMatrix);
  }

  const matrix: unknown[][] = [[...keys], ...toSheetValueMatrix(keys, items)];

  writeMatrixInChunks(sheet, 1, 1, matrix, SHEET_WRITE_CHUNK_ROWS);
  trimSheetGrid(sheet, totalRows, numCols);
}

/**
 * Reads and coerces rows from a `_Stage_*` tab.
 *
 * Reads every physical column (`getMaxColumns()`) and matches `keys` to columns by Row 1 header name
 * only. A key whose header is absent gets its `FieldType` default (`''` for `STRING`, `null` for
 * `NUMBER`/`NULLABLE_STRING`, `false` for `BOOLEAN`); values are never taken from a column at the same
 * position. If a header name appears more than once, the first (left-most) occurrence
 * wins. Rows whose matched cells are all empty are skipped.
 */
function readTableRows<T>(
  sheet: SheetLike | null,
  keys: ReadonlyArray<keyof T>,
  schema: Record<string, FieldType>
): T[] {
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

  const out: T[] = [];
  for (let r = 1; r < allValues.length; r++) {
    const rawRow = allValues[r];
    const obj: Record<string, unknown> = {};
    let hasNonEmptyCell = false;

    for (const key of keys) {
      const k = String(key);
      const colIdx = colIndexByKey.get(k);
      const cell = colIdx === undefined ? undefined : rawRow[colIdx];
      if (cell !== '' && cell !== null && cell !== undefined) {
        hasNonEmptyCell = true;
      }
      obj[k] = coerceStagingCell(cell, schema[k] ?? 'STRING');
    }

    if (hasNonEmptyCell) {
      out.push(obj as T);
    }
  }

  return out;
}

function coerceStagingCell(cell: unknown, fieldType: FieldType): unknown {
  if (fieldType === 'BOOLEAN') {
    return parseSheetBoolean(cell);
  }
  if (fieldType === 'NUMBER') {
    if (cell === '' || cell === null || cell === undefined) return null;
    const num = Number(cell);
    return Number.isFinite(num) ? num : null;
  }
  if (fieldType === 'NULLABLE_STRING') {
    if (cell === '' || cell === null || cell === undefined) return null;
    if (cell instanceof Date) {
      return cell.toISOString().replace('T', ' ').slice(0, 19);
    }
    return unescapeFormulaString(String(cell));
  }
  if (cell === null || cell === undefined) return '';
  if (cell instanceof Date) {
    return cell.toISOString().replace('T', ' ').slice(0, 19);
  }
  return unescapeFormulaString(String(cell));
}
