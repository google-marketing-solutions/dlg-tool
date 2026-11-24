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
  MAX_SHEET_CELL_CHARS,
  joinListForSheetCell,
  readReportRowsFromSheet,
  writeReportRowsToSheet,
} from '../src/appscript/sheets/data-writer';
import {
  REPORT_SCHEMA_TEMPLATE,
  SCHEMA_COLUMN_NAMES,
  SCHEMA_DUMMY_ACCOUNT_NAME,
  createDummyFormattingRow,
  createEmptyReportRow,
} from '../src/shared/schema';
import {
  escapeFormulaString,
  parseSheetBoolean,
  toSheetValueMatrix,
  unescapeFormulaString,
} from '../src/shared/sheets/sheet-utils';
import {
  STAGE_ACCOUNTS_SHEET,
  STAGE_BUDGETS_SHEET,
  STAGE_CAMPAIGNS_SHEET,
  STAGE_RECS_SHEET,
  appendStagingBatch,
  clearMccFromStaging,
  readStagingBatchForMcc,
} from '../src/shared/sheets/staging-writer';
import type {ReportColumnKey, ReportRow} from '../src/shared/types';
import {
  MockSheet,
  createMockSpreadsheetApp,
} from './mocks/google-apps-script.mock';

describe('Bulk Sheet Writer, Grid Expansion, Single-RPC Formatting & Dummy Row', () => {
  it('expands small grids before writing, formats in a single setNumberFormats call, and trims excess rows/cols', () => {
    // Start with a tiny 5x5 sheet to verify grid expansion and trimming
    const sheet = new MockSheet('Data', 100);
    sheet.deleteRows(6, 95);
    sheet.deleteColumns(6, 21);
    expect(sheet.getMaxRows()).toBe(5);
    expect(sheet.getMaxColumns()).toBe(5);

    const row1 = createEmptyReportRow();
    row1.accountId = '123-456-7890';
    row1.accountName = 'Acme UK';
    row1.campaignId = '111';
    row1.campaign30DaysCost = 250.5;

    const row2 = createEmptyReportRow();
    row2.accountId = '123-456-7890';
    row2.accountName = 'Acme UK';
    row2.campaignId = '222';
    row2.campaign30DaysCost = 500.75;

    writeReportRowsToSheet(sheet, [row1, row2]);

    // 1 header row + 1 schema dummy row + 2 data rows = 4 rows, 82 columns
    expect(sheet.getMaxRows()).toBe(4);
    expect(sheet.getMaxColumns()).toBe(SCHEMA_COLUMN_NAMES.length);

    // Single bulk setNumberFormats RPC call
    expect(sheet.setNumberFormatsCallCount).toBe(1);

    const written = sheet.getRange(1, 1, 3, 4).getValues();
    expect(written[0]).toEqual([
      'recommendationNumber',
      'recommendationScope',
      'mccId',
      'mccName',
    ]);
  });

  it('writes the schema dummy row when 0 recommendation rows exist so Data Studio binding never breaks', () => {
    const sheet = new MockSheet('Data', 100);
    writeReportRowsToSheet(sheet, []);

    expect(sheet.getMaxRows()).toBe(2);
    expect(sheet.getMaxColumns()).toBe(SCHEMA_COLUMN_NAMES.length);

    const accountNameIdx = SCHEMA_COLUMN_NAMES.indexOf('accountName') + 1;
    expect(sheet.getRange(2, accountNameIdx).getValue()).toBe(
      '!!! SCHEMA DEFINITION (FILTER ME OUT)'
    );
  });

  it('stages normalized entities (_Stage_Accounts_<mcc>, _Stage_Budgets_<mcc>, _Stage_Campaigns_<mcc>, _Stage_Recs_<mcc>) and shrinks them to 1 row when cleared', () => {
    const ss = createMockSpreadsheetApp().getActiveSpreadsheet();
    appendStagingBatch(
      ss,
      {
        accounts: [
          {
            mccId: '111-222-3333',
            mccName: 'EMEA MCC',
            accountId: '123-456-7890',
            accountName: 'Acme UK',
            currencyCode: 'GBP',
            timestamp: '2026-10-14 00:00:00',
          },
        ],
        budgets: [
          {
            mccId: '111-222-3333',
            accountId: '123-456-7890',
            budgetResourceName: 'customers/1234567890/campaignBudgets/1',
            campaignBudgetName: 'Shared Search Budget',
            campaignBudgetAmount: 100,
            campaignBudgetTotalAmount: null,
            campaignBudgetDeliveryMethod: 'STANDARD',
            campaignBudgetType: 'DAILY',
            campaignBudgetIsShared: true,
          },
        ],
        campaigns: [
          {
            mccId: '111-222-3333',
            accountId: '123-456-7890',
            campaignResourceName: 'customers/1234567890/campaigns/10',
            campaignId: '10',
            campaignName: 'Brand Search',
            campaignType: 'SEARCH',
            campaignSubType: 'N/A',
            campaignIsAiMax: true,
            campaignAiMaxTextCustomizationEnabled: true,
            campaignAiMaxFinalUrlExpansionEnabled: false,
            budgetResourceName: 'customers/1234567890/campaignBudgets/1',
            campaignIsPortfolioBiddingStrategy: false,
            campaignPortfolioBiddingStrategyName: 'N/A',
            portfolioStrategyResourceName: 'N/A',
            portfolioOwnerId: 'N/A',
            isLocallyOwnedPortfolio: false,
            campaignPortfolioIsManagerOwned: false,
            campaignBiddingStrategyType: 'MAXIMIZE_CONVERSIONS',
            campaignBiddingTargetRoas: null,
            campaignBiddingTargetCpa: 15,
            campaignStatsYesterdayCost: 95,
            campaignStats7DaysCost: 650,
            campaign30DaysCost: 2800,
            campaign30DaysConversionsValue: 8400,
            campaign30DaysConversions: 200,
            campaign30DaysSearchRankLostImpressionShare: 0.1,
            campaign30DaysSearchBudgetLostImpressionShare: 0.2,
          },
        ],
        recommendations: [
          {
            mccId: '111-222-3333',
            accountId: '123-456-7890',
            recommendationId: 'customers/1234567890/recommendations/99',
            recommendationType: 'CAMPAIGN_BUDGET',
            campaignResourceName: 'customers/1234567890/campaigns/10',
            recommendationCurrentBudgetAmount: 100,
            recommendationNewBudgetAmount: 150,
            recommendationNewTargetCpa: null,
            recommendationNewTargetRoas: null,
            recommendationBaseCost: 700,
            recommendationPotentialCost: 1050,
            recommendationBaseClicks: null,
            recommendationPotentialClicks: null,
            recommendationBaseConversions: 50,
            recommendationPotentialConversions: 75,
            recommendationBaseConversionsValue: 1500,
            recommendationPotentialConversionsValue: 2250,
            recommendationBaseImpressions: null,
            recommendationPotentialImpressions: null,
            recommendationBaseVideoViews: null,
            recommendationPotentialVideoViews: null,
            moveBudgetSourceBudgetResourceName: '',
            moveBudgetSourceCampaigns: 'N/A',
            moveBudgetSourceTargetCpa: null,
            moveBudgetSourceTargetRoas: null,
          },
        ],
      },
      '111-222-3333'
    );

    const readBack = readStagingBatchForMcc(ss, '1112223333');
    expect(readBack.accounts.length).toBe(1);
    expect(readBack.budgets.length).toBe(1);
    expect(readBack.campaigns.length).toBe(1);
    expect(readBack.recommendations.length).toBe(1);

    // Clear all 4 per-MCC staging sheets -> each shrinks to 1 header row
    clearMccFromStaging(ss, '111-222-3333');
    expect(ss.getSheetByName(`${STAGE_ACCOUNTS_SHEET}_1112223333`)?.getMaxRows()).toBe(1);
    expect(ss.getSheetByName(`${STAGE_BUDGETS_SHEET}_1112223333`)?.getMaxRows()).toBe(1);
    expect(ss.getSheetByName(`${STAGE_CAMPAIGNS_SHEET}_1112223333`)?.getMaxRows()).toBe(1);
    expect(ss.getSheetByName(`${STAGE_RECS_SHEET}_1112223333`)?.getMaxRows()).toBe(1);
  });

  it('chunks large writes (> 1000 rows) into 1000-row RPC slices to prevent SpreadsheetApp payload timeouts', () => {
    const sheet = new MockSheet('Data', 100);
    const rows = Array.from({ length: 1200 }, (_, i) => {
      const r = createEmptyReportRow();
      r.campaignId = String(i + 1);
      return r;
    });

    writeReportRowsToSheet(sheet, rows);

    expect(sheet.getMaxRows()).toBe(1202); // 1 header + 1 schema dummy row + 1200 data rows
    // 1201 data rows (dummy + 1200) split across 2 chunks (1000 + 201) -> 2 setNumberFormats calls
    expect(sheet.setNumberFormatsCallCount).toBe(2);
    expect(sheet.getRange(1202, SCHEMA_COLUMN_NAMES.indexOf('campaignId') + 1).getValue()).toBe(
      '1200'
    );
  });

  it('preserves leading-zero IDs, formula-prefix strings (+Brand, -Search, =Promo), and booleans/numbers across tabular staging and Data round-trips', () => {
    const ss = createMockSpreadsheetApp().getActiveSpreadsheet();
    appendStagingBatch(
      ss,
      {
        accounts: [
          {
            mccId: '001-234-5678',
            mccName: '+EMEA MCC',
            accountId: '009-876-5432',
            accountName: '=Acme Search',
            currencyCode: 'EUR',
            timestamp: '2026-10-14 00:00:00',
          },
        ],
        budgets: [],
        campaigns: [
          {
            mccId: '001-234-5678',
            accountId: '009-876-5432',
            campaignResourceName: 'customers/0098765432/campaigns/0077',
            campaignId: '0077',
            campaignName: '+Brand - Exact',
            campaignType: 'SEARCH',
            campaignSubType: 'N/A',
            campaignIsAiMax: true,
            campaignAiMaxTextCustomizationEnabled: false,
            campaignAiMaxFinalUrlExpansionEnabled: true,
            budgetResourceName: 'customers/0098765432/campaignBudgets/1',
            campaignIsPortfolioBiddingStrategy: false,
            campaignPortfolioBiddingStrategyName: 'N/A',
            portfolioStrategyResourceName: 'N/A',
            portfolioOwnerId: '0012345678',
            isLocallyOwnedPortfolio: false,
            campaignPortfolioIsManagerOwned: false,
            campaignBiddingStrategyType: 'MAXIMIZE_CONVERSIONS',
            campaignBiddingTargetRoas: null,
            campaignBiddingTargetCpa: 12.5,
            campaignStatsYesterdayCost: 10,
            campaignStats7DaysCost: 70,
            campaign30DaysCost: 300,
            campaign30DaysConversionsValue: 900,
            campaign30DaysConversions: 30,
            campaign30DaysSearchRankLostImpressionShare: 0.15,
            campaign30DaysSearchBudgetLostImpressionShare: null,
          },
        ],
        recommendations: [],
      },
      '001-234-5678'
    );

    const readBack = readStagingBatchForMcc(ss, '001-234-5678');
    expect(readBack.accounts[0].mccId).toBe('001-234-5678');
    expect(readBack.accounts[0].mccName).toBe('+EMEA MCC');
    expect(readBack.accounts[0].accountName).toBe('=Acme Search');
    expect(readBack.campaigns[0].campaignId).toBe('0077');
    expect(readBack.campaigns[0].campaignName).toBe('+Brand - Exact');
    expect(readBack.campaigns[0].portfolioOwnerId).toBe('0012345678');
    expect(readBack.campaigns[0].campaignIsAiMax).toBe(true);
    expect(readBack.campaigns[0].campaignAiMaxTextCustomizationEnabled).toBe(false);
    expect(readBack.campaigns[0].campaignBiddingTargetCpa).toBe(12.5);
    expect(readBack.campaigns[0].campaign30DaysSearchBudgetLostImpressionShare).toBeNull();

    const dataSheet = new MockSheet('Data', 100);
    const r = createEmptyReportRow();
    r.mccId = '001-234-5678';
    r.accountId = '009-876-5432';
    r.campaignId = '0077';
    r.campaignName = '+Brand - Exact';
    r.campaignIsAiMax = true;
    r.campaign30DaysCost = 123.45;
    writeReportRowsToSheet(dataSheet, [r]);

    const dataRows = readReportRowsFromSheet(dataSheet);
    expect(dataRows.length).toBe(1);
    expect(dataRows[0].mccId).toBe('001-234-5678');
    expect(dataRows[0].campaignId).toBe('0077');
    expect(dataRows[0].campaignName).toBe('+Brand - Exact');
    expect(dataRows[0].campaignIsAiMax).toBe(true);
    expect(dataRows[0].campaign30DaysCost).toBe(123.45);
  });

  it('caps plain and formula-escaped cells at Google Sheets 50,000-char limit while keeping exact-limit cells intact', () => {
    const sheet = new MockSheet('Data', 10);
    const exactLimitRow = createEmptyReportRow();
    exactLimitRow.portfolioMemberCampaignNames = 'A'.repeat(MAX_SHEET_CELL_CHARS);
    const oversizedRow = createEmptyReportRow();
    oversizedRow.portfolioMemberCampaignNames = 'C'.repeat(MAX_SHEET_CELL_CHARS + 500);
    const formulaAtLimitRow = createEmptyReportRow();
    formulaAtLimitRow.portfolioMemberCampaignNames = `=${'x'.repeat(MAX_SHEET_CELL_CHARS - 1)}`;

    writeReportRowsToSheet(sheet, [exactLimitRow, oversizedRow, formulaAtLimitRow]);

    const colIdx = SCHEMA_COLUMN_NAMES.indexOf('portfolioMemberCampaignNames') + 1;
    expect(sheet.getRange(3, colIdx).getValue()).toBe('A'.repeat(MAX_SHEET_CELL_CHARS));
    expect(sheet.getRange(4, colIdx).getValue()).toBe(`${'C'.repeat(MAX_SHEET_CELL_CHARS - 1)}…`);
    expect(sheet.getRange(5, colIdx).getValue()).toBe(`'=${'x'.repeat(MAX_SHEET_CELL_CHARS - 3)}…`);
  });
});

describe('schema dummy row is always written as Data Row 2', () => {
  const ACCOUNT_NAME_COL = SCHEMA_COLUMN_NAMES.indexOf('accountName') + 1;

  function realRow(campaignId: string, recommendationType = 'CAMPAIGN_BUDGET'): ReportRow {
    const r = createEmptyReportRow();
    r.mccId = '111-222-3333';
    r.accountId = '123-456-7890';
    r.accountName = 'Acme UK';
    r.campaignId = campaignId;
    r.recommendationType = recommendationType;
    r.recommendationCurrentBudgetAmount = 100;
    r.recommendationNewBudgetAmount = 150;
    return r;
  }

  function accountNames(sheet: MockSheet): unknown[] {
    return sheet
      .getRange(1, ACCOUNT_NAME_COL, sheet.getLastRow(), 1)
      .getValues()
      .map((r) => r[0]);
  }

  it('writes header + dummy + N rows, trims the grid to fit, and reads back exactly N rows', () => {
    const sheet = new MockSheet('Data', 100);
    const rows = [realRow('1'), realRow('2'), realRow('3')];
    writeReportRowsToSheet(sheet, rows);

    expect(sheet.getMaxRows()).toBe(5); // header + dummy + 3
    expect(sheet.getLastRow()).toBe(5);
    expect(sheet.getMaxColumns()).toBe(SCHEMA_COLUMN_NAMES.length);
    expect(accountNames(sheet)).toEqual([
      'accountName',
      SCHEMA_DUMMY_ACCOUNT_NAME,
      'Acme UK',
      'Acme UK',
      'Acme UK',
    ]);

    const readBack = readReportRowsFromSheet(sheet);
    expect(readBack.map((r) => r.campaignId)).toEqual(['1', '2', '3']);

    // Shrinking to 1 row trims the grid to header + dummy + 1 row
    writeReportRowsToSheet(sheet, [realRow('9')]);
    expect(sheet.getMaxRows()).toBe(3);
    expect(readReportRowsFromSheet(sheet).map((r) => r.campaignId)).toEqual(['9']);
  });

  it('never stacks dummy rows when rows read from Data (or a dummy row) are written back', () => {
    const sheet = new MockSheet('Data', 100);
    writeReportRowsToSheet(sheet, [realRow('1'), realRow('2')]);
    for (let i = 0; i < 3; i++) {
      const reRead = readReportRowsFromSheet(sheet);
      writeReportRowsToSheet(sheet, [...reRead, createDummyFormattingRow()]);
    }

    expect(sheet.getMaxRows()).toBe(4);
    const names = accountNames(sheet);
    expect(names.filter((n) => n === SCHEMA_DUMMY_ACCOUNT_NAME)).toHaveLength(1);
    expect(names[1]).toBe(SCHEMA_DUMMY_ACCOUNT_NAME);
    expect(readReportRowsFromSheet(sheet)).toHaveLength(2);
  });

  it('keeps RAISE_TARGET_CPA / LOWER_TARGET_ROAS / MOVE_UNUSED_BUDGET columns numeric when no such rows exist', () => {
    const sheet = new MockSheet('Data', 100);
    writeReportRowsToSheet(sheet, [realRow('1'), realRow('2')]);

    const dummy = sheet.getRange(2, 1, 1, SCHEMA_COLUMN_NAMES.length).getValues()[0];
    const real = sheet.getRange(3, 1, 1, SCHEMA_COLUMN_NAMES.length).getValues()[0];
    const typeSpecificNumericKeys: ReportColumnKey[] = [
      'recommendationNewTargetCpa',
      'recommendationCpaDelta',
      'recommendationBaseCpa',
      'recommendationPotentialCpa',
      'recommendationNewTargetRoas',
      'recommendationRoasDelta',
      'moveBudgetAmount',
      'moveBudgetAmountGlobal',
      'moveBudgetSourceBudgetAmount',
      'moveBudgetSourceTargetCpa',
      'moveBudgetSourceTargetRoas',
    ];
    for (const key of typeSpecificNumericKeys) {
      const idx = SCHEMA_COLUMN_NAMES.indexOf(key);
      expect(real[idx], key).toBe(''); // blank in every real row
      expect(typeof dummy[idx], key).toBe('number');
    }
  });

  it('gives every numeric, date, and boolean column a correctly typed dummy value', () => {
    const sheet = new MockSheet('Data', 100);
    writeReportRowsToSheet(sheet, [realRow('1')]);
    const dummy = sheet.getRange(2, 1, 1, SCHEMA_COLUMN_NAMES.length).getValues()[0];
    const colValue = (key: ReportColumnKey) => dummy[SCHEMA_COLUMN_NAMES.indexOf(key)];

    const booleanKeys = SCHEMA_COLUMN_NAMES.filter(
      (k) => typeof REPORT_SCHEMA_TEMPLATE[k].defaultValue === 'boolean'
    );
    const dateKeys = SCHEMA_COLUMN_NAMES.filter((k) =>
      REPORT_SCHEMA_TEMPLATE[k].format.startsWith('yyyy')
    );
    const numberKeys = SCHEMA_COLUMN_NAMES.filter(
      (k) =>
        typeof REPORT_SCHEMA_TEMPLATE[k].defaultValue !== 'boolean' &&
        !REPORT_SCHEMA_TEMPLATE[k].format.startsWith('yyyy') &&
        REPORT_SCHEMA_TEMPLATE[k].format !== '@'
    );
    const stringKeys = SCHEMA_COLUMN_NAMES.filter(
      (k) =>
        typeof REPORT_SCHEMA_TEMPLATE[k].defaultValue !== 'boolean' &&
        REPORT_SCHEMA_TEMPLATE[k].format === '@'
    );

    expect(booleanKeys.every((k) => typeof colValue(k) === 'boolean')).toBe(true);
    expect(
      dateKeys.every((k) => {
        const cell = colValue(k);
        return cell instanceof Date && !Number.isNaN(cell.getTime());
      })
    ).toBe(true);
    expect(numberKeys.every((k) => typeof colValue(k) === 'number')).toBe(true);
    expect(stringKeys.every((k) => typeof colValue(k) === 'string')).toBe(true);
    expect(colValue('recommendationNumber')).toBe(1);
  });

});

describe('regression: Data & staging readers match columns by header name only', () => {
  const MOVE_BUDGET_KEYS = SCHEMA_COLUMN_NAMES.filter((k) => k.startsWith('moveBudget'));
  const RETIRED_HEADERS = Array.from({length: 8}, (_, i) => `retiredLegacyColumn${i + 1}`);

  /** Writes a raw header + data matrix to a sheet, expanding the mock grid as needed. */
  function writeRawMatrix(sheet: MockSheet, matrix: unknown[][]): void {
    const numCols = Math.max(...matrix.map((row) => row.length));
    if (numCols > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
    }
    if (matrix.length > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), matrix.length - sheet.getMaxRows());
    }
    sheet.getRange(1, 1, matrix.length, numCols).setValues(matrix);
  }

  /** Base Data-row values keyed by schema column; unspecified columns stay blank. */
  const BASE_VALUES: Partial<Record<ReportColumnKey, unknown>> = {
    mccId: 'mcc-a',
    accountId: 'acct-a',
    accountName: 'Acme UK',
    campaignId: 'cmp-a',
    campaignName: 'Brand Search',
    campaignPercentageOfBudgetUsedYesterday: 0.95,
    weeklyCostIncrease: 1750,
    weeklyCostIncreaseGlobal: 2100,
    newWeeklyConversions: 12,
    dailyBudget30Days: 250,
    dailyBudgetDelta: 50,
    recommendationCpaDelta: 3,
    recommendationRoasDelta: 0.5,
  };

  it('reads moveBudget* by name from a legacy 89-column Data header where they sit past column 81', () => {
    const realMoveBudget: Partial<Record<ReportColumnKey, unknown>> = {
      moveBudgetAmount: 42,
      moveBudgetSourceCampaigns: 'Donor A, Donor B',
      moveBudgetSourceBudgetName: 'Donor Budget',
      moveBudgetSourceBudgetType: 'DAILY',
      moveBudgetSourceBudgetDeliveryMethod: 'STANDARD',
      moveBudgetSourceBudgetIsShared: true,
      moveBudgetSourceBudgetAmount: 300,
      moveBudgetSourceTargetCpa: 15,
      moveBudgetSourceTargetRoas: 3.5,
    };
    // Legacy layout: the 81-column schema (before moveBudgetAmountGlobal existed) with 8 retired
    // columns inserted before the derived-weekly block, pushing the trailing moveBudget* block to
    // columns 81-89.
    const legacySchema = SCHEMA_COLUMN_NAMES.filter((k) => k !== 'moveBudgetAmountGlobal');
    const legacyMoveBudgetKeys = MOVE_BUDGET_KEYS.filter((k) => k !== 'moveBudgetAmountGlobal');
    const insertAt = legacySchema.indexOf('campaignPercentageOfBudgetUsedYesterday');
    const header: string[] = [
      ...legacySchema.slice(0, insertAt),
      ...RETIRED_HEADERS,
      ...legacySchema.slice(insertAt),
    ];
    expect(header.length).toBe(89);
    // moveBudget* occupy (1-based) columns 81-89 in the legacy layout.
    expect(header.slice(80)).toStrictEqual(legacyMoveBudgetKeys);
    // The positional slot of moveBudgetSourceCampaigns in the legacy schema holds weeklyCostIncrease.
    expect(header[legacySchema.indexOf('moveBudgetSourceCampaigns')]).toBe('weeklyCostIncrease');

    const values: Record<string, unknown> = {...BASE_VALUES, ...realMoveBudget};
    RETIRED_HEADERS.forEach((h) => {
      values[h] = 9999;
    });
    const dataRow = header.map((h) => values[h] ?? '');

    const sheet = new MockSheet('Data', 100);
    writeRawMatrix(sheet, [header, dataRow]);

    const rows = readReportRowsFromSheet(sheet);
    expect(rows.length).toBe(1);
    const row = rows[0];
    for (const key of legacyMoveBudgetKeys) {
      expect(row[key]).toStrictEqual(realMoveBudget[key]);
    }
    // Columns absent from the legacy layout fall back to the schema default.
    expect(row.moveBudgetAmountGlobal).toBeNull();
    expect(row.moveBudgetSourceCampaigns).not.toBe('1750');
    expect(row.moveBudgetSourceBudgetName).not.toBe('2100');
    expect(row.weeklyCostIncrease).toBe(1750);
    expect(row.weeklyCostIncreaseGlobal).toBe(2100);
    expect(row.campaignName).toBe('Brand Search');
  });

  it('returns schema defaults for moveBudget* when their headers are missing from the Data tab', () => {
    // moveBudget* header names replaced in place by retired columns holding 9999.
    const header: string[] = SCHEMA_COLUMN_NAMES.map((k) =>
      k.startsWith('moveBudget') ? `retired_${k}` : k
    );
    const dataRow = header.map((h) =>
      h.startsWith('retired_') ? 9999 : (BASE_VALUES as Record<string, unknown>)[h] ?? ''
    );

    const sheet = new MockSheet('Data', 100);
    writeRawMatrix(sheet, [header, dataRow]);

    const rows = readReportRowsFromSheet(sheet);
    expect(rows.length).toBe(1);
    for (const key of MOVE_BUDGET_KEYS) {
      expect(rows[0][key]).toStrictEqual(REPORT_SCHEMA_TEMPLATE[key].defaultValue);
    }
    expect(rows[0].weeklyCostIncrease).toBe(1750);
  });

  it('returns FieldType defaults for columns missing from a per-MCC staging tab', () => {
    const recHeader = [
      'mccId',
      'accountId',
      'recommendationId',
      'recommendationType',
      'campaignResourceName',
      'recommendationCurrentBudgetAmount',
      'recommendationNewBudgetAmount',
      'recommendationNewTargetCpa',
      'recommendationNewTargetRoas',
      'recommendationBaseCost',
      'recommendationPotentialCost',
      'recommendationBaseClicks',
      'recommendationPotentialClicks',
      'recommendationBaseConversions',
      'recommendationPotentialConversions',
      'recommendationBaseConversionsValue',
      'recommendationPotentialConversionsValue',
      'recommendationBaseImpressions',
      'recommendationPotentialImpressions',
      'recommendationBaseVideoViews',
      'recommendationPotentialVideoViews',
      'moveBudgetSourceBudgetResourceName',
      'moveBudgetSourceCampaigns',
      'moveBudgetSourceTargetCpa',
      'moveBudgetSourceTargetRoas',
    ];
    const values: Record<string, unknown> = {
      mccId: '111-222-3333',
      accountId: '123-456-7890',
      recommendationId: 'customers/1234567890/recommendations/99',
      recommendationType: 'CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/1234567890/campaigns/10',
      recommendationCurrentBudgetAmount: 100,
      recommendationBaseCost: 700,
      moveBudgetSourceBudgetResourceName: 'customers/1234567890/campaignBudgets/7',
      retiredNumberColumn: 9999,
      retiredStringColumn: 'LEAKED_RETIRED_VALUE',
    };
    // Drop a NUMBER and a STRING column, putting a retired column in each former position.
    const header = recHeader.map((h) => {
      if (h === 'recommendationNewBudgetAmount') return 'retiredNumberColumn';
      if (h === 'moveBudgetSourceCampaigns') return 'retiredStringColumn';
      return h;
    });
    const dataRow = header.map((h) => values[h] ?? '');

    const ss = createMockSpreadsheetApp().getActiveSpreadsheet();
    const sheet = ss.insertSheet('_Stage_Recs_1112223333');
    writeRawMatrix(sheet, [header, dataRow]);

    const readBack = readStagingBatchForMcc(ss, '1112223333');
    expect(readBack.recommendations.length).toBe(1);
    const rec = readBack.recommendations[0];
    expect(rec.recommendationNewBudgetAmount).toBeNull();
    expect(rec.moveBudgetSourceCampaigns).toBe('');
    expect(rec.recommendationCurrentBudgetAmount).toBe(100);
    expect(rec.recommendationBaseCost).toBe(700);
    expect(rec.moveBudgetSourceBudgetResourceName).toBe('customers/1234567890/campaignBudgets/7');
    expect(Object.values(rec)).not.toContain(9999);
    expect(Object.values(rec)).not.toContain('LEAKED_RETIRED_VALUE');
  });

  it('uses the first occurrence when a header name is duplicated', () => {
    // Later duplicates of campaignName / weeklyCostIncrease replace two derived columns.
    const header: string[] = SCHEMA_COLUMN_NAMES.map((k) => {
      if (k === 'recommendationRoasDelta') return 'campaignName';
      if (k === 'dailyBudgetDelta') return 'weeklyCostIncrease';
      return k;
    });
    const values: Record<string, unknown> = {...BASE_VALUES, campaignName: 'First Campaign'};
    const dataRow: unknown[] = SCHEMA_COLUMN_NAMES.map((k) => {
      if (k === 'recommendationRoasDelta') return 'Second Campaign';
      if (k === 'dailyBudgetDelta') return 8888;
      return values[k] ?? '';
    });

    const sheet = new MockSheet('Data', 100);
    writeRawMatrix(sheet, [header, dataRow]);

    const rows = readReportRowsFromSheet(sheet);
    expect(rows.length).toBe(1);
    expect(rows[0].campaignName).toBe('First Campaign');
    expect(rows[0].weeklyCostIncrease).toBe(1750);
  });
});

describe('Data date columns round-trip as strings (stale-publish timestamp corruption)', () => {
  const TS_COL = SCHEMA_COLUMN_NAMES.indexOf('timestamp') + 1;
  const DATE_COL = SCHEMA_COLUMN_NAMES.indexOf('mccLastUpdatedDate') + 1;
  /** Row 1 is the header and Row 2 the schema dummy row, so the first real row is Row 3. */
  const FIRST_DATA_ROW = 3;

  /** Time-zone-aware stand-in for `Utilities.formatDate` (patterns used by the Data reader only). */
  function intlFormatDate(date: Date, timeZone: string, pattern: string): string {
    const parts = Object.fromEntries(
      new Intl.DateTimeFormat('en-CA', {
        timeZone,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hourCycle: 'h23',
      })
        .formatToParts(date)
        .map((p) => [p.type, p.value])
    );
    const day = `${parts.year}-${parts.month}-${parts.day}`;
    return pattern === 'yyyy-MM-dd' ? day : `${day} ${parts.hour}:${parts.minute}:${parts.second}`;
  }

  /** Runs `fn` with a time-zone-aware `Utilities` global, restoring the previous value after. */
  function withTzAwareUtilities(fn: () => void): void {
    const g = globalThis as Record<string, unknown>;
    const previous = g.Utilities;
    g.Utilities = {formatDate: intlFormatDate, sleep: () => {}};
    try {
      fn();
    } finally {
      g.Utilities = previous;
    }
  }

  function writeSingleRow(timestamp: string, mccLastUpdatedDate: string): MockSheet {
    const sheet = new MockSheet('Data', 100);
    const row = createEmptyReportRow();
    row.accountName = 'Acme UK';
    row.campaignId = '111';
    row.timestamp = timestamp;
    row.mccLastUpdatedDate = mccLastUpdatedDate;
    writeReportRowsToSheet(sheet, [row]);
    return sheet;
  }

  it('reads Date cells in date-formatted columns back as formatted strings', () => {
    const sheet = writeSingleRow('2026-10-26 13:45:10', '2026-10-26');
    // Sheets parses date strings in date-formatted cells into Date values
    expect(sheet.getRange(FIRST_DATA_ROW, TS_COL).getValue()).toBeInstanceOf(Date);
    expect(sheet.getRange(FIRST_DATA_ROW, DATE_COL).getValue()).toBeInstanceOf(Date);

    withTzAwareUtilities(() => {
      const rows = readReportRowsFromSheet(sheet, 'Etc/UTC');
      expect(rows[0].timestamp).toBe('2026-10-26 13:45:10');
      expect(rows[0].mccLastUpdatedDate).toBe('2026-10-26');
    });
  });

  it('formats Date cells in the given spreadsheet time zone', () => {
    const sheet = writeSingleRow('2026-10-26 00:00:00', '2026-10-26');
    // Midnight wall-clock in New York is 04:00 UTC
    sheet.getRange(FIRST_DATA_ROW, TS_COL).setValue(new Date('2026-10-26T04:00:00Z'));
    sheet.getRange(FIRST_DATA_ROW, DATE_COL).setValue(new Date('2026-10-26T04:00:00Z'));

    withTzAwareUtilities(() => {
      const rows = readReportRowsFromSheet(sheet, 'America/New_York');
      expect(rows[0].timestamp).toBe('2026-10-26 00:00:00');
      expect(rows[0].mccLastUpdatedDate).toBe('2026-10-26');
    });
  });

  it('self-heals epoch-millisecond numbers and drops other numbers to the default', () => {
    const sheet = writeSingleRow('2026-10-26 00:00:00', '2026-10-26');
    sheet.getRange(FIRST_DATA_ROW, TS_COL).setValue(1792972800000);
    sheet.getRange(FIRST_DATA_ROW, DATE_COL).setValue(1792972800000);

    withTzAwareUtilities(() => {
      const rows = readReportRowsFromSheet(sheet, 'Etc/UTC');
      expect(rows[0].timestamp).toBe('2026-10-26 00:00:00');
      expect(rows[0].mccLastUpdatedDate).toBe('2026-10-26');
    });

    sheet.getRange(FIRST_DATA_ROW, TS_COL).setValue(46321);
    sheet.getRange(FIRST_DATA_ROW, DATE_COL).setValue(46321);
    const rows = readReportRowsFromSheet(sheet, 'Etc/UTC');
    expect(rows[0].timestamp).toBe(REPORT_SCHEMA_TEMPLATE.timestamp.defaultValue);
    expect(rows[0].mccLastUpdatedDate).toBe(REPORT_SCHEMA_TEMPLATE.mccLastUpdatedDate.defaultValue);
  });

  it('passes date strings through unchanged and maps empty cells to defaults', () => {
    const sheet = writeSingleRow('', '');
    sheet.getRange(FIRST_DATA_ROW, DATE_COL).setValue('previous run');
    const rows = readReportRowsFromSheet(sheet, 'Etc/UTC');
    expect(rows[0].timestamp).toBe(REPORT_SCHEMA_TEMPLATE.timestamp.defaultValue);
    expect(rows[0].mccLastUpdatedDate).toBe('previous run');
  });
});

describe('sheet cell helpers', () => {
  it('escapes every formula-prefix character, including tab and carriage return, and round-trips it', () => {
    for (const raw of ['=Promo', '+Brand', '-Search', '@Mention', '\tTabbed', '\rReturn']) {
      const escaped = escapeFormulaString(raw);
      expect(escaped).toBe(`'${raw}`);
      expect(unescapeFormulaString(escaped)).toBe(raw);
    }
    expect(escapeFormulaString('Plain name')).toBe('Plain name');
    expect(unescapeFormulaString("'Plain")).toBe("'Plain");
  });

  it('parses checkbox cells as booleans', () => {
    expect([true, 'TRUE', 'true', false, 'FALSE', '', null, undefined, 1].map(parseSheetBoolean)).toEqual([
      true, true, true, false, false, false, false, false, false,
    ]);
  });

  it('builds a value matrix with blanks for null and escaped strings', () => {
    expect(toSheetValueMatrix(['a', 'b', 'c'], [{a: null, b: '=x', c: 3}])).toEqual([['', "'=x", 3]]);
  });

  it('joins campaign lists cleanly at an item boundary with (+N more) when exceeding the limit', () => {
    expect(joinListForSheetCell([])).toBe('');
    expect(joinListForSheetCell(['Alpha', 'Beta'], 50)).toBe('Alpha, Beta');
    expect(joinListForSheetCell(['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon'], 28)).toBe(
      'Alpha, Beta, ... (+3 more)'
    );
  });
});
