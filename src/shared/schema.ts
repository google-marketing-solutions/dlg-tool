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
 * @fileoverview Single source of truth for the 82-column `Data` sheet schema (`REPORT_SCHEMA_TEMPLATE`).
 * Specifies the exact column order, default values, Google Sheets number/text format strings,
 * and the schema dummy row (`createDummyFormattingRow`) that is always written as `Data` Row 2.
 * Bundled only into `dist/sheet/Code.gs` (which builds and formats the final `Data` tab).
 */

import type {ReportColumnKey, ReportRow, SchemaFieldSpec} from './types';

/**
 * Authoritative V2 schema definition matching the 82 columns in the Master Template Sheet's `Data` tab.
 * Column order is guaranteed by `Object.keys(REPORT_SCHEMA_TEMPLATE)`.
 *
 * APPEND-ONLY: add new columns at the END only. Never insert, reorder, or rename existing columns,
 * because Data Studio's Google Sheets connector binds `Data` fields by column position; shifting a
 * column changes the type and data of every field after it in every existing dashboard. (The staging
 * `_Stage_*` tabs are read by header name, so this rule is about the `Data` tab.)
 */
export const REPORT_SCHEMA_TEMPLATE: Record<ReportColumnKey, SchemaFieldSpec> = {
  // --- Ranking, Scope & Multi-MCC Metadata ---
  recommendationNumber: { defaultValue: 1, format: '#,##0', dummyValueForFormatting: 1 },
  recommendationScope: { defaultValue: 'CAMPAIGN', format: '@', dummyValueForFormatting: 'CAMPAIGN' },
  mccId: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: '000-000-0000' },
  mccName: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DUMMY_MCC' },
  mccSyncStatus: { defaultValue: 'OK', format: '@', dummyValueForFormatting: 'OK' },
  mccLastUpdatedDate: { defaultValue: null, format: 'yyyy-mm-dd', dummyValueForFormatting: '2026-10-14' },

  // --- Account & Recommendation Identity ---
  accountId: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: '000-000-0000' },
  accountName: {
    defaultValue: 'N/A',
    format: '@',
    dummyValueForFormatting: '!!! SCHEMA DEFINITION (FILTER ME OUT)',
  },
  timestamp: { defaultValue: null, format: 'yyyy-mm-dd hh:mm:ss', dummyValueForFormatting: '2026-10-14 00:00:00' },
  recommendationId: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DUMMY_RECO_ID' },
  recommendationType: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DUMMY_TYPE' },
  recommendationsDetailsUrl: {
    defaultValue: 'N/A',
    format: '@',
    dummyValueForFormatting: 'https://ads.google.com',
  },

  // --- Campaign Details ---
  campaignUrl: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'https://ads.google.com' },
  campaignId: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DUMMY_SCHEMA_ROW' },
  campaignName: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DUMMY_CAMPAIGN' },
  campaignType: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'SEARCH' },
  campaignSubType: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'N/A' },
  campaignIsAiMax: { defaultValue: false, format: 'General', dummyValueForFormatting: false },
  campaignAiMaxTextCustomizationEnabled: { defaultValue: false, format: 'General', dummyValueForFormatting: false },
  campaignAiMaxFinalUrlExpansionEnabled: { defaultValue: false, format: 'General', dummyValueForFormatting: false },
  currencyCode: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'USD' },
  currencyCodeGlobal: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'USD' },

  // --- Budget Details ---
  campaignBudgetAmount: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaignBudgetTotalAmount: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaignBudgetDeliveryMethod: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'STANDARD' },
  campaignBudgetType: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DAILY' },

  // --- Bidding & Portfolio Details ---
  portfolioCampaignsCount: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  portfolioMemberCampaignNames: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'N/A' },
  campaignBiddingStrategyType: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'MANUAL_CPC' },
  campaignBiddingTargetRoas: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },
  campaignBiddingTargetCpa: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },

  // --- Historical Performance Stats (Yesterday / 7d / 30d) ---
  campaignStatsYesterdayCost: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaignStats7DaysCost: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaign30DaysCost: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaign30DaysConversionsValue: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaign30DaysConversions: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  campaign30DaysRoas: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },
  campaign30DaysCpa: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  campaign30DaysSearchRankLostImpressionShare: {
    defaultValue: null,
    format: '0.00%',
    dummyValueForFormatting: 0.0,
  },
  campaign30DaysSearchBudgetLostImpressionShare: {
    defaultValue: null,
    format: '0.00%',
    dummyValueForFormatting: 0.0,
  },

  // --- Recommendation Targets & 7-Day Native Forecast Impact ---
  recommendationCurrentBudgetAmount: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationCurrentBudgetAmountGlobal: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationNewBudgetAmount: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationNewBudgetAmountGlobal: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationNewTargetCpa: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationNewTargetRoas: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },
  recommendationBaseCost: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationPotentialCost: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationBaseClicks: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationPotentialClicks: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationBaseConversions: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationPotentialConversions: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationBaseConversionsValue: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationBaseConversionsValueGlobal: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationPotentialConversionsValue: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationPotentialConversionsValueGlobal: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationBaseCpa: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationPotentialCpa: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationBaseRoas: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },
  recommendationPotentialRoas: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },
  recommendationBaseImpressions: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationPotentialImpressions: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationBaseVideoViews: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  recommendationPotentialVideoViews: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },

  // --- Derived Weekly & Budget Metrics ---
  campaignPercentageOfBudgetUsedYesterday: {
    defaultValue: null,
    format: '0.00%',
    dummyValueForFormatting: 0.0,
  },
  weeklyCostIncrease: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  weeklyCostIncreaseGlobal: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  newWeeklyConversions: { defaultValue: null, format: '#,##0', dummyValueForFormatting: 0 },
  dailyBudget30Days: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  dailyBudgetDelta: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationCpaDelta: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  recommendationRoasDelta: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },

  // --- MOVE_UNUSED_BUDGET Specifics (Including Source Campaign Targets) ---
  moveBudgetAmount: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  moveBudgetSourceCampaigns: {
    defaultValue: 'N/A',
    format: '@',
    dummyValueForFormatting: 'CAMPAIGN_1, CAMPAIGN_2',
  },
  moveBudgetSourceBudgetName: {
    defaultValue: 'N/A',
    format: '@',
    dummyValueForFormatting: 'SOURCE_BUDGET',
  },
  moveBudgetSourceBudgetType: { defaultValue: 'N/A', format: '@', dummyValueForFormatting: 'DAILY' },
  moveBudgetSourceBudgetDeliveryMethod: {
    defaultValue: 'N/A',
    format: '@',
    dummyValueForFormatting: 'STANDARD',
  },
  moveBudgetSourceBudgetIsShared: { defaultValue: false, format: 'General', dummyValueForFormatting: false },
  moveBudgetSourceBudgetAmount: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  moveBudgetSourceTargetCpa: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
  moveBudgetSourceTargetRoas: { defaultValue: null, format: '0.00', dummyValueForFormatting: 0.0 },

  // --- Appended Columns (82+; APPEND-ONLY, see the rule above) ---
  moveBudgetAmountGlobal: {
    defaultValue: null,
    format: '#,##0.00',
    dummyValueForFormatting: 0.0,
  },
};

/**
 * Ordered list of all 82 column names written to Row 1 of the `Data` tab.
 */
export const SCHEMA_COLUMN_NAMES = Object.keys(REPORT_SCHEMA_TEMPLATE) as ReportColumnKey[];

/**
 * Creates a fresh `ReportRow` populated with each column's `defaultValue` from `REPORT_SCHEMA_TEMPLATE`.
 */
export function createEmptyReportRow(): ReportRow {
  const row: Record<string, unknown> = {};
  for (const key of SCHEMA_COLUMN_NAMES) {
    row[key] = REPORT_SCHEMA_TEMPLATE[key].defaultValue;
  }
  return row as unknown as ReportRow;
}

/**
 * `accountName` of the schema dummy row. Dashboards must filter out rows with this account name
 * (the row also has `recommendationNumber = 1`, so a "top recommendation" filter alone keeps it).
 */
export const SCHEMA_DUMMY_ACCOUNT_NAME = REPORT_SCHEMA_TEMPLATE.accountName
  .dummyValueForFormatting as string;

/** True when `row` is the schema dummy row (matched by its `accountName`). */
export function isDummyFormattingRow(row: Pick<ReportRow, 'accountName'>): boolean {
  return row.accountName === SCHEMA_DUMMY_ACCOUNT_NAME;
}

/**
 * Creates the schema dummy row that `writeReportRowsToSheet` always writes to `Data` Row 2, directly
 * under the header. Every numeric, date, and boolean column holds a typed value, so the Sheets connector
 * in Data Studio infers the right field types even when no real row has a value in a
 * column (for example, no `RAISE_TARGET_CPA` or `MOVE_UNUSED_BUDGET` recommendations today).
 */
export function createDummyFormattingRow(): ReportRow {
  const row: Record<string, unknown> = {};
  for (const key of SCHEMA_COLUMN_NAMES) {
    row[key] = REPORT_SCHEMA_TEMPLATE[key].dummyValueForFormatting;
  }
  return row as unknown as ReportRow;
}
