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
  REPORT_SCHEMA_TEMPLATE,
  SCHEMA_COLUMN_NAMES,
  createDummyFormattingRow,
  createEmptyReportRow,
} from '../src/shared/schema';
import {formatCidHyphenated, normalizeCid} from '../src/shared/utils/cid';

describe('V2 Schema Definition, Types & CID Normalization', () => {
  it('defines deterministic 82-column schema matching the Master Template Sheet Data tab', () => {
    expect(SCHEMA_COLUMN_NAMES).toEqual(Object.keys(REPORT_SCHEMA_TEMPLATE));
    expect(SCHEMA_COLUMN_NAMES.length).toBe(82);

    // Core V2 additions
    expect(SCHEMA_COLUMN_NAMES).toContain('recommendationNumber');
    expect(SCHEMA_COLUMN_NAMES).toContain('recommendationScope');
    expect(SCHEMA_COLUMN_NAMES).toContain('mccId');
    expect(SCHEMA_COLUMN_NAMES).toContain('mccName');
    expect(SCHEMA_COLUMN_NAMES).toContain('mccSyncStatus');
    expect(SCHEMA_COLUMN_NAMES).toContain('mccLastUpdatedDate');
    expect(SCHEMA_COLUMN_NAMES).toContain('portfolioCampaignsCount');
    expect(SCHEMA_COLUMN_NAMES).toContain('portfolioMemberCampaignNames');
    expect(SCHEMA_COLUMN_NAMES).toContain('moveBudgetAmount');
    expect(SCHEMA_COLUMN_NAMES).toContain('moveBudgetAmountGlobal');
    // Appended at the end (APPEND-ONLY rule: Data Studio binds Sheets fields by position)
    expect(SCHEMA_COLUMN_NAMES[SCHEMA_COLUMN_NAMES.length - 1]).toBe('moveBudgetAmountGlobal');
    expect(SCHEMA_COLUMN_NAMES).toContain('moveBudgetSourceTargetCpa');
    expect(SCHEMA_COLUMN_NAMES).toContain('moveBudgetSourceTargetRoas');
  });

  it('keeps the original 81 Data columns in their exact positions (append-only guard)', () => {
    // Data Studio's Google Sheets connector binds fields by column position, so these 81 columns
    // (the layout before moveBudgetAmountGlobal) must never be inserted into, reordered, or renamed.
    // New columns may only be appended after them.
    const ORIGINAL_81_COLUMNS = [
      'recommendationNumber',
      'recommendationScope',
      'mccId',
      'mccName',
      'mccSyncStatus',
      'mccLastUpdatedDate',
      'accountId',
      'accountName',
      'timestamp',
      'recommendationId',
      'recommendationType',
      'recommendationsDetailsUrl',
      'campaignUrl',
      'campaignId',
      'campaignName',
      'campaignType',
      'campaignSubType',
      'campaignIsAiMax',
      'campaignAiMaxTextCustomizationEnabled',
      'campaignAiMaxFinalUrlExpansionEnabled',
      'currencyCode',
      'currencyCodeGlobal',
      'campaignBudgetAmount',
      'campaignBudgetTotalAmount',
      'campaignBudgetDeliveryMethod',
      'campaignBudgetType',
      'portfolioCampaignsCount',
      'portfolioMemberCampaignNames',
      'campaignBiddingStrategyType',
      'campaignBiddingTargetRoas',
      'campaignBiddingTargetCpa',
      'campaignStatsYesterdayCost',
      'campaignStats7DaysCost',
      'campaign30DaysCost',
      'campaign30DaysConversionsValue',
      'campaign30DaysConversions',
      'campaign30DaysRoas',
      'campaign30DaysCpa',
      'campaign30DaysSearchRankLostImpressionShare',
      'campaign30DaysSearchBudgetLostImpressionShare',
      'recommendationCurrentBudgetAmount',
      'recommendationCurrentBudgetAmountGlobal',
      'recommendationNewBudgetAmount',
      'recommendationNewBudgetAmountGlobal',
      'recommendationNewTargetCpa',
      'recommendationNewTargetRoas',
      'recommendationBaseCost',
      'recommendationPotentialCost',
      'recommendationBaseClicks',
      'recommendationPotentialClicks',
      'recommendationBaseConversions',
      'recommendationPotentialConversions',
      'recommendationBaseConversionsValue',
      'recommendationBaseConversionsValueGlobal',
      'recommendationPotentialConversionsValue',
      'recommendationPotentialConversionsValueGlobal',
      'recommendationBaseCpa',
      'recommendationPotentialCpa',
      'recommendationBaseRoas',
      'recommendationPotentialRoas',
      'recommendationBaseImpressions',
      'recommendationPotentialImpressions',
      'recommendationBaseVideoViews',
      'recommendationPotentialVideoViews',
      'campaignPercentageOfBudgetUsedYesterday',
      'weeklyCostIncrease',
      'weeklyCostIncreaseGlobal',
      'newWeeklyConversions',
      'dailyBudget30Days',
      'dailyBudgetDelta',
      'recommendationCpaDelta',
      'recommendationRoasDelta',
      'moveBudgetAmount',
      'moveBudgetSourceCampaigns',
      'moveBudgetSourceBudgetName',
      'moveBudgetSourceBudgetType',
      'moveBudgetSourceBudgetDeliveryMethod',
      'moveBudgetSourceBudgetIsShared',
      'moveBudgetSourceBudgetAmount',
      'moveBudgetSourceTargetCpa',
      'moveBudgetSourceTargetRoas',
    ];
    expect(ORIGINAL_81_COLUMNS.length).toBe(81);
    expect(SCHEMA_COLUMN_NAMES.slice(0, 81)).toEqual(ORIGINAL_81_COLUMNS);
  });

  it('leaves not-applicable dates and move-budget amounts blank (null) by default', () => {
    for (const key of [
      'timestamp',
      'mccLastUpdatedDate',
      'moveBudgetAmount',
      'moveBudgetSourceBudgetAmount',
      'moveBudgetAmountGlobal',
    ] as const) {
      expect(REPORT_SCHEMA_TEMPLATE[key].defaultValue).toBeNull();
    }
    // The dummy formatting row keeps numeric values so Data Studio binds number types
    expect(typeof REPORT_SCHEMA_TEMPLATE.moveBudgetAmount.dummyValueForFormatting).toBe('number');
    expect(typeof REPORT_SCHEMA_TEMPLATE.moveBudgetSourceBudgetAmount.dummyValueForFormatting).toBe(
      'number'
    );
  });

  it('normalizes string and numeric CIDs to 10-digit strings and formats with hyphens', () => {
    expect(normalizeCid('123-456-7890')).toBe('1234567890');
    expect(normalizeCid(' 123 456 7890 ')).toBe('1234567890');
    expect(normalizeCid(1234567890)).toBe('1234567890');
    expect(normalizeCid(12345678)).toBe('0012345678');
    expect(normalizeCid('')).toBe('');
    expect(normalizeCid(' 123-456-7890 ')).toBe('1234567890');
    expect(normalizeCid('0012345678')).toBe('0012345678');
  });

  it('rejects values that are not Customer IDs instead of extracting their digits', () => {
    expect(normalizeCid('MCC 1')).toBe('');
    expect(normalizeCid('Account 2026')).toBe('');
    // Dates have hyphens but not the XXX-XXX-XXXX shape, so they never become 8-digit CIDs
    expect(normalizeCid('2026-10-14')).toBe('');
    expect(normalizeCid('12-345-67890')).toBe('');
    expect(normalizeCid('12345678901')).toBe('');
    expect(normalizeCid(12345678901)).toBe('');
    expect(normalizeCid('000-000-0000')).toBe('');
    expect(normalizeCid(0)).toBe('');
    expect(normalizeCid(-1234567890)).toBe('');
    expect(normalizeCid(null)).toBe('');
    expect(normalizeCid(undefined)).toBe('');
    expect(formatCidHyphenated('MCC 1')).toBe('');
    expect(formatCidHyphenated('2026-10-14')).toBe('');

    expect(formatCidHyphenated('1234567890')).toBe('123-456-7890');
    expect(formatCidHyphenated('123-456-7890')).toBe('123-456-7890');
  });

  it('creates empty and dummy schema rows with all 82 schema keys populated', () => {
    const emptyRow = createEmptyReportRow();
    const dummyRow = createDummyFormattingRow();

    expect(Object.keys(emptyRow)).toEqual(SCHEMA_COLUMN_NAMES);
    expect(Object.keys(dummyRow)).toEqual(SCHEMA_COLUMN_NAMES);
    expect(dummyRow.accountName).toBe('!!! SCHEMA DEFINITION (FILTER ME OUT)');
    expect(dummyRow.campaignId).toBe('DUMMY_SCHEMA_ROW');
  });
});
