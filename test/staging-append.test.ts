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
import {appendStagingBatch, readStagingBatchForMcc} from '../src/shared/sheets/staging-writer';
import type {StagedCampaign, StagedRecommendation, StagingBatch} from '../src/shared/types';
import {
  type MockSheet,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

const MCC = '1112223333';

function rec(id: string): StagedRecommendation {
  return {
    mccId: '111-222-3333',
    accountId: '100-000-0001',
    recommendationId: `customers/1000000001/recommendations/${id}`,
    recommendationType: 'CAMPAIGN_BUDGET',
    campaignResourceName: 'customers/1000000001/campaigns/10',
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
    recommendationBaseConversionsValue: null,
    recommendationPotentialConversionsValue: null,
    recommendationBaseImpressions: null,
    recommendationPotentialImpressions: null,
    recommendationBaseVideoViews: null,
    recommendationPotentialVideoViews: null,
    moveBudgetSourceBudgetResourceName: '',
    moveBudgetSourceCampaigns: 'N/A',
    moveBudgetSourceTargetCpa: null,
    moveBudgetSourceTargetRoas: null,
  };
}

function campaign(): StagedCampaign {
  return {
    mccId: '111-222-3333',
    accountId: '100-000-0001',
    campaignResourceName: 'customers/1000000001/campaigns/10',
    campaignId: '0010',
    campaignName: 'Brand',
    campaignType: 'SEARCH',
    campaignSubType: 'N/A',
    campaignIsAiMax: true,
    campaignAiMaxTextCustomizationEnabled: false,
    campaignAiMaxFinalUrlExpansionEnabled: false,
    budgetResourceName: 'customers/1000000001/campaignBudgets/1',
    campaignIsPortfolioBiddingStrategy: false,
    campaignPortfolioBiddingStrategyName: 'N/A',
    portfolioStrategyResourceName: 'N/A',
    portfolioOwnerId: 'N/A',
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
  };
}

function batch(recs: StagedRecommendation[], campaigns: StagedCampaign[] = []): StagingBatch {
  return {accounts: [], budgets: [], campaigns, recommendations: recs};
}

/** Records the first row of every value write on `sheet`. */
function trackWriteRows(sheet: MockSheet): number[] {
  const rows: number[] = [];
  const original = sheet.getRange.bind(sheet);
  sheet.getRange = ((rowOrA1: number | string, col?: number, numRows?: number, numCols?: number) => {
    const range = original(rowOrA1, col, numRows, numCols);
    const setValues = range.setValues.bind(range);
    range.setValues = (values: unknown[][]) => {
      rows.push(range.getRow());
      return setValues(values);
    };
    return range;
  }) as MockSheet['getRange'];
  return rows;
}

describe('Staging writer: append fast path and column formats', () => {
  function setup() {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({spreadsheetApp});
    const ss = spreadsheetApp.getActiveSpreadsheet();
    appendStagingBatch(ss, batch([rec('1'), rec('2')]), MCC);
    const recsSheet = ss.getSheetByName(`_Stage_Recs_${MCC}`) as MockSheet;
    return {ss, recsSheet};
  }

  it('appends only the new rows after the last row when nothing collides', () => {
    const {ss, recsSheet} = setup();
    const writeRows = trackWriteRows(recsSheet);
    appendStagingBatch(ss, batch([rec('3'), rec('4')]), MCC);
    expect(writeRows).toEqual([4]);
    expect(readStagingBatchForMcc(ss, MCC).recommendations.map((r) => r.recommendationId.split('/').pop())).toEqual([
      '1',
      '2',
      '3',
      '4',
    ]);
  });

  it('rewrites the whole table when a new row updates an existing key', () => {
    const {ss, recsSheet} = setup();
    const writeRows = trackWriteRows(recsSheet);
    const updated = {...rec('2'), recommendationNewBudgetAmount: 999};
    appendStagingBatch(ss, batch([updated, rec('3')]), MCC);
    expect(writeRows).toEqual([1]);
    const recs = readStagingBatchForMcc(ss, MCC).recommendations;
    expect(recs.length).toBe(3);
    expect(recs.find((r) => r.recommendationId.endsWith('/2'))?.recommendationNewBudgetAmount).toBe(999);
  });

  it('rewrites the whole table when the header row does not match the schema', () => {
    const {ss, recsSheet} = setup();
    recsSheet.getRange(1, 2).setValue('legacyAccountIdColumn');
    const writeRows = trackWriteRows(recsSheet);
    appendStagingBatch(ss, batch([rec('3')]), MCC);
    expect(writeRows).toEqual([1]);
    expect(recsSheet.getRange(1, 2).getValue()).toBe('accountId');
  });

  it('formats BOOLEAN columns as General, NUMBER as 0.########, and text as @ (full write and append)', () => {
    const {ss} = setup();
    appendStagingBatch(ss, batch([], [campaign()]), MCC);
    appendStagingBatch(ss, batch([], [{...campaign(), campaignResourceName: 'customers/1000000001/campaigns/11'}]), MCC);
    const sheet = ss.getSheetByName(`_Stage_Campaigns_${MCC}`) as MockSheet;
    const header = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
    const col = (name: string) => header.indexOf(name) + 1;
    const expectFormats = () => {
      for (const row of [2, 3]) {
        expect(sheet.getCellFormat(row, col('campaignIsAiMax'))).toBe('General');
        expect(sheet.getCellFormat(row, col('campaign30DaysCost'))).toBe('0.########');
        expect(sheet.getCellFormat(row, col('campaignId'))).toBe('@');
      }
    };
    expectFormats(); // rows 2 and 3 were written by the append path

    // A key update forces the full rewrite, which must apply the same formats
    appendStagingBatch(ss, batch([], [{...campaign(), campaignName: 'Brand (renamed)'}]), MCC);
    expectFormats();
    const staged = readStagingBatchForMcc(ss, MCC).campaigns;
    expect(staged.map((c) => [c.campaignId, c.campaignIsAiMax])).toEqual([
      ['0010', true],
      ['0010', true],
    ]);
  });
});
