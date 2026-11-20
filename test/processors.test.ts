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
import type {
  BudgetDetails,
  EnrichedCampaignData,
  GaqlRecommendationRow,
  SourceBudgetCampaignInfo,
} from '../src/ads/gaql/fetcher';
import {fetchCampaignsByBudgets} from '../src/ads/gaql/fetcher';
import {
  type BuildAccountRowsContext,
  buildAccountStagingBatch,
} from '../src/ads/processors/recommendations';
import {consolidateRowsByScope} from '../src/appscript/processors/consolidator';
import {joinStagingBatchToReportRows} from '../src/appscript/processors/joiner';
import {assignRecommendationNumbers} from '../src/appscript/processors/ranking';
import {createEmptyReportRow} from '../src/shared/schema';
import type {ReportRow} from '../src/shared/types';
import {createMockAdsApp, installAppsScriptGlobals} from './mocks/google-apps-script.mock';

function buildAccountRecommendationRows(ctx: BuildAccountRowsContext): ReportRow[] {
  return joinStagingBatchToReportRows(buildAccountStagingBatch(ctx));
}

function makeCampaign(overrides: Partial<EnrichedCampaignData> = {}): EnrichedCampaignData {
  return {
    campaignResourceName: overrides.campaignResourceName ?? 'customers/100/campaigns/1',
    campaignId: overrides.campaignId ?? '1',
    campaignName: overrides.campaignName ?? 'Campaign 1',
    campaignType: overrides.campaignType ?? 'SEARCH',
    campaignSubType: overrides.campaignSubType ?? 'UNSPECIFIED',
    merchantId: overrides.merchantId ?? null,
    campaignIsAiMax: overrides.campaignIsAiMax ?? false,
    campaignAiMaxTextCustomizationEnabled: overrides.campaignAiMaxTextCustomizationEnabled ?? false,
    campaignAiMaxFinalUrlExpansionEnabled: overrides.campaignAiMaxFinalUrlExpansionEnabled ?? false,
    campaignBudgetResourceName: overrides.campaignBudgetResourceName ?? 'customers/100/campaignBudgets/10',
    campaignBudgetName: overrides.campaignBudgetName ?? 'Budget 10',
    campaignBudgetAmount: overrides.campaignBudgetAmount ?? 100,
    campaignBudgetTotalAmount: overrides.campaignBudgetTotalAmount ?? null,
    campaignBudgetDeliveryMethod: overrides.campaignBudgetDeliveryMethod ?? 'STANDARD',
    campaignBudgetType: overrides.campaignBudgetType ?? 'DAILY',
    campaignBudgetIsShared: overrides.campaignBudgetIsShared ?? false,
    campaignIsPortfolioBiddingStrategy: overrides.campaignIsPortfolioBiddingStrategy ?? false,
    campaignPortfolioBiddingStrategyName: overrides.campaignPortfolioBiddingStrategyName ?? 'N/A',
    portfolioStrategyResourceName: overrides.portfolioStrategyResourceName ?? 'N/A',
    portfolioOwnerId: overrides.portfolioOwnerId ?? 'N/A',
    isLocallyOwnedPortfolio: overrides.isLocallyOwnedPortfolio ?? false,
    campaignPortfolioIsManagerOwned: overrides.campaignPortfolioIsManagerOwned ?? false,
    campaignBiddingStrategyType: overrides.campaignBiddingStrategyType ?? 'TARGET_CPA',
    campaignBiddingTargetCpa: overrides.campaignBiddingTargetCpa ?? 20,
    campaignBiddingTargetRoas: overrides.campaignBiddingTargetRoas ?? null,
    campaignStatsYesterdayCost: overrides.campaignStatsYesterdayCost ?? 90,
    campaignStats7DaysCost: overrides.campaignStats7DaysCost ?? 600,
    campaign30DaysCost: overrides.campaign30DaysCost ?? 2500,
    campaign30DaysConversionsValue: overrides.campaign30DaysConversionsValue ?? 5000,
    campaign30DaysConversions: overrides.campaign30DaysConversions ?? 100,
    campaign30DaysRoas: overrides.campaign30DaysRoas ?? 2.0,
    campaign30DaysCpa: overrides.campaign30DaysCpa ?? 25,
    campaign30DaysSearchRankLostImpressionShare:
      overrides.campaign30DaysSearchRankLostImpressionShare ?? 0.1,
    campaign30DaysSearchBudgetLostImpressionShare:
      overrides.campaign30DaysSearchBudgetLostImpressionShare ?? 0.2,
  };
}

describe('Recommendation Handlers, Native Forecasts, Scope Consolidation & Ranking', () => {
  it('extracts native 7-day forecasts for FORECASTING/MARGINAL_ROI budget recs and never fabricates null conversionsValue', () => {
    const campRes = 'customers/100/campaigns/1';
    const recRes = 'customers/100/recommendations/rec1';

    const campaignMap = new Map<string, EnrichedCampaignData>([[campRes, makeCampaign()]]);
    const recDetailsMap = new Map<string, GaqlRecommendationRow>([
      [
        recRes,
        {
          resourceName: recRes,
          type: 'FORECASTING_CAMPAIGN_BUDGET',
          campaign: campRes,
          forecastingCampaignBudgetRecommendation: {
            currentBudgetAmountMicros: 100_000_000,
            recommendedBudgetAmountMicros: 150_000_000,
          },
          // Forecast impact provided on top-level rec.impact with null conversionsValue
          impact: {
            baseMetrics: {
              costMicros: 700_000_000,
              clicks: 500,
              conversions: 35,
              conversionsValue: null,
              impressions: 5000,
              videoViews: 0,
            },
            potentialMetrics: {
              costMicros: 1_050_000_000,
              clicks: 700,
              conversions: 50,
              conversionsValue: null,
              impressions: 7000,
              videoViews: 0,
            },
          },
        },
      ],
    ]);

    const rows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: recDetailsMap,
      campaignDataMap: campaignMap,
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });

    expect(rows.length).toBe(1);
    const row = rows[0];
    expect(row.recommendationType).toBe('FORECASTING_CAMPAIGN_BUDGET');
    expect(row.recommendationCurrentBudgetAmount).toBe(100);
    expect(row.recommendationNewBudgetAmount).toBe(150);
    expect(row.recommendationBaseCost).toBe(700);
    expect(row.recommendationPotentialCost).toBe(1050);
    expect(row.weeklyCostIncrease).toBe(350);
    expect(row.newWeeklyConversions).toBe(15);

    // Native conversionsValue is null -> must stay strictly null (zero custom extrapolation!)
    expect(row.recommendationBaseConversionsValue).toBeNull();
    expect(row.recommendationPotentialConversionsValue).toBeNull();
    expect(row.recommendationBaseRoas).toBeNull();
    expect(row.recommendationPotentialRoas).toBeNull();
    // Move-budget amounts only apply to MOVE_UNUSED_BUDGET, so they stay blank (null) here
    expect(row.moveBudgetAmount).toBeNull();
    expect(row.moveBudgetSourceBudgetAmount).toBeNull();
  });

  it('enriches MOVE_UNUSED_BUDGET with source campaign targets and filters by EC Remedy (feed-PMax only)', () => {
    const destCampRes = 'customers/100/campaigns/2';
    const excessBudgetRes = 'customers/100/campaignBudgets/99';
    const recRes = 'customers/100/recommendations/move1';

    const campaignMap = new Map<string, EnrichedCampaignData>([
      [
        destCampRes,
        makeCampaign({
          campaignResourceName: destCampRes,
          campaignId: '2',
          campaignName: 'Non-Feed PMax Dest',
          campaignType: 'PERFORMANCE_MAX',
          merchantId: null, // Non-feed PMax -> retained under EC Remedy!
        }),
      ],
    ]);

    const budgetMap = new Map<string, BudgetDetails>([
      [
        excessBudgetRes,
        {
          budgetResourceName: excessBudgetRes,
          campaignBudgetName: 'Excess Search Budget',
          campaignBudgetAmount: 200,
          campaignBudgetTotalAmount: null,
          campaignBudgetDeliveryMethod: 'STANDARD',
          campaignBudgetType: 'DAILY',
          campaignBudgetIsShared: false,
        },
      ],
    ]);

    const byBudgetMap = new Map<string, SourceBudgetCampaignInfo[]>([
      [
        excessBudgetRes,
        [
          {
            campaignResourceName: 'customers/100/campaigns/3',
            campaignName: 'Source Search Camp',
            campaignType: 'SEARCH',
            merchantId: null,
            targetCpa: 18.5,
            targetRoas: 4.2,
          },
        ],
      ],
    ]);

    const recDetailsMap = new Map<string, GaqlRecommendationRow>([
      [
        recRes,
        {
          resourceName: recRes,
          type: 'MOVE_UNUSED_BUDGET',
          campaign: destCampRes,
          moveUnusedBudgetRecommendation: {
            excessCampaignBudget: excessBudgetRes,
            budgetRecommendation: {
              currentBudgetAmountMicros: 100_000_000,
              recommendedBudgetAmountMicros: 140_000_000,
            },
          },
          impact: {
            baseMetrics: { costMicros: 700_000_000, conversions: 20, conversionsValue: 2000 },
            potentialMetrics: { costMicros: 980_000_000, conversions: 28, conversionsValue: 2800 },
          },
        },
      ],
    ]);

    // 1. With EC Remedy = true, non-feed PMax dest + SEARCH source is RETAINED
    const retainedRows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: true,
      recommendationDetailsMap: recDetailsMap,
      campaignDataMap: campaignMap,
      budgetDataMap: budgetMap,
      budgetToCampaignsMap: byBudgetMap,
    });

    expect(retainedRows.length).toBe(1);
    expect(retainedRows[0].moveBudgetAmount).toBe(40);
    expect(retainedRows[0].moveBudgetSourceBudgetAmount).toBe(200);
    // Pre-FX default mirrors the native amount until applyCurrencyConversionToRow runs
    expect(retainedRows[0].moveBudgetAmountGlobal).toBe(40);
    expect(retainedRows[0].recommendationCurrentBudgetAmount).toBe(100);
    expect(retainedRows[0].recommendationNewBudgetAmount).toBe(140);
    expect(retainedRows[0].moveBudgetSourceCampaigns).toBe('Source Search Camp');
    expect(retainedRows[0].moveBudgetSourceTargetCpa).toBe(18.5);
    expect(retainedRows[0].moveBudgetSourceTargetRoas).toBe(4.2);

    // 2. If the source campaign is a feed-connected PMax (merchantId = '555'), EC Remedy excludes the move rec!
    byBudgetMap.set(excessBudgetRes, [
      {
        campaignResourceName: 'customers/100/campaigns/3',
        campaignName: 'Retail PMax Source',
        campaignType: 'PERFORMANCE_MAX',
        merchantId: '555',
        targetCpa: null,
        targetRoas: 3.5,
      },
    ]);

    const excludedRows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: true,
      recommendationDetailsMap: recDetailsMap,
      campaignDataMap: campaignMap,
      budgetDataMap: budgetMap,
      budgetToCampaignsMap: byBudgetMap,
    });
    expect(excludedRows.length).toBe(0);
  });

  it('consolidates portfolio recommendations into 1 row with summed campaign metrics and cost-weighted Lost IS', () => {
    const camp1 = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/10',
      campaignId: '10',
      campaignName: 'Port Camp A',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'EMEA tROAS Portfolio',
      portfolioStrategyResourceName: 'customers/100/biddingStrategies/777',
      campaign30DaysCost: 1000,
      campaign30DaysConversions: 50,
      campaign30DaysConversionsValue: 3000,
      campaign30DaysSearchRankLostImpressionShare: 0.1,
      campaign30DaysSearchBudgetLostImpressionShare: 0.2,
    });
    const camp2 = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/20',
      campaignId: '20',
      campaignName: 'Port Camp B',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'EMEA tROAS Portfolio',
      portfolioStrategyResourceName: 'customers/100/biddingStrategies/777',
      campaign30DaysCost: 3000,
      campaign30DaysConversions: 150,
      campaign30DaysConversionsValue: 9000,
      campaign30DaysSearchRankLostImpressionShare: 0.3,
      campaign30DaysSearchBudgetLostImpressionShare: 0.4,
    });

    const recRes = 'customers/100/recommendations/portRec1';
    const recDetailsMap = new Map<string, GaqlRecommendationRow>([
      [
        recRes,
        {
          resourceName: recRes,
          type: 'LOWER_TARGET_ROAS',
          campaigns: [camp1.campaignResourceName, camp2.campaignResourceName],
          lowerTargetRoasRecommendation: {
            targetRoasMultiplier: 0.9,
          },
          impact: {
            baseMetrics: { costMicros: 900_000_000, conversions: 45, conversionsValue: 2700 },
            potentialMetrics: { costMicros: 1_100_000_000, conversions: 55, conversionsValue: 3100 },
          },
        },
      ],
    ]);

    const rawRows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'EUR',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: recDetailsMap,
      campaignDataMap: new Map([
        [camp1.campaignResourceName, camp1],
        [camp2.campaignResourceName, camp2],
      ]),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });

    // Raw rows has 2 target campaign rows before consolidation
    expect(rawRows.length).toBe(2);

    const consolidated = consolidateRowsByScope(rawRows);
    expect(consolidated.length).toBe(1);

    const portRow = consolidated[0];
    expect(portRow.recommendationScope).toBe('PORTFOLIO');
    expect(portRow.campaignPortfolioBiddingStrategyName).toBe('EMEA tROAS Portfolio');
    expect(portRow.campaignName).toBe('EMEA tROAS Portfolio');
    expect(portRow.campaignId).toBe('777');
    expect(portRow.campaignUrl).toBe(
      'https://ads.google.com/aw/bidreport/portfolio?strategyId=777&ocid=1000000001'
    );
    expect(portRow.portfolioCampaignsCount).toBe(2);
    expect(portRow.portfolioMemberCampaignNames).toBe('Port Camp A, Port Camp B');

    // Summed 30d historical campaign metrics across the 2 portfolio campaigns
    expect(portRow.campaign30DaysCost).toBe(4000);
    expect(portRow.campaign30DaysConversions).toBe(200);
    expect(portRow.campaign30DaysConversionsValue).toBe(12000);
    expect(portRow.campaign30DaysCpa).toBe(20);
    expect(portRow.campaign30DaysRoas).toBe(3.0);

    // Cost-weighted Lost IS: (1000 * 0.1 + 3000 * 0.3) / 4000 = 1000 / 4000 = 0.25
    expect(portRow.campaign30DaysSearchRankLostImpressionShare).toBeCloseTo(0.25);

    // Recommendation forecast impact is NOT doubled!
    expect(portRow.recommendationBaseCost).toBe(900);
    expect(portRow.recommendationPotentialCost).toBe(1100);
    expect(portRow.weeklyCostIncrease).toBe(200);
  });

  it('ranks recommendationNumber (#1, #2) per entity with budget recommendations prioritized first', () => {
    const camp = makeCampaign({ campaignId: '555', campaignResourceName: 'customers/100/campaigns/555' });
    const recTarget = 'customers/100/recommendations/rTarget';
    const recBudget = 'customers/100/recommendations/rBudget';

    const rawRows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: new Map([
        [
          recTarget,
          {
            resourceName: recTarget,
            type: 'RAISE_TARGET_CPA',
            campaign: camp.campaignResourceName,
            raiseTargetCpaRecommendation: { targetCpaMultiplier: 1.1 },
            impact: {
              baseMetrics: { costMicros: 500_000_000, conversions: 25 },
              potentialMetrics: { costMicros: 600_000_000, conversions: 30 },
            },
          },
        ],
        [
          recBudget,
          {
            resourceName: recBudget,
            type: 'CAMPAIGN_BUDGET',
            campaign: camp.campaignResourceName,
            campaignBudgetRecommendation: {
              currentBudgetAmountMicros: 100_000_000,
              recommendedBudgetAmountMicros: 130_000_000,
            },
            impact: {
              baseMetrics: { costMicros: 500_000_000, conversions: 25 },
              potentialMetrics: { costMicros: 650_000_000, conversions: 32 },
            },
          },
        ],
      ]),
      campaignDataMap: new Map([[camp.campaignResourceName, camp]]),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });

    const ranked = assignRecommendationNumbers(consolidateRowsByScope(rawRows));
    expect(ranked.length).toBe(2);

    const budgetRow = ranked.find((r) => r.recommendationType === 'CAMPAIGN_BUDGET');
    const targetRow = ranked.find((r) => r.recommendationType === 'RAISE_TARGET_CPA');

    // Budget recommendation gets #1; target recommendation gets #2
    expect(budgetRow?.recommendationNumber).toBe(1);
    expect(targetRow?.recommendationNumber).toBe(2);
  });

  it('extracts RAISE_TARGET_CPA and LOWER_TARGET_ROAS from API-native targetAdjustment proto fields', () => {
    const campCpa = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/701',
      campaignId: '701',
      campaignBiddingTargetCpa: null, // Even when campaign-level target is null, currentAverageTargetMicros is used
    });
    const campRoas = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/702',
      campaignId: '702',
      campaignBiddingTargetRoas: 4.0,
    });

    const rows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: new Map([
        [
          'customers/100/recommendations/recCpa',
          {
            resourceName: 'customers/100/recommendations/recCpa',
            type: 'RAISE_TARGET_CPA',
            campaign: campCpa.campaignResourceName,
            raiseTargetCpaRecommendation: {
              targetAdjustment: {
                recommendedTargetMultiplier: 1.25,
                currentAverageTargetMicros: 20_000_000,
              },
            },
            impact: {
              baseMetrics: { costMicros: 500_000_000, conversions: 25 },
              potentialMetrics: { costMicros: 625_000_000, conversions: 30 },
            },
          },
        ],
        [
          'customers/100/recommendations/recRoas',
          {
            resourceName: 'customers/100/recommendations/recRoas',
            type: 'LOWER_TARGET_ROAS',
            campaign: campRoas.campaignResourceName,
            lowerTargetRoasRecommendation: {
              targetAdjustment: {
                recommendedTargetMultiplier: 0.8,
                currentAverageTargetMicros: 5_000_000,
              },
            },
            impact: {
              baseMetrics: { costMicros: 500_000_000, conversions: 25, conversionsValue: 2500 },
              potentialMetrics: { costMicros: 650_000_000, conversions: 32, conversionsValue: 2900 },
            },
          },
        ],
      ]),
      campaignDataMap: new Map([
        [campCpa.campaignResourceName, campCpa],
        [campRoas.campaignResourceName, campRoas],
      ]),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });

    expect(rows.length).toBe(2);
    const cpaRow = rows.find((r) => r.recommendationType === 'RAISE_TARGET_CPA')!;
    const roasRow = rows.find((r) => r.recommendationType === 'LOWER_TARGET_ROAS')!;

    expect(cpaRow.recommendationNewTargetCpa).toBeCloseTo(25); // 20 * 1.25
    expect(roasRow.recommendationNewTargetRoas).toBeCloseTo(4.0); // 5.0 * 0.8
  });

  it('rejects multi-campaign recs with filtered Shopping/Feed-PMax members and MOVE_UNUSED_BUDGET with Shopping donors under EC Remedy', () => {
    const searchCamp = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/801',
      campaignId: '801',
      campaignType: 'SEARCH',
    });
    const shoppingCampRes = 'customers/100/campaigns/802'; // Filtered out by GAQL != 'SHOPPING'
    const excessShoppingBudgetRes = 'customers/100/campaignBudgets/899';

    const rows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: true,
      recommendationDetailsMap: new Map([
        [
          'customers/100/recommendations/mixedPortfolioRec',
          {
            resourceName: 'customers/100/recommendations/mixedPortfolioRec',
            type: 'RAISE_TARGET_CPA',
            // Contains both a Search campaign and a Shopping campaign that was filtered out of campaignDataMap
            campaigns: [searchCamp.campaignResourceName, shoppingCampRes],
            raiseTargetCpaRecommendation: {
              targetAdjustment: { recommendedTargetMultiplier: 1.1 },
            },
          },
        ],
        [
          'customers/100/recommendations/moveFromShoppingDonor',
          {
            resourceName: 'customers/100/recommendations/moveFromShoppingDonor',
            type: 'MOVE_UNUSED_BUDGET',
            campaign: searchCamp.campaignResourceName,
            moveUnusedBudgetRecommendation: {
              excessCampaignBudget: excessShoppingBudgetRes,
              budgetRecommendation: {
                currentBudgetAmountMicros: 100_000_000,
                recommendedBudgetAmountMicros: 150_000_000,
              },
            },
          },
        ],
      ]),
      // Only searchCamp returned by GAQL; shoppingCampRes and excessShoppingBudgetRes's donor campaign were filtered out by != 'SHOPPING'
      campaignDataMap: new Map([[searchCamp.campaignResourceName, searchCamp]]),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });

    expect(rows.length).toBe(0);
  });

  it('keys SHARED_BUDGET by budgetResourceName and ranks budget recs on portfolio campaigns by campaign', () => {
    const campShared1 = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/901',
      campaignId: '901',
      campaignBudgetResourceName: 'customers/100/campaignBudgets/b1',
      campaignBudgetName: 'Shared Search',
      campaignBudgetIsShared: true,
    });
    const campShared2 = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/902',
      campaignId: '902',
      campaignBudgetResourceName: 'customers/100/campaignBudgets/b2',
      campaignBudgetName: 'Shared Search', // Identical display name, distinct budgetResourceName
      campaignBudgetIsShared: true,
    });

    const portCamp = makeCampaign({
      campaignResourceName: 'customers/100/campaigns/903',
      campaignId: '903',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'Core Portfolio',
      portfolioStrategyResourceName: 'customers/100/biddingStrategies/p1',
    });

    const rawRows = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: '100-000-0001',
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: new Map([
        [
          'customers/100/recommendations/sharedRec',
          {
            resourceName: 'customers/100/recommendations/sharedRec',
            type: 'CAMPAIGN_BUDGET',
            campaigns: [campShared1.campaignResourceName, campShared2.campaignResourceName],
            campaignBudgetRecommendation: {
              currentBudgetAmountMicros: 100_000_000,
              recommendedBudgetAmountMicros: 150_000_000,
            },
          },
        ],
        [
          'customers/100/recommendations/portTargetRec',
          {
            resourceName: 'customers/100/recommendations/portTargetRec',
            type: 'RAISE_TARGET_CPA',
            campaigns: [portCamp.campaignResourceName],
            raiseTargetCpaRecommendation: {
              targetAdjustment: { recommendedTargetMultiplier: 1.1 },
            },
          },
        ],
        [
          'customers/100/recommendations/portCampBudgetRec',
          {
            resourceName: 'customers/100/recommendations/portCampBudgetRec',
            type: 'CAMPAIGN_BUDGET',
            campaign: portCamp.campaignResourceName,
            campaignBudgetRecommendation: {
              currentBudgetAmountMicros: 100_000_000,
              recommendedBudgetAmountMicros: 140_000_000,
            },
          },
        ],
      ]),
      campaignDataMap: new Map([
        [campShared1.campaignResourceName, campShared1],
        [campShared2.campaignResourceName, campShared2],
        [portCamp.campaignResourceName, portCamp],
      ]),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });

    const consolidated = consolidateRowsByScope(rawRows);
    const ranked = assignRecommendationNumbers(consolidated);

    // The 2 shared budgets with distinct budgetResourceNames remain 2 separate SHARED_BUDGET rows
    const sharedRows = ranked.filter((r) => r.recommendationScope === 'SHARED_BUDGET');
    expect(sharedRows.length).toBe(2);
    expect(sharedRows.every((r) => r.recommendationNumber === 1)).toBe(true);
    expect(sharedRows.every((r) => r.campaignName === 'Shared Search')).toBe(true);
    expect(sharedRows.map((r) => r.campaignId)).toEqual(['b1', 'b2']);
    expect(sharedRows[0].campaignUrl).toBe(
      'https://ads.google.com/aw/budget/report?budgetId=b1&ocid=1000000001'
    );

    // A budget rec on a portfolio campaign is ranked with its CAMPAIGN; only the PORTFOLIO row uses the portfolio key
    const portRelated = ranked.filter(
      (r) => r.portfolioStrategyResourceName === 'customers/100/biddingStrategies/p1'
    );
    expect(
      portRelated.map((r) => `${r.recommendationScope}|${r.recommendationType}|${r.recommendationNumber}`).sort()
    ).toEqual(['CAMPAIGN|CAMPAIGN_BUDGET|1', 'PORTFOLIO|RAISE_TARGET_CPA|1']);
  });
});

describe('Portfolio keys, budget dedupe and donor targets', () => {
  const OWNER = '999-888-7777';

  /** Rows for one cross-account portfolio LOWER_TARGET_ROAS rec as seen from child account `childCid`. */
  function portfolioRowsForChild(childCid: string, currencyCode: string, campaignIds: string[]): ReportRow[] {
    const child = childCid.replace(/-/g, '');
    const campaigns = campaignIds.map((id) =>
      makeCampaign({
        campaignResourceName: `customers/${child}/campaigns/${id}`,
        campaignId: id,
        campaignName: `Camp ${id}`,
        campaignBudgetResourceName: `customers/${child}/campaignBudgets/${id}`,
        campaignIsPortfolioBiddingStrategy: true,
        campaignPortfolioBiddingStrategyName: 'MCC tROAS Portfolio',
        // GAQL prefixes the strategy with the requesting child CID, not the owner
        portfolioStrategyResourceName: `customers/${child}/biddingStrategies/777`,
        portfolioOwnerId: OWNER.replace(/-/g, ''),
        campaignPortfolioIsManagerOwned: true,
        campaign30DaysCost: 1000,
        campaignStats7DaysCost: 250,
        campaignStatsYesterdayCost: 40,
      })
    );
    const recRes = `customers/${child}/recommendations/sharedRec42`;
    return buildAccountRecommendationRows({
      mccId: OWNER,
      mccName: 'MCC',
      accountId: childCid,
      accountName: `Child ${childCid}`,
      currencyCode,
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: new Map<string, GaqlRecommendationRow>([
        [
          recRes,
          {
            resourceName: recRes,
            type: 'LOWER_TARGET_ROAS',
            campaigns: campaigns.map((c) => c.campaignResourceName),
            lowerTargetRoasRecommendation: {targetRoasMultiplier: 0.9},
            // Google Ads puts the full portfolio-wide forecast on every child-account copy
            impact: {
              baseMetrics: {costMicros: 900_000_000, conversions: 45, conversionsValue: 2700},
              potentialMetrics: {costMicros: 1_100_000_000, conversions: 55, conversionsValue: 3100},
            },
          },
        ],
      ]),
      campaignDataMap: new Map(campaigns.map((c) => [c.campaignResourceName, c])),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });
  }

  it('merges the copies of one cross-account portfolio recommendation into a single PORTFOLIO row', () => {
    const rows = [
      ...portfolioRowsForChild('100-000-0001', 'USD', ['11']),
      ...portfolioRowsForChild('100-000-0002', 'USD', ['21', '22']),
    ];
    const consolidated = consolidateRowsByScope(rows);
    expect(consolidated.length).toBe(1);
    const row = consolidated[0];
    expect(row.recommendationScope).toBe('PORTFOLIO');
    expect(row.campaignId).toBe('777');
    expect(row.campaignUrl).toBe('https://ads.google.com/aw/bidreport/portfolio?strategyId=777&ocid=9998887777');
    expect(row.portfolioCampaignsCount).toBe(3);
    // Forecast impact is counted once, not once per child account
    expect(row.weeklyCostIncrease).toBe(200);
    expect(row.newWeeklyConversions).toBe(10);
    // Historical spend is summed across all member campaigns in both child accounts
    expect(row.campaign30DaysCost).toBe(3000);
    expect(row.campaignStats7DaysCost).toBe(750);
    expect(row.campaignStatsYesterdayCost).toBe(120);
  });

  it('keeps cross-account portfolio copies in different currencies as separate rows', () => {
    const rows = [
      ...portfolioRowsForChild('100-000-0001', 'USD', ['11']),
      ...portfolioRowsForChild('100-000-0002', 'EUR', ['21']),
    ];
    const consolidated = consolidateRowsByScope(rows);
    expect(consolidated.length).toBe(2);
    expect(consolidated.map((r) => r.currencyCode).sort()).toEqual(['EUR', 'USD']);
  });

  it('keeps one row per child account in SPLIT_BY_ACCOUNT mode', () => {
    const rows = [
      ...portfolioRowsForChild('100-000-0001', 'USD', ['11']),
      ...portfolioRowsForChild('100-000-0002', 'USD', ['21']),
    ];
    expect(consolidateRowsByScope(rows, {crossAccountPortfolioMode: 'SPLIT_BY_ACCOUNT'}).length).toBe(2);
  });

  it('sums a budget shared by several portfolio member campaigns only once', () => {
    const sharedBudget = 'customers/1000000001/campaignBudgets/55';
    const rows = portfolioRowsForChild('100-000-0001', 'USD', ['1', '2', '3', '4', '5']).map((r) => ({
      ...r,
      budgetResourceName: sharedBudget,
      campaignBudgetAmount: 100,
    }));
    const [row] = consolidateRowsByScope(rows);
    expect(row.portfolioCampaignsCount).toBe(5);
    expect(row.campaignBudgetAmount).toBe(100);
    expect(row.dailyBudget30Days).toBe(3000);
    expect(row.dailyBudgetDelta).toBe(3000 - 5000);
    expect(row.campaignPercentageOfBudgetUsedYesterday).toBeCloseTo(200 / 100);
  });

  it('numbers rows from two child accounts on the same MCC-owned portfolio in one sequence', () => {
    const makeRow = (child: string, type: string, recId: string): ReportRow => ({
      ...createEmptyReportRow(),
      accountId: `${child.slice(0, 3)}-${child.slice(3, 6)}-${child.slice(6)}`,
      recommendationScope: 'PORTFOLIO',
      recommendationType: type,
      recommendationId: `customers/${child}/recommendations/${recId}`,
      campaignIsPortfolioBiddingStrategy: true,
      portfolioStrategyResourceName: `customers/${child}/biddingStrategies/777`,
      portfolioOwnerId: '9998887777',
    });
    const ranked = assignRecommendationNumbers([
      makeRow('1000000001', 'RAISE_TARGET_CPA', 'a'),
      makeRow('1000000002', 'LOWER_TARGET_ROAS', 'b'),
    ]);
    expect(ranked.map((r) => [r.recommendationType, r.recommendationNumber])).toEqual([
      ['LOWER_TARGET_ROAS', 1],
      ['RAISE_TARGET_CPA', 2],
    ]);
  });

  it('takes MOVE_UNUSED_BUDGET donor targets from the donor portfolio strategy (tCPA and tROAS)', () => {
    const destCampRes = 'customers/100/campaigns/2';
    const excessBudgetRes = 'customers/100/campaignBudgets/99';
    const recRes = 'customers/100/recommendations/move1';

    const runWithDonorStrategy = (strategy: Record<string, unknown>) => {
      installAppsScriptGlobals({
        adsApp: createMockAdsApp({
          searchHandler: (query: string) => [
            {
              campaign: {
                resourceName: 'customers/100/campaigns/3',
                name: 'Portfolio Donor',
                advertisingChannelType: 'SEARCH',
                ...(query.includes('campaign.bidding_strategy')
                  ? {biddingStrategy: 'customers/100/biddingStrategies/5'}
                  : {}),
              },
              accessibleBiddingStrategy: strategy,
              campaignBudget: {resourceName: excessBudgetRes},
            },
          ],
        }),
      });
      const [row] = buildAccountRecommendationRows({
        mccId: '111-222-3333',
        mccName: 'EMEA MCC',
        accountId: '100-000-0001',
        accountName: 'Acme Search',
        currencyCode: 'USD',
        timestamp: '2026-10-14 12:00:00',
        excludeShoppingAndPmax: false,
        recommendationDetailsMap: new Map<string, GaqlRecommendationRow>([
          [
            recRes,
            {
              resourceName: recRes,
              type: 'MOVE_UNUSED_BUDGET',
              campaign: destCampRes,
              moveUnusedBudgetRecommendation: {
                excessCampaignBudget: excessBudgetRes,
                budgetRecommendation: {
                  currentBudgetAmountMicros: 100_000_000,
                  recommendedBudgetAmountMicros: 140_000_000,
                },
              },
            },
          ],
        ]),
        campaignDataMap: new Map([
          [destCampRes, makeCampaign({campaignResourceName: destCampRes, campaignId: '2'})],
        ]),
        budgetDataMap: new Map(),
        budgetToCampaignsMap: fetchCampaignsByBudgets([excessBudgetRes]),
      });
      return row;
    };

    const tcpaRow = runWithDonorStrategy({targetCpa: {targetCpaMicros: 22_000_000}});
    expect(tcpaRow.moveBudgetSourceTargetCpa).toBe(22);
    expect(tcpaRow.moveBudgetSourceTargetRoas).toBeNull();

    const troasRow = runWithDonorStrategy({targetRoas: {targetRoas: 5.5}});
    expect(troasRow.moveBudgetSourceTargetCpa).toBeNull();
    expect(troasRow.moveBudgetSourceTargetRoas).toBe(5.5);
  });
});

describe('Google Ads UI deep links (recommendationsDetailsUrl & campaignUrl)', () => {
  // Recommendation IDs end in base64(<ocid>-<recoTypeId>-<suffix>). Each test notes what its ID decodes to.

  /** Joins one CAMPAIGN_BUDGET recommendation on one campaign through the Ads Script and joiner. */
  function joinCampaignBudgetRec(args: {recommendationId: string; accountId: string; campaignId: string}): ReportRow {
    const campRes = `customers/1000000001/campaigns/${args.campaignId}`;
    const [row] = buildAccountRecommendationRows({
      mccId: '111-222-3333',
      mccName: 'EMEA MCC',
      accountId: args.accountId,
      accountName: 'Acme Search',
      currencyCode: 'USD',
      timestamp: '2026-10-14 12:00:00',
      excludeShoppingAndPmax: false,
      recommendationDetailsMap: new Map<string, GaqlRecommendationRow>([
        [
          args.recommendationId,
          {
            resourceName: args.recommendationId,
            type: 'CAMPAIGN_BUDGET',
            campaign: campRes,
            campaignBudgetRecommendation: {
              currentBudgetAmountMicros: 100_000_000,
              recommendedBudgetAmountMicros: 150_000_000,
            },
          },
        ],
      ]),
      campaignDataMap: new Map([
        [campRes, makeCampaign({campaignResourceName: campRes, campaignId: args.campaignId})],
      ]),
      budgetDataMap: new Map(),
      budgetToCampaignsMap: new Map(),
    });
    return row;
  }

  /** A report row with only the fields consolidation needs; every test sets the fields it depends on. */
  function memberRow(fields: Partial<ReportRow>): ReportRow {
    return {
      ...createEmptyReportRow(),
      currencyCode: 'USD',
      campaignBudgetIsShared: false,
      campaignIsPortfolioBiddingStrategy: false,
      ...fields,
    };
  }

  it('links to the recommendation type and the campaign with the decoded ocid', () => {
    const row = joinCampaignBudgetRec({
      // Decodes to 4455667788-13-1700000000
      recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xMy0xNzAwMDAwMDAw',
      accountId: '100-000-0001',
      campaignId: '11',
    });

    expect(row.recommendationsDetailsUrl).toBe('https://ads.google.com/aw/recommendations?ocid=4455667788&opp=13');
    expect(row.campaignUrl).toBe('https://ads.google.com/aw/overview?campaignId=11&ocid=4455667788');
  });

  it("falls back to the row's accountId (not the resource-name CID) when the ID cannot be decoded", () => {
    const row = joinCampaignBudgetRec({
      recommendationId: 'customers/9999999999/recommendations/rec1',
      accountId: '100-000-0001',
      campaignId: '11',
    });

    expect(row.recommendationsDetailsUrl).toBe('https://ads.google.com/aw/recommendations?ocid=1000000001');
    expect(row.campaignUrl).toBe('https://ads.google.com/aw/overview?campaignId=11&ocid=1000000001');
  });

  it('links a SHARED_BUDGET row to the budget report with the decoded ocid', () => {
    // Decodes to 4455667788-1-1700000000
    const recommendationId = 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xLTE3MDAwMDAwMDA=';
    const sharedBudget = {
      recommendationId,
      recommendationType: 'CAMPAIGN_BUDGET',
      accountId: '100-000-0001',
      campaignBudgetIsShared: true,
      budgetResourceName: 'customers/1000000001/campaignBudgets/55',
    };
    const rows = [memberRow({...sharedBudget, campaignId: '11'}), memberRow({...sharedBudget, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.recommendationScope).toBe('SHARED_BUDGET');
    expect(row.campaignUrl).toBe('https://ads.google.com/aw/budget/report?budgetId=55&ocid=4455667788');
  });

  it("keeps the members' recommendation-type link on a consolidated SHARED_BUDGET row", () => {
    const sharedBudget = {
      // Decodes to 4455667788-1-1700000000
      recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xLTE3MDAwMDAwMDA=',
      recommendationsDetailsUrl: 'https://ads.google.com/aw/recommendations?ocid=4455667788&opp=1',
      recommendationType: 'CAMPAIGN_BUDGET',
      accountId: '100-000-0001',
      campaignBudgetIsShared: true,
      budgetResourceName: 'customers/1000000001/campaignBudgets/55',
    };
    const rows = [memberRow({...sharedBudget, campaignId: '11'}), memberRow({...sharedBudget, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.recommendationsDetailsUrl).toBe('https://ads.google.com/aw/recommendations?ocid=4455667788&opp=1');
  });

  it('links a SHARED_BUDGET row without a budget ID to the shared budgets page with the decoded ocid', () => {
    const sharedBudget = {
      // Decodes to 4455667788-1-1700000000
      recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xLTE3MDAwMDAwMDA=',
      recommendationType: 'CAMPAIGN_BUDGET',
      accountId: '100-000-0001',
      campaignBudgetIsShared: true,
      budgetResourceName: '',
      campaignBudgetName: 'Shared Brand Budget',
    };
    const rows = [memberRow({...sharedBudget, campaignId: '11'}), memberRow({...sharedBudget, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.campaignUrl).toBe('https://ads.google.com/aw/sharedbudgets?ocid=4455667788');
  });

  it('falls back to the CID for a SHARED_BUDGET row when the ID cannot be decoded', () => {
    const sharedBudget = {
      recommendationId: '500004',
      recommendationType: 'MARGINAL_ROI_CAMPAIGN_BUDGET',
      accountId: '100-000-0001',
      campaignBudgetIsShared: true,
      budgetResourceName: 'customers/1000000001/campaignBudgets/55',
    };
    const rows = [memberRow({...sharedBudget, campaignId: '11'}), memberRow({...sharedBudget, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.campaignUrl).toBe('https://ads.google.com/aw/budget/report?budgetId=55&ocid=1000000001');
  });

  it('links a PORTFOLIO owned by the recommendation account to the strategy report with the decoded ocid', () => {
    const portfolio = {
      // Decodes to 4455667788-17-1700000000
      recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xNy0xNzAwMDAwMDAw',
      recommendationType: 'RAISE_TARGET_CPA',
      accountId: '100-000-0001',
      campaignIsPortfolioBiddingStrategy: true,
      portfolioStrategyResourceName: 'customers/1000000001/biddingStrategies/777',
      portfolioOwnerId: '1000000001',
    };
    const rows = [memberRow({...portfolio, campaignId: '11'}), memberRow({...portfolio, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.recommendationScope).toBe('PORTFOLIO');
    expect(row.campaignUrl).toBe('https://ads.google.com/aw/bidreport/portfolio?strategyId=777&ocid=4455667788');
  });

  it('treats a PORTFOLIO with an unknown owner as owned by the recommendation account', () => {
    const portfolio = {
      // Decodes to 4455667788-17-1700000000
      recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xNy0xNzAwMDAwMDAw',
      recommendationType: 'RAISE_TARGET_CPA',
      accountId: '100-000-0001',
      campaignIsPortfolioBiddingStrategy: true,
      portfolioStrategyResourceName: 'customers/1000000001/biddingStrategies/777',
      portfolioOwnerId: 'N/A',
    };
    const rows = [memberRow({...portfolio, campaignId: '11'}), memberRow({...portfolio, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.campaignUrl).toBe('https://ads.google.com/aw/bidreport/portfolio?strategyId=777&ocid=4455667788');
  });

  it('links a PORTFOLIO without a strategy ID to the bid strategies page with the decoded ocid', () => {
    const portfolio = {
      // Decodes to 4455667788-17-1700000000
      recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xNy0xNzAwMDAwMDAw',
      recommendationType: 'RAISE_TARGET_CPA',
      accountId: '100-000-0001',
      campaignIsPortfolioBiddingStrategy: true,
      portfolioStrategyResourceName: 'portfolio-without-a-strategy-id',
      portfolioOwnerId: '1000000001',
    };
    const rows = [memberRow({...portfolio, campaignId: '11'}), memberRow({...portfolio, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.campaignUrl).toBe('https://ads.google.com/aw/bidstrategies?ocid=4455667788');
  });

  it('falls back to the CID for a PORTFOLIO row when the ID cannot be decoded', () => {
    const portfolio = {
      recommendationId: '500007',
      recommendationType: 'LOWER_TARGET_ROAS',
      accountId: '100-000-0001',
      campaignIsPortfolioBiddingStrategy: true,
      portfolioStrategyResourceName: 'customers/1000000001/biddingStrategies/777',
      portfolioOwnerId: '1000000001',
    };
    const rows = [memberRow({...portfolio, campaignId: '11'}), memberRow({...portfolio, campaignId: '12'})];

    const [row] = consolidateRowsByScope(rows);

    expect(row.campaignUrl).toBe('https://ads.google.com/aw/bidreport/portfolio?strategyId=777&ocid=1000000001');
  });

  it('keeps the owner CID for an MCC-owned PORTFOLIO (known limitation: no owner ocid is available)', () => {
    const rows = [
      memberRow({
        // Decodes to 4455667788-17-1700000000: the child account's ocid, not the owner's
        recommendationId: 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xNy0xNzAwMDAwMDAw',
        recommendationType: 'RAISE_TARGET_CPA',
        accountId: '100-000-0001',
        campaignId: '11',
        campaignIsPortfolioBiddingStrategy: true,
        portfolioStrategyResourceName: 'customers/1000000001/biddingStrategies/777',
        portfolioOwnerId: '9998887777',
      }),
    ];

    const [row] = consolidateRowsByScope(rows);

    expect(row.campaignUrl).toBe('https://ads.google.com/aw/bidreport/portfolio?strategyId=777&ocid=9998887777');
  });
});
