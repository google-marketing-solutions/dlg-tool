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
  fetchAllBudgetDetails,
  fetchAllCampaignData,
  fetchAllRecommendationDetails,
  fetchCampaignsByBudgets,
  isCampaignExcludedByEcRemedy,
} from '../src/ads/gaql/fetcher';
import {
  GAQL_IN_CLAUSE_BATCH_SIZE,
  SUPPORTED_RECOMMENDATION_TYPES,
  buildBudgetDetailsQuery,
  buildCampaignCostQuery,
  buildCampaignMetadataQuery,
  buildCampaignMetrics30DaysQuery,
  buildCampaignsByBudgetsQuery,
  buildRecommendationDetailsQuery,
  chunkArray,
} from '../src/ads/gaql/queries';
import {createMockAdsApp, installAppsScriptGlobals} from './mocks/google-apps-script.mock';

describe('GAQL Query Layer, 500-Chunking, Paused Filter & Feed-PMax EC Remedy', () => {
  it('supports all 6 recommendation types including FORECASTING and MARGINAL_ROI budget recs', () => {
    expect(SUPPORTED_RECOMMENDATION_TYPES).toEqual([
      'CAMPAIGN_BUDGET',
      'FORECASTING_CAMPAIGN_BUDGET',
      'MARGINAL_ROI_CAMPAIGN_BUDGET',
      'MOVE_UNUSED_BUDGET',
      'RAISE_TARGET_CPA',
      'LOWER_TARGET_ROAS',
    ]);

    const query = buildRecommendationDetailsQuery();
    expect(query).toContain('recommendation.forecasting_campaign_budget_recommendation');
    expect(query).toContain('recommendation.marginal_roi_campaign_budget_recommendation');
    expect(query).toContain(
      "WHERE recommendation.type IN ('CAMPAIGN_BUDGET','FORECASTING_CAMPAIGN_BUDGET','MARGINAL_ROI_CAMPAIGN_BUDGET','MOVE_UNUSED_BUDGET','RAISE_TARGET_CPA','LOWER_TARGET_ROAS')"
    );
  });

  it('filters out paused campaigns (status = ENABLED), adds != SHOPPING in GAQL when requested, and selects shopping_setting.merchant_id', () => {
    const campaignQuery = buildCampaignMetadataQuery(['customers/1/campaigns/10']);
    expect(campaignQuery).toContain("campaign.status = 'ENABLED'");
    expect(campaignQuery).toContain('campaign.shopping_setting.merchant_id');
    expect(campaignQuery).toContain('campaign_budget.explicitly_shared');
    expect(campaignQuery).not.toContain("campaign.advertising_channel_type != 'SHOPPING'");

    const noShoppingCampaignQuery = buildCampaignMetadataQuery(['customers/1/campaigns/10'], true);
    expect(noShoppingCampaignQuery).toContain("campaign.advertising_channel_type != 'SHOPPING'");

    const budgetCampaignsQuery = buildCampaignsByBudgetsQuery(['customers/1/campaignBudgets/20']);
    expect(budgetCampaignsQuery).toContain("campaign.status = 'ENABLED'");
    expect(budgetCampaignsQuery).toContain('campaign.shopping_setting.merchant_id');

    const noShoppingBudgetCampaignsQuery = buildCampaignsByBudgetsQuery(
      ['customers/1/campaignBudgets/20'],
      true
    );
    expect(noShoppingBudgetCampaignsQuery).toContain(
      "campaign.advertising_channel_type != 'SHOPPING'"
    );
  });

  it('queries only campaign.resource_name and metrics.cost_micros for YESTERDAY and LAST_7_DAYS', () => {
    const yesterdayQuery = buildCampaignCostQuery(['customers/1/campaigns/10'], 'DURING YESTERDAY');
    expect(yesterdayQuery).toContain('metrics.cost_micros');
    expect(yesterdayQuery).not.toContain('metrics.conversions');
    expect(yesterdayQuery).not.toContain('metrics.unique_users');
    expect(yesterdayQuery).not.toContain('metrics.trueview_average_cpv');
  });

  it('excludes SHOPPING and feed-connected PMax (merchant_id != null) while keeping non-feed PMax', () => {
    // When EC remedy is disabled, nothing is excluded
    expect(
      isCampaignExcludedByEcRemedy(
        { campaignType: 'SHOPPING', merchantId: '12345' },
        false
      )
    ).toBe(false);

    // When EC remedy is enabled:
    // 1. Standard SHOPPING is excluded
    expect(
      isCampaignExcludedByEcRemedy(
        { campaignType: 'SHOPPING', merchantId: null },
        true
      )
    ).toBe(true);

    // 2. PMax WITH Merchant Center feed is excluded
    expect(
      isCampaignExcludedByEcRemedy(
        { campaignType: 'PERFORMANCE_MAX', merchantId: '99887766' },
        true
      )
    ).toBe(true);

    // 3. PMax WITHOUT Merchant Center feed (merchantId null/empty) is RETAINED!
    expect(
      isCampaignExcludedByEcRemedy(
        { campaignType: 'PERFORMANCE_MAX', merchantId: null },
        true
      )
    ).toBe(false);
    expect(
      isCampaignExcludedByEcRemedy(
        { campaignType: 'PERFORMANCE_MAX', merchantId: '' },
        true
      )
    ).toBe(false);

    // 4. SEARCH / DEMAND_GEN are retained
    expect(
      isCampaignExcludedByEcRemedy(
        { campaignType: 'SEARCH', merchantId: null },
        true
      )
    ).toBe(false);
  });

  it('fetches recommendation details in a single GAQL query and chunks > 500 resource names for entity lookups', () => {
    const totalIds = 505;
    const budgetResourceNames = Array.from(
      { length: totalIds },
      (_, i) => `customers/123/campaignBudgets/${i + 1}`
    );

    const chunks = chunkArray(budgetResourceNames, GAQL_IN_CLAUSE_BATCH_SIZE);
    expect(chunks.length).toBe(2);
    expect(chunks[0].length).toBe(500);
    expect(chunks[1].length).toBe(5);

    const executedQueries: string[] = [];
    const adsApp = createMockAdsApp({
      searchHandler: (query: string) => {
        executedQueries.push(query);
        if (query.includes('FROM recommendation')) {
          return [
            {
              recommendation: {
                resourceName: 'customers/123/recommendations/1',
                type: 'CAMPAIGN_BUDGET',
              },
            },
          ];
        }
        const matches = Array.from(query.matchAll(/'(customers\/123\/campaignBudgets\/\d+)'/g));
        return matches.map((m) => ({
          campaignBudget: {
            resourceName: m[1],
            amountMicros: 10_000_000,
          },
        }));
      },
    });
    installAppsScriptGlobals({ adsApp });

    const detailsMap = fetchAllRecommendationDetails();
    expect(detailsMap.size).toBe(1);

    const budgetMap = fetchAllBudgetDetails(budgetResourceNames);
    expect(executedQueries.length).toBe(3);
    expect(budgetMap.size).toBe(505);
    expect(adsApp.searchOptionsHistory).toEqual([undefined, undefined, undefined]);
  });

  it('populates null (not 0) for 30d CPA when conversions=0 and for 30d ROAS when cost=0, and extracts source budget targets', () => {
    const campaignRes = 'customers/123/campaigns/111';
    const budgetRes = 'customers/123/campaignBudgets/222';

    const adsApp = createMockAdsApp({
      searchHandler: (query: string) => {
        if (query.includes('FROM campaign_budget')) {
          return [
            {
              campaignBudget: {
                resourceName: budgetRes,
                name: 'Shared Excess Budget',
                amountMicros: 50_000_000,
                period: 'DAILY',
                deliveryMethod: 'STANDARD',
                explicitlyShared: true,
              },
            },
          ];
        }
        if (query.includes('WHERE campaign_budget.resource_name IN')) {
          return [
            {
              campaign: {
                resourceName: 'customers/123/campaigns/999',
                name: 'Source Campaign A',
                advertisingChannelType: 'SEARCH',
                shoppingSetting: { merchantId: null },
                targetCpa: { targetCpaMicros: 15_000_000 },
                targetRoas: { targetRoas: 3.2 },
              },
              campaignBudget: { resourceName: budgetRes },
            },
          ];
        }
        if (query.includes('DURING LAST_30_DAYS')) {
          return [
            {
              campaign: { resourceName: campaignRes },
              metrics: {
                costMicros: 100_000_000,
                conversions: 0,
                conversionsValue: 0,
                searchRankLostImpressionShare: 0.12,
                searchBudgetLostImpressionShare: 0.34,
              },
            },
          ];
        }
        if (query.includes('DURING YESTERDAY')) {
          return [{ campaign: { resourceName: campaignRes }, metrics: { costMicros: 10_000_000 } }];
        }
        if (query.includes('DURING LAST_7_DAYS')) {
          return [{ campaign: { resourceName: campaignRes }, metrics: { costMicros: 70_000_000 } }];
        }
        // Main campaign metadata query
        return [
          {
            customer: { id: '1234567890' },
            campaign: {
              resourceName: campaignRes,
              name: 'Non-Feed PMax Campaign',
              advertisingChannelType: 'PERFORMANCE_MAX',
              advertisingChannelSubType: 'UNSPECIFIED',
              shoppingSetting: { merchantId: null },
              biddingStrategyType: 'MAXIMIZE_CONVERSION_VALUE',
              maximizeConversionValue: { targetRoas: 4.5 },
            },
            campaignBudget: {
              resourceName: budgetRes,
              name: 'Main Budget',
              amountMicros: 20_000_000,
              period: 'DAILY',
              deliveryMethod: 'STANDARD',
              explicitlyShared: false,
            },
          },
        ];
      },
    });
    installAppsScriptGlobals({ adsApp });

    const campaignMap = fetchAllCampaignData([campaignRes]);
    const cData = campaignMap.get(campaignRes);
    expect(cData).toBeDefined();
    expect(cData?.campaign30DaysCost).toBe(100);
    expect(cData?.campaign30DaysConversions).toBe(0);
    // Conversions = 0 -> CPA must be null, not 0!
    expect(cData?.campaign30DaysCpa).toBeNull();
    expect(cData?.campaignStatsYesterdayCost).toBe(10);
    expect(cData?.campaignStats7DaysCost).toBe(70);

    const budgetMap = fetchAllBudgetDetails([budgetRes]);
    expect(budgetMap.get(budgetRes)?.campaignBudgetAmount).toBe(50);

    const byBudgetMap = fetchCampaignsByBudgets([budgetRes]);
    const sourceCampaigns = byBudgetMap.get(budgetRes) ?? [];
    expect(sourceCampaigns.length).toBe(1);
    expect(sourceCampaigns[0].campaignName).toBe('Source Campaign A');
    expect(sourceCampaigns[0].targetCpa).toBe(15);
    expect(sourceCampaigns[0].targetRoas).toBe(3.2);
  });
});

describe('GAQL query builders', () => {
  it('fails fast when any resource-name query builder receives an empty list', () => {
    const builders: Array<[string, () => string]> = [
      ['buildCampaignMetadataQuery', () => buildCampaignMetadataQuery([])],
      ['buildCampaignMetrics30DaysQuery', () => buildCampaignMetrics30DaysQuery([])],
      ['buildCampaignCostQuery', () => buildCampaignCostQuery([], 'DURING YESTERDAY')],
      ['buildBudgetDetailsQuery', () => buildBudgetDetailsQuery([])],
      ['buildCampaignsByBudgetsQuery', () => buildCampaignsByBudgetsQuery([])],
    ];
    for (const [name, build] of builders) {
      expect(build, name).toThrow(`${name}: at least one resource name is required`);
    }
  });

  it('never yields an empty batch from chunkArray, so batched callers never hit the empty-list error', () => {
    expect(chunkArray([])).toEqual([]);
    for (const size of [1, 499, 500, 501, 1000, 1001]) {
      const batches = chunkArray(Array.from({length: size}, (_, i) => `customers/1/campaigns/${i}`));
      expect(batches.every((b) => b.length > 0)).toBe(true);
      expect(batches.reduce((n, b) => n + b.length, 0)).toBe(size);
    }
  });

  it('reads the budget period (not campaign_budget.type) as campaignBudgetType', () => {
    const metadataQuery = buildCampaignMetadataQuery(['customers/1/campaigns/10']);
    expect(metadataQuery).not.toContain('campaign_budget.type');
    expect(metadataQuery).toContain('campaign_budget.period');

    const adsApp = createMockAdsApp({
      searchHandler: () => [
        {campaignBudget: {resourceName: 'customers/1/campaignBudgets/1', period: 'CUSTOM_PERIOD'}},
        {campaignBudget: {resourceName: 'customers/1/campaignBudgets/2', type: 'STANDARD'}},
      ],
    });
    installAppsScriptGlobals({adsApp});
    const budgets = fetchAllBudgetDetails(['customers/1/campaignBudgets/1', 'customers/1/campaignBudgets/2']);
    expect(budgets.get('customers/1/campaignBudgets/1')?.campaignBudgetType).toBe('CUSTOM_PERIOD');
    expect(budgets.get('customers/1/campaignBudgets/2')?.campaignBudgetType).toBe('N/A');
  });

  it('reads donor targets from the portfolio strategy for MOVE_UNUSED_BUDGET donor campaigns', () => {
    const budgetRes = 'customers/123/campaignBudgets/222';
    expect(buildCampaignsByBudgetsQuery([budgetRes])).toContain('campaign.bidding_strategy');

    const adsApp = createMockAdsApp({
      // Like the API, only selected fields come back: bidding_strategy is present only if selected
      searchHandler: (query: string) => {
        const withStrategy = query.includes('campaign.bidding_strategy');
        return [
          {
            campaign: {
              resourceName: 'customers/123/campaigns/1',
              name: 'Portfolio tCPA Donor',
              advertisingChannelType: 'SEARCH',
              ...(withStrategy ? {biddingStrategy: 'customers/123/biddingStrategies/77'} : {}),
            },
            accessibleBiddingStrategy: {targetCpa: {targetCpaMicros: 22_000_000}},
            campaignBudget: {resourceName: budgetRes},
          },
          {
            campaign: {
              resourceName: 'customers/123/campaigns/2',
              name: 'Portfolio tROAS Donor',
              advertisingChannelType: 'SEARCH',
              ...(withStrategy ? {biddingStrategy: 'customers/123/biddingStrategies/88'} : {}),
            },
            accessibleBiddingStrategy: {maximizeConversionValue: {targetRoas: 5.5}},
            campaignBudget: {resourceName: budgetRes},
          },
        ];
      },
    });
    installAppsScriptGlobals({adsApp});

    const donors = fetchCampaignsByBudgets([budgetRes]).get(budgetRes) ?? [];
    expect(donors.map((d) => [d.campaignName, d.targetCpa, d.targetRoas])).toEqual([
      ['Portfolio tCPA Donor', 22, null],
      ['Portfolio tROAS Donor', null, 5.5],
    ]);
  });
});
