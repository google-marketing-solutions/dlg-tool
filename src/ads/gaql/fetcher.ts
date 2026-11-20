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
 * @fileoverview Batch GAQL execution engine and raw Google Ads API response normalizer.
 * Executes queries in 500-ID chunks via `AdsApp.search()`, converts micros to currency units (`fromMicros`),
 * enforces the EC Remedy filter (`isCampaignExcludedByEcRemedy`: excludes `SHOPPING` and feed-connected
 * `PERFORMANCE_MAX` campaigns while retaining non-feed PMax), avoids zero-division on 30d CPA/ROAS (`null` when 0),
 * and extracts portfolio ownership (`isLocallyOwnedPortfolio` vs `campaignPortfolioIsManagerOwned`) and AI Max settings.
 * Bundled only into `dist/ads/Code.gs`.
 */

import {normalizeCid} from '../../shared/utils/cid';
import {
  buildBudgetDetailsQuery,
  buildCampaignCostQuery,
  buildCampaignMetadataQuery,
  buildCampaignMetrics30DaysQuery,
  buildCampaignsByBudgetsQuery,
  buildRecommendationDetailsQuery,
  chunkArray,
} from './queries';

/** Normalized campaign metadata, bidding targets, budget attributes, and historical performance metrics. */
export interface EnrichedCampaignData {
  campaignResourceName: string;
  campaignId: string;
  campaignName: string;
  campaignType: string;
  campaignSubType: string;
  merchantId: string | null;
  campaignIsAiMax: boolean;
  campaignAiMaxTextCustomizationEnabled: boolean;
  campaignAiMaxFinalUrlExpansionEnabled: boolean;
  campaignBudgetResourceName: string | null;
  campaignBudgetName: string;
  campaignBudgetAmount: number | null;
  campaignBudgetTotalAmount: number | null;
  campaignBudgetDeliveryMethod: string;
  campaignBudgetType: string;
  campaignBudgetIsShared: boolean;
  campaignIsPortfolioBiddingStrategy: boolean;
  campaignPortfolioBiddingStrategyName: string;
  portfolioStrategyResourceName: string;
  portfolioOwnerId: string;
  isLocallyOwnedPortfolio: boolean;
  campaignPortfolioIsManagerOwned: boolean;
  campaignBiddingStrategyType: string;
  campaignBiddingTargetCpa: number | null;
  campaignBiddingTargetRoas: number | null;
  campaignStatsYesterdayCost: number | null;
  campaignStats7DaysCost: number | null;
  campaign30DaysCost: number | null;
  campaign30DaysConversionsValue: number | null;
  campaign30DaysConversions: number | null;
  campaign30DaysRoas: number | null;
  campaign30DaysCpa: number | null;
  campaign30DaysSearchRankLostImpressionShare: number | null;
  campaign30DaysSearchBudgetLostImpressionShare: number | null;
}

/** Normalized budget metadata returned by `fetchAllBudgetDetails`. */
export interface BudgetDetails {
  budgetResourceName: string;
  campaignBudgetName: string;
  campaignBudgetAmount: number | null;
  campaignBudgetTotalAmount: number | null;
  campaignBudgetDeliveryMethod: string;
  campaignBudgetType: string;
  campaignBudgetIsShared: boolean;
}

/** Metadata and bidding targets for an enabled donor campaign attached to a `MOVE_UNUSED_BUDGET` source budget. */
export interface SourceBudgetCampaignInfo {
  campaignResourceName: string;
  campaignName: string;
  campaignType: string;
  merchantId: string | null;
  targetCpa: number | null;
  targetRoas: number | null;
}

/** Raw recommendation sub-message structure for budget recommendation types. */
export interface GaqlBudgetRecommendationProto {
  currentBudgetAmountMicros?: number | string | null;
  recommendedBudgetAmountMicros?: number | string | null;
  budgetOptions?: Array<{
    budgetAmountMicros?: number | string | null;
    impact?: GaqlImpactProto;
  }>;
}

/** Raw forecast metric bundle inside a Google Ads recommendation `impact` proto. */
export interface GaqlMetricsProto {
  costMicros?: number | string | null;
  clicks?: number | string | null;
  conversions?: number | string | null;
  conversionsValue?: number | string | null;
  impressions?: number | string | null;
  videoViews?: number | string | null;
}

/** Raw `impact` pair (`baseMetrics` and `potentialMetrics`) on a Google Ads recommendation. */
export interface GaqlImpactProto {
  baseMetrics?: GaqlMetricsProto;
  potentialMetrics?: GaqlMetricsProto;
}

/** Raw recommendation object returned inside a `FROM recommendation` GAQL row. */
export interface GaqlRecommendationRow {
  resourceName?: string;
  type?: string;
  campaign?: string;
  campaigns?: string[];
  campaignBudgetRecommendation?: GaqlBudgetRecommendationProto;
  forecastingCampaignBudgetRecommendation?: GaqlBudgetRecommendationProto;
  marginalRoiCampaignBudgetRecommendation?: GaqlBudgetRecommendationProto;
  moveUnusedBudgetRecommendation?: {
    excessCampaignBudget?: string;
    budgetRecommendation?: GaqlBudgetRecommendationProto;
  };
  raiseTargetCpaRecommendation?: {
    targetCpaMultiplier?: number | string | null;
    targetAdjustment?: {
      recommendedTargetMultiplier?: number | string | null;
      currentAverageTargetMicros?: number | string | null;
    };
  };
  lowerTargetRoasRecommendation?: {
    targetRoasMultiplier?: number | string | null;
    targetAdjustment?: {
      recommendedTargetMultiplier?: number | string | null;
      currentAverageTargetMicros?: number | string | null;
    };
  };
  impact?: GaqlImpactProto;
}

interface GaqlCampaignBudgetProto {
  resourceName?: string;
  name?: string;
  amountMicros?: number | string | null;
  totalAmountMicros?: number | string | null;
  deliveryMethod?: string;
  period?: string;
  explicitlyShared?: boolean;
}

interface GaqlAssetAutomationSettingProto {
  assetAutomationType?: string;
  assetAutomationStatus?: string;
}

interface GaqlCampaignProto {
  resourceName?: string;
  name?: string;
  status?: string;
  advertisingChannelType?: string;
  advertisingChannelSubType?: string;
  shoppingSetting?: { merchantId?: string | number | null };
  aiMaxSetting?: { enableAiMax?: boolean };
  assetAutomationSettings?:
    | GaqlAssetAutomationSettingProto[]
    | Record<string, GaqlAssetAutomationSettingProto>;
  biddingStrategyType?: string;
  biddingStrategy?: string;
  targetCpa?: { targetCpaMicros?: number | string | null };
  maximizeConversions?: { targetCpaMicros?: number | string | null };
  targetRoas?: { targetRoas?: number | null };
  maximizeConversionValue?: { targetRoas?: number | null };
}

interface GaqlAccessibleBiddingStrategyProto {
  name?: string;
  ownerCustomerId?: string | number | null;
  id?: string | number | null;
  type?: string;
  targetCpa?: { targetCpaMicros?: number | string | null };
  maximizeConversions?: { targetCpaMicros?: number | string | null };
  targetRoas?: { targetRoas?: number | null };
  maximizeConversionValue?: { targetRoas?: number | null };
}

/** Structural type of a row returned by `AdsApp.search()`. */
export interface GaqlSearchRow {
  customer?: { id?: string | number | null };
  recommendation?: GaqlRecommendationRow;
  campaign?: GaqlCampaignProto;
  campaignBudget?: GaqlCampaignBudgetProto;
  accessibleBiddingStrategy?: GaqlAccessibleBiddingStrategyProto;
  metrics?: {
    costMicros?: number | string | null;
    conversionsValue?: number | string | null;
    conversions?: number | string | null;
    searchRankLostImpressionShare?: number | null;
    searchBudgetLostImpressionShare?: number | null;
  };
}

/**
 * Converts a Google Ads API `micros` value (1/1,000,000 of a currency unit) into standard currency units,
 * returning `null` if the input is null, empty, or non-finite.
 */
export function fromMicros(micros: number | string | null | undefined): number | null {
  if (micros === null || micros === undefined || micros === '') {
    return null;
  }
  const num = Number(micros);
  return Number.isFinite(num) ? num / 1_000_000 : null;
}

/**
 * EC Remedy ("No Shopping") check:
 * When `excludeShoppingAndPmax === true`, excludes:
 *  - `SHOPPING` campaigns
 *  - `PERFORMANCE_MAX` campaigns ONLY when connected to a Merchant Center feed (`merchantId` is non-empty).
 * Non-feed `PERFORMANCE_MAX` campaigns are retained.
 */
export function isCampaignExcludedByEcRemedy(
  campaign: { campaignType: string; merchantId?: string | number | null },
  excludeShoppingAndPmax: boolean
): boolean {
  if (!excludeShoppingAndPmax) {
    return false;
  }
  const type = String(campaign.campaignType ?? '').toUpperCase();
  if (type === 'SHOPPING') {
    return true;
  }
  if (type === 'PERFORMANCE_MAX') {
    const merchantIdStr = String(campaign.merchantId ?? '').trim();
    return merchantIdStr.length > 0;
  }
  return false;
}

/**
 * Executes a GAQL query via `AdsApp.search(query)` and returns all result rows.
 */
export function executeGaqlSearch(query: string): GaqlSearchRow[] {
  // Omit explicit apiVersion so AdsApp.search automatically uses Google Ads Scripts' active DEFAULT_REPORT_VERSION
  const iterator = AdsApp.search(query);
  const rows: GaqlSearchRow[] = [];
  while (iterator.hasNext()) {
    rows.push(iterator.next() as unknown as GaqlSearchRow);
  }
  return rows;
}

/**
 * Fetches recommendation details for all supported recommendation types in the selected account.
 */
export function fetchAllRecommendationDetails(): Map<string, GaqlRecommendationRow> {
  const detailsMap = new Map<string, GaqlRecommendationRow>();
  const rows = executeGaqlSearch(buildRecommendationDetailsQuery());
  for (const row of rows) {
    if (row.recommendation?.resourceName) {
      detailsMap.set(row.recommendation.resourceName, row.recommendation);
    }
  }
  return detailsMap;
}

/**
 * Fetches enabled campaign metadata, 30d metrics, 7d cost, and yesterday cost in chunks of 500 IDs.
 * When `excludeShoppingAndPmax` is true, excludes `SHOPPING` in the GAQL WHERE clause and drops
 * feed-connected `PERFORMANCE_MAX` campaigns before querying metrics or returning data.
 */
export function fetchAllCampaignData(
  campaignResourceNames: readonly string[],
  excludeShoppingAndPmax = false
): Map<string, EnrichedCampaignData> {
  const campaignMap = new Map<string, EnrichedCampaignData>();
  if (campaignResourceNames.length === 0) {
    return campaignMap;
  }

  for (const batch of chunkArray(campaignResourceNames)) {
    // 1. Metadata query (filters to campaign.status = 'ENABLED' and != 'SHOPPING' when configured)
    const metaRows = executeGaqlSearch(buildCampaignMetadataQuery(batch, excludeShoppingAndPmax));
    for (const row of metaRows) {
      const c = row.campaign ?? {};
      const resName = String(c.resourceName ?? '');
      if (!resName) continue;

      const campaignType = c.advertisingChannelType ?? 'N/A';
      const merchantId = extractMerchantId(c);

      if (isCampaignExcludedByEcRemedy({ campaignType, merchantId }, excludeShoppingAndPmax)) {
        continue;
      }

      const campaignId = resName.split('/')[3] ?? 'N/A';
      const { targetCpa, targetRoas, isPortfolio, portfolioName, portfolioRes, ownerId, isLocal, isManager } =
        extractBiddingDetails(row);
      const { isAiMax, textOptIn, urlOptIn } = extractAiMaxDetails(c);
      const budget = extractBudgetDetails(row.campaignBudget);

      campaignMap.set(resName, {
        campaignResourceName: resName,
        campaignId,
        campaignName: c.name ?? 'N/A',
        campaignType,
        campaignSubType: c.advertisingChannelSubType ?? 'N/A',
        merchantId,
        campaignIsAiMax: isAiMax,
        campaignAiMaxTextCustomizationEnabled: textOptIn,
        campaignAiMaxFinalUrlExpansionEnabled: urlOptIn,
        campaignBudgetResourceName: row.campaignBudget?.resourceName ?? null,
        campaignBudgetName: budget.campaignBudgetName,
        campaignBudgetAmount: budget.campaignBudgetAmount,
        campaignBudgetTotalAmount: budget.campaignBudgetTotalAmount,
        campaignBudgetDeliveryMethod: budget.campaignBudgetDeliveryMethod,
        campaignBudgetType: budget.campaignBudgetType,
        campaignBudgetIsShared: budget.campaignBudgetIsShared,
        campaignIsPortfolioBiddingStrategy: isPortfolio,
        campaignPortfolioBiddingStrategyName: portfolioName,
        portfolioStrategyResourceName: portfolioRes,
        portfolioOwnerId: ownerId,
        isLocallyOwnedPortfolio: isLocal,
        campaignPortfolioIsManagerOwned: isManager,
        campaignBiddingStrategyType: c.biddingStrategyType ?? 'N/A',
        campaignBiddingTargetCpa: targetCpa,
        campaignBiddingTargetRoas: targetRoas,
        campaignStatsYesterdayCost: null,
        campaignStats7DaysCost: null,
        campaign30DaysCost: null,
        campaign30DaysConversionsValue: null,
        campaign30DaysConversions: null,
        campaign30DaysRoas: null,
        campaign30DaysCpa: null,
        campaign30DaysSearchRankLostImpressionShare: null,
        campaign30DaysSearchBudgetLostImpressionShare: null,
      });
    }

    const enabledBatchNames = batch.filter((name) => campaignMap.has(name));
    if (enabledBatchNames.length === 0) continue;

    // 2. 30-day metrics. CPA/ROAS are null (never 0 or Infinity) when the denominator is 0.
    for (const row of executeGaqlSearch(buildCampaignMetrics30DaysQuery(enabledBatchNames))) {
      const entry = campaignMap.get(row.campaign?.resourceName ?? '');
      if (!entry || !row.metrics) continue;
      const cost = fromMicros(row.metrics.costMicros ?? 0) ?? 0;
      const convValue = Number(row.metrics.conversionsValue ?? 0);
      const convs = Number(row.metrics.conversions ?? 0);

      entry.campaign30DaysCost = cost;
      entry.campaign30DaysConversionsValue = convValue;
      entry.campaign30DaysConversions = convs;
      entry.campaign30DaysRoas = cost > 0 ? convValue / cost : null;
      entry.campaign30DaysCpa = convs > 0 ? cost / convs : null;
      entry.campaign30DaysSearchRankLostImpressionShare =
        row.metrics.searchRankLostImpressionShare ?? null;
      entry.campaign30DaysSearchBudgetLostImpressionShare =
        row.metrics.searchBudgetLostImpressionShare ?? null;
    }

    // 3. Yesterday Cost
    for (const row of executeGaqlSearch(buildCampaignCostQuery(enabledBatchNames, 'DURING YESTERDAY'))) {
      const entry = campaignMap.get(row.campaign?.resourceName ?? '');
      if (entry && row.metrics) {
        entry.campaignStatsYesterdayCost = fromMicros(row.metrics.costMicros ?? 0);
      }
    }

    // 4. Last 7 Days Cost
    for (const row of executeGaqlSearch(buildCampaignCostQuery(enabledBatchNames, 'DURING LAST_7_DAYS'))) {
      const entry = campaignMap.get(row.campaign?.resourceName ?? '');
      if (entry && row.metrics) {
        entry.campaignStats7DaysCost = fromMicros(row.metrics.costMicros ?? 0);
      }
    }
  }

  return campaignMap;
}

/**
 * Fetches budget details in chunks of 500 IDs for `MOVE_UNUSED_BUDGET` source budgets.
 */
export function fetchAllBudgetDetails(
  budgetResourceNames: readonly string[]
): Map<string, BudgetDetails> {
  const budgetMap = new Map<string, BudgetDetails>();
  for (const batch of chunkArray(budgetResourceNames)) {
    for (const row of executeGaqlSearch(buildBudgetDetailsQuery(batch))) {
      const resName = row.campaignBudget?.resourceName;
      if (resName) {
        budgetMap.set(resName, extractBudgetDetails(row.campaignBudget));
      }
    }
  }
  return budgetMap;
}

/**
 * Fetches enabled donor campaigns attached to `MOVE_UNUSED_BUDGET` source budgets in chunks of 500 IDs.
 */
export function fetchCampaignsByBudgets(
  budgetResourceNames: readonly string[],
  excludeShoppingAndPmax = false
): Map<string, SourceBudgetCampaignInfo[]> {
  const byBudget = new Map<string, SourceBudgetCampaignInfo[]>();
  for (const batch of chunkArray(budgetResourceNames)) {
    for (const row of executeGaqlSearch(
      buildCampaignsByBudgetsQuery(batch, excludeShoppingAndPmax)
    )) {
      const budgetRes = row.campaignBudget?.resourceName;
      const c = row.campaign;
      if (!budgetRes || !c) continue;

      const campaignType = c.advertisingChannelType ?? 'N/A';
      const merchantId = extractMerchantId(c);

      if (isCampaignExcludedByEcRemedy({ campaignType, merchantId }, excludeShoppingAndPmax)) {
        continue;
      }

      const { targetCpa, targetRoas } = extractBiddingDetails(row);
      const list = byBudget.get(budgetRes) ?? [];
      list.push({
        campaignResourceName: c.resourceName ?? '',
        campaignName: c.name ?? 'N/A',
        campaignType,
        merchantId,
        targetCpa,
        targetRoas,
      });
      byBudget.set(budgetRes, list);
    }
  }
  return byBudget;
}

/** Merchant Center ID linked through `campaign.shopping_setting.merchant_id`, or `null` when there is no feed. */
function extractMerchantId(c: GaqlCampaignProto): string | null {
  const id = c.shoppingSetting?.merchantId;
  return id === undefined || id === null ? null : String(id);
}

function extractBudgetDetails(b: GaqlCampaignBudgetProto | undefined): BudgetDetails {
  if (!b) {
    return {
      budgetResourceName: '',
      campaignBudgetName: 'N/A',
      campaignBudgetAmount: null,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'N/A',
      campaignBudgetType: 'N/A',
      campaignBudgetIsShared: false,
    };
  }
  return {
    budgetResourceName: b.resourceName ?? '',
    campaignBudgetName: b.name || 'N/A',
    campaignBudgetAmount: fromMicros(b.amountMicros),
    campaignBudgetTotalAmount: fromMicros(b.totalAmountMicros),
    campaignBudgetDeliveryMethod: b.deliveryMethod || 'N/A',
    // Budget period (DAILY / CUSTOM_PERIOD); campaign_budget.type is not selected
    campaignBudgetType: b.period || 'N/A',
    campaignBudgetIsShared: Boolean(b.explicitlyShared),
  };
}

function extractBiddingDetails(row: GaqlSearchRow) {
  const c = row.campaign ?? {};
  const abs = row.accessibleBiddingStrategy;
  const isPortfolio = Boolean(c.biddingStrategy);
  const portfolioRes = c.biddingStrategy ?? 'N/A';
  const portfolioName = abs?.name ?? 'N/A';

  const ownerCid = normalizeCid(abs?.ownerCustomerId);
  const customerCid = normalizeCid(row.customer?.id);
  const isLocal = Boolean(isPortfolio && ownerCid && customerCid && ownerCid === customerCid);
  const isManager = Boolean(isPortfolio && ownerCid && customerCid && ownerCid !== customerCid);

  const rawCpaMicros = isPortfolio
    ? abs?.targetCpa?.targetCpaMicros ?? abs?.maximizeConversions?.targetCpaMicros
    : c.targetCpa?.targetCpaMicros ?? c.maximizeConversions?.targetCpaMicros;

  const rawRoas = isPortfolio
    ? abs?.targetRoas?.targetRoas ?? abs?.maximizeConversionValue?.targetRoas
    : c.targetRoas?.targetRoas ?? c.maximizeConversionValue?.targetRoas;

  return {
    isPortfolio,
    portfolioRes,
    portfolioName,
    ownerId: abs?.ownerCustomerId ? String(abs.ownerCustomerId) : 'N/A',
    isLocal,
    isManager,
    targetCpa: fromMicros(rawCpaMicros),
    targetRoas: typeof rawRoas === 'number' && Number.isFinite(rawRoas) ? rawRoas : null,
  };
}

function extractAiMaxDetails(campaign: GaqlCampaignProto) {
  const isAiMax = Boolean(campaign.aiMaxSetting?.enableAiMax);
  const rawSettings = campaign.assetAutomationSettings ?? [];
  const settings: GaqlAssetAutomationSettingProto[] = Array.isArray(rawSettings)
    ? rawSettings
    : typeof rawSettings === 'object' && rawSettings !== null
      ? Object.values(rawSettings)
      : [];

  let textOptIn = false;
  let urlOptIn = false;
  for (const s of settings) {
    if (s.assetAutomationType === 'TEXT_ASSET_AUTOMATION') {
      textOptIn = s.assetAutomationStatus === 'OPTED_IN';
    } else if (s.assetAutomationType === 'FINAL_URL_EXPANSION_TEXT_ASSET_AUTOMATION') {
      urlOptIn = s.assetAutomationStatus === 'OPTED_IN';
    }
  }
  return { isAiMax, textOptIn, urlOptIn };
}
