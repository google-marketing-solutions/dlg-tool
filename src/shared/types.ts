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
 * @fileoverview Core TypeScript type definitions and data contracts for DLG Dashboard v2.
 * Defines:
 * - `ReportRow` (the 82-column schema written to the `Data` tab for Data Studio),
 * - `StagedAccount`, `StagedBudget`, `StagedCampaign`, `StagedRecommendation` (the 4 normalized
 *   atomic staging tables written to per-MCC `_Stage_*` tabs),
 * - `MccIdentity` and `ChildAccountInput` (MCC and child account discovery payloads), and
 * - `GlobalConfig` (the user settings read from the `Config` tab and locked in `_RunState`).
 * Erased at compile time (zero runtime footprint in `dist/ads/Code.gs` or `dist/sheet/Code.gs`).
 */

/**
 * Scope classification for a row in the `Data` sheet (`CAMPAIGN`, `PORTFOLIO`, or `SHARED_BUDGET`).
 */
export type RecommendationScope = 'CAMPAIGN' | 'PORTFOLIO' | 'SHARED_BUDGET';

/**
 * Supported Google Ads budget and bidding recommendation types in DLG v2.
 */
export type RecommendationType =
  | 'CAMPAIGN_BUDGET'
  | 'FORECASTING_CAMPAIGN_BUDGET'
  | 'MARGINAL_ROI_CAMPAIGN_BUDGET'
  | 'MOVE_UNUSED_BUDGET'
  | 'RAISE_TARGET_CPA'
  | 'LOWER_TARGET_ROAS';

/** Display identity of a Google Ads Manager Account (MCC). */
export interface MccIdentity {
  mccId: string;
  mccName: string;
}

/** Discovered Google Ads child account metadata under an MCC. */
export interface ChildAccountInput {
  customerId: string;
  name: string;
  currencyCode: string;
}

/**
 * Specification for a single column in the `Data` sheet schema (`REPORT_SCHEMA_TEMPLATE`).
 */
export interface SchemaFieldSpec<T = string | number | boolean | null> {
  defaultValue: T;
  format: string;
  dummyValueForFormatting: string | number | boolean;
}

/**
 * Full in-memory representation of a recommendation row before and after consolidation,
 * including both the 82 published `Data` columns and transient internal join keys.
 */
export interface ReportRow {
  recommendationNumber: number | null;
  recommendationScope: RecommendationScope;
  mccId: string;
  mccName: string;
  mccSyncStatus: string;
  mccLastUpdatedDate: string | null;
  accountId: string;
  accountName: string;
  timestamp: string | null;
  recommendationId: string;
  recommendationType: string;
  recommendationsDetailsUrl: string;
  campaignUrl: string;
  campaignId: string;
  campaignName: string;
  campaignType: string;
  campaignSubType: string;
  campaignIsAiMax: boolean;
  campaignAiMaxTextCustomizationEnabled: boolean;
  campaignAiMaxFinalUrlExpansionEnabled: boolean;
  currencyCode: string;
  currencyCodeGlobal: string;
  campaignBudgetAmount: number | null;
  campaignBudgetTotalAmount: number | null;
  campaignBudgetDeliveryMethod: string;
  campaignBudgetType: string;
  portfolioCampaignsCount: number | null;
  portfolioMemberCampaignNames: string;
  campaignBiddingStrategyType: string;
  campaignBiddingTargetRoas: number | null;
  campaignBiddingTargetCpa: number | null;
  campaignStatsYesterdayCost: number | null;
  campaignStats7DaysCost: number | null;
  campaign30DaysCost: number | null;
  campaign30DaysConversionsValue: number | null;
  campaign30DaysConversions: number | null;
  campaign30DaysRoas: number | null;
  campaign30DaysCpa: number | null;
  campaign30DaysSearchRankLostImpressionShare: number | null;
  campaign30DaysSearchBudgetLostImpressionShare: number | null;
  recommendationCurrentBudgetAmount: number | null;
  recommendationCurrentBudgetAmountGlobal: number | null;
  recommendationNewBudgetAmount: number | null;
  recommendationNewBudgetAmountGlobal: number | null;
  recommendationNewTargetCpa: number | null;
  recommendationNewTargetRoas: number | null;
  recommendationBaseCost: number | null;
  recommendationPotentialCost: number | null;
  recommendationBaseClicks: number | null;
  recommendationPotentialClicks: number | null;
  recommendationBaseConversions: number | null;
  recommendationPotentialConversions: number | null;
  recommendationBaseConversionsValue: number | null;
  recommendationBaseConversionsValueGlobal: number | null;
  recommendationPotentialConversionsValue: number | null;
  recommendationPotentialConversionsValueGlobal: number | null;
  recommendationBaseCpa: number | null;
  recommendationPotentialCpa: number | null;
  recommendationBaseRoas: number | null;
  recommendationPotentialRoas: number | null;
  recommendationBaseImpressions: number | null;
  recommendationPotentialImpressions: number | null;
  recommendationBaseVideoViews: number | null;
  recommendationPotentialVideoViews: number | null;
  campaignPercentageOfBudgetUsedYesterday: number | null;
  weeklyCostIncrease: number | null;
  weeklyCostIncreaseGlobal: number | null;
  newWeeklyConversions: number | null;
  dailyBudget30Days: number | null;
  dailyBudgetDelta: number | null;
  recommendationCpaDelta: number | null;
  recommendationRoasDelta: number | null;
  moveBudgetAmount: number | null;
  moveBudgetSourceCampaigns: string;
  moveBudgetSourceBudgetName: string;
  moveBudgetSourceBudgetType: string;
  moveBudgetSourceBudgetDeliveryMethod: string;
  moveBudgetSourceBudgetIsShared: boolean;
  moveBudgetSourceBudgetAmount: number | null;
  moveBudgetSourceTargetCpa: number | null;
  moveBudgetSourceTargetRoas: number | null;
  moveBudgetAmountGlobal: number | null;
  /**
   * Non-schema internal identifiers carried in memory for deterministic shared-budget and
   * portfolio consolidation, URL construction, and ranking (excluded from the final `Data` tab).
   */
  budgetResourceName?: string;
  campaignBudgetName?: string;
  campaignBudgetIsShared?: boolean;
  campaignIsPortfolioBiddingStrategy?: boolean;
  campaignPortfolioBiddingStrategyName?: string;
  portfolioStrategyResourceName?: string;
  portfolioOwnerId?: string;
}

/**
 * Union of the 82 column keys written to the `Data` tab (excluding internal-only join fields).
 */
export type ReportColumnKey = Exclude<
  keyof ReportRow,
  | 'budgetResourceName'
  | 'campaignBudgetName'
  | 'campaignBudgetIsShared'
  | 'campaignIsPortfolioBiddingStrategy'
  | 'campaignPortfolioBiddingStrategyName'
  | 'portfolioStrategyResourceName'
  | 'portfolioOwnerId'
>;

/**
 * Normalized account metadata row written to `_Stage_Accounts_<mccId>`.
 */
export interface StagedAccount {
  mccId: string;
  mccName: string;
  accountId: string;
  accountName: string;
  currencyCode: string;
  timestamp: string;
}

/**
 * Normalized campaign budget metadata row written to `_Stage_Budgets_<mccId>`.
 */
export interface StagedBudget {
  mccId: string;
  accountId: string;
  budgetResourceName: string;
  campaignBudgetName: string;
  campaignBudgetAmount: number | null;
  campaignBudgetTotalAmount: number | null;
  campaignBudgetDeliveryMethod: string;
  campaignBudgetType: string;
  campaignBudgetIsShared: boolean;
}

/**
 * Normalized campaign metadata and historical performance row written to `_Stage_Campaigns_<mccId>`.
 */
export interface StagedCampaign {
  mccId: string;
  accountId: string;
  campaignResourceName: string;
  campaignId: string;
  campaignName: string;
  campaignType: string;
  campaignSubType: string;
  campaignIsAiMax: boolean;
  campaignAiMaxTextCustomizationEnabled: boolean;
  campaignAiMaxFinalUrlExpansionEnabled: boolean;
  budgetResourceName: string;
  campaignIsPortfolioBiddingStrategy: boolean;
  campaignPortfolioBiddingStrategyName: string;
  portfolioStrategyResourceName: string;
  portfolioOwnerId: string;
  isLocallyOwnedPortfolio: boolean;
  campaignPortfolioIsManagerOwned: boolean;
  campaignBiddingStrategyType: string;
  campaignBiddingTargetRoas: number | null;
  campaignBiddingTargetCpa: number | null;
  campaignStatsYesterdayCost: number | null;
  campaignStats7DaysCost: number | null;
  campaign30DaysCost: number | null;
  campaign30DaysConversionsValue: number | null;
  campaign30DaysConversions: number | null;
  campaign30DaysSearchRankLostImpressionShare: number | null;
  campaign30DaysSearchBudgetLostImpressionShare: number | null;
}

/**
 * Normalized recommendation and 7-day forecast impact row written to `_Stage_Recs_<mccId>`.
 */
export interface StagedRecommendation {
  mccId: string;
  accountId: string;
  recommendationId: string;
  recommendationType: string;
  campaignResourceName: string;
  recommendationCurrentBudgetAmount: number | null;
  recommendationNewBudgetAmount: number | null;
  recommendationNewTargetCpa: number | null;
  recommendationNewTargetRoas: number | null;
  recommendationBaseCost: number | null;
  recommendationPotentialCost: number | null;
  recommendationBaseClicks: number | null;
  recommendationPotentialClicks: number | null;
  recommendationBaseConversions: number | null;
  recommendationPotentialConversions: number | null;
  recommendationBaseConversionsValue: number | null;
  recommendationPotentialConversionsValue: number | null;
  recommendationBaseImpressions: number | null;
  recommendationPotentialImpressions: number | null;
  recommendationBaseVideoViews: number | null;
  recommendationPotentialVideoViews: number | null;
  moveBudgetSourceBudgetResourceName: string;
  moveBudgetSourceCampaigns: string;
  moveBudgetSourceTargetCpa: number | null;
  moveBudgetSourceTargetRoas: number | null;
}

/**
 * Container for the 4 normalized staging tables flushed per MCC checkpoint batch.
 */
export interface StagingBatch {
  accounts: StagedAccount[];
  budgets: StagedBudget[];
  campaigns: StagedCampaign[];
  recommendations: StagedRecommendation[];
}

/**
 * Configuration settings read from the `Config` tab and locked daily in `_RunState`
 * (`DAILY_CONFIG_LOCK_<date>`).
 */
export interface GlobalConfig {
  targetCurrency: string;
  excludeShoppingAndPmax: boolean;
  /**
   * Normalized 10-digit selected child account CIDs keyed by normalized 10-digit parent MCC ID.
   */
  selectedCidsByMcc?: Record<string, string[]>;
}
