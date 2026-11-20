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
 * @fileoverview Google Ads Query Language (GAQL) query builders and 500-ID batch chunker (`chunkArray`).
 * Supports all 6 budget and bidding recommendation types (`CAMPAIGN_BUDGET`, `FORECASTING_CAMPAIGN_BUDGET`,
 * `MARGINAL_ROI_CAMPAIGN_BUDGET`, `MOVE_UNUSED_BUDGET`, `RAISE_TARGET_CPA`, `LOWER_TARGET_ROAS`),
 * enforces `campaign.status = 'ENABLED'`, blocks `SHOPPING` directly in the GAQL `WHERE` clause when
 * `excludeShoppingAndPmax` is enabled, and uses lightweight cost-only queries for 1d and 7d spend.
 * Bundled only into `dist/ads/Code.gs`.
 */

import type {RecommendationType} from '../../shared/types';

/** Maximum number of resource names included in a single GAQL `WHERE ... IN (...)` clause. */
export const GAQL_IN_CLAUSE_BATCH_SIZE = 500;

/** List of all 6 budget and bidding recommendation types queried by DLG v2. */
export const SUPPORTED_RECOMMENDATION_TYPES: RecommendationType[] = [
  'CAMPAIGN_BUDGET',
  'FORECASTING_CAMPAIGN_BUDGET',
  'MARGINAL_ROI_CAMPAIGN_BUDGET',
  'MOVE_UNUSED_BUDGET',
  'RAISE_TARGET_CPA',
  'LOWER_TARGET_ROAS',
];

/**
 * Splits an array into batches of at most `batchSize` items to avoid exceeding GAQL `IN (...)` limits.
 */
export function chunkArray<T>(
  items: readonly T[],
  batchSize = GAQL_IN_CLAUSE_BATCH_SIZE
): T[][] {
  if (items.length === 0) return [];
  const chunks: T[][] = [];
  for (let i = 0; i < items.length; i += batchSize) {
    chunks.push(items.slice(i, i + batchSize));
  }
  return chunks;
}

/**
 * Formats resource names for a GAQL `IN (...)` clause. Throws on an empty list, because `IN ()` is invalid
 * GAQL: builders fail fast instead of sending a broken query. (Callers batch with `chunkArray`, which never
 * yields an empty batch.)
 */
function formatInClause(resourceNames: readonly string[], builderName: string): string {
  if (resourceNames.length === 0) {
    throw new Error(`${builderName}: at least one resource name is required (got an empty list).`);
  }
  return resourceNames.map((name) => `'${name}'`).join(',');
}

/**
 * GAQL Query 1: Recommendation details for all 6 supported recommendation types.
 */
export function buildRecommendationDetailsQuery(): string {
  const typeList = SUPPORTED_RECOMMENDATION_TYPES.map((t) => `'${t}'`).join(',');
  return `
    SELECT
      recommendation.resource_name,
      recommendation.type,
      recommendation.campaign,
      recommendation.campaigns,
      recommendation.campaign_budget_recommendation,
      recommendation.forecasting_campaign_budget_recommendation,
      recommendation.marginal_roi_campaign_budget_recommendation,
      recommendation.move_unused_budget_recommendation,
      recommendation.raise_target_cpa_recommendation,
      recommendation.lower_target_roas_recommendation,
      recommendation.impact
    FROM recommendation
    WHERE recommendation.type IN (${typeList})
  `.trim();
}

/**
 * GAQL Query 2: Core campaign, bidding strategy, budget, and Merchant Center feed (`merchant_id`) metadata.
 * Filters strictly to `ENABLED` campaigns, and excludes `SHOPPING` directly in GAQL when
 * `excludeShoppingAndPmax` is true (regulatory requirement).
 */
export function buildCampaignMetadataQuery(
  campaignResourceNames: readonly string[],
  excludeShoppingAndPmax = false
): string {
  const shoppingFilter = excludeShoppingAndPmax
    ? "\n      AND campaign.advertising_channel_type != 'SHOPPING'"
    : '';
  return `
    SELECT
      customer.id,
      campaign.resource_name,
      campaign.name,
      campaign.status,
      campaign.advertising_channel_type,
      campaign.advertising_channel_sub_type,
      campaign.shopping_setting.merchant_id,
      campaign.ai_max_setting.enable_ai_max,
      campaign.asset_automation_settings,
      campaign.bidding_strategy_type,
      campaign.bidding_strategy,
      campaign.maximize_conversions.target_cpa_micros,
      campaign.maximize_conversion_value.target_roas,
      campaign.target_cpa.target_cpa_micros,
      campaign.target_roas.target_roas,
      campaign_budget.resource_name,
      campaign_budget.amount_micros,
      campaign_budget.total_amount_micros,
      campaign_budget.delivery_method,
      campaign_budget.name,
      campaign_budget.period,
      campaign_budget.explicitly_shared,
      accessible_bidding_strategy.name,
      accessible_bidding_strategy.owner_customer_id,
      accessible_bidding_strategy.id,
      accessible_bidding_strategy.maximize_conversions.target_cpa_micros,
      accessible_bidding_strategy.maximize_conversion_value.target_roas,
      accessible_bidding_strategy.target_cpa.target_cpa_micros,
      accessible_bidding_strategy.target_roas.target_roas,
      accessible_bidding_strategy.type
    FROM campaign
    WHERE campaign.resource_name IN (${formatInClause(campaignResourceNames, 'buildCampaignMetadataQuery')})
      AND campaign.status = 'ENABLED'${shoppingFilter}
  `.trim();
}

/**
 * GAQL Query 3: 30-day performance metrics (only fields used in V2).
 */
export function buildCampaignMetrics30DaysQuery(campaignResourceNames: readonly string[]): string {
  return `
    SELECT
      campaign.resource_name,
      metrics.cost_micros,
      metrics.conversions_value,
      metrics.conversions,
      metrics.search_rank_lost_impression_share,
      metrics.search_budget_lost_impression_share
    FROM campaign
    WHERE campaign.resource_name IN (${formatInClause(campaignResourceNames, 'buildCampaignMetrics30DaysQuery')})
      AND segments.date DURING LAST_30_DAYS
  `.trim();
}

/**
 * GAQL Queries 4 & 5: Lightweight cost-only query for `DURING YESTERDAY` and `DURING LAST_7_DAYS` (cost only, to keep these queries cheap).
 */
export function buildCampaignCostQuery(
  campaignResourceNames: readonly string[],
  dateCondition: 'DURING YESTERDAY' | 'DURING LAST_7_DAYS'
): string {
  return `
    SELECT
      campaign.resource_name,
      metrics.cost_micros
    FROM campaign
    WHERE campaign.resource_name IN (${formatInClause(campaignResourceNames, 'buildCampaignCostQuery')})
      AND segments.date ${dateCondition}
  `.trim();
}

/**
 * GAQL Query 6a: Budget details for `MOVE_UNUSED_BUDGET` source budgets.
 */
export function buildBudgetDetailsQuery(budgetResourceNames: readonly string[]): string {
  return `
    SELECT
      campaign_budget.resource_name,
      campaign_budget.amount_micros,
      campaign_budget.total_amount_micros,
      campaign_budget.name,
      campaign_budget.period,
      campaign_budget.delivery_method,
      campaign_budget.explicitly_shared
    FROM campaign_budget
    WHERE campaign_budget.resource_name IN (${formatInClause(budgetResourceNames, 'buildBudgetDetailsQuery')})
  `.trim();
}

/**
 * GAQL Query 6b: Enabled campaigns attached to `MOVE_UNUSED_BUDGET` source budgets,
 * including their tCPA / tROAS targets (shown as donor targets) and `shopping_setting.merchant_id`
 * (used by the EC Remedy filter).
 * Selects `campaign.bidding_strategy` so donor campaigns on a portfolio strategy read their targets from
 * `accessible_bidding_strategy`.
 */
export function buildCampaignsByBudgetsQuery(
  budgetResourceNames: readonly string[],
  excludeShoppingAndPmax = false
): string {
  const shoppingFilter = excludeShoppingAndPmax
    ? "\n      AND campaign.advertising_channel_type != 'SHOPPING'"
    : '';
  return `
    SELECT
      campaign.resource_name,
      campaign.name,
      campaign.status,
      campaign.advertising_channel_type,
      campaign.shopping_setting.merchant_id,
      campaign.bidding_strategy,
      campaign.maximize_conversions.target_cpa_micros,
      campaign.maximize_conversion_value.target_roas,
      campaign.target_cpa.target_cpa_micros,
      campaign.target_roas.target_roas,
      accessible_bidding_strategy.maximize_conversions.target_cpa_micros,
      accessible_bidding_strategy.maximize_conversion_value.target_roas,
      accessible_bidding_strategy.target_cpa.target_cpa_micros,
      accessible_bidding_strategy.target_roas.target_roas,
      campaign_budget.resource_name
    FROM campaign
    WHERE campaign_budget.resource_name IN (${formatInClause(budgetResourceNames, 'buildCampaignsByBudgetsQuery')})
      AND campaign.status = 'ENABLED'${shoppingFilter}
  `.trim();
}
