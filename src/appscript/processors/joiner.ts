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
 * @fileoverview In-memory relational joiner (`joinStagingBatchToReportRows`) and derived metric calculator.
 * Joins the 4 normalized staging tables (`_Stage_Accounts`, `_Stage_Budgets`, `_Stage_Campaigns`, `_Stage_Recs`)
 * by composite key (`mccId::accountId::resourceName`) into 82-column `ReportRow` objects and computes
 * derived ratios and deltas (`weeklyCostIncrease`, `newWeeklyConversions`, `dailyBudget30Days`, etc.).
 * Bundled only into `dist/sheet/Code.gs`.
 */

import {createEmptyReportRow} from '../../shared/schema';
import type {
  ReportRow,
  StagedAccount,
  StagedBudget,
  StagedCampaign,
  StagingBatch,
} from '../../shared/types';
import {normalizeCid} from '../../shared/utils/cid';
import {buildCampaignUrl, buildRecommendationsDetailsUrl} from './ads-ui-urls';

/**
 * Joins normalized staging tables (`accounts`, `budgets`, `campaigns`, `recommendations`)
 * into full 82-column `ReportRow` objects and computes all ratios/deltas in memory.
 * Runs inside `dist/sheet/Code.gs` when publishing `_Stage_*` to `Data`.
 */
export function joinStagingBatchToReportRows(batch: StagingBatch): ReportRow[] {
  const accountMap = new Map<string, StagedAccount>();
  for (const acc of batch.accounts) {
    accountMap.set(`${normalizeCid(acc.mccId)}::${normalizeCid(acc.accountId)}`, acc);
  }

  const budgetMap = new Map<string, StagedBudget>();
  for (const b of batch.budgets) {
    budgetMap.set(`${normalizeCid(b.mccId)}::${normalizeCid(b.accountId)}::${b.budgetResourceName}`, b);
  }

  const campaignMap = new Map<string, StagedCampaign>();
  for (const c of batch.campaigns) {
    campaignMap.set(`${normalizeCid(c.mccId)}::${normalizeCid(c.accountId)}::${c.campaignResourceName}`, c);
  }

  const rows: ReportRow[] = [];
  for (const rec of batch.recommendations) {
    const normMccCid = normalizeCid(rec.mccId);
    const normAccountCid = normalizeCid(rec.accountId);
    const account = accountMap.get(`${normMccCid}::${normAccountCid}`);
    const campaign = campaignMap.get(`${normMccCid}::${normAccountCid}::${rec.campaignResourceName}`);
    if (!campaign) continue;

    const budget = budgetMap.get(`${normMccCid}::${normAccountCid}::${campaign.budgetResourceName}`);
    const row = createEmptyReportRow();

    row.mccId = rec.mccId;
    row.mccName = account?.mccName ?? 'N/A';
    row.accountId = rec.accountId;
    row.accountName = account?.accountName ?? 'N/A';
    row.currencyCode = account?.currencyCode ?? 'USD';
    row.currencyCodeGlobal = row.currencyCode;
    row.timestamp = account?.timestamp ?? null;
    row.recommendationId = rec.recommendationId;
    row.recommendationType = rec.recommendationType;
    row.recommendationsDetailsUrl = buildRecommendationsDetailsUrl(rec.recommendationId, rec.accountId);

    // Campaign metadata (Google Ads UI links use the ocid decoded from the recommendation ID)
    row.campaignUrl = buildCampaignUrl(campaign.campaignId, rec.recommendationId, rec.accountId);
    row.campaignId = campaign.campaignId;
    row.campaignName = campaign.campaignName;
    row.campaignType = campaign.campaignType;
    row.campaignSubType = campaign.campaignSubType;
    row.campaignIsAiMax = campaign.campaignIsAiMax;
    row.campaignAiMaxTextCustomizationEnabled = campaign.campaignAiMaxTextCustomizationEnabled;
    row.campaignAiMaxFinalUrlExpansionEnabled = campaign.campaignAiMaxFinalUrlExpansionEnabled;

    // Campaign budget (joined from _Stage_Budgets)
    row.budgetResourceName = campaign.budgetResourceName;
    row.campaignBudgetName = budget?.campaignBudgetName ?? 'N/A';
    row.campaignBudgetAmount = budget?.campaignBudgetAmount ?? null;
    row.campaignBudgetTotalAmount = budget?.campaignBudgetTotalAmount ?? null;
    row.campaignBudgetDeliveryMethod = budget?.campaignBudgetDeliveryMethod ?? 'N/A';
    row.campaignBudgetType = budget?.campaignBudgetType ?? 'N/A';
    row.campaignBudgetIsShared = budget?.campaignBudgetIsShared ?? false;

    // Bidding & Portfolio metadata
    row.campaignIsPortfolioBiddingStrategy = campaign.campaignIsPortfolioBiddingStrategy;
    row.campaignPortfolioBiddingStrategyName = campaign.campaignPortfolioBiddingStrategyName;
    row.portfolioStrategyResourceName = campaign.portfolioStrategyResourceName;
    row.portfolioOwnerId = campaign.portfolioOwnerId;
    row.campaignBiddingStrategyType = campaign.campaignBiddingStrategyType;
    row.campaignBiddingTargetCpa = campaign.campaignBiddingTargetCpa;
    row.campaignBiddingTargetRoas = campaign.campaignBiddingTargetRoas;

    // Historical metrics & derived 30d CPA / ROAS
    row.campaignStatsYesterdayCost = campaign.campaignStatsYesterdayCost;
    row.campaignStats7DaysCost = campaign.campaignStats7DaysCost;
    row.campaign30DaysCost = campaign.campaign30DaysCost;
    row.campaign30DaysConversionsValue = campaign.campaign30DaysConversionsValue;
    row.campaign30DaysConversions = campaign.campaign30DaysConversions;
    const cost30d = campaign.campaign30DaysCost ?? 0;
    const conv30d = campaign.campaign30DaysConversions ?? 0;
    const convVal30d = campaign.campaign30DaysConversionsValue ?? 0;
    row.campaign30DaysRoas =
      campaign.campaign30DaysCost !== null && cost30d > 0 ? convVal30d / cost30d : null;
    row.campaign30DaysCpa =
      campaign.campaign30DaysConversions !== null && conv30d > 0 ? cost30d / conv30d : null;
    row.campaign30DaysSearchRankLostImpressionShare =
      campaign.campaign30DaysSearchRankLostImpressionShare;
    row.campaign30DaysSearchBudgetLostImpressionShare =
      campaign.campaign30DaysSearchBudgetLostImpressionShare;

    // Recommendation target & forecast impact metrics
    row.recommendationCurrentBudgetAmount = rec.recommendationCurrentBudgetAmount;
    row.recommendationCurrentBudgetAmountGlobal = rec.recommendationCurrentBudgetAmount;
    row.recommendationNewBudgetAmount = rec.recommendationNewBudgetAmount;
    row.recommendationNewBudgetAmountGlobal = rec.recommendationNewBudgetAmount;
    row.recommendationNewTargetCpa = rec.recommendationNewTargetCpa;
    row.recommendationNewTargetRoas = rec.recommendationNewTargetRoas;

    row.recommendationBaseCost = rec.recommendationBaseCost;
    row.recommendationPotentialCost = rec.recommendationPotentialCost;
    row.recommendationBaseClicks = rec.recommendationBaseClicks;
    row.recommendationPotentialClicks = rec.recommendationPotentialClicks;
    row.recommendationBaseConversions = rec.recommendationBaseConversions;
    row.recommendationPotentialConversions = rec.recommendationPotentialConversions;
    row.recommendationBaseConversionsValue = rec.recommendationBaseConversionsValue;
    row.recommendationBaseConversionsValueGlobal = rec.recommendationBaseConversionsValue;
    row.recommendationPotentialConversionsValue = rec.recommendationPotentialConversionsValue;
    row.recommendationPotentialConversionsValueGlobal = rec.recommendationPotentialConversionsValue;

    row.recommendationBaseCpa =
      rec.recommendationBaseCost !== null &&
      rec.recommendationBaseConversions !== null &&
      rec.recommendationBaseConversions > 0
        ? rec.recommendationBaseCost / rec.recommendationBaseConversions
        : null;
    row.recommendationPotentialCpa =
      rec.recommendationPotentialCost !== null &&
      rec.recommendationPotentialConversions !== null &&
      rec.recommendationPotentialConversions > 0
        ? rec.recommendationPotentialCost / rec.recommendationPotentialConversions
        : null;
    row.recommendationBaseRoas =
      rec.recommendationBaseCost !== null &&
      rec.recommendationBaseCost > 0 &&
      rec.recommendationBaseConversionsValue !== null
        ? rec.recommendationBaseConversionsValue / rec.recommendationBaseCost
        : null;
    row.recommendationPotentialRoas =
      rec.recommendationPotentialCost !== null &&
      rec.recommendationPotentialCost > 0 &&
      rec.recommendationPotentialConversionsValue !== null
        ? rec.recommendationPotentialConversionsValue / rec.recommendationPotentialCost
        : null;

    row.recommendationBaseImpressions = rec.recommendationBaseImpressions;
    row.recommendationPotentialImpressions = rec.recommendationPotentialImpressions;
    row.recommendationBaseVideoViews = rec.recommendationBaseVideoViews;
    row.recommendationPotentialVideoViews = rec.recommendationPotentialVideoViews;

    // Move Unused Budget source fields (joined from _Stage_Budgets)
    if (rec.recommendationType === 'MOVE_UNUSED_BUDGET') {
      if (
        rec.recommendationCurrentBudgetAmount !== null &&
        rec.recommendationNewBudgetAmount !== null
      ) {
        row.moveBudgetAmount =
          rec.recommendationNewBudgetAmount - rec.recommendationCurrentBudgetAmount;
      }
      const sourceBudget = rec.moveBudgetSourceBudgetResourceName
        ? budgetMap.get(
            `${normMccCid}::${normAccountCid}::${rec.moveBudgetSourceBudgetResourceName}`
          )
        : undefined;
      if (sourceBudget) {
        row.moveBudgetSourceBudgetName = sourceBudget.campaignBudgetName;
        row.moveBudgetSourceBudgetType = sourceBudget.campaignBudgetType;
        row.moveBudgetSourceBudgetDeliveryMethod = sourceBudget.campaignBudgetDeliveryMethod;
        row.moveBudgetSourceBudgetIsShared = sourceBudget.campaignBudgetIsShared;
        row.moveBudgetSourceBudgetAmount = sourceBudget.campaignBudgetAmount;
      }
      if (rec.moveBudgetSourceCampaigns && rec.moveBudgetSourceCampaigns !== 'N/A') {
        row.moveBudgetSourceCampaigns = rec.moveBudgetSourceCampaigns;
      }
      row.moveBudgetSourceTargetCpa = rec.moveBudgetSourceTargetCpa;
      row.moveBudgetSourceTargetRoas = rec.moveBudgetSourceTargetRoas;
    }

    populateDerivedRowMetrics(row);
    rows.push(row);
  }

  return rows;
}

/**
 * Populates derived budget utilization, weekly deltas, and CPA/ROAS deltas on a `ReportRow` in place.
 */
export function populateDerivedRowMetrics(row: ReportRow): void {
  const currentBudget = row.recommendationCurrentBudgetAmount ?? row.campaignBudgetAmount;
  if (
    typeof currentBudget === 'number' &&
    currentBudget > 0 &&
    typeof row.campaignStatsYesterdayCost === 'number'
  ) {
    row.campaignPercentageOfBudgetUsedYesterday = row.campaignStatsYesterdayCost / currentBudget;
  } else {
    row.campaignPercentageOfBudgetUsedYesterday = null;
  }

  row.weeklyCostIncrease = subtractNullable(row.recommendationPotentialCost, row.recommendationBaseCost);
  row.weeklyCostIncreaseGlobal = row.weeklyCostIncrease;
  // Pre-FX default; applyCurrencyConversionToRow overwrites it with the converted value.
  row.moveBudgetAmountGlobal = row.moveBudgetAmount;
  row.newWeeklyConversions = subtractNullable(
    row.recommendationPotentialConversions,
    row.recommendationBaseConversions
  );

  if (typeof row.campaignBudgetAmount === 'number') {
    row.dailyBudget30Days = row.campaignBudgetAmount * 30;
    if (typeof row.campaign30DaysCost === 'number') {
      row.dailyBudgetDelta = row.dailyBudget30Days - row.campaign30DaysCost;
    }
  }

  row.recommendationCpaDelta = subtractNullable(
    row.recommendationBaseCpa,
    row.recommendationPotentialCpa
  );
  row.recommendationRoasDelta = subtractNullable(
    row.recommendationBaseRoas,
    row.recommendationPotentialRoas
  );
}

function subtractNullable(a: number | null, b: number | null): number | null {
  if (typeof a === 'number' && typeof b === 'number') {
    return a - b;
  }
  return null;
}
