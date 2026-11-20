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
 * @fileoverview Per-account recommendation extractor and normalized staging builder (`buildAccountStagingBatch`).
 * Deduplicates campaigns and budgets across recommendations for a single child account, applies the dual
 * EC Remedy check (on both target campaigns and `MOVE_UNUSED_BUDGET` donor campaigns), matches the recommended
 * budget option to extract strictly API-native 7-day forecast impact metrics, and resolves donor campaign tCPA/tROAS.
 * Bundled only into `dist/ads/Code.gs`.
 */

import type {
  StagedAccount,
  StagedBudget,
  StagedCampaign,
  StagedRecommendation,
  StagingBatch,
} from '../../shared/types';
import {
  type BudgetDetails,
  type EnrichedCampaignData,
  type GaqlBudgetRecommendationProto,
  type GaqlImpactProto,
  type GaqlRecommendationRow,
  type SourceBudgetCampaignInfo,
  fromMicros,
  isCampaignExcludedByEcRemedy,
} from '../gaql/fetcher';

/** Input context for building an account's normalized staging batch (`buildAccountStagingBatch`). */
export interface BuildAccountRowsContext {
  mccId: string;
  mccName: string;
  accountId: string;
  accountName: string;
  currencyCode: string;
  timestamp: string;
  excludeShoppingAndPmax: boolean;
  recommendationDetailsMap: Map<string, GaqlRecommendationRow>;
  campaignDataMap: Map<string, EnrichedCampaignData>;
  budgetDataMap: Map<string, BudgetDetails>;
  budgetToCampaignsMap: Map<string, SourceBudgetCampaignInfo[]>;
}

/**
 * Builds normalized staging entities (`StagedAccount`, `StagedBudget[]`, `StagedCampaign[]`, `StagedRecommendation[]`)
 * for a single child account inside `Code.gs`, deduplicating campaigns and budgets across multiple recommendations.
 */
export function buildAccountStagingBatch(
  ctx: BuildAccountRowsContext
): StagingBatch {
  const budgetsByKey = new Map<string, StagedBudget>();
  const campaignsByKey = new Map<string, StagedCampaign>();
  const stagedRecs: StagedRecommendation[] = [];

  for (const [recResourceName, recDetails] of ctx.recommendationDetailsMap.entries()) {
    const type = String(recDetails.type ?? '');
    const targetCampaignNames = extractTargetCampaignResourceNames(recDetails);
    if (targetCampaignNames.length === 0) continue;

    const targetCampaigns = targetCampaignNames
      .map((name) => ctx.campaignDataMap.get(name))
      .filter((c): c is EnrichedCampaignData => c !== undefined);

    if (targetCampaigns.length === 0) continue;

    // EC Remedy Check 1: If any target campaign was filtered out (or is excluded by EC Remedy), skip this recommendation
    if (
      ctx.excludeShoppingAndPmax &&
      targetCampaigns.length < targetCampaignNames.length
    ) {
      continue;
    }
    if (
      targetCampaigns.some((c) => isCampaignExcludedByEcRemedy(c, ctx.excludeShoppingAndPmax))
    ) {
      continue;
    }

    // EC Remedy Check 2: For MOVE_UNUSED_BUDGET, also verify source budget's enabled campaigns
    const excessBudgetRes = String(
      recDetails.moveUnusedBudgetRecommendation?.excessCampaignBudget ?? ''
    );
    const sourceCampaigns = excessBudgetRes
      ? ctx.budgetToCampaignsMap.get(excessBudgetRes) ?? []
      : [];
    if (type === 'MOVE_UNUSED_BUDGET') {
      if (ctx.excludeShoppingAndPmax && excessBudgetRes && sourceCampaigns.length === 0) {
        continue;
      }
      if (
        sourceCampaigns.some((sc) => isCampaignExcludedByEcRemedy(sc, ctx.excludeShoppingAndPmax))
      ) {
        continue;
      }
    }

    const impactStats = extractNativeImpactMetrics(recDetails);

    // Stage source budget if MOVE_UNUSED_BUDGET
    if (type === 'MOVE_UNUSED_BUDGET' && excessBudgetRes) {
      const sourceBudget = ctx.budgetDataMap.get(excessBudgetRes);
      if (sourceBudget && !budgetsByKey.has(excessBudgetRes)) {
        budgetsByKey.set(excessBudgetRes, {
          mccId: ctx.mccId,
          accountId: ctx.accountId,
          budgetResourceName: excessBudgetRes,
          campaignBudgetName: sourceBudget.campaignBudgetName,
          campaignBudgetAmount: sourceBudget.campaignBudgetAmount,
          campaignBudgetTotalAmount: sourceBudget.campaignBudgetTotalAmount,
          campaignBudgetDeliveryMethod: sourceBudget.campaignBudgetDeliveryMethod,
          campaignBudgetType: sourceBudget.campaignBudgetType,
          campaignBudgetIsShared: sourceBudget.campaignBudgetIsShared,
        });
      }
    }

    for (const campaign of targetCampaigns) {
      const budgetResName =
        campaign.campaignBudgetResourceName ||
        `customers/${ctx.accountId}/campaignBudgets/camp_${campaign.campaignId}`;

      if (!budgetsByKey.has(budgetResName)) {
        budgetsByKey.set(budgetResName, {
          mccId: ctx.mccId,
          accountId: ctx.accountId,
          budgetResourceName: budgetResName,
          campaignBudgetName: campaign.campaignBudgetName,
          campaignBudgetAmount: campaign.campaignBudgetAmount,
          campaignBudgetTotalAmount: campaign.campaignBudgetTotalAmount,
          campaignBudgetDeliveryMethod: campaign.campaignBudgetDeliveryMethod,
          campaignBudgetType: campaign.campaignBudgetType,
          campaignBudgetIsShared: campaign.campaignBudgetIsShared,
        });
      }

      if (!campaignsByKey.has(campaign.campaignResourceName)) {
        campaignsByKey.set(campaign.campaignResourceName, {
          mccId: ctx.mccId,
          accountId: ctx.accountId,
          campaignResourceName: campaign.campaignResourceName,
          campaignId: campaign.campaignId,
          campaignName: campaign.campaignName,
          campaignType: campaign.campaignType,
          campaignSubType: campaign.campaignSubType,
          campaignIsAiMax: campaign.campaignIsAiMax,
          campaignAiMaxTextCustomizationEnabled: campaign.campaignAiMaxTextCustomizationEnabled,
          campaignAiMaxFinalUrlExpansionEnabled: campaign.campaignAiMaxFinalUrlExpansionEnabled,
          budgetResourceName: budgetResName,
          campaignIsPortfolioBiddingStrategy: campaign.campaignIsPortfolioBiddingStrategy,
          campaignPortfolioBiddingStrategyName: campaign.campaignPortfolioBiddingStrategyName,
          portfolioStrategyResourceName: campaign.portfolioStrategyResourceName,
          portfolioOwnerId: campaign.portfolioOwnerId,
          isLocallyOwnedPortfolio: campaign.isLocallyOwnedPortfolio,
          campaignPortfolioIsManagerOwned: campaign.campaignPortfolioIsManagerOwned,
          campaignBiddingStrategyType: campaign.campaignBiddingStrategyType,
          campaignBiddingTargetRoas: campaign.campaignBiddingTargetRoas,
          campaignBiddingTargetCpa: campaign.campaignBiddingTargetCpa,
          campaignStatsYesterdayCost: campaign.campaignStatsYesterdayCost,
          campaignStats7DaysCost: campaign.campaignStats7DaysCost,
          campaign30DaysCost: campaign.campaign30DaysCost,
          campaign30DaysConversionsValue: campaign.campaign30DaysConversionsValue,
          campaign30DaysConversions: campaign.campaign30DaysConversions,
          campaign30DaysSearchRankLostImpressionShare:
            campaign.campaign30DaysSearchRankLostImpressionShare,
          campaign30DaysSearchBudgetLostImpressionShare:
            campaign.campaign30DaysSearchBudgetLostImpressionShare,
        });
      }

      const typeDetails = extractTypeSpecificStagingDetails(
        type,
        recDetails,
        campaign,
        sourceCampaigns
      );

      stagedRecs.push({
        mccId: ctx.mccId,
        accountId: ctx.accountId,
        recommendationId: recResourceName,
        recommendationType: type,
        campaignResourceName: campaign.campaignResourceName,
        recommendationCurrentBudgetAmount: typeDetails.recommendationCurrentBudgetAmount,
        recommendationNewBudgetAmount: typeDetails.recommendationNewBudgetAmount,
        recommendationNewTargetCpa: typeDetails.recommendationNewTargetCpa,
        recommendationNewTargetRoas: typeDetails.recommendationNewTargetRoas,
        recommendationBaseCost: impactStats.recommendationBaseCost,
        recommendationPotentialCost: impactStats.recommendationPotentialCost,
        recommendationBaseClicks: impactStats.recommendationBaseClicks,
        recommendationPotentialClicks: impactStats.recommendationPotentialClicks,
        recommendationBaseConversions: impactStats.recommendationBaseConversions,
        recommendationPotentialConversions: impactStats.recommendationPotentialConversions,
        recommendationBaseConversionsValue: impactStats.recommendationBaseConversionsValue,
        recommendationPotentialConversionsValue:
          impactStats.recommendationPotentialConversionsValue,
        recommendationBaseImpressions: impactStats.recommendationBaseImpressions,
        recommendationPotentialImpressions: impactStats.recommendationPotentialImpressions,
        recommendationBaseVideoViews: impactStats.recommendationBaseVideoViews,
        recommendationPotentialVideoViews: impactStats.recommendationPotentialVideoViews,
        moveBudgetSourceBudgetResourceName: typeDetails.moveBudgetSourceBudgetResourceName,
        moveBudgetSourceCampaigns: typeDetails.moveBudgetSourceCampaigns,
        moveBudgetSourceTargetCpa: typeDetails.moveBudgetSourceTargetCpa,
        moveBudgetSourceTargetRoas: typeDetails.moveBudgetSourceTargetRoas,
      });
    }
  }

  const accounts: StagedAccount[] =
    stagedRecs.length > 0
      ? [
          {
            mccId: ctx.mccId,
            mccName: ctx.mccName,
            accountId: ctx.accountId,
            accountName: ctx.accountName,
            currencyCode: ctx.currencyCode,
            timestamp: ctx.timestamp,
          },
        ]
      : [];

  return {
    accounts,
    budgets: Array.from(budgetsByKey.values()),
    campaigns: Array.from(campaignsByKey.values()),
    recommendations: stagedRecs,
  };
}

/**
 * Extracts target campaign resource names from either `recommendation.campaigns` (multi-campaign / portfolio)
 * or `recommendation.campaign` (single campaign).
 */
export function extractTargetCampaignResourceNames(recDetails: GaqlRecommendationRow): string[] {
  if (Array.isArray(recDetails.campaigns) && recDetails.campaigns.length > 0) {
    return recDetails.campaigns.map(String);
  }
  if (recDetails.campaign) {
    return [String(recDetails.campaign)];
  }
  return [];
}

/**
 * The budget sub-recommendation of a budget-type recommendation. The API's `recommendation` details are a
 * proto oneof, so at most one of these fields is set.
 */
function getBudgetRecommendation(
  recDetails: GaqlRecommendationRow
): GaqlBudgetRecommendationProto | undefined {
  return (
    recDetails.campaignBudgetRecommendation ??
    recDetails.forecastingCampaignBudgetRecommendation ??
    recDetails.marginalRoiCampaignBudgetRecommendation ??
    recDetails.moveUnusedBudgetRecommendation?.budgetRecommendation
  );
}

function extractNativeImpactMetrics(recDetails: GaqlRecommendationRow) {
  const budgetSubRec = getBudgetRecommendation(recDetails);

  let impactPair: GaqlImpactProto | null = null;

  if (budgetSubRec && Array.isArray(budgetSubRec.budgetOptions)) {
    const matchedOption = budgetSubRec.budgetOptions.find(
      (opt) =>
        opt.budgetAmountMicros !== undefined &&
        opt.budgetAmountMicros !== null &&
        budgetSubRec.recommendedBudgetAmountMicros !== undefined &&
        budgetSubRec.recommendedBudgetAmountMicros !== null &&
        Number(opt.budgetAmountMicros) === Number(budgetSubRec.recommendedBudgetAmountMicros)
    );
    if (matchedOption?.impact) {
      impactPair = matchedOption.impact;
    }
  }

  if (!impactPair && recDetails.impact) {
    impactPair = recDetails.impact;
  }

  const base = impactPair?.baseMetrics;
  const potential = impactPair?.potentialMetrics;

  return {
    recommendationBaseCost: fromMicros(base?.costMicros),
    recommendationPotentialCost: fromMicros(potential?.costMicros),
    recommendationBaseClicks: toNullableNumber(base?.clicks),
    recommendationPotentialClicks: toNullableNumber(potential?.clicks),
    recommendationBaseConversions: toNullableNumber(base?.conversions),
    recommendationPotentialConversions: toNullableNumber(potential?.conversions),
    recommendationBaseConversionsValue: toNullableNumber(base?.conversionsValue),
    recommendationPotentialConversionsValue: toNullableNumber(potential?.conversionsValue),
    recommendationBaseImpressions: toNullableNumber(base?.impressions),
    recommendationPotentialImpressions: toNullableNumber(potential?.impressions),
    recommendationBaseVideoViews: toNullableNumber(base?.videoViews),
    recommendationPotentialVideoViews: toNullableNumber(potential?.videoViews),
  };
}

function extractTypeSpecificStagingDetails(
  type: string,
  recDetails: GaqlRecommendationRow,
  campaign: EnrichedCampaignData,
  sourceCampaigns: SourceBudgetCampaignInfo[]
) {
  let recommendationCurrentBudgetAmount: number | null = null;
  let recommendationNewBudgetAmount: number | null = null;
  let recommendationNewTargetCpa: number | null = null;
  let recommendationNewTargetRoas: number | null = null;
  let moveBudgetSourceBudgetResourceName = '';
  let moveBudgetSourceCampaigns = 'N/A';
  let moveBudgetSourceTargetCpa: number | null = null;
  let moveBudgetSourceTargetRoas: number | null = null;

  if (
    type === 'CAMPAIGN_BUDGET' ||
    type === 'FORECASTING_CAMPAIGN_BUDGET' ||
    type === 'MARGINAL_ROI_CAMPAIGN_BUDGET'
  ) {
    const budgetRec = getBudgetRecommendation(recDetails);
    if (budgetRec) {
      recommendationCurrentBudgetAmount = fromMicros(budgetRec.currentBudgetAmountMicros);
      recommendationNewBudgetAmount = fromMicros(budgetRec.recommendedBudgetAmountMicros);
    }
  } else if (type === 'MOVE_UNUSED_BUDGET') {
    const moveRec = recDetails.moveUnusedBudgetRecommendation;
    const budgetRec = getBudgetRecommendation(recDetails);
    if (budgetRec) {
      recommendationCurrentBudgetAmount = fromMicros(budgetRec.currentBudgetAmountMicros);
      recommendationNewBudgetAmount = fromMicros(budgetRec.recommendedBudgetAmountMicros);
    }
    moveBudgetSourceBudgetResourceName = String(moveRec?.excessCampaignBudget ?? '');
    if (sourceCampaigns.length > 0) {
      moveBudgetSourceCampaigns = sourceCampaigns.map((sc) => sc.campaignName).join(', ');
      const { targetCpa, targetRoas } = resolveSourceBudgetTargets(sourceCampaigns);
      moveBudgetSourceTargetCpa = targetCpa;
      moveBudgetSourceTargetRoas = targetRoas;
    }
  } else if (type === 'RAISE_TARGET_CPA') {
    const cpaAdj = recDetails.raiseTargetCpaRecommendation?.targetAdjustment;
    const mult = toNullableNumber(
      cpaAdj?.recommendedTargetMultiplier ??
        recDetails.raiseTargetCpaRecommendation?.targetCpaMultiplier
    );
    const baseCpa =
      fromMicros(cpaAdj?.currentAverageTargetMicros) ?? campaign.campaignBiddingTargetCpa;
    if (mult !== null && baseCpa !== null) {
      recommendationNewTargetCpa = baseCpa * mult;
    }
  } else if (type === 'LOWER_TARGET_ROAS') {
    const roasAdj = recDetails.lowerTargetRoasRecommendation?.targetAdjustment;
    const mult = toNullableNumber(
      roasAdj?.recommendedTargetMultiplier ??
        recDetails.lowerTargetRoasRecommendation?.targetRoasMultiplier
    );
    const baseRoas =
      fromMicros(roasAdj?.currentAverageTargetMicros) ?? campaign.campaignBiddingTargetRoas;
    if (mult !== null && baseRoas !== null) {
      recommendationNewTargetRoas = baseRoas * mult;
    }
  }

  return {
    recommendationCurrentBudgetAmount,
    recommendationNewBudgetAmount,
    recommendationNewTargetCpa,
    recommendationNewTargetRoas,
    moveBudgetSourceBudgetResourceName,
    moveBudgetSourceCampaigns,
    moveBudgetSourceTargetCpa,
    moveBudgetSourceTargetRoas,
  };
}

/**
 * Donor targets for a `MOVE_UNUSED_BUDGET` source budget: the shared tCPA / tROAS when every enabled donor
 * campaign has the same value, otherwise `null` (a mixed set has no single meaningful target).
 */
function resolveSourceBudgetTargets(
  sourceCampaigns: SourceBudgetCampaignInfo[]
): { targetCpa: number | null; targetRoas: number | null } {
  const cpas = sourceCampaigns
    .map((c) => c.targetCpa)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  const roases = sourceCampaigns
    .map((c) => c.targetRoas)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));

  const uniqueCpas = Array.from(new Set(cpas));
  const uniqueRoases = Array.from(new Set(roases));
  return {
    targetCpa: uniqueCpas.length === 1 ? uniqueCpas[0] : null,
    targetRoas: uniqueRoases.length === 1 ? uniqueRoases[0] : null,
  };
}

function toNullableNumber(val: unknown): number | null {
  if (val === null || val === undefined || val === '') return null;
  const num = Number(val);
  return Number.isFinite(num) ? num : null;
}
