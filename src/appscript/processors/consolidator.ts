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
 * @fileoverview Recommendation scope classifier and multi-campaign row consolidator (`consolidateRowsByScope`).
 * Classifies each row into `CAMPAIGN`, `PORTFOLIO`, or `SHARED_BUDGET` scope, and merges multi-campaign
 * portfolio and shared-budget recommendations into 1 row per recommendation—summing historical member-campaign
 * metrics and computing cost-weighted average Lost Impression Share without double-counting forecast impact.
 * Bundled only into `dist/sheet/Code.gs`.
 */

import type {ReportRow} from '../../shared/types';
import {normalizeCid} from '../../shared/utils/cid';
import {resolveAdsUiOcid} from './ads-ui-urls';
import {extractResourceTrailingId, getPortfolioGroupKey, getPortfolioOwnerCid} from './entity-keys';
import {populateDerivedRowMetrics} from './joiner';

/** Grouping mode for portfolio bidding strategies shared across multiple child accounts. */
export type CrossAccountPortfolioMode = 'CONSOLIDATED_STRATEGY' | 'SPLIT_BY_ACCOUNT';

/** Options controlling how portfolio and shared-budget recommendations are consolidated. */
export interface ConsolidationOptions {
  crossAccountPortfolioMode?: CrossAccountPortfolioMode;
}

/**
 * Classifies each row's `recommendationScope` (`CAMPAIGN`, `PORTFOLIO`, `SHARED_BUDGET`)
 * and consolidates multi-campaign portfolio or shared-budget recommendations into 1 row per recommendation
 * while summing historical member-campaign metrics.
 */
export function consolidateRowsByScope(
  rows: readonly ReportRow[],
  options: ConsolidationOptions = {}
): ReportRow[] {
  const crossAccountMode = options.crossAccountPortfolioMode ?? 'CONSOLIDATED_STRATEGY';

  const portfolioGroups = new Map<string, ReportRow[]>();
  const sharedBudgetGroups = new Map<string, ReportRow[]>();
  const campaignRows: ReportRow[] = [];

  for (const row of rows) {
    const isPortfolio =
      row.campaignIsPortfolioBiddingStrategy === true &&
      row.portfolioStrategyResourceName !== 'N/A' &&
      Boolean(row.portfolioStrategyResourceName);

    const isTargetOrMultiCampaignPortfolioRec =
      isPortfolio &&
      (row.recommendationType === 'RAISE_TARGET_CPA' ||
        row.recommendationType === 'LOWER_TARGET_ROAS');

    if (isTargetOrMultiCampaignPortfolioRec) {
      // Owner CID + trailing strategy/recommendation IDs + currency, so the copies Google Ads forks into each
      // child account merge into one row (see entity-keys.ts)
      const groupKey = getPortfolioGroupKey(row, crossAccountMode === 'SPLIT_BY_ACCOUNT');
      const list = portfolioGroups.get(groupKey) ?? [];
      list.push(row);
      portfolioGroups.set(groupKey, list);
      continue;
    }

    const isSharedBudgetRec =
      row.campaignBudgetIsShared === true &&
      (row.recommendationType === 'CAMPAIGN_BUDGET' ||
        row.recommendationType === 'FORECASTING_CAMPAIGN_BUDGET' ||
        row.recommendationType === 'MARGINAL_ROI_CAMPAIGN_BUDGET');

    if (isSharedBudgetRec) {
      const budgetKey = row.budgetResourceName || row.campaignBudgetName;
      const groupKey = `${row.recommendationId}::${row.accountId}::${budgetKey}`;
      const list = sharedBudgetGroups.get(groupKey) ?? [];
      list.push(row);
      sharedBudgetGroups.set(groupKey, list);
      continue;
    }

    const clone: ReportRow = { ...row, recommendationScope: 'CAMPAIGN' };
    campaignRows.push(clone);
  }

  const out: ReportRow[] = [...campaignRows];

  for (const group of portfolioGroups.values()) {
    out.push(mergeGroupRows(group, 'PORTFOLIO'));
  }

  for (const group of sharedBudgetGroups.values()) {
    out.push(mergeGroupRows(group, 'SHARED_BUDGET'));
  }

  return out;
}

function mergeGroupRows(
  group: ReportRow[],
  scope: 'PORTFOLIO' | 'SHARED_BUDGET'
): ReportRow {
  // Deduplicate member campaigns by campaignId + accountId so cross-account or repeated targets are summed once
  const uniqueCampaigns = new Map<string, ReportRow>();
  for (const r of group) {
    uniqueCampaigns.set(`${r.accountId}::${r.campaignId}`, r);
  }
  const members = Array.from(uniqueCampaigns.values());
  const base = { ...members[0] };

  base.recommendationScope = scope;
  const memberNames = members.map((m) => m.campaignName);
  base.portfolioCampaignsCount = members.length;
  base.portfolioMemberCampaignNames = memberNames.join(', ');

  if (scope === 'PORTFOLIO') {
    base.campaignName = base.campaignPortfolioBiddingStrategyName || base.campaignName;
    const strategyId = extractResourceTrailingId(base.portfolioStrategyResourceName, 'biddingStrategies');
    base.campaignId = strategyId || 'N/A';
    // The recommendation's decoded ocid identifies its own account, so it only applies when that account owns
    // the strategy; manager-owned strategies keep the owner CID (no owner ocid is available)
    const ownerCid = getPortfolioOwnerCid(base);
    const ocid =
      ownerCid === normalizeCid(base.accountId) ? resolveAdsUiOcid(base.recommendationId, ownerCid) : ownerCid;
    base.campaignUrl = strategyId
      ? `https://ads.google.com/aw/bidreport/portfolio?strategyId=${strategyId}&ocid=${ocid}`
      : `https://ads.google.com/aw/bidstrategies?ocid=${ocid}`;
  } else if (scope === 'SHARED_BUDGET') {
    base.campaignName = base.campaignBudgetName || base.campaignName;
    const budgetId = extractResourceTrailingId(base.budgetResourceName, 'campaignBudgets');
    base.campaignId = budgetId || 'N/A';
    // Shared budgets belong to the recommendation's own account, so its decoded ocid applies
    const ocid = resolveAdsUiOcid(base.recommendationId, base.accountId);
    base.campaignUrl = budgetId
      ? `https://ads.google.com/aw/budget/report?budgetId=${budgetId}&ocid=${ocid}`
      : `https://ads.google.com/aw/sharedbudgets?ocid=${ocid}`;
  }

  // Sum historical 30d/7d/yesterday metrics across distinct member campaigns
  base.campaignStatsYesterdayCost = sumNullable(members.map((m) => m.campaignStatsYesterdayCost));
  base.campaignStats7DaysCost = sumNullable(members.map((m) => m.campaignStats7DaysCost));
  base.campaign30DaysCost = sumNullable(members.map((m) => m.campaign30DaysCost));
  base.campaign30DaysConversions = sumNullable(members.map((m) => m.campaign30DaysConversions));
  base.campaign30DaysConversionsValue = sumNullable(
    members.map((m) => m.campaign30DaysConversionsValue)
  );

  // For PORTFOLIO scope, sum each distinct budget once (member campaigns can share a budget); for
  // SHARED_BUDGET, keep the single shared budget amount
  if (scope === 'PORTFOLIO') {
    const distinctBudgets = new Map<string, number | null>();
    for (const m of members) {
      distinctBudgets.set(`${m.accountId}::${m.budgetResourceName || m.campaignId}`, m.campaignBudgetAmount);
    }
    base.campaignBudgetAmount = sumNullable(Array.from(distinctBudgets.values()));
  }

  const totalCost30d = base.campaign30DaysCost ?? 0;
  const totalConv30d = base.campaign30DaysConversions ?? 0;
  const totalConvVal30d = base.campaign30DaysConversionsValue ?? 0;

  base.campaign30DaysCpa = totalConv30d > 0 ? totalCost30d / totalConv30d : null;
  base.campaign30DaysRoas = totalCost30d > 0 ? totalConvVal30d / totalCost30d : null;

  base.campaign30DaysSearchRankLostImpressionShare = costWeightedAverage(
    members,
    (m) => m.campaign30DaysSearchRankLostImpressionShare
  );
  base.campaign30DaysSearchBudgetLostImpressionShare = costWeightedAverage(
    members,
    (m) => m.campaign30DaysSearchBudgetLostImpressionShare
  );

  populateDerivedRowMetrics(base);
  return base;
}

function sumNullable(values: Array<number | null>): number | null {
  let hasNumber = false;
  let total = 0;
  for (const v of values) {
    if (typeof v === 'number' && Number.isFinite(v)) {
      hasNumber = true;
      total += v;
    }
  }
  return hasNumber ? total : null;
}

function costWeightedAverage(
  members: ReportRow[],
  selector: (row: ReportRow) => number | null
): number | null {
  let weightedSum = 0;
  let weightTotal = 0;
  for (const m of members) {
    const metric = selector(m);
    const cost = m.campaign30DaysCost ?? 0;
    if (typeof metric === 'number' && Number.isFinite(metric) && cost > 0) {
      weightedSum += cost * metric;
      weightTotal += cost;
    }
  }
  return weightTotal > 0 ? weightedSum / weightTotal : null;
}
