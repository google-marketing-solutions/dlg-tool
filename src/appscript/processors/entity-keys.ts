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
 * @fileoverview Shared entity keys for consolidation (`consolidator.ts`) and ranking (`ranking.ts`).
 * GAQL prefixes `recommendation.resource_name` and `campaign.bidding_strategy` with the requesting child
 * account (`customers/<childCid>/...`), even for a portfolio strategy owned by a manager account, and Google
 * Ads copies cross-account portfolio recommendations into every child account with the same trailing ID.
 * Keys are therefore built from the strategy owner and trailing IDs, never from raw resource names.
 * Bundled only into `dist/sheet/Code.gs`.
 */

import type {ReportRow} from '../../shared/types';
import {normalizeCid} from '../../shared/utils/cid';

/**
 * Returns the trailing ID after `/<collectionName>/` in a resource name
 * (e.g. `customers/1/biddingStrategies/77` -> `77`), or `''` when absent.
 */
export function extractResourceTrailingId(
  resourceName: string | undefined,
  collectionName: string
): string {
  if (!resourceName || resourceName === 'N/A') return '';
  const marker = `/${collectionName}/`;
  const idx = resourceName.indexOf(marker);
  if (idx === -1) return '';
  return resourceName.slice(idx + marker.length).trim();
}

/** Owner CID of a row's portfolio strategy: `portfolioOwnerId`, falling back to the row's account. */
export function getPortfolioOwnerCid(row: Pick<ReportRow, 'portfolioOwnerId' | 'accountId'>): string {
  return normalizeCid(row.portfolioOwnerId) || normalizeCid(row.accountId);
}

/**
 * Identity of a portfolio bidding strategy across child accounts: owner CID + trailing strategy ID.
 * Falls back to the raw resource name when it has no `/biddingStrategies/<id>` segment.
 */
export function getPortfolioEntityKey(
  row: Pick<ReportRow, 'portfolioOwnerId' | 'accountId' | 'portfolioStrategyResourceName'>
): string {
  const strategyId = extractResourceTrailingId(row.portfolioStrategyResourceName, 'biddingStrategies');
  return `${getPortfolioOwnerCid(row)}::${strategyId || row.portfolioStrategyResourceName || ''}`;
}

/**
 * Consolidation key for one portfolio recommendation: the portfolio entity, the trailing recommendation ID
 * (identical on every child-account copy), and `currencyCode` (so local-currency spend is never summed across
 * currencies). `perAccount` (the `SPLIT_BY_ACCOUNT` mode) also keys by account.
 */
export function getPortfolioGroupKey(
  row: Pick<
    ReportRow,
    'portfolioOwnerId' | 'accountId' | 'portfolioStrategyResourceName' | 'recommendationId' | 'currencyCode'
  >,
  perAccount = false
): string {
  const recId =
    extractResourceTrailingId(row.recommendationId, 'recommendations') ||
    String(row.recommendationId ?? '').split('/').pop() ||
    '';
  const key = `${getPortfolioEntityKey(row)}::${recId}::${row.currencyCode}`;
  return perAccount ? `${key}::${normalizeCid(row.accountId)}` : key;
}
