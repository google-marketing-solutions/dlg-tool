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
 * @fileoverview Per-entity recommendation ranker (`assignRecommendationNumbers`).
 * Partitions recommendations by their `recommendationScope` entity (`PORTFOLIO`, `SHARED_BUDGET`, or `CAMPAIGN`) and
 * assigns `recommendationNumber` (`1, 2, 3...`) sorted alphabetically (`A–Z`) by `recommendationType` (then
 * `recommendationId`), so custom Data Studio reports can filter to `recommendationNumber = 1` to count each entity's
 * historical spend once.
 * Bundled only into `dist/sheet/Code.gs`.
 */

import type {ReportRow} from '../../shared/types';
import {getPortfolioEntityKey} from './entity-keys';

/**
 * Assigns `recommendationNumber` (`1, 2, 3...`) partitioned by target entity
 * (owner CID + trailing strategy ID for PORTFOLIO, `budgetResourceName` for SHARED_BUDGET,
 * or `campaignId` for CAMPAIGN), ordered alphabetically (`A–Z`) by `recommendationType` and `recommendationId`.
 */
export function assignRecommendationNumbers(rows: readonly ReportRow[]): ReportRow[] {
  const byPartition = new Map<string, ReportRow[]>();

  for (const row of rows) {
    const clone = { ...row };
    const partitionKey = getPartitionKey(clone);
    const list = byPartition.get(partitionKey) ?? [];
    list.push(clone);
    byPartition.set(partitionKey, list);
  }

  const result: ReportRow[] = [];
  for (const group of byPartition.values()) {
    group.sort((a, b) => {
      const typeCmp = a.recommendationType.localeCompare(b.recommendationType);
      if (typeCmp !== 0) {
        return typeCmp;
      }
      return a.recommendationId.localeCompare(b.recommendationId);
    });

    for (let i = 0; i < group.length; i++) {
      group[i].recommendationNumber = i + 1;
      result.push(group[i]);
    }
  }

  return result;
}

function getPartitionKey(row: ReportRow): string {
  if (row.recommendationScope === 'PORTFOLIO') {
    // Child accounts see an MCC-owned portfolio under their own CID, so key by owner + strategy ID
    return `PORTFOLIO::${getPortfolioEntityKey(row)}`;
  }
  if (row.recommendationScope === 'SHARED_BUDGET') {
    const budgetKey = row.budgetResourceName || row.campaignBudgetName;
    return `SHARED_BUDGET::${row.accountId}::${budgetKey}`;
  }
  return `CAMPAIGN::${row.accountId}::${row.campaignId}`;
}
