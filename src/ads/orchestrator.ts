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
 * @fileoverview Google Ads Script orchestrator (`dist/ads/Code.gs`).
 * 1. Discovers child accounts under the current MCC and writes discovery + checkpoint state to the
 *    MCC's hidden tab (`_MCC_<normMccId>`).
 * 2. If `DAILY_CONFIG_LOCK_<date>` has not been written to `_RunState` yet (or does not include this
 *    MCC yet, or the script is running in Preview mode), writes the discovered accounts to
 *    `_MCC_<normMccId>` and exits until the Sheet App Script locks today's config in `_RunState`.
 * 3. Reads the frozen daily config from `_RunState` (never reading `Config` directly), runs the GAQL
 *    queries, and flushes batches to the MCC's staging tabs (`_Stage_*_<normMccId>`).
 * 4. Exits idle hourly runs once today's run is already `COMPLETED`.
 */

import {
  type FailedAccountInfo,
  MAX_ACCOUNT_ATTEMPTS,
  type MccCheckpointState,
  getMccCheckpoint,
  writeMccSheetState,
} from '../shared/sheets/mcc-state';
import {
  type SpreadsheetLike,
  readLockedDailyConfig,
} from '../shared/sheets/sheet-utils';
import {
  appendStagingBatch,
  clearMccFromStaging,
  createEmptyStagingBatch,
} from '../shared/sheets/staging-writer';
import type {ChildAccountInput, GlobalConfig, MccIdentity, StagingBatch} from '../shared/types';
import {formatCidHyphenated, normalizeCid} from '../shared/utils/cid';
import {
  fetchAllBudgetDetails,
  fetchAllCampaignData,
  fetchAllRecommendationDetails,
  fetchCampaignsByBudgets,
} from './gaql/fetcher';
import {
  buildAccountStagingBatch,
  extractTargetCampaignResourceNames,
} from './processors/recommendations';

/** Minimum remaining execution time in seconds (`4 min`) before saving a checkpoint and pausing until the next hourly trigger. */
export const MIN_REMAINING_TIME_SECONDS = 240;
/** Number of child accounts processed between intermediate `_Stage_*` and `_MCC_*` flushes. */
export const FLUSH_EVERY_N_ACCOUNTS = 25;
/** Number of staged recommendations accumulated in memory before forcing an intermediate flush. */
export const FLUSH_EVERY_N_RECS = 500;

/** Configuration options passed to `runAdsScriptOrchestrator`. */
export interface OrchestratorOptions {
  spreadsheetUrl: string;
  timeZone?: string;
}

/** Thrown when `Code.gs` runs outside a Google Ads manager (MCC) account. */
export const NOT_A_MANAGER_ACCOUNT_ERROR =
  'DLG Dashboard: this Google Ads Script must run in a Google Ads manager (MCC) account.';

/** Maximum length of an account error stored in the `_MCC_<id>` checkpoint (one Sheets cell holds 50,000 chars). */
export const MAX_STORED_ERROR_LENGTH = 200;

/** Identity and date of the current Google Ads Script execution. */
interface RunEnvironment {
  spreadsheet: SpreadsheetLike;
  mcc: MccIdentity;
  normMccId: string;
  dateStr: string;
  isPreview: boolean;
}

/** Today's frozen config from `_RunState`, with this MCC's locked child accounts. */
interface DailyLockForMcc {
  config: GlobalConfig;
  lockedCids: string[];
}

/** Everything `saveProgress` needs to write staging and the MCC's `_MCC_<id>` state. */
interface MccRunContext {
  spreadsheet: SpreadsheetLike;
  mcc: MccIdentity;
  normMccId: string;
  dateStr: string;
  childAccounts: ChildAccountInput[];
}

/** Today's progress: processed (or given-up) accounts and accounts that failed. */
interface RunProgress {
  completed: Set<string>;
  failed: Record<string, FailedAccountInfo>;
}

type ProcessOutcome = { status: 'PAUSED_FOR_TIME' } | { status: 'FINISHED'; pendingBatch: StagingBatch };

/**
 * Coordinates MCC discovery, hourly checkpoint/resume loop, and staging (`_Stage_*_<normMccId>`
 * and `_MCC_<normMccId>`).
 */
export function runAdsScriptOrchestrator(options: OrchestratorOptions): void {
  const env = createRunEnvironment(options);

  // Read today's frozen config lock from _RunState (written by the Sheet App Script)
  const lock = readDailyLockForMcc(env);

  // 1. Idle hourly runs: once today's run is COMPLETED, exit. Accounts linked or renamed later in
  //    the day are picked up the next day or after resetting the checkpoint.
  let checkpoint: MccCheckpointState | undefined;
  if (!env.isPreview && lock) {
    checkpoint = getMccCheckpoint(env.spreadsheet, env.dateStr, env.normMccId);
    if (checkpoint.status === 'COMPLETED' && checkpoint.dateStr === env.dateStr) {
      Logger.log(
        `[DLG v2] MCC ${env.mcc.mccId} already marked COMPLETED for ${env.dateStr}. Nothing to do (use "🚀 DLG Dashboard v2 -> Apply Config Changes Now" in the Google Sheet to re-run).`
      );
      return;
    }
  }

  // 2. Discover child accounts under this MCC
  const { childAccounts, accountEntityMap } = discoverChildAccountsUnderCurrentMcc();
  const ctx: MccRunContext = {
    spreadsheet: env.spreadsheet,
    mcc: env.mcc,
    normMccId: env.normMccId,
    dateStr: env.dateStr,
    childAccounts,
  };

  // 3. Preview, First-Run Discovery, or waiting for DAILY_CONFIG_LOCK_<date> in _RunState:
  //    write discovery payload to _MCC_<normMccId> and exit until the Sheet script locks today's config.
  if (env.isPreview || !lock || !checkpoint) {
    registerDiscovery(ctx, env.isPreview, lock !== null);
    return;
  }

  // 4. Start / continue today's run using strictly the frozen config from _RunState
  const selectedCids = lock.lockedCids.filter((cid) => accountEntityMap.has(cid));
  const progress = startOrResumeRun(ctx, checkpoint, selectedCids.length);

  // 5. Process remaining selected accounts
  const outcome = processSelectedAccounts(
    ctx,
    progress,
    selectedCids,
    accountEntityMap,
    lock.config.excludeShoppingAndPmax
  );
  if (outcome.status === 'PAUSED_FOR_TIME') return;

  // 6. Retry pending accounts on the next run, or mark the MCC COMPLETED
  finalizeRun(ctx, progress, selectedCids, outcome.pendingBatch);
}

function createRunEnvironment(options: OrchestratorOptions): RunEnvironment {
  const mccAccount = AdsApp.currentAccount();
  const mccId = mccAccount.getCustomerId();
  return {
    spreadsheet: SpreadsheetApp.openByUrl(options.spreadsheetUrl),
    mcc: { mccId, mccName: mccAccount.getName() || 'MCC' },
    normMccId: normalizeCid(mccId),
    dateStr: Utilities.formatDate(new Date(), options.timeZone || 'UTC', 'yyyy-MM-dd'),
    isPreview: AdsApp.getExecutionInfo().isPreview(),
  };
}

/** Today's locked config, or `null` when it is not locked yet or does not include this MCC. */
function readDailyLockForMcc(env: RunEnvironment): DailyLockForMcc | null {
  const config = readLockedDailyConfig(env.spreadsheet, env.dateStr);
  const lockedCids = config?.selectedCidsByMcc?.[env.normMccId];
  return config && Array.isArray(lockedCids) ? { config, lockedCids } : null;
}

/** Writes the discovered accounts to `_MCC_<id>` (no checkpoint change) so the Sheet can register them. */
function registerDiscovery(ctx: MccRunContext, isPreview: boolean, isLockedInRunState: boolean): void {
  const { mccId, mccName } = ctx.mcc;
  const msg = `🟢 MCC ${mccId} (${mccName}) registered (${ctx.childAccounts.length} accounts). Review checkboxes & Target Currency in Config; full run will execute on next trigger.`;
  Logger.log(
    `[DLG v2] Discovery / waiting for daily lock (isPreview=${isPreview}, isLockedInRunState=${isLockedInRunState}): discovered ${ctx.childAccounts.length} child accounts under MCC ${mccId} (${mccName}).`
  );
  writeMccSheetState(ctx.spreadsheet, {
    mcc: ctx.mcc,
    childAccounts: ctx.childAccounts,
    statusMessage: msg,
  });
}

/**
 * Loads today's progress from `checkpoint`. Only a fresh start of the day clears yesterday's staging: once
 * today has recorded progress (completed OR failed accounts), staging holds today's data and must survive
 * retries, even if every account failed.
 */
function startOrResumeRun(
  ctx: MccRunContext,
  checkpoint: MccCheckpointState,
  selectedCount: number
): RunProgress {
  const progress: RunProgress = {
    completed: new Set(checkpoint.completedCids),
    failed: { ...checkpoint.failedCids },
  };
  Logger.log(
    `[DLG v2] MCC ${ctx.mcc.mccId} (${ctx.mcc.mccName}) on ${ctx.dateStr}: ${selectedCount} selected account(s) (of ${ctx.childAccounts.length} discovered). Checkpoint status: ${checkpoint.status} (${checkpoint.completedCids.length} completed, ${Object.keys(progress.failed).length} failed).`
  );
  if (checkpoint.completedCids.length === 0 && Object.keys(progress.failed).length === 0) {
    clearMccFromStaging(ctx.spreadsheet, ctx.normMccId);
  }
  return progress;
}

/**
 * Processes every selected account not yet completed, flushing every `FLUSH_EVERY_N_ACCOUNTS` accounts or
 * `FLUSH_EVERY_N_RECS` recommendations. Saves progress and pauses when less than
 * `MIN_REMAINING_TIME_SECONDS` of execution time is left.
 */
function processSelectedAccounts(
  ctx: MccRunContext,
  progress: RunProgress,
  selectedCids: readonly string[],
  accountEntityMap: ReadonlyMap<string, GoogleAdsScripts.AdsManagerApp.ManagedAccount>,
  excludeShoppingAndPmax: boolean
): ProcessOutcome {
  let batch = createEmptyStagingBatch();
  let unFlushedAccountCount = 0;

  for (const cid of selectedCids) {
    if (progress.completed.has(cid)) continue;

    // Flush staging data if timer is less than the minimum amount of seconds and exit.
    const remainingSeconds = Number(AdsApp.getExecutionInfo().getRemainingTime());
    if (remainingSeconds < MIN_REMAINING_TIME_SECONDS) {
      saveProgress(
        ctx,
        batch,
        progress,
        'IN_PROGRESS',
        `⏳ IN PROGRESS — MCC ${ctx.mcc.mccId}: ${progress.completed.size}/${selectedCids.length} accounts processed. Will resume on next hourly trigger.`
      );
      return { status: 'PAUSED_FOR_TIME' };
    }

    processOneAccount(ctx, progress, cid, accountEntityMap.get(cid)!, excludeShoppingAndPmax, batch);
    unFlushedAccountCount++;

    if (
      unFlushedAccountCount >= FLUSH_EVERY_N_ACCOUNTS ||
      batch.recommendations.length >= FLUSH_EVERY_N_RECS
    ) {
      saveProgress(
        ctx,
        batch,
        progress,
        'IN_PROGRESS',
        `⏳ IN PROGRESS — MCC ${ctx.mcc.mccId}: ${progress.completed.size}/${selectedCids.length} accounts processed.`
      );
      batch = createEmptyStagingBatch();
      unFlushedAccountCount = 0;
    }
  }
  return { status: 'FINISHED', pendingBatch: batch };
}

/**
 * Processes one account into `batch`. On failure the account stays out of `completed` (retried on the next
 * hourly run) until it has failed in `MAX_ACCOUNT_ATTEMPTS` runs, after which it is given up for today.
 */
function processOneAccount(
  ctx: MccRunContext,
  progress: RunProgress,
  cid: string,
  accountEntity: GoogleAdsScripts.AdsManagerApp.ManagedAccount,
  excludeShoppingAndPmax: boolean,
  batch: StagingBatch
): void {
  try {
    const accountBatch = processSingleChildAccountToBatch(
      accountEntity,
      ctx.mcc.mccId,
      ctx.mcc.mccName,
      `${ctx.dateStr} 00:00:00`,
      excludeShoppingAndPmax
    );
    batch.accounts.push(...accountBatch.accounts);
    batch.budgets.push(...accountBatch.budgets);
    batch.campaigns.push(...accountBatch.campaigns);
    batch.recommendations.push(...accountBatch.recommendations);
    delete progress.failed[cid];
    progress.completed.add(cid);
    Logger.log(
      `[DLG v2] Processed account ${formatCidHyphenated(cid)}: ${accountBatch.recommendations.length} recommendation(s) staged.`
    );
  } catch (err) {
    const errMsg = truncate(err instanceof Error ? err.message : String(err), MAX_STORED_ERROR_LENGTH);
    const attempts = (progress.failed[cid]?.attempts ?? 0) + 1;
    progress.failed[cid] = { error: errMsg, attempts };
    if (attempts >= MAX_ACCOUNT_ATTEMPTS) {
      progress.completed.add(cid);
    }
    Logger.log(
      `Error processing child account ${cid} under MCC ${ctx.mcc.mccId} (attempt ${attempts}/${MAX_ACCOUNT_ATTEMPTS}): ${errMsg}`
    );
  }
}

/**
 * Failed accounts still below `MAX_ACCOUNT_ATTEMPTS` keep the MCC `IN_PROGRESS` so the next run retries them;
 * otherwise flushes the remaining batch and marks the MCC `COMPLETED` for today.
 */
function finalizeRun(
  ctx: MccRunContext,
  progress: RunProgress,
  selectedCids: readonly string[],
  pendingBatch: StagingBatch
): void {
  const retryCount = selectedCids.filter((cid) => !progress.completed.has(cid)).length;
  if (retryCount > 0) {
    saveProgress(
      ctx,
      pendingBatch,
      progress,
      'IN_PROGRESS',
      `⏳ IN PROGRESS — MCC ${ctx.mcc.mccId}: ${progress.completed.size}/${selectedCids.length} accounts processed; retrying ${retryCount} failed account(s) on the next hourly trigger.`
    );
    return;
  }

  const givenUpCount = Object.keys(progress.failed).length;
  saveProgress(
    ctx,
    pendingBatch,
    progress,
    'COMPLETED',
    `⏳ STAGED FOR PUBLISH (${ctx.dateStr}) — MCC ${ctx.mcc.mccId} completed (${progress.completed.size} accounts${givenUpCount > 0 ? `, ${givenUpCount} failed` : ''}).`
  );
}

function truncate(text: string, maxLength: number): string {
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function discoverChildAccountsUnderCurrentMcc() {
  const childAccounts: ChildAccountInput[] = [];
  const accountEntityMap = new Map<string, GoogleAdsScripts.AdsManagerApp.ManagedAccount>();

  if (typeof AdsManagerApp === 'undefined') {
    throw new Error(NOT_A_MANAGER_ACCOUNT_ERROR);
  }

  const iterator = AdsManagerApp.accounts().get();
  while (iterator.hasNext()) {
    const acc = iterator.next();
    const rawCid = String(acc.getCustomerId() ?? '');
    const normCid = normalizeCid(rawCid);
    if (!normCid) continue;

    accountEntityMap.set(normCid, acc);
    childAccounts.push({
      customerId: rawCid,
      name: String(acc.getName() || 'Account'),
      currencyCode: String(acc.getCurrencyCode() || 'USD'),
    });
  }

  return { childAccounts, accountEntityMap };
}

function processSingleChildAccountToBatch(
  accountEntity: GoogleAdsScripts.AdsManagerApp.ManagedAccount,
  mccId: string,
  mccName: string,
  timestamp: string,
  excludeShoppingAndPmax: boolean
): StagingBatch {
  // Switch the Ads context once; every query below runs against this account.
  AdsManagerApp.select(accountEntity);

  const accountId = accountEntity.getCustomerId();
  const accountName = accountEntity.getName() || 'Account';
  const currencyCode = accountEntity.getCurrencyCode() || 'USD';

  const recommendationDetailsMap = fetchAllRecommendationDetails();
  if (recommendationDetailsMap.size === 0) {
    return createEmptyStagingBatch();
  }

  const campaignResourceNames = new Set<string>();
  const budgetResourceNames = new Set<string>();

  for (const details of recommendationDetailsMap.values()) {
    for (const cName of extractTargetCampaignResourceNames(details)) {
      campaignResourceNames.add(cName);
    }
    const excessBudget = details.moveUnusedBudgetRecommendation?.excessCampaignBudget;
    if (excessBudget) {
      budgetResourceNames.add(String(excessBudget));
    }
  }

  const campaignDataMap = fetchAllCampaignData(
    Array.from(campaignResourceNames),
    excludeShoppingAndPmax
  );
  const budgetDataMap = fetchAllBudgetDetails(Array.from(budgetResourceNames));
  const budgetToCampaignsMap = fetchCampaignsByBudgets(
    Array.from(budgetResourceNames),
    excludeShoppingAndPmax
  );

  return buildAccountStagingBatch({
    mccId,
    mccName,
    accountId,
    accountName,
    currencyCode,
    timestamp,
    excludeShoppingAndPmax,
    recommendationDetailsMap,
    campaignDataMap,
    budgetDataMap,
    budgetToCampaignsMap,
  });
}

/**
 * Appends `batch` to the MCC's staging tabs, then writes the checkpoint and status message to `_MCC_<id>`.
 * A `COMPLETED` status also records today as the MCC's last completed date.
 */
function saveProgress(
  ctx: MccRunContext,
  batch: StagingBatch,
  progress: RunProgress,
  status: 'IN_PROGRESS' | 'COMPLETED',
  statusMessage: string
): void {
  appendStagingBatch(ctx.spreadsheet, batch, ctx.normMccId);

  const lastCompletedDate = status === 'COMPLETED' ? ctx.dateStr : undefined;
  const checkpoint: MccCheckpointState = {
    mccId: ctx.normMccId,
    dateStr: ctx.dateStr,
    status,
    completedCids: Array.from(progress.completed),
    ...(Object.keys(progress.failed).length > 0 ? { failedCids: progress.failed } : {}),
    ...(lastCompletedDate ? { lastCompletedDate } : {}),
  };

  writeMccSheetState(ctx.spreadsheet, {
    mcc: ctx.mcc,
    childAccounts: ctx.childAccounts,
    statusMessage,
    checkpoint,
    ...(lastCompletedDate ? { lastCompletedDate } : {}),
  });
}
