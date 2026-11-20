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
 * @fileoverview Final join, consolidation, ranking, FX conversion, and `Data` tab publisher (`publishFromStagingToData`).
 * Runs inside the Google Sheet's bound Apps Script (`dist/sheet/Code.gs`) via its periodic 15-minute trigger
 * (`runScheduledPublish`) or on-demand from the Sheet menu. The whole publish runs under `withSheetLock`
 * (`LockService.getDocumentLock()`): it syncs discovered MCCs/accounts from `_MCC_*` into `Config`, publishes each
 * active MCC as soon as its run completes (pending MCCs keep their previous snapshot, marked stale), builds the
 * 82-column `Data` table from `_Stage_*`, and records a deterministic `PUBLISH_SIGNATURE_<date>` so idle triggers
 * exit without writing anything.
 */

import {
  type MccCheckpointState,
  getMccCheckpoint,
} from '../shared/sheets/mcc-state';
import {
  type SheetLike,
  type SpreadsheetLike,
  getOrCreateSheet,
} from '../shared/sheets/sheet-utils';
import {readStagingBatchForMcc} from '../shared/sheets/staging-writer';
import type {GlobalConfig, ReportRow} from '../shared/types';
import {DLG_VERSION} from '../shared/version';
import {formatCidHyphenated, normalizeCid} from '../shared/utils/cid';
import {resolveSpreadsheetDate} from '../shared/utils/date';
import {
  CONFIG_SHEET_NAME,
  type MccSelectionSummary,
  type MccSyncResult,
  STATUS_BANNER_CELL,
  getMccSelectionSummary,
  getSelectedCidsForMcc,
  readGlobalConfig,
  syncAllMccSheetsToConfig,
} from './config-sheet';
import {applyCurrencyConversionToRow, resolveDailyFxRates} from './fx-rates';
import {consolidateRowsByScope} from './processors/consolidator';
import {joinStagingBatchToReportRows} from './processors/joiner';
import {assignRecommendationNumbers} from './processors/ranking';
import {
  getOrLockDailyConfig,
  readLockedDailyConfig,
  readRunStateMap,
  setRunStateValues,
  withSheetLock,
} from './run-state';
import {
  DATA_SHEET_NAME,
  readReportRowsFromSheet,
  writeReportRowsToSheet,
} from './sheets/data-writer';

/** Options for invoking `publishFromStagingToData`. */
export interface PublishOptions {
  spreadsheet?: SpreadsheetLike;
  /**
   * Publish even when nothing changed (skips the NO_OP / signature short-circuits). Never marks an MCC
   * fresh: readiness always comes from today's `_MCC_` checkpoint.
   */
  force?: boolean;
}

/** Maximum length of an account error shown in `mccSyncStatus`. */
const MAX_STATUS_ERROR_LENGTH = 80;
/** Maximum number of accounts listed by name in the "account under several MCCs" banner warning. */
const MAX_LISTED_DUPLICATE_ACCOUNTS = 3;

/** An MCC with at least one selected child account today. */
interface ActiveMcc extends MccSelectionSummary {
  selectedCids: string[];
}

/** Everything one publish needs, computed from `Config`, `_RunState`, and the `_MCC_` checkpoints. */
interface PublishPlan {
  dateStr: string;
  dailyConfig: GlobalConfig;
  activeMccs: ActiveMcc[];
  stateByMcc: Map<string, MccCheckpointState>;
  readyMccIds: Set<string>;
  hasExistingData: boolean;
  /** Setup problems shown in the status banner (they do not change the published rows). */
  setupWarnings: string[];
  /** Deterministic signature of the publish inputs, so idle triggers exit with zero sheet writes. */
  signature: string;
}

/** The rows to publish, plus the warnings collected while building them. */
interface AssembledRows {
  rows: ReportRow[];
  staleWarnings: string[];
  failedAccountWarnings: string[];
}

/**
 * Executed by the Sheet-bound Apps Script (`dist/sheet/Code.gs`) on its periodic trigger
 * or on-demand when the user clicks "Apply Config Changes Now".
 * Syncs discovered MCCs/accounts from `_MCC_*` to `Config`, joins the normalized `_Stage_*` sheets
 * for each completed MCC as soon as it finishes, marks any active MCC whose run has not completed today
 * as stale (preserving its previous snapshot in `Data`), consolidates by scope, ranks recommendations,
 * applies daily `GOOGLEFINANCE` FX rates (resolved once for all MCCs), and writes the final `Data` sheet.
 * Everything runs inside `withSheetLock`; idle ticks still perform zero sheet writes.
 *
 * Failed accounts: child accounts that an otherwise COMPLETED MCC gave up on (`failedCids`) keep their previous
 * `Data` rows with a warning sync status, and the status banner shows how many accounts failed.
 */
export function publishFromStagingToData(options: PublishOptions = {}): void {
  const spreadsheet: SpreadsheetLike =
    options.spreadsheet ?? SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = resolveSpreadsheetDate(spreadsheet);
  const force = Boolean(options.force);

  withSheetLock(spreadsheet, 'SHEET_PUBLISH', () => {
    const configSheet = getOrCreateSheet(spreadsheet, CONFIG_SHEET_NAME, { index: 0 });
    const syncInfo = syncConfigFromMccTabs(spreadsheet, configSheet);
    const runStateMap = readRunStateMap(spreadsheet);

    const plan = buildPublishPlan(spreadsheet, configSheet, dateStr, runStateMap, syncInfo);
    if (!plan || shouldSkipPublish(plan, runStateMap, force)) {
      return;
    }

    const dataSheet = getOrCreateSheet(spreadsheet, DATA_SHEET_NAME);
    // Sheets returns date cells as Dates built in the spreadsheet time zone; format them back in it.
    const existingDataRows = readReportRowsFromSheet(
      dataSheet,
      spreadsheet.getSpreadsheetTimeZone() || undefined
    );
    const assembled = assembleRows(spreadsheet, plan, existingDataRows);
    const publishedRows = convertCurrencies(spreadsheet, plan, assembled.rows);

    writeReportRowsToSheet(dataSheet, publishedRows);
    recordPublish(spreadsheet, plan, publishedRows.length, assembled);
    updateStatusBanner(configSheet, plan, publishedRows.length, assembled);
  });
}

/**
 * Syncs newly discovered MCCs & child accounts from `_MCC_<mccId>` tabs into `Config` (zero writes when all
 * MCCs and CIDs are already registered) and shows the latest registration message for a new MCC.
 */
function syncConfigFromMccTabs(spreadsheet: SpreadsheetLike, configSheet: SheetLike): MccSyncResult {
  const syncInfo = syncAllMccSheetsToConfig(spreadsheet, configSheet);
  if (syncInfo.newlyRegisteredMccs > 0 && syncInfo.latestStatusMessages.length > 0) {
    setStatusBannerIfChanged(
      configSheet,
      syncInfo.latestStatusMessages[syncInfo.latestStatusMessages.length - 1]
    );
  }
  return syncInfo;
}

/**
 * Builds today's publish plan, or returns `null` when no MCC has selected accounts. Locks today's config in
 * `_RunState` as soon as an MCC is active (even while runs are `IN_PROGRESS`).
 */
function buildPublishPlan(
  spreadsheet: SpreadsheetLike,
  configSheet: SheetLike,
  dateStr: string,
  runStateMap: ReadonlyMap<string, string>,
  syncInfo: MccSyncResult
): PublishPlan | null {
  const summaries = getMccSelectionSummary(configSheet);
  if (summaries.length === 0) {
    return null;
  }

  // Today's frozen GlobalConfig (settings + selected CIDs per MCC). If not locked yet, read live Config.
  const lockedConfig = readLockedDailyConfig(spreadsheet, dateStr, runStateMap);
  const dailyConfig = lockedConfig ?? readGlobalConfig(configSheet);
  const activeMccs: ActiveMcc[] = summaries
    .map((s) => ({
      ...s,
      selectedCids:
        dailyConfig.selectedCidsByMcc?.[s.mccId] ?? getSelectedCidsForMcc(configSheet, s.mccId),
    }))
    .filter((s) => s.selectedCids.length > 0);
  if (activeMccs.length === 0) {
    return null;
  }
  if (!lockedConfig) {
    getOrLockDailyConfig(spreadsheet, dateStr, () => dailyConfig);
  }

  const stateByMcc = new Map<string, MccCheckpointState>(
    activeMccs.map((a) => [a.mccId, getMccCheckpoint(spreadsheet, dateStr, a.mccId)])
  );
  // Fresh only when today's checkpoint is COMPLETED (force never pretends a reset MCC is fresh)
  const readyMccIds = new Set(
    activeMccs
      .filter((a) => {
        const st = stateByMcc.get(a.mccId);
        return st?.status === 'COMPLETED' && st.dateStr === dateStr;
      })
      .map((a) => a.mccId)
  );

  const existingDataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  const setupWarnings = buildSetupWarnings(syncInfo);
  const signature = JSON.stringify({
    dateStr,
    targetCurrency: dailyConfig.targetCurrency,
    excludeShoppingAndPmax: dailyConfig.excludeShoppingAndPmax,
    mccs: activeMccs.map((a) => {
      const st = stateByMcc.get(a.mccId);
      return {
        mccId: a.mccId,
        selectedCids: a.selectedCids,
        isReadyToday: readyMccIds.has(a.mccId),
        lastCompletedDate: st?.lastCompletedDate ?? null,
        failedCids: Object.keys(st?.failedCids ?? {}).sort(),
      };
    }),
    setupWarnings,
  });

  return {
    dateStr,
    dailyConfig,
    activeMccs,
    stateByMcc,
    readyMccIds,
    hasExistingData: existingDataSheet !== null && existingDataSheet.getLastRow() > 1,
    setupWarnings,
    signature,
  };
}

/**
 * Banner warnings for setup problems: accounts discovered under more than one registered MCC, and MCCs
 * running a Google Ads Script copy from a different version than this Sheet.
 */
function buildSetupWarnings(syncInfo: MccSyncResult): string[] {
  const warnings: string[] = [];
  const duplicates = syncInfo.duplicateAccounts;
  if (duplicates.length > 0) {
    const listed = duplicates
      .slice(0, MAX_LISTED_DUPLICATE_ACCOUNTS)
      .map((d) => `${formatCidHyphenated(d.cid)} (MCCs ${d.mccIds.map(formatCidHyphenated).join(', ')})`);
    const more =
      duplicates.length > MAX_LISTED_DUPLICATE_ACCOUNTS
        ? ` and ${duplicates.length - MAX_LISTED_DUPLICATE_ACCOUNTS} more`
        : '';
    warnings.push(
      `⚠️ ${duplicates.length} account(s) are under more than one MCC — install the Google Ads Script only in top-level MCCs: ${listed.join(', ')}${more}`
    );
  }
  if (syncInfo.outdatedScripts.length > 0) {
    const listed = syncInfo.outdatedScripts
      .map((o) => `${formatCidHyphenated(o.mccId)} (${o.version ? `v${o.version}` : 'older version'})`)
      .join(', ');
    warnings.push(
      `⚠️ Google Ads Script out of date in MCC(s) ${listed} — re-copy it with 📋 Copy Google Ads Script (this Sheet runs v${DLG_VERSION})`
    );
  }
  return warnings;
}

/**
 * True when this tick must not publish: no MCC has finished today and there is no previous `Data` to mark
 * stale (do not create an empty `Data` sheet mid-checkpoint), or nothing changed since the last publish.
 * `force` overrides both.
 */
function shouldSkipPublish(
  plan: PublishPlan,
  runStateMap: ReadonlyMap<string, string>,
  force: boolean
): boolean {
  if (force) {
    return false;
  }
  if (plan.readyMccIds.size === 0 && !plan.hasExistingData) {
    return true;
  }
  const previousSignature = runStateMap.get(`PUBLISH_SIGNATURE_${plan.dateStr}`);
  const publishStatus = runStateMap.get(`PUBLISH_STATUS_${plan.dateStr}`);
  const pendingCount = plan.activeMccs.length - plan.readyMccIds.size;
  return (
    previousSignature === plan.signature ||
    (publishStatus === 'PUBLISHED' && pendingCount === 0 && !previousSignature)
  );
}

/** Builds the rows of every active MCC: fresh rows for MCCs completed today, previous snapshots otherwise. */
function assembleRows(
  spreadsheet: SpreadsheetLike,
  plan: PublishPlan,
  existingDataRows: readonly ReportRow[]
): AssembledRows {
  const out: AssembledRows = { rows: [], staleWarnings: [], failedAccountWarnings: [] };
  for (const mcc of plan.activeMccs) {
    if (plan.readyMccIds.has(mcc.mccId)) {
      out.rows.push(...buildFreshMccRows(spreadsheet, plan, mcc));
      out.rows.push(...buildFailedAccountRows(plan, mcc, existingDataRows, out.failedAccountWarnings));
    } else {
      out.staleWarnings.push(`${formatCidHyphenated(mcc.mccId)} (${mcc.mccName})`);
      out.rows.push(...buildStaleMccRows(spreadsheet, plan, mcc, existingDataRows));
    }
  }
  return out;
}

/** True when `row` belongs to a selected account and passes today's Shopping filter. */
function keepRow(plan: PublishPlan, row: ReportRow, selectedCids: ReadonlySet<string>): boolean {
  if (!selectedCids.has(normalizeCid(row.accountId))) return false;
  if (plan.dailyConfig.excludeShoppingAndPmax && row.campaignType === 'SHOPPING') return false;
  return true;
}

/**
 * Fresh completed run today: rows from the staged batch (even if 0 recommendations — never resurrect old
 * `Data`). Accounts given up today are excluded here and handled by `buildFailedAccountRows`.
 */
function buildFreshMccRows(spreadsheet: SpreadsheetLike, plan: PublishPlan, mcc: ActiveMcc): ReportRow[] {
  const selectedCids = new Set(mcc.selectedCids);
  const failedCids = new Set(Object.keys(plan.stateByMcc.get(mcc.mccId)?.failedCids ?? {}));
  const stagedBatch = readStagingBatchForMcc(spreadsheet, mcc.mccId);
  if (stagedBatch.recommendations.length === 0) {
    return [];
  }
  const joined = joinStagingBatchToReportRows(stagedBatch).filter(
    (r) => keepRow(plan, r, selectedCids) && !failedCids.has(normalizeCid(r.accountId))
  );
  return assignRecommendationNumbers(consolidateRowsByScope(joined)).map((r) => ({
    ...r,
    timestamp: r.timestamp || `${plan.dateStr} 00:00:00`,
    mccSyncStatus: 'OK',
    mccLastUpdatedDate: plan.dateStr,
  }));
}

/**
 * Accounts that an otherwise COMPLETED MCC gave up on today keep their previous `Data` rows with a warning
 * (nothing if they had none). Appends each such account to `warnings`.
 */
function buildFailedAccountRows(
  plan: PublishPlan,
  mcc: ActiveMcc,
  existingDataRows: readonly ReportRow[],
  warnings: string[]
): ReportRow[] {
  const selectedCids = new Set(mcc.selectedCids);
  const failedInfo = plan.stateByMcc.get(mcc.mccId)?.failedCids ?? {};
  const rows: ReportRow[] = [];
  for (const cid of Object.keys(failedInfo).filter((c) => selectedCids.has(c))) {
    const hyphenatedCid = formatCidHyphenated(cid);
    warnings.push(hyphenatedCid);
    const shortError = truncate(failedInfo[cid].error, MAX_STATUS_ERROR_LENGTH);
    for (const r of existingDataRows) {
      if (normalizeCid(r.mccId) !== mcc.mccId || normalizeCid(r.accountId) !== cid) continue;
      if (!keepRow(plan, r, selectedCids)) continue;
      rows.push({
        ...r,
        mccSyncStatus: `WARNING: Account ${hyphenatedCid} failed (${shortError}); showing snapshot from ${r.mccLastUpdatedDate ?? 'a previous run'}`,
      });
    }
  }
  return rows;
}

/**
 * Pending/stale MCC: its previous rows in `Data`, or `_Stage_*` while that still holds the previous completed
 * run (only when today's run has made no progress; once today's run has flushed completed or failed
 * accounts, staging holds a partial batch, so the MCC shows nothing until its run completes).
 */
function buildStaleMccRows(
  spreadsheet: SpreadsheetLike,
  plan: PublishPlan,
  mcc: ActiveMcc,
  existingDataRows: readonly ReportRow[]
): ReportRow[] {
  const state = plan.stateByMcc.get(mcc.mccId)!;
  const selectedCids = new Set(mcc.selectedCids);
  let rawRows: readonly ReportRow[] = existingDataRows.filter((r) => normalizeCid(r.mccId) === mcc.mccId);

  const hasProgressToday =
    state.completedCids.length > 0 || Object.keys(state.failedCids ?? {}).length > 0;
  if (rawRows.length === 0 && !hasProgressToday) {
    const stagedBatch = readStagingBatchForMcc(spreadsheet, mcc.mccId);
    if (stagedBatch.recommendations.length > 0) {
      rawRows = assignRecommendationNumbers(
        consolidateRowsByScope(joinStagingBatchToReportRows(stagedBatch))
      );
    }
  }

  const lastDate = state.lastCompletedDate ?? 'previous run';
  return rawRows
    .filter((r) => keepRow(plan, r, selectedCids))
    .map((r) => ({
      ...r,
      timestamp: r.timestamp || `${state.lastCompletedDate ?? plan.dateStr} 00:00:00`,
      mccSyncStatus: `WARNING: Stale snapshot from ${lastDate}`,
      // Date column: never write the 'previous run' label; leave it blank when the date is unknown
      mccLastUpdatedDate: state.lastCompletedDate ?? null,
    }));
}

/** Resolves FX once for all MCCs' rows and fills the `*Global` columns. */
function convertCurrencies(
  spreadsheet: SpreadsheetLike,
  plan: PublishPlan,
  rows: ReportRow[]
): ReportRow[] {
  if (rows.length === 0) {
    return rows;
  }
  const fxRates = resolveDailyFxRates({
    spreadsheet,
    dateStr: plan.dateStr,
    sourceCurrencies: Array.from(new Set(rows.map((r) => r.currencyCode))),
    targetCurrency: plan.dailyConfig.targetCurrency,
  });
  return rows.map((r) => applyCurrencyConversionToRow(r, plan.dailyConfig.targetCurrency, fxRates));
}

/** Records today's publish signature and status in `_RunState`. */
function recordPublish(
  spreadsheet: SpreadsheetLike,
  plan: PublishPlan,
  publishedCount: number,
  assembled: AssembledRows
): void {
  const hasDataWarnings =
    assembled.staleWarnings.length > 0 || assembled.failedAccountWarnings.length > 0;
  setRunStateValues(
    spreadsheet,
    [
      { key: `PUBLISH_SIGNATURE_${plan.dateStr}`, value: plan.signature, notes: `Signature for ${plan.dateStr}` },
      {
        key: `PUBLISH_STATUS_${plan.dateStr}`,
        value: hasDataWarnings ? 'PUBLISHED_WITH_STALE_WARNING' : 'PUBLISHED',
        notes: `Published ${publishedCount} rows on ${plan.dateStr}`,
      },
    ],
    plan.dateStr
  );
}

/** Shows the publish result, data warnings, and setup warnings in the `Config` status banner. */
function updateStatusBanner(
  configSheet: SheetLike,
  plan: PublishPlan,
  publishedCount: number,
  assembled: AssembledRows
): void {
  const parts: string[] = [];
  if (assembled.staleWarnings.length > 0) {
    parts.push(
      `⚠️ WARNING — Published (${plan.dateStr}) with stale snapshot for pending MCC(s): ${assembled.staleWarnings.join(', ')}`
    );
  }
  if (assembled.failedAccountWarnings.length > 0) {
    parts.push(
      `⚠️ ${assembled.failedAccountWarnings.length} account(s) failed in today's run (showing their previous snapshot): ${assembled.failedAccountWarnings.join(', ')}`
    );
  }
  if (parts.length === 0) {
    parts.push(
      `🟢 PUBLISHED (${plan.dateStr}) — ${publishedCount} recommendations across ${plan.activeMccs.length} active MCC(s).`
    );
  }
  parts.push(...plan.setupWarnings);
  setStatusBannerIfChanged(configSheet, parts.join(' | '));
}

function truncate(text: string, maxLength: number): string {
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function setStatusBannerIfChanged(configSheet: SheetLike, nextMessage: string): void {
  const cell = configSheet.getRange(STATUS_BANNER_CELL);
  const currentMessage = String(cell.getValue() ?? '');
  if (currentMessage !== nextMessage) {
    cell.setValue(nextMessage);
  }
}
