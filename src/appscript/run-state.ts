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
 * @fileoverview Publisher state, concurrency lock, and auto-compaction manager for `_RunState`.
 * Bundled only into the Sheet-bound Apps Script (`dist/sheet/Code.gs`), which is the sole writer
 * of `_RunState`, `Config`, and `Data`.
 * Provides:
 * - `withSheetLock`: Re-entrant wrapper around `LockService.getDocumentLock()`, available only in the
 *   Sheet-bound script (`dist/ads/Code.gs` never takes it).
 * - `setRunStateValue` / `setRunStateValues` / `deleteRunStateKeysByPrefix`: Atomic compacted writers
 *   that prune expired date-scoped keys (`DAILY_CONFIG_LOCK_`, `FX_RATES_`, `PUBLISH_STATUS_`,
 *   `PUBLISH_SIGNATURE_`).
 * - `getOrLockDailyConfig`, `resetTodayPublisherState`, and `resetTodayRunState`.
 */

import {
  MCC_SHEET_PREFIX,
  clearCheckpointRowInMccSheet,
  readMccSheetState,
} from '../shared/sheets/mcc-state';
import {
  RUN_STATE_SHEET_NAME,
  type SheetLike,
  type SpreadsheetLike,
  getOrCreateSheet,
  getRunStateValue,
  readLockedDailyConfig,
  readRunStateMap,
} from '../shared/sheets/sheet-utils';
import {serializeLockedDailyConfig} from '../shared/contract';
import type {GlobalConfig} from '../shared/types';
import {normalizeCid} from '../shared/utils/cid';

export {
  RUN_STATE_SHEET_NAME,
  getRunStateValue,
  readLockedDailyConfig,
  readRunStateMap,
};

/** Maximum time in milliseconds `withSheetLock` waits for the document lock. */
export const DOCUMENT_LOCK_TIMEOUT_MS = 30_000;

const DATE_SCOPED_KEY_REGEX =
  /^(DAILY_CONFIG_LOCK_|FX_RATES_|PUBLISH_STATUS_|PUBLISH_SIGNATURE_)(\d{4}-\d{2}-\d{2})/;

const activeSheetLocks = new WeakSet<SpreadsheetLike>();

type RunStateRow = [string, string, string, string];

function getOrCreateRunStateSheet(spreadsheet: SpreadsheetLike): SheetLike {
  let sheet = spreadsheet.getSheetByName(RUN_STATE_SHEET_NAME);
  if (!sheet) {
    sheet = getOrCreateSheet(spreadsheet, RUN_STATE_SHEET_NAME, { hidden: true });
    sheet.getRange(1, 1, 1, 4).setNumberFormat('@');
    sheet.getRange(1, 1, 1, 4).setValues([['Key', 'Value', 'UpdatedAt', 'Notes']]);
  }
  return sheet;
}

/**
 * Runs `fn` while holding Apps Script's `LockService.getDocumentLock()`, so concurrent executions of the
 * Sheet-bound script (scheduled publish, menu actions) never interleave writes to `Config`, `_RunState`, or
 * `Data`. Pending writes are flushed (`SpreadsheetApp.flush()`) before the lock is released. Re-entrant: nested
 * calls for the same spreadsheet run `fn` directly without re-acquiring.
 * Throws when no document lock is available (for example inside `Code.gs`, which must never call this).
 */
export function withSheetLock<T>(
  spreadsheet: SpreadsheetLike,
  ownerId: string,
  fn: () => T
): T {
  if (activeSheetLocks.has(spreadsheet)) {
    return fn();
  }

  const docLock =
    typeof LockService !== 'undefined' && typeof LockService.getDocumentLock === 'function'
      ? LockService.getDocumentLock()
      : null;
  if (!docLock) {
    throw new Error(
      `withSheetLock (${ownerId}) must run in the Sheet-bound script (LockService document lock unavailable).`
    );
  }

  docLock.waitLock(DOCUMENT_LOCK_TIMEOUT_MS);
  activeSheetLocks.add(spreadsheet);
  try {
    return fn();
  } finally {
    try {
      // Apps Script batches writes: flush before releasing so the next lock holder reads fresh values
      if (typeof SpreadsheetApp !== 'undefined') {
        SpreadsheetApp.flush();
      }
    } finally {
      activeSheetLocks.delete(spreadsheet);
      docLock.releaseLock();
    }
  }
}

/**
 * Returns true if `key` is a date-scoped operational entry (`DAILY_CONFIG_LOCK_`, `FX_RATES_`,
 * `PUBLISH_STATUS_`, or `PUBLISH_SIGNATURE_`) belonging to a date other than
 * `activeDateStr`. Multi-day keys (`FX_LAST_GOOD_*`) are never expired.
 */
export function isExpiredDateScopedKey(key: string, activeDateStr: string): boolean {
  const match = DATE_SCOPED_KEY_REGEX.exec(key);
  if (!match) return false;
  return match[2] !== activeDateStr;
}

/**
 * Atomically reads `_RunState` inside `withSheetLock`, optionally prunes expired date-scoped keys
 * from previous days, applies `mutator`, and writes the compacted table back in a single `setValues()` call.
 */
function mutateAndCompactRunState(
  spreadsheet: SpreadsheetLike,
  mutator: (rows: RunStateRow[], nowIso: string) => RunStateRow[],
  activeDateStr?: string
): void {
  withSheetLock(spreadsheet, 'RUN_STATE', () => {
    const sheet = getOrCreateRunStateSheet(spreadsheet);
    const prevLastRow = Math.max(1, sheet.getLastRow());
    const nowIso = new Date().toISOString();

    const currentRows: RunStateRow[] = [];
    if (prevLastRow >= 2) {
      const rawRows = sheet.getRange(2, 1, prevLastRow - 1, 4).getValues();
      for (const r of rawRows) {
        const k = String(r[0] ?? '').trim();
        if (!k) continue;
        if (activeDateStr && isExpiredDateScopedKey(k, activeDateStr)) continue;
        currentRows.push([k, String(r[1] ?? ''), String(r[2] ?? ''), String(r[3] ?? '')]);
      }
    }

    const nextRows = mutator(currentRows, nowIso);
    const targetRowCount = 1 + nextRows.length;
    const minSheetRows = Math.max(1, targetRowCount);
    const writeRowCount = Math.max(prevLastRow, minSheetRows);

    if (writeRowCount > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), writeRowCount - sheet.getMaxRows() + 5);
    }

    const matrix: RunStateRow[] = [['Key', 'Value', 'UpdatedAt', 'Notes'], ...nextRows];
    while (matrix.length < writeRowCount) {
      matrix.push(['', '', '', '']);
    }

    const range = sheet.getRange(1, 1, matrix.length, 4);
    range.setNumberFormat('@');
    range.setValues(matrix);

    if (prevLastRow > minSheetRows) {
      sheet.deleteRows(minSheetRows + 1, prevLastRow - minSheetRows);
    }
  });
}

function upsertRow(
  rows: RunStateRow[],
  key: string,
  value: string,
  nowIso: string,
  notes = ''
): void {
  const existingIdx = rows.findIndex((r) => r[0] === key);
  if (existingIdx >= 0) {
    rows[existingIdx] = [key, value, nowIso, notes];
  } else {
    rows.push([key, value, nowIso, notes]);
  }
}

/**
 * Atomically upserts several key-value entries in `_RunState` in a single read-and-rewrite
 * (`mutateAndCompactRunState`) and compacts expired date-scoped rows. No-op for an empty list.
 */
export function setRunStateValues(
  spreadsheet: SpreadsheetLike,
  entries: ReadonlyArray<{key: string; value: string; notes?: string}>,
  activeDateStr?: string
): void {
  if (entries.length === 0) return;
  mutateAndCompactRunState(
    spreadsheet,
    (rows, nowIso) => {
      for (const e of entries) {
        upsertRow(rows, e.key, e.value, nowIso, e.notes ?? '');
      }
      return rows;
    },
    activeDateStr
  );
}

/**
 * Atomically upserts a key-value entry in `_RunState` and compacts expired date-scoped rows.
 */
export function setRunStateValue(
  spreadsheet: SpreadsheetLike,
  key: string,
  value: string,
  notes = '',
  activeDateStr?: string
): void {
  mutateAndCompactRunState(
    spreadsheet,
    (rows, nowIso) => {
      upsertRow(rows, key, value, nowIso, notes);
      return rows;
    },
    activeDateStr
  );
}

/**
 * Atomically deletes all keys in `_RunState` matching `prefix` and compacts remaining rows.
 */
export function deleteRunStateKeysByPrefix(
  spreadsheet: SpreadsheetLike,
  prefix: string,
  activeDateStr?: string
): void {
  mutateAndCompactRunState(
    spreadsheet,
    (rows) => rows.filter((r) => !r[0].startsWith(prefix)),
    activeDateStr
  );
}

/**
 * Locks `GlobalConfig` (`targetCurrency`, `excludeShoppingAndPmax`, and `selectedCidsByMcc`) in
 * `_RunState` for `dateStr` (`YYYY-MM-DD`) so the Ads Script reads the frozen daily config from
 * `_RunState` and mid-day edits in `Config` do not take effect until tomorrow (or until "Apply
 * Config Changes Now"). Also prunes any expired date-scoped keys from previous days when locking
 * a new day.
 */
export function getOrLockDailyConfig(
  spreadsheet: SpreadsheetLike,
  dateStr: string,
  readCurrentConfig: () => GlobalConfig
): GlobalConfig {
  const existing = readLockedDailyConfig(spreadsheet, dateStr);
  if (existing) {
    return existing;
  }

  const currentConfig = readCurrentConfig();
  setRunStateValue(
    spreadsheet,
    `DAILY_CONFIG_LOCK_${dateStr}`,
    serializeLockedDailyConfig(currentConfig),
    `Locked daily config for ${dateStr}`,
    dateStr
  );
  return currentConfig;
}

/**
 * Clears today's publisher-side `_RunState` entries (daily config lock, FX cache, publish status/signature) in a
 * single compacted write, keeping every `_MCC_<id>` checkpoint. Used when only publisher-side settings changed.
 */
export function resetTodayPublisherState(spreadsheet: SpreadsheetLike, dateStr: string): void {
  const prefixes = [
    `DAILY_CONFIG_LOCK_${dateStr}`,
    `FX_RATES_${dateStr}`,
    `PUBLISH_STATUS_${dateStr}`,
    `PUBLISH_SIGNATURE_${dateStr}`,
  ];
  mutateAndCompactRunState(
    spreadsheet,
    (rows) => rows.filter((r) => !prefixes.some((p) => r[0].startsWith(p))),
    dateStr
  );
}

/**
 * Clears today's publisher state in `_RunState` and clears today's checkpoint row in every
 * `_MCC_<normMccId>` sheet (used when seeding synthetic demo data or performing a full reset).
 */
export function resetTodayRunState(spreadsheet: SpreadsheetLike, dateStr: string): void {
  resetTodayPublisherState(spreadsheet, dateStr);

  // Also clear today's checkpoint row inside each per-MCC _MCC_<normMccId> sheet
  const sheets = spreadsheet.getSheets();
  const nowIso = new Date().toISOString();
  for (const s of sheets) {
    const name = s.getName();
    if (!name.startsWith(MCC_SHEET_PREFIX)) continue;
    const normMcc = normalizeCid(name.slice(MCC_SHEET_PREFIX.length));
    if (!normMcc) continue;
    const payload = readMccSheetState(spreadsheet, normMcc);
    if (payload?.checkpoint?.dateStr === dateStr) {
      clearCheckpointRowInMccSheet(s, nowIso);
    }
  }
}
