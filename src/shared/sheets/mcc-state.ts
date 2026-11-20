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
 * @fileoverview Per-MCC hidden state tab (`_MCC_<normMccId>`) reader, writer, and checkpoint manager.
 * Each MCC's Google Ads Script (`dist/ads/Code.gs`) writes its discovered child accounts, checkpoint state,
 * and status message exclusively to its own `_MCC_<normMccId>` sheet, while the Sheet-bound script
 * (`dist/sheet/Code.gs`) reads these tabs to sync discovered accounts to `Config`, check daily completion,
 * and clear the `CHECKPOINT` row when resetting an MCC.
 */

import {parseMccCheckpoint, serializeMccCheckpoint} from '../contract';
import type {ChildAccountInput, MccIdentity} from '../types';
import {DLG_VERSION} from '../version';
import {formatCidHyphenated, normalizeCid} from '../utils/cid';
import {type SheetLike, type SpreadsheetLike, getOrCreateSheet} from './sheet-utils';

/** Prefix for per-MCC hidden state tabs (`_MCC_<normMccId>`) written by `Code.gs`. */
export const MCC_SHEET_PREFIX = '_MCC_';
/** Runs in which a failing child account is retried before it is given up for the day. */
export const MAX_ACCOUNT_ATTEMPTS = 3;

/** A child account that failed during today's run. */
export interface FailedAccountInfo {
  /** Last error message. */
  error: string;
  /** Number of separate runs in which the account failed today. */
  attempts: number;
}

/** Checkpoint progress state for a single MCC on a given date (`YYYY-MM-DD`). */
export interface MccCheckpointState {
  mccId: string;
  dateStr: string;
  status: 'IN_PROGRESS' | 'COMPLETED';
  completedCids: string[];
  /**
   * Accounts that failed today, keyed by normalized CID. Failed accounts stay out of `completedCids` until
   * they succeed or are given up after `MAX_ACCOUNT_ATTEMPTS` runs (then they are in both).
   */
  failedCids?: Record<string, FailedAccountInfo>;
  lastCompletedDate?: string;
}

/** Payload stored in a dedicated per-MCC hidden sheet (`_MCC_<normMccId>`). */
export interface MccSheetPayload {
  mcc: MccIdentity;
  /**
   * Discovered child accounts. Omit to keep the accounts already stored in `_MCC_<id>`; pass `[]` to
   * clear them (for example when every account was unlinked from the MCC).
   */
  childAccounts?: ChildAccountInput[];
  statusMessage?: string;
  checkpoint?: MccCheckpointState;
  lastCompletedDate?: string;
  /**
   * Version of the script that last wrote the tab (read-only: `writeMccSheetState` always records the
   * running bundle's `DLG_VERSION`). Missing for tabs written before versions were recorded.
   */
  scriptVersion?: string;
}

/**
 * Returns the hidden per-MCC state sheet name (`_MCC_<10DigitMccId>`) for `mccId`.
 */
export function getMccSheetName(mccId: string): string {
  return `${MCC_SHEET_PREFIX}${normalizeCid(mccId)}`;
}

/**
 * Writes an MCC's discovered child accounts, checkpoint state, and latest status banner message
 * into its dedicated hidden `_MCC_<normMccId>` sheet.
 * Because each MCC script writes only to its own `_MCC_<normMccId>` sheet, concurrent MCC executions
 * never contend on `Config` or `_RunState`.
 */
export function writeMccSheetState(
  spreadsheet: SpreadsheetLike,
  payload: MccSheetPayload
): void {
  const normMcc = normalizeCid(payload.mcc.mccId);
  const formattedMcc = formatCidHyphenated(normMcc);
  const sheetName = getMccSheetName(normMcc);
  const sheet = getOrCreateSheet(spreadsheet, sheetName, { hidden: true });

  const existing = readMccSheetState(spreadsheet, normMcc);
  const checkpoint = payload.checkpoint ?? existing?.checkpoint;
  const lastCompletedDate =
    payload.lastCompletedDate ??
    (checkpoint?.status === 'COMPLETED' ? checkpoint.dateStr : undefined) ??
    existing?.lastCompletedDate ??
    '';
  const statusMessage = payload.statusMessage ?? existing?.statusMessage ?? '';
  const childAccounts = payload.childAccounts ?? existing?.childAccounts ?? [];

  const nowIso = new Date().toISOString();
  const rows: string[][] = [
    ['Key', 'Value', 'Extra1', 'Extra2'],
    ['MCC_META', formattedMcc, payload.mcc.mccName, nowIso],
    ['CHECKPOINT', checkpoint ? serializeMccCheckpoint(checkpoint) : '', checkpoint?.dateStr ?? '', nowIso],
    ['LAST_COMPLETED_DATE', lastCompletedDate, '', nowIso],
    ['STATUS_MESSAGE', statusMessage, '', nowIso],
    ['SCRIPT_VERSION', DLG_VERSION, '', nowIso],
  ];

  for (const acc of childAccounts) {
    rows.push([
      'CHILD_ACCOUNT',
      formatCidHyphenated(acc.customerId),
      acc.name,
      String(acc.currencyCode ?? 'USD').trim().toUpperCase(),
    ]);
  }

  const prevLastRow = Math.max(1, sheet.getLastRow());
  const writeRowCount = Math.max(prevLastRow, rows.length);
  if (writeRowCount > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), writeRowCount - sheet.getMaxRows() + 5);
  }

  while (rows.length < writeRowCount) {
    rows.push(['', '', '', '']);
  }

  const range = sheet.getRange(1, 1, rows.length, 4);
  range.setNumberFormat('@');
  range.setValues(rows);

  if (prevLastRow > rows.length) {
    sheet.deleteRows(rows.length + 1, prevLastRow - rows.length);
  }
}

/**
 * Reads an MCC's discovery & checkpoint state from its dedicated `_MCC_<normMccId>` sheet.
 */
export function readMccSheetState(
  spreadsheet: SpreadsheetLike,
  mccId: string
): MccSheetPayload | null {
  const normMcc = normalizeCid(mccId);
  const sheet = spreadsheet.getSheetByName(getMccSheetName(normMcc));
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rawRows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  let formattedMcc = formatCidHyphenated(normMcc);
  let mccName = 'MCC';
  let checkpoint: MccCheckpointState | undefined;
  let lastCompletedDate: string | undefined;
  let statusMessage: string | undefined;
  let scriptVersion: string | undefined;
  const childAccounts: ChildAccountInput[] = [];

  for (const r of rawRows) {
    const key = String(r[0] ?? '').trim();
    const rawVal = r[1];
    const val =
      rawVal instanceof Date
        ? rawVal.toISOString().slice(0, 10)
        : String(rawVal ?? '').trim();
    const extra1 = String(r[2] ?? '').trim();
    const extra2 = String(r[3] ?? '').trim();

    if (key === 'MCC_META') {
      if (val) formattedMcc = val;
      if (extra1) mccName = extra1;
    } else if (key === 'CHECKPOINT' && val) {
      checkpoint = parseMccCheckpoint(val) ?? undefined;
    } else if (key === 'SCRIPT_VERSION' && val) {
      scriptVersion = val;
    } else if (key === 'LAST_COMPLETED_DATE' && val) {
      lastCompletedDate = val;
    } else if (key === 'STATUS_MESSAGE' && val) {
      statusMessage = val;
    } else if (key === 'CHILD_ACCOUNT' && val) {
      childAccounts.push({
        customerId: val,
        name: extra1 || 'Account',
        currencyCode: extra2 || 'USD',
      });
    }
  }

  return {
    mcc: { mccId: formattedMcc, mccName },
    childAccounts,
    statusMessage,
    checkpoint,
    lastCompletedDate,
    scriptVersion,
  };
}

/**
 * Reads the daily checkpoint for `mccId` on `dateStr` from the MCC's dedicated `_MCC_<normMccId>` sheet
 * (written by `Code.gs`). Returns an `IN_PROGRESS` state with no completed accounts when the sheet has no
 * checkpoint for `dateStr` (or does not exist).
 */
export function getMccCheckpoint(
  spreadsheet: SpreadsheetLike,
  dateStr: string,
  mccId: string
): MccCheckpointState {
  const normMcc = normalizeCid(mccId);
  const mccPayload = readMccSheetState(spreadsheet, normMcc);
  const lastCompletedDate =
    mccPayload?.lastCompletedDate ?? mccPayload?.checkpoint?.lastCompletedDate;
  if (mccPayload?.checkpoint && mccPayload.checkpoint.dateStr === dateStr) {
    return {
      ...mccPayload.checkpoint,
      mccId: normMcc,
      lastCompletedDate,
    };
  }
  return {
    mccId: normMcc,
    dateStr,
    status: 'IN_PROGRESS',
    completedCids: [],
    lastCompletedDate,
  };
}

/**
 * Clears the `CHECKPOINT` row of an MCC's `_MCC_<normMccId>` sheet, keeping `LAST_COMPLETED_DATE`
 * and all other rows. No-op when the sheet or row does not exist.
 */
export function clearMccSheetCheckpoint(spreadsheet: SpreadsheetLike, mccId: string): void {
  const sheet = spreadsheet.getSheetByName(getMccSheetName(normalizeCid(mccId)));
  if (!sheet) return;
  clearCheckpointRowInMccSheet(sheet, new Date().toISOString());
}

/** Finds the `CHECKPOINT` row by its key in column A (never by position) and blanks its value. */
export function clearCheckpointRowInMccSheet(sheet: SheetLike, nowIso: string): void {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  const keys = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0] ?? '').trim() === 'CHECKPOINT') {
      sheet.getRange(i + 1, 1, 1, 4).setValues([['CHECKPOINT', '', '', nowIso]]);
      return;
    }
  }
}
