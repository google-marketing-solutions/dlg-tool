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
 * @fileoverview Versioned contract for the JSON the two scripts exchange through Sheet cells: the per-MCC
 * checkpoint (`_MCC_<id>` `CHECKPOINT` row, written by `dist/ads/Code.gs`) and the daily config lock
 * (`_RunState` `DAILY_CONFIG_LOCK_<date>`, written by `dist/sheet/Code.gs`). The Google Ads Script is pasted
 * by hand into each MCC while the Sheet script is updated with `clasp push`, so the two can run different
 * versions: every payload carries `schemaVersion`, and readers validate the shape instead of trusting a cast.
 */

import type {MccCheckpointState} from './sheets/mcc-state';
import type {GlobalConfig} from './types';
import {normalizeCid} from './utils/cid';

/** Version of the checkpoint and daily-lock JSON written by this build. Payloads without one are version 1. */
export const CONTRACT_VERSION = 1;

type JsonRecord = Record<string, unknown>;

function isRecord(value: unknown): value is JsonRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((v) => typeof v === 'string');
}

/** Parses `raw` as a JSON object with a supported `schemaVersion`, or returns `null` (and logs why). */
function parseVersionedRecord(raw: string, label: string, log: (msg: string) => void): JsonRecord | null {
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    log(`[DLG v2] ${label} is not valid JSON; ignoring it.`);
    return null;
  }
  if (!isRecord(value)) {
    log(`[DLG v2] ${label} is not a JSON object; ignoring it.`);
    return null;
  }
  const version = value['schemaVersion'] ?? 1;
  if (version !== CONTRACT_VERSION) {
    log(
      `[DLG v2] ${label} has schemaVersion ${String(version)}, but this script reads version ${CONTRACT_VERSION}. Re-copy the Google Ads Script from the Google Sheet menu (📋 Copy Google Ads Script) so both scripts match.`
    );
    return null;
  }
  return value;
}

/** Serializes a checkpoint with the current `schemaVersion`. */
export function serializeMccCheckpoint(checkpoint: MccCheckpointState): string {
  return JSON.stringify({ schemaVersion: CONTRACT_VERSION, ...checkpoint });
}

/**
 * Parses a `CHECKPOINT` cell. Returns `null` for malformed JSON, an unexpected shape, or an unsupported
 * `schemaVersion` (the MCC then starts today's run again, which is safe because staging writes are idempotent).
 */
export function parseMccCheckpoint(
  raw: string,
  log: (msg: string) => void = defaultLog
): MccCheckpointState | null {
  const v = parseVersionedRecord(raw, 'MCC CHECKPOINT', log);
  if (!v) return null;
  const status = v['status'];
  if (
    typeof v['mccId'] !== 'string' ||
    typeof v['dateStr'] !== 'string' ||
    (status !== 'IN_PROGRESS' && status !== 'COMPLETED') ||
    !isStringArray(v['completedCids'])
  ) {
    log('[DLG v2] MCC CHECKPOINT has an unexpected shape; ignoring it.');
    return null;
  }
  const failedCids = parseFailedCids(v['failedCids']);
  const lastCompletedDate = v['lastCompletedDate'];
  return {
    mccId: v['mccId'],
    dateStr: v['dateStr'],
    status,
    completedCids: v['completedCids'].map(normalizeCid).filter(Boolean),
    ...(failedCids ? { failedCids } : {}),
    ...(typeof lastCompletedDate === 'string' && lastCompletedDate ? { lastCompletedDate } : {}),
  };
}

function parseFailedCids(value: unknown): MccCheckpointState['failedCids'] {
  if (!isRecord(value)) return undefined;
  const out: NonNullable<MccCheckpointState['failedCids']> = {};
  for (const [rawCid, info] of Object.entries(value)) {
    const cid = normalizeCid(rawCid);
    if (!cid || !isRecord(info)) continue;
    const attempts = Number(info['attempts']);
    out[cid] = {
      error: String(info['error'] ?? ''),
      attempts: Number.isFinite(attempts) && attempts > 0 ? attempts : 1,
    };
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** Serializes a daily config lock with the current `schemaVersion`. */
export function serializeLockedDailyConfig(config: GlobalConfig): string {
  return JSON.stringify({ schemaVersion: CONTRACT_VERSION, ...config });
}

/**
 * Parses a `DAILY_CONFIG_LOCK_<date>` value. Returns `null` for malformed JSON, an unexpected shape, or an
 * unsupported `schemaVersion` (the Google Ads Script then waits for a lock it can read).
 */
export function parseLockedDailyConfig(
  raw: string,
  log: (msg: string) => void = defaultLog
): GlobalConfig | null {
  const v = parseVersionedRecord(raw, 'DAILY_CONFIG_LOCK', log);
  if (!v) return null;
  if (typeof v['targetCurrency'] !== 'string' || typeof v['excludeShoppingAndPmax'] !== 'boolean') {
    log('[DLG v2] DAILY_CONFIG_LOCK has an unexpected shape; ignoring it.');
    return null;
  }
  const config: GlobalConfig = {
    targetCurrency: v['targetCurrency'],
    excludeShoppingAndPmax: v['excludeShoppingAndPmax'],
  };
  const rawSelected = v['selectedCidsByMcc'];
  if (rawSelected !== undefined) {
    if (!isRecord(rawSelected)) {
      log('[DLG v2] DAILY_CONFIG_LOCK selectedCidsByMcc is not an object; ignoring the lock.');
      return null;
    }
    const selectedCidsByMcc: Record<string, string[]> = {};
    for (const [rawMcc, cids] of Object.entries(rawSelected)) {
      const mcc = normalizeCid(rawMcc);
      if (!mcc || !isStringArray(cids)) continue;
      selectedCidsByMcc[mcc] = cids.map(normalizeCid).filter(Boolean);
    }
    config.selectedCidsByMcc = selectedCidsByMcc;
  }
  return config;
}

function defaultLog(msg: string): void {
  if (typeof Logger !== 'undefined') {
    Logger.log(msg);
  }
}
