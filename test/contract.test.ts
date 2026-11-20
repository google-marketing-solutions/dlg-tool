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
 */

import {describe, expect, it} from 'vitest';
import {
  CONTRACT_VERSION,
  parseLockedDailyConfig,
  parseMccCheckpoint,
  serializeLockedDailyConfig,
  serializeMccCheckpoint,
} from '../src/shared/contract';
import type {MccCheckpointState} from '../src/shared/sheets/mcc-state';

const CHECKPOINT: MccCheckpointState = {
  mccId: '1112223333',
  dateStr: '2026-10-14',
  status: 'IN_PROGRESS',
  completedCids: ['1000000001'],
  failedCids: {'1000000002': {error: 'boom', attempts: 2}},
};

describe('cross-script contract', () => {
  it('round-trips a checkpoint and stamps the schema version', () => {
    const raw = serializeMccCheckpoint(CHECKPOINT);
    expect(JSON.parse(raw).schemaVersion).toBe(CONTRACT_VERSION);
    expect(parseMccCheckpoint(raw)).toEqual(CHECKPOINT);
  });

  it('reads checkpoints written before versioning as version 1', () => {
    expect(parseMccCheckpoint(JSON.stringify(CHECKPOINT))).toEqual(CHECKPOINT);
  });

  it('rejects checkpoints from a newer contract, malformed JSON, or the wrong shape', () => {
    const logs: string[] = [];
    const log = (m: string) => logs.push(m);
    expect(parseMccCheckpoint(JSON.stringify({...CHECKPOINT, schemaVersion: CONTRACT_VERSION + 1}), log)).toBeNull();
    expect(logs.at(-1)).toContain('Re-copy the Google Ads Script');
    expect(parseMccCheckpoint('{not json', log)).toBeNull();
    expect(parseMccCheckpoint(JSON.stringify({...CHECKPOINT, status: 'DONE'}), log)).toBeNull();
    expect(parseMccCheckpoint(JSON.stringify({...CHECKPOINT, completedCids: 'all'}), log)).toBeNull();
  });

  it('normalizes CIDs and drops malformed failed-account entries', () => {
    const parsed = parseMccCheckpoint(
      JSON.stringify({
        ...CHECKPOINT,
        completedCids: ['100-000-0001', 'not-a-cid'],
        failedCids: {'100-000-0002': {error: 'x', attempts: 3}, bogus: {error: 'y', attempts: 1}, '1000000003': 'nope'},
      })
    );
    expect(parsed?.completedCids).toEqual(['1000000001']);
    expect(parsed?.failedCids).toEqual({'1000000002': {error: 'x', attempts: 3}});
  });

  it('round-trips a daily config lock, normalizing selected CIDs', () => {
    const raw = serializeLockedDailyConfig({
      targetCurrency: 'EUR',
      excludeShoppingAndPmax: true,
      selectedCidsByMcc: {'111-222-3333': ['100-000-0001', 'junk']},
    });
    expect(parseLockedDailyConfig(raw)).toEqual({
      targetCurrency: 'EUR',
      excludeShoppingAndPmax: true,
      selectedCidsByMcc: {'1112223333': ['1000000001']},
    });
  });

  it('rejects daily config locks from a newer contract or with the wrong shape', () => {
    const log = () => {};
    const base = {targetCurrency: 'EUR', excludeShoppingAndPmax: false};
    expect(parseLockedDailyConfig(JSON.stringify({...base, schemaVersion: CONTRACT_VERSION + 1}), log)).toBeNull();
    expect(parseLockedDailyConfig(JSON.stringify({...base, excludeShoppingAndPmax: 'yes'}), log)).toBeNull();
    expect(parseLockedDailyConfig(JSON.stringify({...base, selectedCidsByMcc: []}), log)).toBeNull();
    expect(parseLockedDailyConfig(JSON.stringify(base), log)).toEqual(base);
  });
});
