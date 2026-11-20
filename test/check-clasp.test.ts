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

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {afterEach, describe, expect, it} from 'vitest';
import {checkClaspConfig} from '../scripts/check-clasp.mjs';

describe('push:sheet clasp guard (scripts/check-clasp.mjs)', () => {
  const dirs: string[] = [];
  afterEach(() => {
    for (const d of dirs.splice(0)) fs.rmSync(d, {recursive: true, force: true});
  });

  function repoWith(claspJson: string | null): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dlg-clasp-check-'));
    dirs.push(dir);
    if (claspJson !== null) fs.writeFileSync(path.join(dir, '.clasp.json'), claspJson);
    return dir;
  }

  it('accepts rootDir dist/sheet (including ./dist/sheet/)', () => {
    expect(checkClaspConfig(repoWith('{"scriptId": "x", "rootDir": "dist/sheet"}'))).toBeNull();
    expect(checkClaspConfig(repoWith('{"scriptId": "x", "rootDir": "./dist/sheet/"}'))).toBeNull();
  });

  it('refuses when .clasp.json is missing and points to the example', () => {
    expect(checkClaspConfig(repoWith(null))).toMatch(/not found.*\.clasp\.json\.example/);
  });

  it('refuses a missing, empty, or wrong rootDir', () => {
    for (const json of [
      '{"scriptId": "x"}',
      '{"scriptId": "x", "rootDir": ""}',
      '{"scriptId": "x", "rootDir": "."}',
      '{"scriptId": "x", "rootDir": "dist"}',
      '{"scriptId": "x", "rootDir": "src"}',
    ]) {
      expect(checkClaspConfig(repoWith(json))).toMatch(/Refusing to push.*\.clasp\.json\.example/);
    }
  });

  it('refuses invalid JSON', () => {
    expect(checkClaspConfig(repoWith('{not json'))).toMatch(/not valid JSON/);
  });

  it('matches the committed .clasp.json.example', () => {
    const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
    const example = fs.readFileSync(path.join(repoRoot, '.clasp.json.example'), 'utf8');
    expect(checkClaspConfig(repoWith(example))).toBeNull();
  });
});
