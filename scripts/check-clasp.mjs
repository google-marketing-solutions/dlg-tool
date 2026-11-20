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
 * @fileoverview Safety guard run by `npm run push:sheet` before `clasp push --force`.
 * Refuses to push unless `.clasp.json` exists and its `rootDir` resolves to `dist/sheet`, so a
 * misconfigured `.clasp.json` can never push raw `src/` / `test/` files over the Sheet's script.
 */

import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/**
 * Returns an error message when `.clasp.json` at `repoRoot` is missing or does not point at
 * `dist/sheet`, or `null` when it is safe to push.
 * @param {string} repoRoot
 * @return {string | null}
 */
export function checkClaspConfig(repoRoot) {
  const claspJson = path.join(repoRoot, '.clasp.json');
  const hint = 'Copy .clasp.json.example to .clasp.json (it sets "rootDir": "dist/sheet").';
  if (!fs.existsSync(claspJson)) {
    return `.clasp.json not found at ${claspJson}. ${hint}`;
  }
  /** @type {{rootDir?: unknown}} */
  let config;
  try {
    config = JSON.parse(fs.readFileSync(claspJson, 'utf8'));
  } catch (e) {
    return `.clasp.json is not valid JSON (${e instanceof Error ? e.message : String(e)}). ${hint}`;
  }
  const rootDir = typeof config.rootDir === 'string' ? config.rootDir : '';
  const resolved = path.resolve(repoRoot, rootDir);
  if (!rootDir || resolved !== path.join(repoRoot, 'dist', 'sheet')) {
    return (
      `.clasp.json "rootDir" is ${JSON.stringify(config.rootDir ?? null)}, which resolves to ` +
      `${resolved}, not dist/sheet. Refusing to push: clasp would upload the wrong files and ` +
      `overwrite the Sheet's script. ${hint}`
    );
  }
  return null;
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const error = checkClaspConfig(REPO_ROOT);
  if (error) {
    console.error(`push:sheet aborted: ${error}`);
    process.exit(1);
  }
  console.log('.clasp.json OK (rootDir -> dist/sheet).');
}
