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
 * @fileoverview Build version of the bundle (`package.json` `version`, injected by `scripts/build.mjs`).
 * Each MCC's Google Ads Script records its version in `_MCC_<id>` so the Sheet can flag outdated copies.
 */

declare const __DLG_VERSION__: string | undefined;

/** Version of this bundle, or `dev` when running unbundled (unit tests). */
export const DLG_VERSION: string =
  typeof __DLG_VERSION__ !== 'undefined' ? __DLG_VERSION__ : 'dev';
