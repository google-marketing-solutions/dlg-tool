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
 * @fileoverview ISO 4217 currency code validation shared by `dist/ads/Code.gs` and `dist/sheet/Code.gs`.
 */

const ISO_4217_RE = /^[A-Z]{3}$/;

/**
 * Returns the upper-cased 3-letter currency code in `raw`, or `null` when `raw` is empty, a placeholder such
 * as `NONE`, or anything else that is not a 3-letter code. Currency codes end up inside Sheets formulas
 * (`GOOGLEFINANCE`), so only validated codes may be used there.
 */
export function normalizeCurrencyCode(raw: unknown): string | null {
  const code = String(raw ?? '').trim().toUpperCase();
  return ISO_4217_RE.test(code) ? code : null;
}
