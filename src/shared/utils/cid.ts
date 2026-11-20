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
 * @fileoverview Google Ads Customer ID (CID) formatting and normalization utilities.
 * Ensures all Map/Set lookups use canonical 10-digit strings (`normalizeCid`: `"1234567890"`)
 * while UI tables and `Data` rows display hyphenated IDs (`formatCidHyphenated`: `"123-456-7890"`).
 * Shared by both `dist/ads/Code.gs` and `dist/sheet/Code.gs`.
 */

/** Canonical hyphenated Google Ads Customer ID shape (`XXX-XXX-XXXX`). */
const HYPHENATED_CID_RE = /^\d{3}-\d{3}-\d{4}$/;

/**
 * Normalizes a Google Ads Customer ID into a canonical 10-digit string
 * (e.g. `"123-456-7890"` -> `"1234567890"`, `12345678` -> `"0012345678"`).
 *
 * Accepts only real CIDs: digits with optional whitespace (up to 10 digits; shorter values are
 * left-padded because Sheets strips leading zeros from numeric cells), or the exact hyphenated
 * `XXX-XXX-XXXX` form. Anything else (labels such as `"MCC 1"`, dates such as `"2026-10-14"`,
 * numbers with more than 10 digits, or an all-zero ID) returns `''`.
 */
export function normalizeCid(rawCid: unknown): string {
  if (rawCid === null || rawCid === undefined) {
    return '';
  }
  const trimmed = String(rawCid).trim();
  // Hyphens are only stripped from the exact XXX-XXX-XXXX shape, so dates like 2026-10-14 are rejected
  const digits = HYPHENATED_CID_RE.test(trimmed)
    ? trimmed.replace(/-/g, '')
    : trimmed.replace(/\s+/g, '');
  if (!/^\d{1,10}$/.test(digits) || /^0+$/.test(digits)) {
    return '';
  }
  return digits.padStart(10, '0');
}

/**
 * Formats a Customer ID into the standard `XXX-XXX-XXXX` display format, or returns `''` when the
 * input is not a valid CID (see `normalizeCid`).
 */
export function formatCidHyphenated(rawCid: unknown): string {
  const digits = normalizeCid(rawCid);
  if (digits.length !== 10) {
    return digits;
  }
  return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 10)}`;
}
