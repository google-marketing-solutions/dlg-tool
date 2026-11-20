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
 * @fileoverview Date formatting utilities shared by `dist/ads/Code.gs` and `dist/sheet/Code.gs`.
 */

/**
 * Resolves the current calendar date (`YYYY-MM-DD`) in the spreadsheet's timezone
 * (falling back to `'UTC'` when `getSpreadsheetTimeZone()` returns `null` or empty).
 */
export function resolveSpreadsheetDate(spreadsheet: {
  getSpreadsheetTimeZone(): string | null;
}): string {
  return Utilities.formatDate(
    new Date(),
    spreadsheet.getSpreadsheetTimeZone() || 'UTC',
    'yyyy-MM-dd'
  );
}
