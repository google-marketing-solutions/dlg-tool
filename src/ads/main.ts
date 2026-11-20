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
 * @fileoverview Top-level entrypoint for the Google Ads Script (`dist/ads/Code.gs`).
 * Validates that `SPREADSHEET_URL` and `SPREADSHEET_TIMEZONE` have been injected by the Google Sheet
 * and invokes `runAdsScriptOrchestrator`.
 */

import {runAdsScriptOrchestrator} from './orchestrator';

declare const SPREADSHEET_URL: string | undefined;
declare const SPREADSHEET_TIMEZONE: string | undefined;

/**
 * Entrypoint for the DLG v2 Google Ads Script (`Code.gs`).
 */
export function main(): void {
  const url = typeof SPREADSHEET_URL !== 'undefined' ? SPREADSHEET_URL : '';
  if (!url || url === 'YOUR_SPREADSHEET_URL_HERE') {
    throw new Error(
      'SPREADSHEET_URL is not configured. Use "📋 Copy Google Ads Script" in your copy of the Master Template Google Sheet.'
    );
  }
  const timeZone = typeof SPREADSHEET_TIMEZONE !== 'undefined' ? SPREADSHEET_TIMEZONE : '';
  if (!timeZone || timeZone === 'YOUR_SPREADSHEET_TIMEZONE_HERE') {
    throw new Error(
      'SPREADSHEET_TIMEZONE is not configured. Use "📋 Copy Google Ads Script" in your copy of the Master Template Google Sheet.'
    );
  }
  runAdsScriptOrchestrator({ spreadsheetUrl: url, timeZone });
}
