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
 * @fileoverview Daily-locked currency exchange rate resolver and row converter.
 * Resolves FX rates in scratch cell `_RunState!G1` from today's daily cache
 * (`FX_RATES_<date>_<target>`), live `=GOOGLEFINANCE`
 * (with retry backoff and `FX_LAST_GOOD_<src>_<target>` persistence only on live success), fallback to
 * `FX_LAST_GOOD_<src>_<target>`, and finally `null` (never fabricating a `1.0` rate).
 * Bundled only into `dist/sheet/Code.gs`.
 */

import {
  type SpreadsheetLike,
  getOrCreateSheet,
} from '../shared/sheets/sheet-utils';
import type {ReportRow} from '../shared/types';
import {normalizeCurrencyCode} from '../shared/utils/currency';
import {
  RUN_STATE_SHEET_NAME,
  getRunStateValue,
  readRunStateMap,
  setRunStateValues,
} from './run-state';

/** Maximum attempts when evaluating a live `=GOOGLEFINANCE` formula in `_RunState!G1`. */
export const FX_MAX_RETRIES = 3;
/** Scratch cell in `_RunState` used to evaluate `=GOOGLEFINANCE("CURRENCY:<SRC><DST>")`. */
export const FX_SCRATCH_CELL = 'G1';

/** Options for resolving and caching daily FX rates in `_RunState`. */
export interface ResolveFxOptions {
  spreadsheet: SpreadsheetLike;
  dateStr: string;
  sourceCurrencies: string[];
  targetCurrency: string;
}

/**
 * Resolves and locks daily FX rates from each `sourceCurrency` to `targetCurrency`.
 * Resolution precedence:
 * 1. Identity (`src === target` -> `1`)
 * 2. Daily cache (`FX_RATES_<date>_<target>`)
 * 3. Live `=GOOGLEFINANCE` in `_RunState!G1` (updates `FX_LAST_GOOD_<src>_<target>` only on live success)
 * 4. Last known good rate (`FX_LAST_GOOD_<src>_<target>`)
 * 5. `null` (leaves `*Global` monetary columns as `null` rather than fabricating `1.0`).
 * Only 3-letter currency codes are used: a target that is not one (empty, `NONE`) keeps native values
 * (rate `1`), and a source that is not one gets a `null` rate, so no formula is ever built from other input.
 * All `_RunState` updates (`FX_LAST_GOOD_*` and the daily cache) are flushed in a single write at the end,
 * and nothing is written when no rate changed.
 */
export function resolveDailyFxRates(options: ResolveFxOptions): Record<string, number | null> {
  const target = normalizeCurrencyCode(options.targetCurrency);
  const uniqueSources = Array.from(
    new Set(
      options.sourceCurrencies
        .map(normalizeCurrencyCode)
        .filter((c): c is string => c !== null)
    )
  );

  if (!target) {
    return Object.fromEntries(uniqueSources.map((c) => [c, 1]));
  }

  const dailyCacheKey = `FX_RATES_${options.dateStr}_${target}`;
  const stateMap = readRunStateMap(options.spreadsheet);
  const cachedJson = getRunStateValue(options.spreadsheet, dailyCacheKey, stateMap);
  const cachedMap: Record<string, number | null> = cachedJson ? safeParseRates(cachedJson) : {};

  let updated = false;
  const lastGoodUpdates: Array<{key: string; value: string}> = [];
  for (const src of uniqueSources) {
    if (src === target) {
      if (cachedMap[src] !== 1) {
        cachedMap[src] = 1;
        updated = true;
      }
      continue;
    }

    // 1. Daily cached rate for today
    if (cachedMap[src]) {
      continue;
    }

    // 2. Live GOOGLEFINANCE with fallback to FX_LAST_GOOD_<src>_<target>
    const { rate, fromLiveGoogleFinance } = fetchSinglePairRate(
      options.spreadsheet,
      src,
      target,
      stateMap
    );
    cachedMap[src] = rate;
    // Only live GOOGLEFINANCE results update FX_LAST_GOOD (never overrides or fallback values)
    if (fromLiveGoogleFinance && rate !== null) {
      lastGoodUpdates.push({key: `FX_LAST_GOOD_${src}_${target}`, value: String(rate)});
    }
    updated = true;
  }

  if (updated) {
    setRunStateValues(
      options.spreadsheet,
      [...lastGoodUpdates, {key: dailyCacheKey, value: JSON.stringify(cachedMap)}],
      options.dateStr
    );
  }

  return cachedMap;
}

function fetchSinglePairRate(
  spreadsheet: SpreadsheetLike,
  fromCurrency: string,
  toCurrency: string,
  stateMap?: ReadonlyMap<string, string>
): { rate: number | null; fromLiveGoogleFinance: boolean } {
  if (!normalizeCurrencyCode(fromCurrency) || !normalizeCurrencyCode(toCurrency)) {
    return { rate: null, fromLiveGoogleFinance: false };
  }
  const sheet = getOrCreateSheet(spreadsheet, RUN_STATE_SHEET_NAME, { hidden: true });
  const scratchRange = sheet.getRange(FX_SCRATCH_CELL);
  const formula = `=GOOGLEFINANCE("CURRENCY:${fromCurrency}${toCurrency}")`;

  try {
    // Set the formula once and poll it, so an in-flight "Loading..." evaluation gets time to resolve
    scratchRange.setFormula(formula);
    for (let attempt = 1; attempt <= FX_MAX_RETRIES; attempt++) {
      SpreadsheetApp.flush();

      const rawVal = scratchRange.getValue();
      const numericVal = Number(rawVal);
      if (Number.isFinite(numericVal) && numericVal > 0) {
        return { rate: numericVal, fromLiveGoogleFinance: true };
      }

      if (attempt < FX_MAX_RETRIES) {
        Utilities.sleep(attempt * 1000);
      }
    }
  } finally {
    scratchRange.setValue('');
  }

  const lastKnownGood = Number(
    getRunStateValue(spreadsheet, `FX_LAST_GOOD_${fromCurrency}_${toCurrency}`, stateMap)
  );
  if (Number.isFinite(lastKnownGood) && lastKnownGood > 0) {
    return { rate: lastKnownGood, fromLiveGoogleFinance: false };
  }

  return { rate: null, fromLiveGoogleFinance: false };
}

function safeParseRates(json: string): Record<string, number | null> {
  try {
    return JSON.parse(json) as Record<string, number | null>;
  } catch {
    return {};
  }
}

/**
 * Converts all `*Global` fields in a row into `targetCurrency`. When the FX rate is `null`
 * (unavailable), `*Global` monetary columns are set to `null` (rather than using `1.0` to avoid
 * miscalculations). Returns the same row with converted values.
 */
export function applyCurrencyConversionToRow(
  row: ReportRow,
  targetCurrency: string,
  rates: Record<string, number | null>
): ReportRow {
  const normTarget = normalizeCurrencyCode(targetCurrency);
  const srcCurr = (row.currencyCode || 'USD').trim().toUpperCase();

  if (!normTarget) {
    return syncGlobalColumnsWithRate(row, srcCurr, 1);
  }

  const resolvedRate = srcCurr === normTarget ? 1 : (rates[srcCurr] ?? null);
  const validRate =
    typeof resolvedRate === 'number' && Number.isFinite(resolvedRate) && resolvedRate > 0
      ? resolvedRate
      : null;

  return syncGlobalColumnsWithRate(row, normTarget, validRate);
}

function syncGlobalColumnsWithRate(
  row: ReportRow,
  currencyCodeGlobal: string,
  rate: number | null
): ReportRow {
  return {
    ...row,
    currencyCodeGlobal,
    recommendationCurrentBudgetAmountGlobal: multiplyNullable(
      row.recommendationCurrentBudgetAmount,
      rate
    ),
    recommendationNewBudgetAmountGlobal: multiplyNullable(
      row.recommendationNewBudgetAmount,
      rate
    ),
    recommendationBaseConversionsValueGlobal: multiplyNullable(
      row.recommendationBaseConversionsValue,
      rate
    ),
    recommendationPotentialConversionsValueGlobal: multiplyNullable(
      row.recommendationPotentialConversionsValue,
      rate
    ),
    weeklyCostIncreaseGlobal: multiplyNullable(row.weeklyCostIncrease, rate),
    moveBudgetAmountGlobal: multiplyNullable(row.moveBudgetAmount, rate),
  };
}

/**
 * Multiplies a nullable monetary value by `rate` while preserving `null`, and rounds to 2 decimal places
 * so that row-level and MCC-level sums in Data Studio never suffer from floating-point penny drift.
 * In JS/TS, `null * rate` would coerce `null` to `0`, falsely turning missing/N/A fields into 0.00.
 */
function multiplyNullable(val: number | null, rate: number | null): number | null {
  if (rate === null || !Number.isFinite(rate)) {
    return null;
  }
  if (typeof val !== 'number' || !Number.isFinite(val)) {
    return null;
  }
  return Math.round((val * rate + Number.EPSILON) * 100) / 100;
}
