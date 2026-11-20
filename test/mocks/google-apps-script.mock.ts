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

/**
 * In-memory mocks for Google Apps Script services (`SpreadsheetApp`, `AdsApp`, `AdsManagerApp`).
 * Simulates real Google Sheets cell coercion (stripping leading zeros on numeric strings when format !== '@'),
 * row/column shifting (values, formats, and validations) on `deleteRows`/`deleteColumns`, `DataValidation.getCriteriaValues()`, and 3-arg `ui.alert(title, prompt, buttons)`.
 */

/** In-memory mock for a Google Sheets `DataValidation` rule. */
export interface MockDataValidation {
  type: 'BOOLEAN' | 'VALUE_IN_LIST';
  values?: string[];
  allowInvalid?: boolean;
  getCriteriaValues(): unknown[];
}

/** In-memory mock for a Google Sheets `Range`. */
export class MockRange {
  constructor(
    private readonly sheet: MockSheet,
    private readonly startRow: number,
    private readonly startCol: number,
    private readonly numRows: number,
    private readonly numCols: number
  ) {}

  getRow(): number {
    return this.startRow;
  }

  getColumn(): number {
    return this.startCol;
  }

  getNumRows(): number {
    return this.numRows;
  }

  getNumColumns(): number {
    return this.numCols;
  }

  getSheet(): MockSheet {
    return this.sheet;
  }

  getValue(): unknown {
    return this.sheet.getCellValue(this.startRow, this.startCol);
  }

  setValue(value: unknown): this {
    this.sheet.valueWriteCount++;
    this.sheet.setCellValue(this.startRow, this.startCol, value);
    return this;
  }

  getValues(): unknown[][] {
    const out: unknown[][] = [];
    for (let r = 0; r < this.numRows; r++) {
      const row: unknown[] = [];
      for (let c = 0; c < this.numCols; c++) {
        row.push(this.sheet.getCellValue(this.startRow + r, this.startCol + c));
      }
      out.push(row);
    }
    return out;
  }

  setValues(values: unknown[][]): this {
    this.sheet.valueWriteCount++;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        this.sheet.setCellValue(this.startRow + r, this.startCol + c, values[r][c]);
      }
    }
    return this;
  }

  setFormula(formula: string): this {
    this.sheet.setCellFormula(this.startRow, this.startCol, formula);
    return this;
  }

  setNumberFormat(format: string): this {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.setCellFormat(this.startRow + r, this.startCol + c, format);
      }
    }
    return this;
  }

  setNumberFormats(formats: string[][]): this {
    this.sheet.setNumberFormatsCallCount++;
    for (let r = 0; r < formats.length; r++) {
      for (let c = 0; c < formats[r].length; c++) {
        this.sheet.setCellFormat(this.startRow + r, this.startCol + c, formats[r][c]);
      }
    }
    return this;
  }

  setDataValidation(
    rule: MockDataValidation | GoogleAppsScript.Spreadsheet.DataValidation | null
  ): this {
    this.sheet.setDataValidationCallCount++;
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.setCellValidation(this.startRow + r, this.startCol + c, rule as MockDataValidation | null);
      }
    }
    return this;
  }

  getDataValidation(): MockDataValidation | null {
    return this.sheet.getCellValidation(this.startRow, this.startCol);
  }

  insertCheckboxes(): this {
    return this.setDataValidation({
      type: 'BOOLEAN',
      getCriteriaValues: () => [],
    });
  }

  clearContent(): this {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.setCellValue(this.startRow + r, this.startCol + c, '');
      }
    }
    return this;
  }
}

/** In-memory mock for a Google Sheets `Sheet` tab. */
export class MockSheet {
  private cells = new Map<string, unknown>();
  private formats = new Map<string, string>();
  private validations = new Map<string, MockDataValidation | null>();
  private maxRows = 100;
  private maxCols = 26;
  private hidden = false;
  setNumberFormatsCallCount = 0;
  setDataValidationCallCount = 0;
  /** Number of `Range.setValue` / `setValues` calls on this sheet (for zero-write assertions). */
  valueWriteCount = 0;
  /** Evaluates a formula when it is set (`setFormula`). */
  formulaEvaluator?: (formula: string) => unknown;
  /**
   * Optional read-side evaluator for formula cells: called on every read with a 1-based read index since
   * the formula was set, to simulate values such as `Loading...` that resolve on a later read.
   */
  formulaReadEvaluator?: (formula: string, readIndex: number) => unknown;
  private formulas = new Map<string, string>();
  private formulaReads = new Map<string, number>();

  constructor(
    private name: string,
    private readonly sheetId: number
  ) {}

  private key(row: number, col: number): string {
    return `${row}:${col}`;
  }

  getName(): string {
    return this.name;
  }

  setName(name: string): this {
    this.name = name;
    return this;
  }

  getSheetId(): number {
    return this.sheetId;
  }

  hideSheet(): this {
    this.hidden = true;
    return this;
  }

  isSheetHidden(): boolean {
    return this.hidden;
  }

  getMaxRows(): number {
    return this.maxRows;
  }

  getMaxColumns(): number {
    return this.maxCols;
  }

  getLastRow(): number {
    let maxRow = 0;
    for (const [k, v] of this.cells.entries()) {
      if (v !== '' && v !== null && v !== undefined) {
        const r = Number(k.split(':')[0]);
        if (r > maxRow) maxRow = r;
      }
    }
    return maxRow;
  }

  getLastColumn(): number {
    let maxCol = 0;
    for (const [k, v] of this.cells.entries()) {
      if (v !== '' && v !== null && v !== undefined) {
        const c = Number(k.split(':')[1]);
        if (c > maxCol) maxCol = c;
      }
    }
    return maxCol;
  }

  insertRowsAfter(afterPosition: number, howMany: number): this {
    if (afterPosition < 1 || afterPosition > this.maxRows) {
      throw new Error(`insertRowsAfter out of bounds: ${afterPosition} (maxRows=${this.maxRows})`);
    }
    this.maxRows += howMany;
    return this;
  }

  insertColumnsAfter(afterPosition: number, howMany: number): this {
    if (afterPosition < 1 || afterPosition > this.maxCols) {
      throw new Error(`insertColumnsAfter out of bounds: ${afterPosition} (maxCols=${this.maxCols})`);
    }
    this.maxCols += howMany;
    return this;
  }

  deleteRows(rowPosition: number, howMany: number): this {
    this.maxRows = Math.max(1, this.maxRows - howMany);
    // Values, formats, and validations move together, like real Sheets rows
    const shiftRow = (r: number, c: number) =>
      r < rowPosition ? this.key(r, c) : r >= rowPosition + howMany ? this.key(r - howMany, c) : null;
    this.cells = this.shiftKeys(this.cells, shiftRow);
    this.formats = this.shiftKeys(this.formats, shiftRow);
    this.validations = this.shiftKeys(this.validations, shiftRow);
    return this;
  }

  deleteColumns(colPosition: number, howMany: number): this {
    this.maxCols = Math.max(1, this.maxCols - howMany);
    // Columns to the right of the deleted range shift left by `howMany`, mirroring `deleteRows`
    const shiftCol = (r: number, c: number) =>
      c < colPosition ? this.key(r, c) : c >= colPosition + howMany ? this.key(r, c - howMany) : null;
    this.cells = this.shiftKeys(this.cells, shiftCol);
    this.formats = this.shiftKeys(this.formats, shiftCol);
    this.validations = this.shiftKeys(this.validations, shiftCol);
    return this;
  }

  /** Re-keys a `row:col` map; entries whose new key is `null` (deleted rows/columns) are dropped. */
  private shiftKeys<V>(
    source: Map<string, V>,
    nextKey: (row: number, col: number) => string | null
  ): Map<string, V> {
    const next = new Map<string, V>();
    for (const [k, v] of source.entries()) {
      const [rStr, cStr] = k.split(':');
      const key = nextKey(Number(rStr), Number(cStr));
      if (key !== null) next.set(key, v);
    }
    return next;
  }

  clear(): this {
    this.cells.clear();
    this.formats.clear();
    this.validations.clear();
    return this;
  }

  getRange(rowOrA1: number | string, col?: number, numRows = 1, numCols = 1): MockRange {
    if (typeof rowOrA1 === 'string') {
      const parsed = this.parseA1(rowOrA1);
      return this.getRange(parsed.row, parsed.col, parsed.numRows, parsed.numCols);
    }
    if (rowOrA1 + numRows - 1 > this.maxRows || (col ?? 1) + numCols - 1 > this.maxCols) {
      throw new Error(
        `Range out of grid bounds: (${rowOrA1},${col},${numRows},${numCols}) exceeds (${this.maxRows},${this.maxCols})`
      );
    }
    return new MockRange(this, rowOrA1, col ?? 1, numRows, numCols);
  }

  getCellValue(row: number, col: number): unknown {
    const formula = this.formulas.get(this.key(row, col));
    if (formula !== undefined && this.formulaReadEvaluator) {
      const readIndex = (this.formulaReads.get(this.key(row, col)) ?? 0) + 1;
      this.formulaReads.set(this.key(row, col), readIndex);
      return this.formulaReadEvaluator(formula, readIndex);
    }
    const val = this.cells.get(this.key(row, col));
    return val === undefined ? '' : val;
  }

  setCellValue(row: number, col: number, value: unknown): void {
    this.formulas.delete(this.key(row, col));
    const fmt = this.formats.get(this.key(row, col));
    if (typeof value === 'string' && value.length > 0 && fmt !== '@') {
      // Simulate Google Sheets stripping leading apostrophe and coercing unformatted numeric strings
      if (value.startsWith("'")) {
        this.cells.set(this.key(row, col), value.slice(1));
        return;
      }
      if (/^-?\d+(\.\d+)?$/.test(value)) {
        const num = Number(value);
        if (Number.isFinite(num)) {
          this.cells.set(this.key(row, col), num);
          return;
        }
      }
      // Simulate Google Sheets parsing date strings in date-formatted cells into Date values
      // (`getValues()` then returns a Date). The mock spreadsheet's wall clock is UTC.
      const dateMatch = /^(\d{4}-\d{2}-\d{2})(?: (\d{2}:\d{2}:\d{2}))?$/.exec(value);
      if (fmt?.startsWith('yyyy') && dateMatch) {
        this.cells.set(this.key(row, col), new Date(`${dateMatch[1]}T${dateMatch[2] ?? '00:00:00'}Z`));
        return;
      }
    }
    this.cells.set(this.key(row, col), value);
  }

  setCellFormula(row: number, col: number, formula: string): void {
    this.formulas.set(this.key(row, col), formula);
    this.formulaReads.delete(this.key(row, col));
    const evaluated = this.formulaEvaluator ? this.formulaEvaluator(formula) : formula;
    this.cells.set(this.key(row, col), evaluated);
  }

  setCellFormat(row: number, col: number, format: string): void {
    this.formats.set(this.key(row, col), format);
  }

  getCellFormat(row: number, col: number): string | undefined {
    return this.formats.get(this.key(row, col));
  }

  setCellValidation(row: number, col: number, rule: MockDataValidation | null): void {
    this.validations.set(this.key(row, col), rule);
  }

  getCellValidation(row: number, col: number): MockDataValidation | null {
    return this.validations.get(this.key(row, col)) ?? null;
  }

  private parseA1(a1: string): { row: number; col: number; numRows: number; numCols: number } {
    const parts = a1.split(':');
    const start = this.parseCellRef(parts[0]);
    if (parts.length === 1) {
      return { row: start.row, col: start.col, numRows: 1, numCols: 1 };
    }
    const end = this.parseCellRef(parts[1]);
    return {
      row: start.row,
      col: start.col,
      numRows: end.row - start.row + 1,
      numCols: end.col - start.col + 1,
    };
  }

  private parseCellRef(ref: string): { row: number; col: number } {
    const match = /^([A-Z]+)(\d+)$/i.exec(ref.trim());
    if (!match) {
      throw new Error(`Invalid A1 reference: ${ref}`);
    }
    const letters = match[1].toUpperCase();
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return { row: Number(match[2]), col };
  }
}

/** In-memory mock for a Google Sheets `Spreadsheet` workbook. */
export class MockSpreadsheet {
  private sheets: MockSheet[] = [];
  private nextSheetId = 100;
  private timeZone: string | null = 'America/New_York';

  constructor(
    private readonly url: string,
    private readonly id = 'mock-spreadsheet-id'
  ) {}

  getUrl(): string {
    return this.url;
  }

  getId(): string {
    return this.id;
  }

  getSpreadsheetTimeZone(): string | null {
    return this.timeZone;
  }

  setSpreadsheetTimeZone(tz: string | null): void {
    this.timeZone = tz;
  }

  getSheetByName(name: string): MockSheet | null {
    return this.sheets.find((s) => s.getName() === name) ?? null;
  }

  getSheets(): MockSheet[] {
    return [...this.sheets];
  }

  insertSheet(name: string, indexOrOptions?: number): MockSheet {
    // Real Sheets throws on duplicate names; production code goes through getOrCreateSheet
    if (this.getSheetByName(name)) {
      throw new Error(
        `A sheet with the name "${name}" already exists. Please enter another name.`
      );
    }
    // Test convenience: a number >= 100 is used as an explicit sheet ID (gid) instead of an index
    const explicitId =
      typeof indexOrOptions === 'number' && indexOrOptions >= 100 ? indexOrOptions : null;
    if (explicitId !== null && this.sheets.some((s) => s.getSheetId() === explicitId)) {
      throw new Error(`Sheet ID ${explicitId} is already in use.`);
    }
    let sheetId = explicitId;
    if (sheetId === null) {
      // Auto-generated IDs skip any ID already in use, so they never collide with explicit IDs
      while (this.sheets.some((s) => s.getSheetId() === this.nextSheetId)) {
        this.nextSheetId++;
      }
      sheetId = this.nextSheetId++;
    }
    const sheet = new MockSheet(name, sheetId);
    if (typeof indexOrOptions === 'number' && indexOrOptions >= 0 && indexOrOptions <= this.sheets.length) {
      this.sheets.splice(indexOrOptions, 0, sheet);
    } else {
      this.sheets.push(sheet);
    }
    return sheet;
  }

  deleteSheet(sheet: MockSheet): void {
    this.sheets = this.sheets.filter((s) => s !== sheet);
  }
}

/**
 * Creates an in-memory `SpreadsheetApp` mock with support for `openByUrl`, `getActiveSpreadsheet`,
 * `newDataValidation`, and `getUi` (menus, modal dialogs, and 3-arg alerts).
 */
export function createMockSpreadsheetApp(
  defaultUrl = 'https://docs.google.com/spreadsheets/d/mock-spreadsheet-id/edit'
) {
  const spreadsheets = new Map<string, MockSpreadsheet>();
  const active = new MockSpreadsheet(defaultUrl);
  spreadsheets.set(defaultUrl, active);
  const alerts: Array<{ title: string; prompt?: string; buttons?: unknown }> = [];
  const modals: Array<{ html: string; title: string }> = [];
  const menuItems: Array<{ label: string; fnName: string }> = [];
  /** Button returned by the next 3-arg `ui.alert` call (`'YES'`, `'NO'`, or `'OK'`). */
  const ui = { alertResponse: 'YES' };

  return {
    alerts,
    modals,
    menuItems,
    ui,
    openByUrl(url: string): MockSpreadsheet {
      let ss = spreadsheets.get(url);
      if (!ss) {
        ss = new MockSpreadsheet(url);
        spreadsheets.set(url, ss);
      }
      return ss;
    },
    getActiveSpreadsheet(): MockSpreadsheet {
      return active;
    },
    flush(): void {},
    newDataValidation() {
      let values: string[] = [];
      let allowInvalid = true;
      return {
        requireValueInList(list: string[]) {
          values = [...list];
          return this;
        },
        requireCheckbox() {
          return {
            build(): MockDataValidation {
              return { type: 'BOOLEAN', getCriteriaValues: () => [] };
            },
          };
        },
        setAllowInvalid(allow: boolean) {
          allowInvalid = allow;
          return this;
        },
        build(): MockDataValidation {
          const snapshot = [...values];
          return {
            type: 'VALUE_IN_LIST',
            values: snapshot,
            allowInvalid,
            getCriteriaValues: () => [snapshot],
          };
        },
      };
    },
    getUi() {
      return {
        ButtonSet: {
          OK: 'OK',
          OK_CANCEL: 'OK_CANCEL',
          YES_NO: 'YES_NO',
        },
        Button: {
          OK: 'OK',
          CANCEL: 'CANCEL',
          YES: 'YES',
          NO: 'NO',
        },
        createMenu(_title: string) {
          const menu = {
            addItem(label: string, fnName: string) {
              menuItems.push({ label, fnName });
              return menu;
            },
            addSeparator() {
              return menu;
            },
            addToUi() {},
          };
          return menu;
        },
        showModalDialog(htmlOutput: { getContent?: () => string } | unknown, title: string) {
          const maybeOutput = htmlOutput as { getContent?: () => string } | null | undefined;
          const html =
            typeof maybeOutput?.getContent === 'function'
              ? maybeOutput.getContent()
              : String(htmlOutput ?? '');
          modals.push({ html, title });
        },
        alert(title: string, prompt?: string, buttons?: unknown) {
          if (arguments.length === 2) {
            throw new Error(
              'Invalid Ui.alert(String, String) overload: Apps Script requires 1 argument or 3 arguments (title, prompt, ButtonSet).'
            );
          }
          alerts.push({ title, prompt, buttons });
          return ui.alertResponse;
        },
      };
    },
  };
}

/** Descriptor for a simulated Google Ads account in `createMockAdsApp`. */
export interface MockAdsAccount {
  customerId: string;
  name: string;
  currencyCode: string;
  timeZone?: string;
}

/**
 * Creates an in-memory `AdsApp` and `AdsManagerApp` mock supporting `currentAccount`,
 * `getExecutionInfo`, `search`, and `accounts().get()`.
 */
export function createMockAdsApp(options: {
  mccId?: string;
  mccName?: string;
  currencyCode?: string;
  remainingTime?: number;
  isPreview?: boolean;
  searchHandler?: (query: string, accountId: string, searchOptions?: { apiVersion?: string }) => unknown[];
  childAccounts?: MockAdsAccount[];
} = {}) {
  const mccAccount: MockAdsAccount = {
    customerId: options.mccId ?? '111-222-3333',
    name: options.mccName ?? 'Default Mock MCC',
    currencyCode: options.currencyCode ?? 'USD',
    timeZone: 'America/New_York',
  };
  let selectedAccount: MockAdsAccount = mccAccount;
  let remainingTime = options.remainingTime ?? 1800;
  const isPreview = options.isPreview ?? false;
  const selectCounts = new Map<string, number>();
  const searchOptionsHistory: Array<{ apiVersion?: string } | undefined> = [];

  const toAccountEntity = (acc: MockAdsAccount) => ({
    getCustomerId: () => acc.customerId,
    getName: () => acc.name,
    getCurrencyCode: () => acc.currencyCode,
    getTimeZone: () => acc.timeZone ?? 'America/New_York',
  });

  const adsApp = {
    currentAccount() {
      return toAccountEntity(mccAccount);
    },
    getExecutionInfo() {
      return {
        isPreview: () => isPreview,
        getRemainingTime: () => remainingTime,
      };
    },
    setRemainingTime(seconds: number) {
      remainingTime = seconds;
    },
    search(query: string, searchOptions?: { apiVersion?: string }) {
      searchOptionsHistory.push(searchOptions);
      const rows = options.searchHandler
        ? options.searchHandler(query, selectedAccount.customerId, searchOptions)
        : [];
      let index = 0;
      return {
        hasNext: () => index < rows.length,
        next: () => rows[index++],
        totalNumEntities: () => rows.length,
      };
    },
    selectCounts,
    searchOptionsHistory,
  };

  const adsManagerApp = {
    accounts() {
      const list = (options.childAccounts ?? []).map(toAccountEntity);
      return {
        withIds(ids: string[]) {
          const normalizedSet = new Set(ids.map((id) => id.replace(/\D/g, '')));
          const filtered = (options.childAccounts ?? [])
            .filter((a) => normalizedSet.has(a.customerId.replace(/\D/g, '')))
            .map(toAccountEntity);
          return {
            get() {
              let idx = 0;
              return {
                hasNext: () => idx < filtered.length,
                next: () => filtered[idx++],
                totalNumEntities: () => filtered.length,
              };
            },
          };
        },
        get() {
          let idx = 0;
          return {
            hasNext: () => idx < list.length,
            next: () => list[idx++],
            totalNumEntities: () => list.length,
          };
        },
      };
    },
    select(accountEntity: ReturnType<typeof toAccountEntity>) {
      const cid = accountEntity.getCustomerId();
      selectCounts.set(cid, (selectCounts.get(cid) ?? 0) + 1);
      selectedAccount = {
        customerId: cid,
        name: accountEntity.getName(),
        currencyCode: accountEntity.getCurrencyCode(),
        timeZone: accountEntity.getTimeZone(),
      };
    },
  };

  return Object.assign(adsApp, { adsManagerApp });
}

/**
 * Creates an in-memory `LockService` mock whose `getDocumentLock()` tracks acquisitions and releases.
 * Like Apps Script, `waitLock` on a lock this execution already holds returns immediately.
 */
export function createMockLockService() {
  const stats = {acquireCount: 0, releaseCount: 0, held: false};
  const lock = {
    waitLock(_timeoutMs: number): void {
      if (stats.held) return;
      stats.held = true;
      stats.acquireCount++;
    },
    tryLock(_timeoutMs: number): boolean {
      lock.waitLock(_timeoutMs);
      return true;
    },
    releaseLock(): void {
      if (!stats.held) return;
      stats.held = false;
      stats.releaseCount++;
    },
    hasLock(): boolean {
      return stats.held;
    },
  };
  return {stats, getDocumentLock: () => lock};
}

/** In-memory `PropertiesService` with separate script and document property stores. */
export function createMockPropertiesService() {
  const makeStore = () => {
    const values = new Map<string, string>();
    return {
      values,
      getProperty: (key: string): string | null => values.get(key) ?? null,
      setProperty(key: string, value: string) {
        values.set(key, value);
        return this;
      },
    };
  };
  const scriptProperties = makeStore();
  const documentProperties = makeStore();
  return {
    scriptProperties,
    documentProperties,
    getScriptProperties: () => scriptProperties,
    getDocumentProperties: () => documentProperties,
  };
}

/**
 * Installs `SpreadsheetApp`, `AdsApp`, `AdsManagerApp`, `LockService`, `PropertiesService`, `HtmlService`, `Logger`,
 * and `Utilities` onto `globalThis` for unit and integration tests. Services not passed in (including `ScriptApp` and
 * `__EMBEDDED_ADS_SCRIPT__`) are removed or replaced with fresh mocks, so mocks from a previous test never leak into
 * the next one.
 */
export function installAppsScriptGlobals(mocks: {
  spreadsheetApp?: ReturnType<typeof createMockSpreadsheetApp>;
  adsApp?: ReturnType<typeof createMockAdsApp>;
  /** Defaults to a fresh `createMockLockService()`; pass `null` to simulate no `LockService` (Code.gs). */
  lockService?: ReturnType<typeof createMockLockService> | null;
  /** Defaults to a fresh `createMockPropertiesService()`. */
  propertiesService?: ReturnType<typeof createMockPropertiesService>;
}): {
  lockService: ReturnType<typeof createMockLockService> | null;
  propertiesService: ReturnType<typeof createMockPropertiesService>;
} {
  const g = globalThis as Record<string, unknown>;
  delete g.ScriptApp;
  delete g.__EMBEDDED_ADS_SCRIPT__;
  const propertiesService = mocks.propertiesService ?? createMockPropertiesService();
  g.PropertiesService = propertiesService;
  const lockService =
    mocks.lockService === undefined ? createMockLockService() : mocks.lockService;
  if (lockService) {
    g.LockService = lockService;
  } else {
    delete g.LockService;
  }
  if (mocks.spreadsheetApp) {
    g.SpreadsheetApp = mocks.spreadsheetApp;
  } else {
    delete g.SpreadsheetApp;
  }
  if (mocks.adsApp) {
    g.AdsApp = mocks.adsApp;
    g.AdsManagerApp = mocks.adsApp.adsManagerApp;
  } else {
    delete g.AdsApp;
    delete g.AdsManagerApp;
  }
  g.HtmlService = {
    createHtmlOutput(html: string) {
      const out = {
        getContent: () => html,
        setWidth: () => out,
        setHeight: () => out,
      };
      return out;
    },
  };
  g.Logger = {
    log: (_msg: string) => {},
  };
  g.Utilities = {
    formatDate: (date: Date, tz: string, format: string) => {
      if (typeof tz !== 'string' || tz.length === 0) {
        throw new Error('Exception: Invalid argument: timeZone. Should be of type: String');
      }
      const iso = date.toISOString();
      if (format === 'yyyy-MM-dd') return iso.slice(0, 10);
      if (format === 'H') return String(date.getUTCHours());
      return iso.replace('T', ' ').slice(0, 19);
    },
    sleep: (_ms: number) => {},
  };
  return {lockService, propertiesService};
}
