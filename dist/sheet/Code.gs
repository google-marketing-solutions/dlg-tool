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
 * @OnlyCurrentDoc
 * DLG Dashboard — Master Google Sheet Bound Script
 */
"use strict";
var DlgSheetBundle = (() => {
  var __defProp = Object.defineProperty;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __export = (target, all) => {
    for (var name in all)
      __defProp(target, name, { get: all[name], enumerable: true });
  };
  var __copyProps = (to, from, except, desc) => {
    if (from && typeof from === "object" || typeof from === "function") {
      for (let key of __getOwnPropNames(from))
        if (!__hasOwnProp.call(to, key) && key !== except)
          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
    }
    return to;
  };
  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

  // src/appscript/sheet-ui.ts
  var sheet_ui_exports = {};
  __export(sheet_ui_exports, {
    DEMO_TOOLS_PROPERTY: () => DEMO_TOOLS_PROPERTY,
    EMBEDDED_ADS_SCRIPT_TEMPLATE: () => EMBEDDED_ADS_SCRIPT_TEMPLATE,
    PUBLISHER_HEARTBEAT_PROPERTY: () => PUBLISHER_HEARTBEAT_PROPERTY,
    PUBLISHER_HEARTBEAT_STALE_MS: () => PUBLISHER_HEARTBEAT_STALE_MS,
    applyConfigChangesNow: () => applyConfigChangesNow,
    buildDataStudioCreationUrlForActiveSheet: () => buildDataStudioCreationUrlForActiveSheet,
    ensureAutoPublishTriggerInstalled: () => ensureAutoPublishTriggerInstalled,
    ensureDataSheetInitialized: () => ensureDataSheetInitialized,
    generateAdsScriptForCurrentSheet: () => generateAdsScriptForCurrentSheet,
    isDemoToolsEnabled: () => isDemoToolsEnabled,
    onEdit: () => onEdit,
    onOpen: () => onOpen,
    openDataStudioCreator: () => openDataStudioCreator,
    openSavedDashboard: () => openSavedDashboard,
    publishFromStagingToData: () => publishFromStagingToData,
    runScheduledPublish: () => runScheduledPublish,
    seedSyntheticDemoData: () => seedSyntheticDemoData,
    seedSyntheticDemoDataFromMenu: () => seedSyntheticDemoDataFromMenu,
    showCopyAdsScriptModal: () => showCopyAdsScriptModal
  });

  // src/shared/utils/cid.ts
  var HYPHENATED_CID_RE = /^\d{3}-\d{3}-\d{4}$/;
  function normalizeCid(rawCid) {
    if (rawCid === null || rawCid === void 0) {
      return "";
    }
    const trimmed = String(rawCid).trim();
    const digits = HYPHENATED_CID_RE.test(trimmed) ? trimmed.replace(/-/g, "") : trimmed.replace(/\s+/g, "");
    if (!/^\d{1,10}$/.test(digits) || /^0+$/.test(digits)) {
      return "";
    }
    return digits.padStart(10, "0");
  }
  function formatCidHyphenated(rawCid) {
    const digits = normalizeCid(rawCid);
    if (digits.length !== 10) {
      return digits;
    }
    return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 10)}`;
  }

  // src/shared/contract.ts
  var CONTRACT_VERSION = 1;
  function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
  }
  function isStringArray(value) {
    return Array.isArray(value) && value.every((v) => typeof v === "string");
  }
  function parseVersionedRecord(raw, label, log) {
    let value;
    try {
      value = JSON.parse(raw);
    } catch {
      log(`[DLG v2] ${label} is not valid JSON; ignoring it.`);
      return null;
    }
    if (!isRecord(value)) {
      log(`[DLG v2] ${label} is not a JSON object; ignoring it.`);
      return null;
    }
    const version = value["schemaVersion"] ?? 1;
    if (version !== CONTRACT_VERSION) {
      log(
        `[DLG v2] ${label} has schemaVersion ${String(version)}, but this script reads version ${CONTRACT_VERSION}. Re-copy the Google Ads Script from the Google Sheet menu (\u{1F4CB} Copy Google Ads Script) so both scripts match.`
      );
      return null;
    }
    return value;
  }
  function serializeMccCheckpoint(checkpoint) {
    return JSON.stringify({ schemaVersion: CONTRACT_VERSION, ...checkpoint });
  }
  function parseMccCheckpoint(raw, log = defaultLog) {
    const v = parseVersionedRecord(raw, "MCC CHECKPOINT", log);
    if (!v)
      return null;
    const status = v["status"];
    if (typeof v["mccId"] !== "string" || typeof v["dateStr"] !== "string" || status !== "IN_PROGRESS" && status !== "COMPLETED" || !isStringArray(v["completedCids"])) {
      log("[DLG v2] MCC CHECKPOINT has an unexpected shape; ignoring it.");
      return null;
    }
    const failedCids = parseFailedCids(v["failedCids"]);
    const lastCompletedDate = v["lastCompletedDate"];
    return {
      mccId: v["mccId"],
      dateStr: v["dateStr"],
      status,
      completedCids: v["completedCids"].map(normalizeCid).filter(Boolean),
      ...failedCids ? { failedCids } : {},
      ...typeof lastCompletedDate === "string" && lastCompletedDate ? { lastCompletedDate } : {}
    };
  }
  function parseFailedCids(value) {
    if (!isRecord(value))
      return void 0;
    const out = {};
    for (const [rawCid, info] of Object.entries(value)) {
      const cid = normalizeCid(rawCid);
      if (!cid || !isRecord(info))
        continue;
      const attempts = Number(info["attempts"]);
      out[cid] = {
        error: String(info["error"] ?? ""),
        attempts: Number.isFinite(attempts) && attempts > 0 ? attempts : 1
      };
    }
    return Object.keys(out).length > 0 ? out : void 0;
  }
  function serializeLockedDailyConfig(config) {
    return JSON.stringify({ schemaVersion: CONTRACT_VERSION, ...config });
  }
  function parseLockedDailyConfig(raw, log = defaultLog) {
    const v = parseVersionedRecord(raw, "DAILY_CONFIG_LOCK", log);
    if (!v)
      return null;
    if (typeof v["targetCurrency"] !== "string" || typeof v["excludeShoppingAndPmax"] !== "boolean") {
      log("[DLG v2] DAILY_CONFIG_LOCK has an unexpected shape; ignoring it.");
      return null;
    }
    const config = {
      targetCurrency: v["targetCurrency"],
      excludeShoppingAndPmax: v["excludeShoppingAndPmax"]
    };
    const rawSelected = v["selectedCidsByMcc"];
    if (rawSelected !== void 0) {
      if (!isRecord(rawSelected)) {
        log("[DLG v2] DAILY_CONFIG_LOCK selectedCidsByMcc is not an object; ignoring the lock.");
        return null;
      }
      const selectedCidsByMcc = {};
      for (const [rawMcc, cids] of Object.entries(rawSelected)) {
        const mcc = normalizeCid(rawMcc);
        if (!mcc || !isStringArray(cids))
          continue;
        selectedCidsByMcc[mcc] = cids.map(normalizeCid).filter(Boolean);
      }
      config.selectedCidsByMcc = selectedCidsByMcc;
    }
    return config;
  }
  function defaultLog(msg) {
    if (typeof Logger !== "undefined") {
      Logger.log(msg);
    }
  }

  // src/shared/version.ts
  var DLG_VERSION = true ? "2.0.0" : "dev";

  // src/shared/sheets/sheet-utils.ts
  var RUN_STATE_SHEET_NAME = "_RunState";
  var SHEET_WRITE_CHUNK_ROWS = 1e3;
  function getOrCreateSheet(spreadsheet, name, options = {}) {
    let sheet = spreadsheet.getSheetByName(name);
    if (sheet) {
      return sheet;
    }
    sheet = options.index !== void 0 ? spreadsheet.insertSheet(name, options.index) : spreadsheet.insertSheet(name);
    if (options.hidden) {
      sheet.hideSheet();
    }
    return sheet;
  }
  function readRunStateMap(spreadsheet) {
    const map = /* @__PURE__ */ new Map();
    const sheet = spreadsheet.getSheetByName(RUN_STATE_SHEET_NAME);
    if (!sheet)
      return map;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2)
      return map;
    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (const row of rows) {
      const k = String(row[0] ?? "").trim();
      if (k) {
        map.set(k, String(row[1] ?? ""));
      }
    }
    return map;
  }
  function getRunStateValue(spreadsheet, key, stateMap) {
    const map = stateMap ?? readRunStateMap(spreadsheet);
    return map.get(key) ?? null;
  }
  function readLockedDailyConfig(spreadsheet, dateStr, stateMap) {
    const lockedJson = getRunStateValue(spreadsheet, `DAILY_CONFIG_LOCK_${dateStr}`, stateMap);
    return lockedJson ? parseLockedDailyConfig(lockedJson) : null;
  }
  var FORMULA_PREFIX_RE = /^[=+\-@\t\r]/;
  var ESCAPED_FORMULA_PREFIX_RE = /^'[=+\-@\t\r]/;
  function escapeFormulaString(val) {
    return FORMULA_PREFIX_RE.test(val) ? `'${val}` : val;
  }
  function unescapeFormulaString(val) {
    return ESCAPED_FORMULA_PREFIX_RE.test(val) ? val.slice(1) : val;
  }
  function parseSheetBoolean(value) {
    return value === true || String(value).toUpperCase() === "TRUE";
  }
  function toSheetValueMatrix(columns, items) {
    return items.map((item) => {
      const rec = item;
      return columns.map((k) => {
        const v = rec[k];
        if (v === null || v === void 0)
          return "";
        return typeof v === "string" ? escapeFormulaString(v) : v;
      });
    });
  }
  function writeMatrixInChunks(sheet, startRow, startCol, matrix, chunkRows) {
    if (matrix.length === 0)
      return;
    const numCols = matrix[0].length;
    for (let offset = 0; offset < matrix.length; offset += chunkRows) {
      const slice = matrix.slice(offset, offset + chunkRows).map((r) => [...r]);
      sheet.getRange(startRow + offset, startCol, slice.length, numCols).setValues(slice);
    }
  }
  function trimSheetGrid(sheet, targetRows, targetCols) {
    const updatedMaxRows = sheet.getMaxRows();
    if (updatedMaxRows > targetRows) {
      sheet.deleteRows(targetRows + 1, updatedMaxRows - targetRows);
    }
    const updatedMaxCols = sheet.getMaxColumns();
    if (updatedMaxCols > targetCols) {
      sheet.deleteColumns(targetCols + 1, updatedMaxCols - targetCols);
    }
  }

  // src/shared/sheets/mcc-state.ts
  var MCC_SHEET_PREFIX = "_MCC_";
  function getMccSheetName(mccId) {
    return `${MCC_SHEET_PREFIX}${normalizeCid(mccId)}`;
  }
  function writeMccSheetState(spreadsheet, payload) {
    const normMcc = normalizeCid(payload.mcc.mccId);
    const formattedMcc = formatCidHyphenated(normMcc);
    const sheetName = getMccSheetName(normMcc);
    const sheet = getOrCreateSheet(spreadsheet, sheetName, { hidden: true });
    const existing = readMccSheetState(spreadsheet, normMcc);
    const checkpoint = payload.checkpoint ?? existing?.checkpoint;
    const lastCompletedDate = payload.lastCompletedDate ?? (checkpoint?.status === "COMPLETED" ? checkpoint.dateStr : void 0) ?? existing?.lastCompletedDate ?? "";
    const statusMessage = payload.statusMessage ?? existing?.statusMessage ?? "";
    const childAccounts = payload.childAccounts ?? existing?.childAccounts ?? [];
    const nowIso = (/* @__PURE__ */ new Date()).toISOString();
    const rows = [
      ["Key", "Value", "Extra1", "Extra2"],
      ["MCC_META", formattedMcc, payload.mcc.mccName, nowIso],
      ["CHECKPOINT", checkpoint ? serializeMccCheckpoint(checkpoint) : "", checkpoint?.dateStr ?? "", nowIso],
      ["LAST_COMPLETED_DATE", lastCompletedDate, "", nowIso],
      ["STATUS_MESSAGE", statusMessage, "", nowIso],
      ["SCRIPT_VERSION", DLG_VERSION, "", nowIso]
    ];
    for (const acc of childAccounts) {
      rows.push([
        "CHILD_ACCOUNT",
        formatCidHyphenated(acc.customerId),
        acc.name,
        String(acc.currencyCode ?? "USD").trim().toUpperCase()
      ]);
    }
    const prevLastRow = Math.max(1, sheet.getLastRow());
    const writeRowCount = Math.max(prevLastRow, rows.length);
    if (writeRowCount > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), writeRowCount - sheet.getMaxRows() + 5);
    }
    while (rows.length < writeRowCount) {
      rows.push(["", "", "", ""]);
    }
    const range = sheet.getRange(1, 1, rows.length, 4);
    range.setNumberFormat("@");
    range.setValues(rows);
    if (prevLastRow > rows.length) {
      sheet.deleteRows(rows.length + 1, prevLastRow - rows.length);
    }
  }
  function readMccSheetState(spreadsheet, mccId) {
    const normMcc = normalizeCid(mccId);
    const sheet = spreadsheet.getSheetByName(getMccSheetName(normMcc));
    if (!sheet)
      return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2)
      return null;
    const rawRows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let formattedMcc = formatCidHyphenated(normMcc);
    let mccName = "MCC";
    let checkpoint;
    let lastCompletedDate;
    let statusMessage;
    let scriptVersion;
    const childAccounts = [];
    for (const r of rawRows) {
      const key = String(r[0] ?? "").trim();
      const rawVal = r[1];
      const val = rawVal instanceof Date ? rawVal.toISOString().slice(0, 10) : String(rawVal ?? "").trim();
      const extra1 = String(r[2] ?? "").trim();
      const extra2 = String(r[3] ?? "").trim();
      if (key === "MCC_META") {
        if (val)
          formattedMcc = val;
        if (extra1)
          mccName = extra1;
      } else if (key === "CHECKPOINT" && val) {
        checkpoint = parseMccCheckpoint(val) ?? void 0;
      } else if (key === "SCRIPT_VERSION" && val) {
        scriptVersion = val;
      } else if (key === "LAST_COMPLETED_DATE" && val) {
        lastCompletedDate = val;
      } else if (key === "STATUS_MESSAGE" && val) {
        statusMessage = val;
      } else if (key === "CHILD_ACCOUNT" && val) {
        childAccounts.push({
          customerId: val,
          name: extra1 || "Account",
          currencyCode: extra2 || "USD"
        });
      }
    }
    return {
      mcc: { mccId: formattedMcc, mccName },
      childAccounts,
      statusMessage,
      checkpoint,
      lastCompletedDate,
      scriptVersion
    };
  }
  function getMccCheckpoint(spreadsheet, dateStr, mccId) {
    const normMcc = normalizeCid(mccId);
    const mccPayload = readMccSheetState(spreadsheet, normMcc);
    const lastCompletedDate = mccPayload?.lastCompletedDate ?? mccPayload?.checkpoint?.lastCompletedDate;
    if (mccPayload?.checkpoint && mccPayload.checkpoint.dateStr === dateStr) {
      return {
        ...mccPayload.checkpoint,
        mccId: normMcc,
        lastCompletedDate
      };
    }
    return {
      mccId: normMcc,
      dateStr,
      status: "IN_PROGRESS",
      completedCids: [],
      lastCompletedDate
    };
  }
  function clearMccSheetCheckpoint(spreadsheet, mccId) {
    const sheet = spreadsheet.getSheetByName(getMccSheetName(normalizeCid(mccId)));
    if (!sheet)
      return;
    clearCheckpointRowInMccSheet(sheet, (/* @__PURE__ */ new Date()).toISOString());
  }
  function clearCheckpointRowInMccSheet(sheet, nowIso) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 1)
      return;
    const keys = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] ?? "").trim() === "CHECKPOINT") {
        sheet.getRange(i + 1, 1, 1, 4).setValues([["CHECKPOINT", "", "", nowIso]]);
        return;
      }
    }
  }

  // src/shared/utils/date.ts
  function resolveSpreadsheetDate(spreadsheet) {
    return Utilities.formatDate(
      /* @__PURE__ */ new Date(),
      spreadsheet.getSpreadsheetTimeZone() || "UTC",
      "yyyy-MM-dd"
    );
  }

  // src/shared/utils/currency.ts
  var ISO_4217_RE = /^[A-Z]{3}$/;
  function normalizeCurrencyCode(raw) {
    const code = String(raw ?? "").trim().toUpperCase();
    return ISO_4217_RE.test(code) ? code : null;
  }

  // src/appscript/config-sheet.ts
  var SETUP_SHEET_NAME = "Setup";
  var SAVED_DASHBOARD_URL_CELL = "D13";
  var CONFIG_SHEET_NAME = "Config";
  var STATUS_BANNER_CELL = "B2";
  var TARGET_CURRENCY_CELL = "B5";
  var GLOBAL_SETTINGS_RANGE = "B5:B6";
  var GLOBAL_SELECT_ALL_ROW = 9;
  var GLOBAL_SELECT_ALL_CELL = `A${GLOBAL_SELECT_ALL_ROW}`;
  var MCC_TABLE_START_ROW = 11;
  var MCC_TABLE_END_ROW = 18;
  var MAX_SUPPORTED_MCCS = MCC_TABLE_END_ROW - MCC_TABLE_START_ROW + 1;
  var CID_TABLE_START_ROW = 22;
  var CONFIG_TABLE_WIDTH = 6;
  var MCC_COL = {
    SELECT_ALL: 0,
    MCC_ID: 1,
    NAME: 2,
    DISCOVERED_COUNT: 3,
    SELECTED_COUNT: 4,
    STATUS: 5
  };
  var CID_COL = {
    INCLUDE: 0,
    CID: 1,
    NAME: 2,
    PARENT_MCC_ID: 3,
    PARENT_MCC_NAME: 4,
    CURRENCY: 5
  };
  var MCC_TABLE_HEADERS = [
    "Select All in MCC",
    "MCC ID",
    "MCC Name",
    "Discovered Accounts",
    "Selected Accounts",
    "Sync Status"
  ];
  var CID_TABLE_HEADERS = [
    "Include",
    "Account CID",
    "Account Name",
    "Parent MCC ID",
    "Parent MCC Name",
    "Currency"
  ];
  function sheetCol(offset) {
    return offset + 1;
  }
  function sheetColLetter(offset) {
    return String.fromCharCode("A".charCodeAt(0) + offset);
  }
  function assertConfigLayout(configSheet) {
    const checks = [
      { row: MCC_TABLE_START_ROW - 1, expected: MCC_TABLE_HEADERS },
      { row: CID_TABLE_START_ROW - 1, expected: CID_TABLE_HEADERS }
    ];
    const actualRows = checks.map(
      (c) => configSheet.getRange(c.row, 1, 1, CONFIG_TABLE_WIDTH).getValues()[0].map((v) => String(v ?? "").trim())
    );
    if (actualRows.every((cells) => cells.every((v) => v === ""))) {
      return;
    }
    checks.forEach((check, i) => {
      check.expected.forEach((header, col) => {
        const actual = actualRows[i][col] ?? "";
        if (actual !== header) {
          throw new Error(
            `Config layout changed: expected "${header}" in ${sheetColLetter(col)}${check.row}, found "${actual}". Restore the Master Template layout of the Config tab (do not insert or delete rows or columns above or inside its tables).`
          );
        }
      });
    });
  }
  function readGlobalConfig(configSheet) {
    const values = configSheet.getRange(GLOBAL_SETTINGS_RANGE).getValues();
    const targetCurrency = normalizeCurrencyCode(values[0]?.[0]) ?? getFirstRegisteredCurrency(configSheet);
    const excludeRaw = values[1]?.[0];
    const excludeShoppingAndPmax = parseSheetBoolean(excludeRaw);
    const selectedCidsByMcc = {};
    for (const summary of getMccSelectionSummary(configSheet)) {
      selectedCidsByMcc[summary.mccId] = getSelectedCidsForMcc(configSheet, summary.mccId);
    }
    return {
      targetCurrency,
      excludeShoppingAndPmax,
      selectedCidsByMcc
    };
  }
  function getFirstRegisteredCurrency(configSheet) {
    const lastRow = configSheet.getLastRow();
    if (lastRow >= CID_TABLE_START_ROW) {
      const cidCount = lastRow - CID_TABLE_START_ROW + 1;
      const currRows = configSheet.getRange(CID_TABLE_START_ROW, sheetCol(CID_COL.CURRENCY), cidCount, 1).getValues();
      for (const r of currRows) {
        const curr = normalizeCurrencyCode(r[0]);
        if (curr) {
          return curr;
        }
      }
    }
    return "USD";
  }
  function syncDiscoveredAccountsToConfig(configSheet, mcc, childAccounts) {
    const { isNewlyRegisteredMcc, defaultIncludeForNewCids } = syncMccTableRow(configSheet, mcc);
    const diff = reconcileCidRows(configSheet, mcc, childAccounts, defaultIncludeForNewCids);
    applyCidTableChanges(configSheet, diff);
    updateTargetCurrencyDropdown(configSheet, [...diff.retainedRows, ...diff.rowsToAppend]);
    return {
      isNewlyRegisteredMcc,
      newlyAddedCidsCount: diff.rowsToAppend.length,
      removedCidsCount: diff.removedCidsCount
    };
  }
  function syncAllMccSheetsToConfig(spreadsheet, configSheet) {
    assertConfigLayout(configSheet);
    const sheets = spreadsheet.getSheets();
    let newlyRegisteredMccs = 0;
    let newlyAddedCids = 0;
    const latestStatusMessages = [];
    const mccIdsByCid = /* @__PURE__ */ new Map();
    const outdatedScripts = [];
    for (const s of sheets) {
      const name = s.getName();
      if (!name.startsWith(MCC_SHEET_PREFIX))
        continue;
      const normMcc = normalizeCid(name.slice(MCC_SHEET_PREFIX.length));
      if (!normMcc)
        continue;
      const payload = readMccSheetState(spreadsheet, normMcc);
      if (!payload)
        continue;
      const syncRes = syncDiscoveredAccountsToConfig(
        configSheet,
        payload.mcc,
        payload.childAccounts ?? []
      );
      if (syncRes.isNewlyRegisteredMcc)
        newlyRegisteredMccs++;
      newlyAddedCids += syncRes.newlyAddedCidsCount;
      if (payload.statusMessage) {
        latestStatusMessages.push(payload.statusMessage);
      }
      if (payload.scriptVersion !== DLG_VERSION) {
        outdatedScripts.push({ mccId: normMcc, version: payload.scriptVersion ?? null });
      }
      for (const acc of payload.childAccounts ?? []) {
        const cid = normalizeCid(acc.customerId);
        if (!cid)
          continue;
        const mccIds = mccIdsByCid.get(cid) ?? [];
        if (!mccIds.includes(normMcc))
          mccIds.push(normMcc);
        mccIdsByCid.set(cid, mccIds);
      }
    }
    const duplicateAccounts = Array.from(mccIdsByCid.entries()).filter(([, mccIds]) => mccIds.length > 1).map(([cid, mccIds]) => ({ cid, mccIds: [...mccIds].sort() })).sort((x, y) => x.cid.localeCompare(y.cid));
    outdatedScripts.sort((x, y) => x.mccId.localeCompare(y.mccId));
    return { newlyRegisteredMccs, newlyAddedCids, latestStatusMessages, duplicateAccounts, outdatedScripts };
  }
  function syncMccTableRow(configSheet, mcc) {
    const normalizedMccId = normalizeCid(mcc.mccId);
    const formattedMccId = formatCidHyphenated(mcc.mccId);
    const mccTableValues = configSheet.getRange(MCC_TABLE_START_ROW, 1, MAX_SUPPORTED_MCCS, CONFIG_TABLE_WIDTH).getValues();
    const registeredMccIds = mccTableValues.map((row) => normalizeCid(row[MCC_COL.MCC_ID])).filter(Boolean);
    const mccRowIndex = registeredMccIds.indexOf(normalizedMccId);
    if (mccRowIndex === -1) {
      if (registeredMccIds.length >= MAX_SUPPORTED_MCCS) {
        throw new Error(
          `Maximum of ${MAX_SUPPORTED_MCCS} MCCs already registered in Config (Rows ${MCC_TABLE_START_ROW}-${MCC_TABLE_END_ROW}).`
        );
      }
      const sheetRow = MCC_TABLE_START_ROW + registeredMccIds.length;
      const mccIdCell = `${sheetColLetter(MCC_COL.MCC_ID)}${sheetRow}`;
      const selectedCountCell = `${sheetColLetter(MCC_COL.SELECTED_COUNT)}${sheetRow}`;
      const parentCol = sheetColLetter(CID_COL.PARENT_MCC_ID);
      const includeCol = sheetColLetter(CID_COL.INCLUDE);
      const parentRange = `$${parentCol}$${CID_TABLE_START_ROW}:$${parentCol}`;
      const includeRange = `$${includeCol}$${CID_TABLE_START_ROW}:$${includeCol}`;
      const row = new Array(CONFIG_TABLE_WIDTH);
      row[MCC_COL.SELECT_ALL] = false;
      row[MCC_COL.MCC_ID] = formattedMccId;
      row[MCC_COL.NAME] = escapeFormulaString(mcc.mccName);
      row[MCC_COL.DISCOVERED_COUNT] = `=IF(${mccIdCell}="", "", COUNTIF(${parentRange}, ${mccIdCell}))`;
      row[MCC_COL.SELECTED_COUNT] = `=IF(${mccIdCell}="", "", COUNTIFS(${parentRange}, ${mccIdCell}, ${includeRange}, TRUE))`;
      row[MCC_COL.STATUS] = `=IF(${mccIdCell}="", "", IF(${selectedCountCell}=0, "\u23F8\uFE0F INACTIVE (0 Selected \u2014 Skipped)", "\u{1F7E2} ACTIVE (" & ${selectedCountCell} & " Selected)"))`;
      configSheet.getRange(sheetRow, sheetCol(MCC_COL.MCC_ID), 1, 2).setNumberFormat("@");
      configSheet.getRange(sheetRow, 1, 1, CONFIG_TABLE_WIDTH).setValues([row]);
      configSheet.getRange(sheetRow, sheetCol(MCC_COL.SELECT_ALL), 1, 1).insertCheckboxes();
      return { isNewlyRegisteredMcc: true, defaultIncludeForNewCids: false };
    }
    const mccRow = mccTableValues[mccRowIndex];
    const defaultIncludeForNewCids = parseSheetBoolean(mccRow[MCC_COL.SELECT_ALL]);
    if (mcc.mccName && readTextCell(mccRow[MCC_COL.NAME]) !== mcc.mccName) {
      writeTextCell(configSheet, MCC_TABLE_START_ROW + mccRowIndex, sheetCol(MCC_COL.NAME), mcc.mccName);
    }
    return { isNewlyRegisteredMcc: false, defaultIncludeForNewCids };
  }
  function reconcileCidRows(configSheet, mcc, childAccounts, defaultIncludeForNewCids) {
    const normalizedMccId = normalizeCid(mcc.mccId);
    const formattedMccId = formatCidHyphenated(mcc.mccId);
    const liveAccountsByCid = /* @__PURE__ */ new Map();
    for (const acc of childAccounts) {
      const normCid = normalizeCid(acc.customerId);
      if (normCid && !liveAccountsByCid.has(normCid)) {
        liveAccountsByCid.set(normCid, acc);
      }
    }
    const lastRow = Math.max(configSheet.getLastRow(), CID_TABLE_START_ROW - 1);
    const existingRowsCount = Math.max(0, lastRow - CID_TABLE_START_ROW + 1);
    const existingRows = existingRowsCount > 0 ? configSheet.getRange(CID_TABLE_START_ROW, 1, existingRowsCount, CONFIG_TABLE_WIDTH).getValues() : [];
    const existingCids = /* @__PURE__ */ new Set();
    const retainedRows = [];
    const renamedCells = [];
    let removedCidsCount = 0;
    let nextWriteRow = CID_TABLE_START_ROW;
    for (let i = 0; i < existingRows.length; i++) {
      const row = existingRows[i];
      const cid = normalizeCid(row[CID_COL.CID]);
      const parentMcc = normalizeCid(row[CID_COL.PARENT_MCC_ID]);
      if (!cid) {
        continue;
      }
      if (parentMcc === normalizedMccId && !liveAccountsByCid.has(cid)) {
        removedCidsCount++;
        continue;
      }
      existingCids.add(cid);
      const liveName = liveAccountsByCid.get(cid)?.name;
      if (liveName && readTextCell(row[CID_COL.NAME]) !== liveName) {
        row[CID_COL.NAME] = liveName;
        renamedCells.push({ row: CID_TABLE_START_ROW + i, col: sheetCol(CID_COL.NAME), value: liveName });
      }
      if (parentMcc === normalizedMccId && mcc.mccName && readTextCell(row[CID_COL.PARENT_MCC_NAME]) !== mcc.mccName) {
        row[CID_COL.PARENT_MCC_NAME] = mcc.mccName;
        renamedCells.push({
          row: CID_TABLE_START_ROW + i,
          col: sheetCol(CID_COL.PARENT_MCC_NAME),
          value: mcc.mccName
        });
      }
      row[CID_COL.NAME] = escapeFormulaString(String(row[CID_COL.NAME] ?? ""));
      row[CID_COL.PARENT_MCC_NAME] = escapeFormulaString(String(row[CID_COL.PARENT_MCC_NAME] ?? ""));
      retainedRows.push(row);
      nextWriteRow = CID_TABLE_START_ROW + i + 1;
    }
    const rowsToAppend = [];
    for (const [normCid, acc] of liveAccountsByCid) {
      if (existingCids.has(normCid)) {
        continue;
      }
      rowsToAppend.push(
        toCidRowValues({
          include: defaultIncludeForNewCids,
          cid: formatCidHyphenated(normCid),
          name: acc.name,
          parentMccId: formattedMccId,
          parentMccName: mcc.mccName,
          currency: String(acc.currencyCode ?? "").trim().toUpperCase()
        })
      );
    }
    return {
      existingRowsCount: existingRows.length,
      retainedRows,
      rowsToAppend,
      renamedCells,
      removedCidsCount,
      nextWriteRow
    };
  }
  function applyCidTableChanges(configSheet, diff) {
    if (diff.removedCidsCount > 0) {
      rewriteCidTable(configSheet, [...diff.retainedRows, ...diff.rowsToAppend], diff.existingRowsCount);
      return;
    }
    for (const cell of diff.renamedCells) {
      writeTextCell(configSheet, cell.row, cell.col, cell.value);
    }
    if (diff.rowsToAppend.length > 0) {
      writeFormattedCidRows(configSheet, diff.nextWriteRow, diff.rowsToAppend);
    }
  }
  function rewriteCidTable(configSheet, combinedRows, existingRowsCount) {
    const totalToWrite = Math.max(existingRowsCount, combinedRows.length);
    if (totalToWrite === 0) {
      return;
    }
    ensureSheetMaxRows(configSheet, CID_TABLE_START_ROW + totalToWrite - 1);
    const paddedRows = [...combinedRows];
    while (paddedRows.length < totalToWrite) {
      paddedRows.push(new Array(CONFIG_TABLE_WIDTH).fill(""));
    }
    if (combinedRows.length > 0) {
      formatCidTextColumns(configSheet, CID_TABLE_START_ROW, combinedRows.length);
    }
    configSheet.getRange(CID_TABLE_START_ROW, 1, totalToWrite, CONFIG_TABLE_WIDTH).setValues(paddedRows);
    if (combinedRows.length > 0) {
      configSheet.getRange(CID_TABLE_START_ROW, sheetCol(CID_COL.INCLUDE), combinedRows.length, 1).insertCheckboxes();
    }
    const keptRows = Math.max(combinedRows.length, 1);
    if (existingRowsCount > keptRows) {
      configSheet.deleteRows(CID_TABLE_START_ROW + keptRows, existingRowsCount - keptRows);
    }
  }
  function writeFormattedCidRows(configSheet, startRow, rows) {
    ensureSheetMaxRows(configSheet, startRow + rows.length - 1);
    formatCidTextColumns(configSheet, startRow, rows.length);
    configSheet.getRange(startRow, 1, rows.length, CONFIG_TABLE_WIDTH).setValues(rows);
    configSheet.getRange(startRow, sheetCol(CID_COL.INCLUDE), rows.length, 1).insertCheckboxes();
  }
  function formatCidTextColumns(configSheet, startRow, numRows) {
    configSheet.getRange(startRow, sheetCol(CID_COL.CID), numRows, CONFIG_TABLE_WIDTH - 1).setNumberFormat("@");
  }
  function toCidRowValues(r) {
    const row = new Array(CONFIG_TABLE_WIDTH);
    row[CID_COL.INCLUDE] = r.include;
    row[CID_COL.CID] = r.cid;
    row[CID_COL.NAME] = escapeFormulaString(r.name);
    row[CID_COL.PARENT_MCC_ID] = r.parentMccId;
    row[CID_COL.PARENT_MCC_NAME] = escapeFormulaString(r.parentMccName);
    row[CID_COL.CURRENCY] = r.currency;
    return row;
  }
  function writeTextCell(configSheet, row, col, value) {
    const cell = configSheet.getRange(row, col);
    cell.setNumberFormat("@");
    cell.setValue(escapeFormulaString(value));
  }
  function readTextCell(value) {
    return unescapeFormulaString(String(value ?? ""));
  }
  function ensureSheetMaxRows(configSheet, neededMaxRow) {
    if (neededMaxRow > configSheet.getMaxRows()) {
      configSheet.insertRowsAfter(configSheet.getMaxRows(), neededMaxRow - configSheet.getMaxRows());
    }
  }
  function updateTargetCurrencyDropdown(configSheet, activeCidRows) {
    const currencies = /* @__PURE__ */ new Set();
    let firstRegisteredCurrency = "";
    for (const row of activeCidRows) {
      const curr = normalizeCurrencyCode(row[CID_COL.CURRENCY]);
      if (curr) {
        currencies.add(curr);
        if (!firstRegisteredCurrency) {
          firstRegisteredCurrency = curr;
        }
      }
    }
    const cell = configSheet.getRange(TARGET_CURRENCY_CELL);
    if (!normalizeCurrencyCode(cell.getValue()) && firstRegisteredCurrency) {
      cell.setValue(firstRegisteredCurrency);
    }
    const existingValidation = cell.getDataValidation();
    const rawExistingList = existingValidation?.getCriteriaValues()?.[0] ?? [];
    const previousOptions = Array.isArray(rawExistingList) ? rawExistingList.map(String) : [];
    const selected = normalizeCurrencyCode(cell.getValue());
    if (selected) {
      currencies.add(selected);
    }
    const options = Array.from(currencies).sort();
    if (options.length === 0) {
      return;
    }
    if (previousOptions.length === options.length && previousOptions.every((val, idx) => val === options[idx])) {
      return;
    }
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(options).setAllowInvalid(false).build();
    cell.setDataValidation(rule);
  }
  function getSelectedCidsForMcc(configSheet, mccId) {
    const targetMcc = normalizeCid(mccId);
    const lastRow = configSheet.getLastRow();
    if (lastRow < CID_TABLE_START_ROW) {
      return [];
    }
    const rows = configSheet.getRange(CID_TABLE_START_ROW, 1, lastRow - CID_TABLE_START_ROW + 1, CONFIG_TABLE_WIDTH).getValues();
    const selected = [];
    for (const row of rows) {
      const include = parseSheetBoolean(row[CID_COL.INCLUDE]);
      const cid = normalizeCid(row[CID_COL.CID]);
      const parentMcc = normalizeCid(row[CID_COL.PARENT_MCC_ID]);
      if (!include || !cid || parentMcc !== targetMcc) {
        continue;
      }
      selected.push(cid);
    }
    return selected;
  }
  function getMccSelectionSummary(configSheet) {
    const mccRows = configSheet.getRange(MCC_TABLE_START_ROW, 1, MAX_SUPPORTED_MCCS, CONFIG_TABLE_WIDTH).getValues();
    const lastRow = configSheet.getLastRow();
    const cidRows = lastRow >= CID_TABLE_START_ROW ? configSheet.getRange(CID_TABLE_START_ROW, 1, lastRow - CID_TABLE_START_ROW + 1, CONFIG_TABLE_WIDTH).getValues() : [];
    const summaries = [];
    for (const mccRow of mccRows) {
      const mccId = normalizeCid(mccRow[MCC_COL.MCC_ID]);
      if (!mccId)
        continue;
      const selectAllChecked = parseSheetBoolean(mccRow[MCC_COL.SELECT_ALL]);
      const mccName = readTextCell(mccRow[MCC_COL.NAME]).trim();
      let discoveredCount = 0;
      let selectedCount = 0;
      for (const cidRow of cidRows) {
        const cid = normalizeCid(cidRow[CID_COL.CID]);
        const parentMcc = normalizeCid(cidRow[CID_COL.PARENT_MCC_ID]);
        if (!cid || parentMcc !== mccId)
          continue;
        discoveredCount++;
        const included = parseSheetBoolean(cidRow[CID_COL.INCLUDE]);
        if (included) {
          selectedCount++;
        }
      }
      summaries.push({
        mccId,
        mccName,
        selectAllChecked,
        discoveredCount,
        selectedCount,
        isActive: selectedCount > 0
      });
    }
    return summaries;
  }

  // src/shared/sheets/staging-writer.ts
  var STAGE_ACCOUNTS_SHEET = "_Stage_Accounts";
  var STAGE_BUDGETS_SHEET = "_Stage_Budgets";
  var STAGE_CAMPAIGNS_SHEET = "_Stage_Campaigns";
  var STAGE_RECS_SHEET = "_Stage_Recs";
  function getMccStagingSheetNames(mccId) {
    const norm = normalizeCid(mccId);
    return {
      accounts: `${STAGE_ACCOUNTS_SHEET}_${norm}`,
      budgets: `${STAGE_BUDGETS_SHEET}_${norm}`,
      campaigns: `${STAGE_CAMPAIGNS_SHEET}_${norm}`,
      recommendations: `${STAGE_RECS_SHEET}_${norm}`
    };
  }
  var STAGED_ACCOUNT_SCHEMA = {
    mccId: "STRING",
    mccName: "STRING",
    accountId: "STRING",
    accountName: "STRING",
    currencyCode: "STRING",
    timestamp: "STRING"
  };
  var STAGED_BUDGET_SCHEMA = {
    mccId: "STRING",
    accountId: "STRING",
    budgetResourceName: "STRING",
    campaignBudgetName: "STRING",
    campaignBudgetAmount: "NUMBER",
    campaignBudgetTotalAmount: "NUMBER",
    campaignBudgetDeliveryMethod: "STRING",
    campaignBudgetType: "STRING",
    campaignBudgetIsShared: "BOOLEAN"
  };
  var STAGED_CAMPAIGN_SCHEMA = {
    mccId: "STRING",
    accountId: "STRING",
    campaignResourceName: "STRING",
    campaignId: "STRING",
    campaignName: "STRING",
    campaignType: "STRING",
    campaignSubType: "STRING",
    campaignIsAiMax: "BOOLEAN",
    campaignAiMaxTextCustomizationEnabled: "BOOLEAN",
    campaignAiMaxFinalUrlExpansionEnabled: "BOOLEAN",
    budgetResourceName: "STRING",
    campaignIsPortfolioBiddingStrategy: "BOOLEAN",
    campaignPortfolioBiddingStrategyName: "STRING",
    portfolioStrategyResourceName: "STRING",
    portfolioOwnerId: "STRING",
    isLocallyOwnedPortfolio: "BOOLEAN",
    campaignPortfolioIsManagerOwned: "BOOLEAN",
    campaignBiddingStrategyType: "STRING",
    campaignBiddingTargetRoas: "NUMBER",
    campaignBiddingTargetCpa: "NUMBER",
    campaignStatsYesterdayCost: "NUMBER",
    campaignStats7DaysCost: "NUMBER",
    campaign30DaysCost: "NUMBER",
    campaign30DaysConversionsValue: "NUMBER",
    campaign30DaysConversions: "NUMBER",
    campaign30DaysSearchRankLostImpressionShare: "NUMBER",
    campaign30DaysSearchBudgetLostImpressionShare: "NUMBER"
  };
  var STAGED_REC_SCHEMA = {
    mccId: "STRING",
    accountId: "STRING",
    recommendationId: "STRING",
    recommendationType: "STRING",
    campaignResourceName: "STRING",
    recommendationCurrentBudgetAmount: "NUMBER",
    recommendationNewBudgetAmount: "NUMBER",
    recommendationNewTargetCpa: "NUMBER",
    recommendationNewTargetRoas: "NUMBER",
    recommendationBaseCost: "NUMBER",
    recommendationPotentialCost: "NUMBER",
    recommendationBaseClicks: "NUMBER",
    recommendationPotentialClicks: "NUMBER",
    recommendationBaseConversions: "NUMBER",
    recommendationPotentialConversions: "NUMBER",
    recommendationBaseConversionsValue: "NUMBER",
    recommendationPotentialConversionsValue: "NUMBER",
    recommendationBaseImpressions: "NUMBER",
    recommendationPotentialImpressions: "NUMBER",
    recommendationBaseVideoViews: "NUMBER",
    recommendationPotentialVideoViews: "NUMBER",
    moveBudgetSourceBudgetResourceName: "STRING",
    moveBudgetSourceCampaigns: "STRING",
    moveBudgetSourceTargetCpa: "NUMBER",
    moveBudgetSourceTargetRoas: "NUMBER"
  };
  var STAGED_ACCOUNT_KEYS = Object.keys(STAGED_ACCOUNT_SCHEMA);
  var STAGED_BUDGET_KEYS = Object.keys(STAGED_BUDGET_SCHEMA);
  var STAGED_CAMPAIGN_KEYS = Object.keys(STAGED_CAMPAIGN_SCHEMA);
  var STAGED_REC_KEYS = Object.keys(STAGED_REC_SCHEMA);
  function appendStagingBatch(workbook, batch, mccId) {
    const names = getMccStagingSheetNames(mccId);
    upsertStagingTable(
      getOrCreateHiddenSheet(workbook, names.accounts),
      STAGED_ACCOUNT_KEYS,
      STAGED_ACCOUNT_SCHEMA,
      batch.accounts,
      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}`
    );
    upsertStagingTable(
      getOrCreateHiddenSheet(workbook, names.budgets),
      STAGED_BUDGET_KEYS,
      STAGED_BUDGET_SCHEMA,
      batch.budgets,
      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.budgetResourceName}`
    );
    upsertStagingTable(
      getOrCreateHiddenSheet(workbook, names.campaigns),
      STAGED_CAMPAIGN_KEYS,
      STAGED_CAMPAIGN_SCHEMA,
      batch.campaigns,
      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.campaignResourceName}`
    );
    upsertStagingTable(
      getOrCreateHiddenSheet(workbook, names.recommendations),
      STAGED_REC_KEYS,
      STAGED_REC_SCHEMA,
      batch.recommendations,
      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.recommendationId}::${r.campaignResourceName}`
    );
  }
  function readStagingBatchForMcc(workbook, mccId) {
    const normMcc = normalizeCid(mccId);
    const perMccNames = getMccStagingSheetNames(normMcc);
    const accounts = readTableRows(
      workbook.getSheetByName(perMccNames.accounts),
      STAGED_ACCOUNT_KEYS,
      STAGED_ACCOUNT_SCHEMA
    ).filter((r) => normalizeCid(r.mccId) === normMcc);
    const budgets = readTableRows(
      workbook.getSheetByName(perMccNames.budgets),
      STAGED_BUDGET_KEYS,
      STAGED_BUDGET_SCHEMA
    ).filter((r) => normalizeCid(r.mccId) === normMcc);
    const campaigns = readTableRows(
      workbook.getSheetByName(perMccNames.campaigns),
      STAGED_CAMPAIGN_KEYS,
      STAGED_CAMPAIGN_SCHEMA
    ).filter((r) => normalizeCid(r.mccId) === normMcc);
    const recommendations = readTableRows(
      workbook.getSheetByName(perMccNames.recommendations),
      STAGED_REC_KEYS,
      STAGED_REC_SCHEMA
    ).filter((r) => normalizeCid(r.mccId) === normMcc);
    return { accounts, budgets, campaigns, recommendations };
  }
  function clearMccFromStaging(workbook, mccId) {
    const normMcc = normalizeCid(mccId);
    const perMccNames = getMccStagingSheetNames(normMcc);
    const perMccSpecs = [
      [perMccNames.accounts, STAGED_ACCOUNT_KEYS, STAGED_ACCOUNT_SCHEMA],
      [perMccNames.budgets, STAGED_BUDGET_KEYS, STAGED_BUDGET_SCHEMA],
      [perMccNames.campaigns, STAGED_CAMPAIGN_KEYS, STAGED_CAMPAIGN_SCHEMA],
      [perMccNames.recommendations, STAGED_REC_KEYS, STAGED_REC_SCHEMA]
    ];
    for (const [name, keys, schema] of perMccSpecs) {
      const sheet = workbook.getSheetByName(name);
      if (sheet) {
        writeTableRows(sheet, keys, schema, []);
      }
    }
  }
  function getOrCreateHiddenSheet(workbook, name) {
    return getOrCreateSheet(workbook, name, { hidden: true });
  }
  function upsertStagingTable(sheet, keys, schema, newItems, keyFn) {
    const lastRow = sheet.getLastRow();
    if (newItems.length === 0 && lastRow > 0)
      return;
    const existing = readTableRows(sheet, keys, schema);
    const existingKeys = new Set(existing.map(keyFn));
    const newKeys = /* @__PURE__ */ new Set();
    let collides = false;
    for (const item of newItems) {
      const k = keyFn(item);
      if (existingKeys.has(k) || newKeys.has(k)) {
        collides = true;
        break;
      }
      newKeys.add(k);
    }
    if (lastRow > 0 && !collides && headerMatchesKeys(sheet, keys)) {
      appendTableRows(sheet, keys, schema, newItems, lastRow + 1);
      return;
    }
    const mergedMap = /* @__PURE__ */ new Map();
    for (const item of existing) {
      mergedMap.set(keyFn(item), item);
    }
    for (const item of newItems) {
      mergedMap.set(keyFn(item), item);
    }
    writeTableRows(sheet, keys, schema, Array.from(mergedMap.values()));
  }
  function headerMatchesKeys(sheet, keys) {
    const maxCols = sheet.getMaxColumns();
    if (maxCols < keys.length)
      return false;
    const header = sheet.getRange(1, 1, 1, maxCols).getValues()[0] ?? [];
    for (let c = 0; c < maxCols; c++) {
      const cell = String(header[c] ?? "").trim();
      if (cell !== (c < keys.length ? keys[c] : ""))
        return false;
    }
    return true;
  }
  function formatForFieldType(fieldType) {
    switch (fieldType) {
      case "NUMBER":
        return "0.########";
      case "BOOLEAN":
        return "General";
      default:
        return "@";
    }
  }
  function appendTableRows(sheet, keys, schema, items, startRow) {
    if (items.length === 0)
      return;
    const numCols = keys.length;
    const lastNeededRow = startRow + items.length - 1;
    if (sheet.getMaxRows() < lastNeededRow) {
      sheet.insertRowsAfter(sheet.getMaxRows(), lastNeededRow - sheet.getMaxRows());
    }
    const rowFormats = keys.map((k) => formatForFieldType(schema[k]));
    for (let offset = 0; offset < items.length; offset += SHEET_WRITE_CHUNK_ROWS) {
      const count = Math.min(SHEET_WRITE_CHUNK_ROWS, items.length - offset);
      const fmtMatrix = Array.from({ length: count }, () => [...rowFormats]);
      sheet.getRange(startRow + offset, 1, count, numCols).setNumberFormats(fmtMatrix);
    }
    writeMatrixInChunks(sheet, startRow, 1, toSheetValueMatrix(keys, items), SHEET_WRITE_CHUNK_ROWS);
  }
  function writeTableRows(sheet, keys, schema, items) {
    const numCols = keys.length;
    const totalRows = items.length + 1;
    if (sheet.getMaxColumns() < numCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < totalRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
    }
    const rowFormats = keys.map((k) => formatForFieldType(schema[k]));
    for (let offset = 0; offset < totalRows; offset += SHEET_WRITE_CHUNK_ROWS) {
      const count = Math.min(SHEET_WRITE_CHUNK_ROWS, totalRows - offset);
      const fmtMatrix = Array.from(
        { length: count },
        () => offset === 0 ? keys.map(() => "@") : [...rowFormats]
      );
      if (offset === 0 && count > 1) {
        for (let r = 1; r < count; r++) {
          fmtMatrix[r] = [...rowFormats];
        }
      }
      sheet.getRange(1 + offset, 1, count, numCols).setNumberFormats(fmtMatrix);
    }
    const matrix = [[...keys], ...toSheetValueMatrix(keys, items)];
    writeMatrixInChunks(sheet, 1, 1, matrix, SHEET_WRITE_CHUNK_ROWS);
    trimSheetGrid(sheet, totalRows, numCols);
  }
  function readTableRows(sheet, keys, schema) {
    if (!sheet)
      return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2)
      return [];
    const numCols = sheet.getMaxColumns();
    if (numCols < 1)
      return [];
    const allValues = sheet.getRange(1, 1, lastRow, numCols).getValues();
    const headerRow = allValues[0] ?? [];
    const colIndexByKey = /* @__PURE__ */ new Map();
    for (let c = 0; c < headerRow.length; c++) {
      const headerName = String(headerRow[c] ?? "").trim();
      if (headerName && !colIndexByKey.has(headerName)) {
        colIndexByKey.set(headerName, c);
      }
    }
    const out = [];
    for (let r = 1; r < allValues.length; r++) {
      const rawRow = allValues[r];
      const obj = {};
      let hasNonEmptyCell = false;
      for (const key of keys) {
        const k = String(key);
        const colIdx = colIndexByKey.get(k);
        const cell = colIdx === void 0 ? void 0 : rawRow[colIdx];
        if (cell !== "" && cell !== null && cell !== void 0) {
          hasNonEmptyCell = true;
        }
        obj[k] = coerceStagingCell(cell, schema[k] ?? "STRING");
      }
      if (hasNonEmptyCell) {
        out.push(obj);
      }
    }
    return out;
  }
  function coerceStagingCell(cell, fieldType) {
    if (fieldType === "BOOLEAN") {
      return parseSheetBoolean(cell);
    }
    if (fieldType === "NUMBER") {
      if (cell === "" || cell === null || cell === void 0)
        return null;
      const num = Number(cell);
      return Number.isFinite(num) ? num : null;
    }
    if (fieldType === "NULLABLE_STRING") {
      if (cell === "" || cell === null || cell === void 0)
        return null;
      if (cell instanceof Date) {
        return cell.toISOString().replace("T", " ").slice(0, 19);
      }
      return unescapeFormulaString(String(cell));
    }
    if (cell === null || cell === void 0)
      return "";
    if (cell instanceof Date) {
      return cell.toISOString().replace("T", " ").slice(0, 19);
    }
    return unescapeFormulaString(String(cell));
  }

  // src/appscript/run-state.ts
  var DOCUMENT_LOCK_TIMEOUT_MS = 3e4;
  var DATE_SCOPED_KEY_REGEX = /^(DAILY_CONFIG_LOCK_|FX_RATES_|PUBLISH_STATUS_|PUBLISH_SIGNATURE_)(\d{4}-\d{2}-\d{2})/;
  var activeSheetLocks = /* @__PURE__ */ new WeakSet();
  function getOrCreateRunStateSheet(spreadsheet) {
    let sheet = spreadsheet.getSheetByName(RUN_STATE_SHEET_NAME);
    if (!sheet) {
      sheet = getOrCreateSheet(spreadsheet, RUN_STATE_SHEET_NAME, { hidden: true });
      sheet.getRange(1, 1, 1, 4).setNumberFormat("@");
      sheet.getRange(1, 1, 1, 4).setValues([["Key", "Value", "UpdatedAt", "Notes"]]);
    }
    return sheet;
  }
  function withSheetLock(spreadsheet, ownerId, fn) {
    if (activeSheetLocks.has(spreadsheet)) {
      return fn();
    }
    const docLock = typeof LockService !== "undefined" && typeof LockService.getDocumentLock === "function" ? LockService.getDocumentLock() : null;
    if (!docLock) {
      throw new Error(
        `withSheetLock (${ownerId}) must run in the Sheet-bound script (LockService document lock unavailable).`
      );
    }
    docLock.waitLock(DOCUMENT_LOCK_TIMEOUT_MS);
    activeSheetLocks.add(spreadsheet);
    try {
      return fn();
    } finally {
      try {
        if (typeof SpreadsheetApp !== "undefined") {
          SpreadsheetApp.flush();
        }
      } finally {
        activeSheetLocks.delete(spreadsheet);
        docLock.releaseLock();
      }
    }
  }
  function isExpiredDateScopedKey(key, activeDateStr) {
    const match = DATE_SCOPED_KEY_REGEX.exec(key);
    if (!match)
      return false;
    return match[2] !== activeDateStr;
  }
  function mutateAndCompactRunState(spreadsheet, mutator, activeDateStr) {
    withSheetLock(spreadsheet, "RUN_STATE", () => {
      const sheet = getOrCreateRunStateSheet(spreadsheet);
      const prevLastRow = Math.max(1, sheet.getLastRow());
      const nowIso = (/* @__PURE__ */ new Date()).toISOString();
      const currentRows = [];
      if (prevLastRow >= 2) {
        const rawRows = sheet.getRange(2, 1, prevLastRow - 1, 4).getValues();
        for (const r of rawRows) {
          const k = String(r[0] ?? "").trim();
          if (!k)
            continue;
          if (activeDateStr && isExpiredDateScopedKey(k, activeDateStr))
            continue;
          currentRows.push([k, String(r[1] ?? ""), String(r[2] ?? ""), String(r[3] ?? "")]);
        }
      }
      const nextRows = mutator(currentRows, nowIso);
      const targetRowCount = 1 + nextRows.length;
      const minSheetRows = Math.max(1, targetRowCount);
      const writeRowCount = Math.max(prevLastRow, minSheetRows);
      if (writeRowCount > sheet.getMaxRows()) {
        sheet.insertRowsAfter(sheet.getMaxRows(), writeRowCount - sheet.getMaxRows() + 5);
      }
      const matrix = [["Key", "Value", "UpdatedAt", "Notes"], ...nextRows];
      while (matrix.length < writeRowCount) {
        matrix.push(["", "", "", ""]);
      }
      const range = sheet.getRange(1, 1, matrix.length, 4);
      range.setNumberFormat("@");
      range.setValues(matrix);
      if (prevLastRow > minSheetRows) {
        sheet.deleteRows(minSheetRows + 1, prevLastRow - minSheetRows);
      }
    });
  }
  function upsertRow(rows, key, value, nowIso, notes = "") {
    const existingIdx = rows.findIndex((r) => r[0] === key);
    if (existingIdx >= 0) {
      rows[existingIdx] = [key, value, nowIso, notes];
    } else {
      rows.push([key, value, nowIso, notes]);
    }
  }
  function setRunStateValues(spreadsheet, entries, activeDateStr) {
    if (entries.length === 0)
      return;
    mutateAndCompactRunState(
      spreadsheet,
      (rows, nowIso) => {
        for (const e of entries) {
          upsertRow(rows, e.key, e.value, nowIso, e.notes ?? "");
        }
        return rows;
      },
      activeDateStr
    );
  }
  function setRunStateValue(spreadsheet, key, value, notes = "", activeDateStr) {
    mutateAndCompactRunState(
      spreadsheet,
      (rows, nowIso) => {
        upsertRow(rows, key, value, nowIso, notes);
        return rows;
      },
      activeDateStr
    );
  }
  function getOrLockDailyConfig(spreadsheet, dateStr, readCurrentConfig) {
    const existing = readLockedDailyConfig(spreadsheet, dateStr);
    if (existing) {
      return existing;
    }
    const currentConfig = readCurrentConfig();
    setRunStateValue(
      spreadsheet,
      `DAILY_CONFIG_LOCK_${dateStr}`,
      serializeLockedDailyConfig(currentConfig),
      `Locked daily config for ${dateStr}`,
      dateStr
    );
    return currentConfig;
  }
  function resetTodayPublisherState(spreadsheet, dateStr) {
    const prefixes = [
      `DAILY_CONFIG_LOCK_${dateStr}`,
      `FX_RATES_${dateStr}`,
      `PUBLISH_STATUS_${dateStr}`,
      `PUBLISH_SIGNATURE_${dateStr}`
    ];
    mutateAndCompactRunState(
      spreadsheet,
      (rows) => rows.filter((r) => !prefixes.some((p) => r[0].startsWith(p))),
      dateStr
    );
  }

  // src/appscript/fx-rates.ts
  var FX_MAX_RETRIES = 3;
  var FX_SCRATCH_CELL = "G1";
  function resolveDailyFxRates(options) {
    const target = normalizeCurrencyCode(options.targetCurrency);
    const uniqueSources = Array.from(
      new Set(
        options.sourceCurrencies.map(normalizeCurrencyCode).filter((c) => c !== null)
      )
    );
    if (!target) {
      return Object.fromEntries(uniqueSources.map((c) => [c, 1]));
    }
    const dailyCacheKey = `FX_RATES_${options.dateStr}_${target}`;
    const stateMap = readRunStateMap(options.spreadsheet);
    const cachedJson = getRunStateValue(options.spreadsheet, dailyCacheKey, stateMap);
    const cachedMap = cachedJson ? safeParseRates(cachedJson) : {};
    let updated = false;
    const lastGoodUpdates = [];
    for (const src of uniqueSources) {
      if (src === target) {
        if (cachedMap[src] !== 1) {
          cachedMap[src] = 1;
          updated = true;
        }
        continue;
      }
      if (cachedMap[src]) {
        continue;
      }
      const { rate, fromLiveGoogleFinance } = fetchSinglePairRate(
        options.spreadsheet,
        src,
        target,
        stateMap
      );
      cachedMap[src] = rate;
      if (fromLiveGoogleFinance && rate !== null) {
        lastGoodUpdates.push({ key: `FX_LAST_GOOD_${src}_${target}`, value: String(rate) });
      }
      updated = true;
    }
    if (updated) {
      setRunStateValues(
        options.spreadsheet,
        [...lastGoodUpdates, { key: dailyCacheKey, value: JSON.stringify(cachedMap) }],
        options.dateStr
      );
    }
    return cachedMap;
  }
  function fetchSinglePairRate(spreadsheet, fromCurrency, toCurrency, stateMap) {
    if (!normalizeCurrencyCode(fromCurrency) || !normalizeCurrencyCode(toCurrency)) {
      return { rate: null, fromLiveGoogleFinance: false };
    }
    const sheet = getOrCreateSheet(spreadsheet, RUN_STATE_SHEET_NAME, { hidden: true });
    const scratchRange = sheet.getRange(FX_SCRATCH_CELL);
    const formula = `=GOOGLEFINANCE("CURRENCY:${fromCurrency}${toCurrency}")`;
    try {
      scratchRange.setFormula(formula);
      for (let attempt = 1; attempt <= FX_MAX_RETRIES; attempt++) {
        SpreadsheetApp.flush();
        const rawVal = scratchRange.getValue();
        const numericVal = Number(rawVal);
        if (Number.isFinite(numericVal) && numericVal > 0) {
          return { rate: numericVal, fromLiveGoogleFinance: true };
        }
        if (attempt < FX_MAX_RETRIES) {
          Utilities.sleep(attempt * 1e3);
        }
      }
    } finally {
      scratchRange.setValue("");
    }
    const lastKnownGood = Number(
      getRunStateValue(spreadsheet, `FX_LAST_GOOD_${fromCurrency}_${toCurrency}`, stateMap)
    );
    if (Number.isFinite(lastKnownGood) && lastKnownGood > 0) {
      return { rate: lastKnownGood, fromLiveGoogleFinance: false };
    }
    return { rate: null, fromLiveGoogleFinance: false };
  }
  function safeParseRates(json) {
    try {
      return JSON.parse(json);
    } catch {
      return {};
    }
  }
  function applyCurrencyConversionToRow(row, targetCurrency, rates) {
    const normTarget = normalizeCurrencyCode(targetCurrency);
    const srcCurr = (row.currencyCode || "USD").trim().toUpperCase();
    if (!normTarget) {
      return syncGlobalColumnsWithRate(row, srcCurr, 1);
    }
    const resolvedRate = srcCurr === normTarget ? 1 : rates[srcCurr] ?? null;
    const validRate = typeof resolvedRate === "number" && Number.isFinite(resolvedRate) && resolvedRate > 0 ? resolvedRate : null;
    return syncGlobalColumnsWithRate(row, normTarget, validRate);
  }
  function syncGlobalColumnsWithRate(row, currencyCodeGlobal, rate) {
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
      moveBudgetAmountGlobal: multiplyNullable(row.moveBudgetAmount, rate)
    };
  }
  function multiplyNullable(val, rate) {
    if (rate === null || !Number.isFinite(rate)) {
      return null;
    }
    if (typeof val !== "number" || !Number.isFinite(val)) {
      return null;
    }
    return Math.round((val * rate + Number.EPSILON) * 100) / 100;
  }

  // src/shared/utils/base64.ts
  var BASE64_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var BASE64_LOOKUP = (() => {
    const lookup = { "-": 62, "_": 63 };
    for (let i = 0; i < BASE64_ALPHABET.length; i++) {
      lookup[BASE64_ALPHABET[i]] = i;
    }
    return lookup;
  })();
  function decodeBase64ToAscii(input) {
    if (typeof input !== "string")
      return null;
    const data = input.replace(/={1,2}$/, "");
    if (data.length % 4 === 1)
      return null;
    if (data.length !== input.length && input.length % 4 !== 0)
      return null;
    let out = "";
    let buffer = 0;
    let bits = 0;
    for (let i = 0; i < data.length; i++) {
      const ch = data[i];
      if (!Object.prototype.hasOwnProperty.call(BASE64_LOOKUP, ch))
        return null;
      buffer = buffer << 6 | BASE64_LOOKUP[ch];
      bits += 6;
      if (bits >= 8) {
        bits -= 8;
        const byte = buffer >> bits & 255;
        buffer &= (1 << bits) - 1;
        if (byte < 32 || byte > 126)
          return null;
        out += String.fromCharCode(byte);
      }
    }
    return out;
  }

  // src/appscript/processors/ads-ui-urls.ts
  var ADS_UI_BASE_URL = "https://ads.google.com/aw";
  function parseRecommendationIdInfo(recommendationId) {
    if (!recommendationId)
      return null;
    const encoded = String(recommendationId).split("/").pop()?.trim() ?? "";
    if (!encoded)
      return null;
    const decoded = decodeBase64ToAscii(encoded);
    if (!decoded)
      return null;
    const [ocid, recoTypeId] = decoded.split("-");
    if (!/^\d+$/.test(ocid ?? "") || /^0+$/.test(ocid) || !/^\d+$/.test(recoTypeId ?? ""))
      return null;
    return { ocid, recoTypeId };
  }
  function resolveAdsUiOcid(recommendationId, fallbackCid) {
    return parseRecommendationIdInfo(recommendationId)?.ocid ?? normalizeCid(fallbackCid);
  }
  function buildRecommendationsDetailsUrl(recommendationId, accountCid) {
    const info = parseRecommendationIdInfo(recommendationId);
    if (!info)
      return `${ADS_UI_BASE_URL}/recommendations?ocid=${normalizeCid(accountCid)}`;
    return `${ADS_UI_BASE_URL}/recommendations?ocid=${info.ocid}&opp=${info.recoTypeId}`;
  }
  function buildCampaignUrl(campaignId, recommendationId, accountCid) {
    return `${ADS_UI_BASE_URL}/overview?campaignId=${campaignId}&ocid=${resolveAdsUiOcid(recommendationId, accountCid)}`;
  }

  // src/appscript/processors/entity-keys.ts
  function extractResourceTrailingId(resourceName, collectionName) {
    if (!resourceName || resourceName === "N/A")
      return "";
    const marker = `/${collectionName}/`;
    const idx = resourceName.indexOf(marker);
    if (idx === -1)
      return "";
    return resourceName.slice(idx + marker.length).trim();
  }
  function getPortfolioOwnerCid(row) {
    return normalizeCid(row.portfolioOwnerId) || normalizeCid(row.accountId);
  }
  function getPortfolioEntityKey(row) {
    const strategyId = extractResourceTrailingId(row.portfolioStrategyResourceName, "biddingStrategies");
    return `${getPortfolioOwnerCid(row)}::${strategyId || row.portfolioStrategyResourceName || ""}`;
  }
  function getPortfolioGroupKey(row, perAccount = false) {
    const recId = extractResourceTrailingId(row.recommendationId, "recommendations") || String(row.recommendationId ?? "").split("/").pop() || "";
    const key = `${getPortfolioEntityKey(row)}::${recId}::${row.currencyCode}`;
    return perAccount ? `${key}::${normalizeCid(row.accountId)}` : key;
  }

  // src/shared/schema.ts
  var REPORT_SCHEMA_TEMPLATE = {
    // --- Ranking, Scope & Multi-MCC Metadata ---
    recommendationNumber: { defaultValue: 1, format: "#,##0", dummyValueForFormatting: 1 },
    recommendationScope: { defaultValue: "CAMPAIGN", format: "@", dummyValueForFormatting: "CAMPAIGN" },
    mccId: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "000-000-0000" },
    mccName: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DUMMY_MCC" },
    mccSyncStatus: { defaultValue: "OK", format: "@", dummyValueForFormatting: "OK" },
    mccLastUpdatedDate: { defaultValue: null, format: "yyyy-mm-dd", dummyValueForFormatting: "2026-10-14" },
    // --- Account & Recommendation Identity ---
    accountId: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "000-000-0000" },
    accountName: {
      defaultValue: "N/A",
      format: "@",
      dummyValueForFormatting: "!!! SCHEMA DEFINITION (FILTER ME OUT)"
    },
    timestamp: { defaultValue: null, format: "yyyy-mm-dd hh:mm:ss", dummyValueForFormatting: "2026-10-14 00:00:00" },
    recommendationId: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DUMMY_RECO_ID" },
    recommendationType: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DUMMY_TYPE" },
    recommendationsDetailsUrl: {
      defaultValue: "N/A",
      format: "@",
      dummyValueForFormatting: "https://ads.google.com"
    },
    // --- Campaign Details ---
    campaignUrl: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "https://ads.google.com" },
    campaignId: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DUMMY_SCHEMA_ROW" },
    campaignName: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DUMMY_CAMPAIGN" },
    campaignType: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "SEARCH" },
    campaignSubType: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "N/A" },
    campaignIsAiMax: { defaultValue: false, format: "General", dummyValueForFormatting: false },
    campaignAiMaxTextCustomizationEnabled: { defaultValue: false, format: "General", dummyValueForFormatting: false },
    campaignAiMaxFinalUrlExpansionEnabled: { defaultValue: false, format: "General", dummyValueForFormatting: false },
    currencyCode: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "USD" },
    currencyCodeGlobal: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "USD" },
    // --- Budget Details ---
    campaignBudgetAmount: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaignBudgetTotalAmount: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaignBudgetDeliveryMethod: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "STANDARD" },
    campaignBudgetType: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DAILY" },
    // --- Bidding & Portfolio Details ---
    portfolioCampaignsCount: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    portfolioMemberCampaignNames: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "N/A" },
    campaignBiddingStrategyType: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "MANUAL_CPC" },
    campaignBiddingTargetRoas: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    campaignBiddingTargetCpa: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    // --- Historical Performance Stats (Yesterday / 7d / 30d) ---
    campaignStatsYesterdayCost: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaignStats7DaysCost: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaign30DaysCost: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaign30DaysConversionsValue: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaign30DaysConversions: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    campaign30DaysRoas: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    campaign30DaysCpa: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    campaign30DaysSearchRankLostImpressionShare: {
      defaultValue: null,
      format: "0.00%",
      dummyValueForFormatting: 0
    },
    campaign30DaysSearchBudgetLostImpressionShare: {
      defaultValue: null,
      format: "0.00%",
      dummyValueForFormatting: 0
    },
    // --- Recommendation Targets & 7-Day Native Forecast Impact ---
    recommendationCurrentBudgetAmount: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationCurrentBudgetAmountGlobal: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationNewBudgetAmount: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationNewBudgetAmountGlobal: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationNewTargetCpa: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationNewTargetRoas: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    recommendationBaseCost: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationPotentialCost: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationBaseClicks: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationPotentialClicks: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationBaseConversions: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationPotentialConversions: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationBaseConversionsValue: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationBaseConversionsValueGlobal: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationPotentialConversionsValue: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationPotentialConversionsValueGlobal: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationBaseCpa: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationPotentialCpa: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationBaseRoas: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    recommendationPotentialRoas: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    recommendationBaseImpressions: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationPotentialImpressions: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationBaseVideoViews: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    recommendationPotentialVideoViews: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    // --- Derived Weekly & Budget Metrics ---
    campaignPercentageOfBudgetUsedYesterday: {
      defaultValue: null,
      format: "0.00%",
      dummyValueForFormatting: 0
    },
    weeklyCostIncrease: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    weeklyCostIncreaseGlobal: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    newWeeklyConversions: { defaultValue: null, format: "#,##0", dummyValueForFormatting: 0 },
    dailyBudget30Days: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    dailyBudgetDelta: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationCpaDelta: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    recommendationRoasDelta: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    // --- MOVE_UNUSED_BUDGET Specifics (Including Source Campaign Targets) ---
    moveBudgetAmount: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    moveBudgetSourceCampaigns: {
      defaultValue: "N/A",
      format: "@",
      dummyValueForFormatting: "CAMPAIGN_1, CAMPAIGN_2"
    },
    moveBudgetSourceBudgetName: {
      defaultValue: "N/A",
      format: "@",
      dummyValueForFormatting: "SOURCE_BUDGET"
    },
    moveBudgetSourceBudgetType: { defaultValue: "N/A", format: "@", dummyValueForFormatting: "DAILY" },
    moveBudgetSourceBudgetDeliveryMethod: {
      defaultValue: "N/A",
      format: "@",
      dummyValueForFormatting: "STANDARD"
    },
    moveBudgetSourceBudgetIsShared: { defaultValue: false, format: "General", dummyValueForFormatting: false },
    moveBudgetSourceBudgetAmount: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    moveBudgetSourceTargetCpa: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    },
    moveBudgetSourceTargetRoas: { defaultValue: null, format: "0.00", dummyValueForFormatting: 0 },
    // --- Appended Columns (82+; APPEND-ONLY, see the rule above) ---
    moveBudgetAmountGlobal: {
      defaultValue: null,
      format: "#,##0.00",
      dummyValueForFormatting: 0
    }
  };
  var SCHEMA_COLUMN_NAMES = Object.keys(REPORT_SCHEMA_TEMPLATE);
  function createEmptyReportRow() {
    const row = {};
    for (const key of SCHEMA_COLUMN_NAMES) {
      row[key] = REPORT_SCHEMA_TEMPLATE[key].defaultValue;
    }
    return row;
  }
  var SCHEMA_DUMMY_ACCOUNT_NAME = REPORT_SCHEMA_TEMPLATE.accountName.dummyValueForFormatting;
  function isDummyFormattingRow(row) {
    return row.accountName === SCHEMA_DUMMY_ACCOUNT_NAME;
  }
  function createDummyFormattingRow() {
    const row = {};
    for (const key of SCHEMA_COLUMN_NAMES) {
      row[key] = REPORT_SCHEMA_TEMPLATE[key].dummyValueForFormatting;
    }
    return row;
  }

  // src/appscript/processors/joiner.ts
  function joinStagingBatchToReportRows(batch) {
    const accountMap = /* @__PURE__ */ new Map();
    for (const acc of batch.accounts) {
      accountMap.set(`${normalizeCid(acc.mccId)}::${normalizeCid(acc.accountId)}`, acc);
    }
    const budgetMap = /* @__PURE__ */ new Map();
    for (const b of batch.budgets) {
      budgetMap.set(`${normalizeCid(b.mccId)}::${normalizeCid(b.accountId)}::${b.budgetResourceName}`, b);
    }
    const campaignMap = /* @__PURE__ */ new Map();
    for (const c of batch.campaigns) {
      campaignMap.set(`${normalizeCid(c.mccId)}::${normalizeCid(c.accountId)}::${c.campaignResourceName}`, c);
    }
    const rows = [];
    for (const rec of batch.recommendations) {
      const normMccCid = normalizeCid(rec.mccId);
      const normAccountCid = normalizeCid(rec.accountId);
      const account = accountMap.get(`${normMccCid}::${normAccountCid}`);
      const campaign = campaignMap.get(`${normMccCid}::${normAccountCid}::${rec.campaignResourceName}`);
      if (!campaign)
        continue;
      const budget = budgetMap.get(`${normMccCid}::${normAccountCid}::${campaign.budgetResourceName}`);
      const row = createEmptyReportRow();
      row.mccId = rec.mccId;
      row.mccName = account?.mccName ?? "N/A";
      row.accountId = rec.accountId;
      row.accountName = account?.accountName ?? "N/A";
      row.currencyCode = account?.currencyCode ?? "USD";
      row.currencyCodeGlobal = row.currencyCode;
      row.timestamp = account?.timestamp ?? null;
      row.recommendationId = rec.recommendationId;
      row.recommendationType = rec.recommendationType;
      row.recommendationsDetailsUrl = buildRecommendationsDetailsUrl(rec.recommendationId, rec.accountId);
      row.campaignUrl = buildCampaignUrl(campaign.campaignId, rec.recommendationId, rec.accountId);
      row.campaignId = campaign.campaignId;
      row.campaignName = campaign.campaignName;
      row.campaignType = campaign.campaignType;
      row.campaignSubType = campaign.campaignSubType;
      row.campaignIsAiMax = campaign.campaignIsAiMax;
      row.campaignAiMaxTextCustomizationEnabled = campaign.campaignAiMaxTextCustomizationEnabled;
      row.campaignAiMaxFinalUrlExpansionEnabled = campaign.campaignAiMaxFinalUrlExpansionEnabled;
      row.budgetResourceName = campaign.budgetResourceName;
      row.campaignBudgetName = budget?.campaignBudgetName ?? "N/A";
      row.campaignBudgetAmount = budget?.campaignBudgetAmount ?? null;
      row.campaignBudgetTotalAmount = budget?.campaignBudgetTotalAmount ?? null;
      row.campaignBudgetDeliveryMethod = budget?.campaignBudgetDeliveryMethod ?? "N/A";
      row.campaignBudgetType = budget?.campaignBudgetType ?? "N/A";
      row.campaignBudgetIsShared = budget?.campaignBudgetIsShared ?? false;
      row.campaignIsPortfolioBiddingStrategy = campaign.campaignIsPortfolioBiddingStrategy;
      row.campaignPortfolioBiddingStrategyName = campaign.campaignPortfolioBiddingStrategyName;
      row.portfolioStrategyResourceName = campaign.portfolioStrategyResourceName;
      row.portfolioOwnerId = campaign.portfolioOwnerId;
      row.campaignBiddingStrategyType = campaign.campaignBiddingStrategyType;
      row.campaignBiddingTargetCpa = campaign.campaignBiddingTargetCpa;
      row.campaignBiddingTargetRoas = campaign.campaignBiddingTargetRoas;
      row.campaignStatsYesterdayCost = campaign.campaignStatsYesterdayCost;
      row.campaignStats7DaysCost = campaign.campaignStats7DaysCost;
      row.campaign30DaysCost = campaign.campaign30DaysCost;
      row.campaign30DaysConversionsValue = campaign.campaign30DaysConversionsValue;
      row.campaign30DaysConversions = campaign.campaign30DaysConversions;
      const cost30d = campaign.campaign30DaysCost ?? 0;
      const conv30d = campaign.campaign30DaysConversions ?? 0;
      const convVal30d = campaign.campaign30DaysConversionsValue ?? 0;
      row.campaign30DaysRoas = campaign.campaign30DaysCost !== null && cost30d > 0 ? convVal30d / cost30d : null;
      row.campaign30DaysCpa = campaign.campaign30DaysConversions !== null && conv30d > 0 ? cost30d / conv30d : null;
      row.campaign30DaysSearchRankLostImpressionShare = campaign.campaign30DaysSearchRankLostImpressionShare;
      row.campaign30DaysSearchBudgetLostImpressionShare = campaign.campaign30DaysSearchBudgetLostImpressionShare;
      row.recommendationCurrentBudgetAmount = rec.recommendationCurrentBudgetAmount;
      row.recommendationCurrentBudgetAmountGlobal = rec.recommendationCurrentBudgetAmount;
      row.recommendationNewBudgetAmount = rec.recommendationNewBudgetAmount;
      row.recommendationNewBudgetAmountGlobal = rec.recommendationNewBudgetAmount;
      row.recommendationNewTargetCpa = rec.recommendationNewTargetCpa;
      row.recommendationNewTargetRoas = rec.recommendationNewTargetRoas;
      row.recommendationBaseCost = rec.recommendationBaseCost;
      row.recommendationPotentialCost = rec.recommendationPotentialCost;
      row.recommendationBaseClicks = rec.recommendationBaseClicks;
      row.recommendationPotentialClicks = rec.recommendationPotentialClicks;
      row.recommendationBaseConversions = rec.recommendationBaseConversions;
      row.recommendationPotentialConversions = rec.recommendationPotentialConversions;
      row.recommendationBaseConversionsValue = rec.recommendationBaseConversionsValue;
      row.recommendationBaseConversionsValueGlobal = rec.recommendationBaseConversionsValue;
      row.recommendationPotentialConversionsValue = rec.recommendationPotentialConversionsValue;
      row.recommendationPotentialConversionsValueGlobal = rec.recommendationPotentialConversionsValue;
      row.recommendationBaseCpa = rec.recommendationBaseCost !== null && rec.recommendationBaseConversions !== null && rec.recommendationBaseConversions > 0 ? rec.recommendationBaseCost / rec.recommendationBaseConversions : null;
      row.recommendationPotentialCpa = rec.recommendationPotentialCost !== null && rec.recommendationPotentialConversions !== null && rec.recommendationPotentialConversions > 0 ? rec.recommendationPotentialCost / rec.recommendationPotentialConversions : null;
      row.recommendationBaseRoas = rec.recommendationBaseCost !== null && rec.recommendationBaseCost > 0 && rec.recommendationBaseConversionsValue !== null ? rec.recommendationBaseConversionsValue / rec.recommendationBaseCost : null;
      row.recommendationPotentialRoas = rec.recommendationPotentialCost !== null && rec.recommendationPotentialCost > 0 && rec.recommendationPotentialConversionsValue !== null ? rec.recommendationPotentialConversionsValue / rec.recommendationPotentialCost : null;
      row.recommendationBaseImpressions = rec.recommendationBaseImpressions;
      row.recommendationPotentialImpressions = rec.recommendationPotentialImpressions;
      row.recommendationBaseVideoViews = rec.recommendationBaseVideoViews;
      row.recommendationPotentialVideoViews = rec.recommendationPotentialVideoViews;
      if (rec.recommendationType === "MOVE_UNUSED_BUDGET") {
        if (rec.recommendationCurrentBudgetAmount !== null && rec.recommendationNewBudgetAmount !== null) {
          row.moveBudgetAmount = rec.recommendationNewBudgetAmount - rec.recommendationCurrentBudgetAmount;
        }
        const sourceBudget = rec.moveBudgetSourceBudgetResourceName ? budgetMap.get(
          `${normMccCid}::${normAccountCid}::${rec.moveBudgetSourceBudgetResourceName}`
        ) : void 0;
        if (sourceBudget) {
          row.moveBudgetSourceBudgetName = sourceBudget.campaignBudgetName;
          row.moveBudgetSourceBudgetType = sourceBudget.campaignBudgetType;
          row.moveBudgetSourceBudgetDeliveryMethod = sourceBudget.campaignBudgetDeliveryMethod;
          row.moveBudgetSourceBudgetIsShared = sourceBudget.campaignBudgetIsShared;
          row.moveBudgetSourceBudgetAmount = sourceBudget.campaignBudgetAmount;
        }
        if (rec.moveBudgetSourceCampaigns && rec.moveBudgetSourceCampaigns !== "N/A") {
          row.moveBudgetSourceCampaigns = rec.moveBudgetSourceCampaigns;
        }
        row.moveBudgetSourceTargetCpa = rec.moveBudgetSourceTargetCpa;
        row.moveBudgetSourceTargetRoas = rec.moveBudgetSourceTargetRoas;
      }
      populateDerivedRowMetrics(row);
      rows.push(row);
    }
    return rows;
  }
  function populateDerivedRowMetrics(row) {
    const currentBudget = row.recommendationCurrentBudgetAmount ?? row.campaignBudgetAmount;
    if (typeof currentBudget === "number" && currentBudget > 0 && typeof row.campaignStatsYesterdayCost === "number") {
      row.campaignPercentageOfBudgetUsedYesterday = row.campaignStatsYesterdayCost / currentBudget;
    } else {
      row.campaignPercentageOfBudgetUsedYesterday = null;
    }
    row.weeklyCostIncrease = subtractNullable(row.recommendationPotentialCost, row.recommendationBaseCost);
    row.weeklyCostIncreaseGlobal = row.weeklyCostIncrease;
    row.moveBudgetAmountGlobal = row.moveBudgetAmount;
    row.newWeeklyConversions = subtractNullable(
      row.recommendationPotentialConversions,
      row.recommendationBaseConversions
    );
    if (typeof row.campaignBudgetAmount === "number") {
      row.dailyBudget30Days = row.campaignBudgetAmount * 30;
      if (typeof row.campaign30DaysCost === "number") {
        row.dailyBudgetDelta = row.dailyBudget30Days - row.campaign30DaysCost;
      }
    }
    row.recommendationCpaDelta = subtractNullable(
      row.recommendationBaseCpa,
      row.recommendationPotentialCpa
    );
    row.recommendationRoasDelta = subtractNullable(
      row.recommendationBaseRoas,
      row.recommendationPotentialRoas
    );
  }
  function subtractNullable(a, b) {
    if (typeof a === "number" && typeof b === "number") {
      return a - b;
    }
    return null;
  }

  // src/appscript/processors/consolidator.ts
  function consolidateRowsByScope(rows, options = {}) {
    const crossAccountMode = options.crossAccountPortfolioMode ?? "CONSOLIDATED_STRATEGY";
    const portfolioGroups = /* @__PURE__ */ new Map();
    const sharedBudgetGroups = /* @__PURE__ */ new Map();
    const campaignRows = [];
    for (const row of rows) {
      const isPortfolio = row.campaignIsPortfolioBiddingStrategy === true && row.portfolioStrategyResourceName !== "N/A" && Boolean(row.portfolioStrategyResourceName);
      const isTargetOrMultiCampaignPortfolioRec = isPortfolio && (row.recommendationType === "RAISE_TARGET_CPA" || row.recommendationType === "LOWER_TARGET_ROAS");
      if (isTargetOrMultiCampaignPortfolioRec) {
        const groupKey = getPortfolioGroupKey(row, crossAccountMode === "SPLIT_BY_ACCOUNT");
        const list = portfolioGroups.get(groupKey) ?? [];
        list.push(row);
        portfolioGroups.set(groupKey, list);
        continue;
      }
      const isSharedBudgetRec = row.campaignBudgetIsShared === true && (row.recommendationType === "CAMPAIGN_BUDGET" || row.recommendationType === "FORECASTING_CAMPAIGN_BUDGET" || row.recommendationType === "MARGINAL_ROI_CAMPAIGN_BUDGET");
      if (isSharedBudgetRec) {
        const budgetKey = row.budgetResourceName || row.campaignBudgetName;
        const groupKey = `${row.recommendationId}::${row.accountId}::${budgetKey}`;
        const list = sharedBudgetGroups.get(groupKey) ?? [];
        list.push(row);
        sharedBudgetGroups.set(groupKey, list);
        continue;
      }
      const clone = { ...row, recommendationScope: "CAMPAIGN" };
      campaignRows.push(clone);
    }
    const out = [...campaignRows];
    for (const group of portfolioGroups.values()) {
      out.push(mergeGroupRows(group, "PORTFOLIO"));
    }
    for (const group of sharedBudgetGroups.values()) {
      out.push(mergeGroupRows(group, "SHARED_BUDGET"));
    }
    return out;
  }
  function mergeGroupRows(group, scope) {
    const uniqueCampaigns = /* @__PURE__ */ new Map();
    for (const r of group) {
      uniqueCampaigns.set(`${r.accountId}::${r.campaignId}`, r);
    }
    const members = Array.from(uniqueCampaigns.values());
    const base = { ...members[0] };
    base.recommendationScope = scope;
    const memberNames = members.map((m) => m.campaignName);
    base.portfolioCampaignsCount = members.length;
    base.portfolioMemberCampaignNames = memberNames.join(", ");
    if (scope === "PORTFOLIO") {
      base.campaignName = base.campaignPortfolioBiddingStrategyName || base.campaignName;
      const strategyId = extractResourceTrailingId(base.portfolioStrategyResourceName, "biddingStrategies");
      base.campaignId = strategyId || "N/A";
      const ownerCid = getPortfolioOwnerCid(base);
      const ocid = ownerCid === normalizeCid(base.accountId) ? resolveAdsUiOcid(base.recommendationId, ownerCid) : ownerCid;
      base.campaignUrl = strategyId ? `https://ads.google.com/aw/bidreport/portfolio?strategyId=${strategyId}&ocid=${ocid}` : `https://ads.google.com/aw/bidstrategies?ocid=${ocid}`;
    } else if (scope === "SHARED_BUDGET") {
      base.campaignName = base.campaignBudgetName || base.campaignName;
      const budgetId = extractResourceTrailingId(base.budgetResourceName, "campaignBudgets");
      base.campaignId = budgetId || "N/A";
      const ocid = resolveAdsUiOcid(base.recommendationId, base.accountId);
      base.campaignUrl = budgetId ? `https://ads.google.com/aw/budget/report?budgetId=${budgetId}&ocid=${ocid}` : `https://ads.google.com/aw/sharedbudgets?ocid=${ocid}`;
    }
    base.campaignStatsYesterdayCost = sumNullable(members.map((m) => m.campaignStatsYesterdayCost));
    base.campaignStats7DaysCost = sumNullable(members.map((m) => m.campaignStats7DaysCost));
    base.campaign30DaysCost = sumNullable(members.map((m) => m.campaign30DaysCost));
    base.campaign30DaysConversions = sumNullable(members.map((m) => m.campaign30DaysConversions));
    base.campaign30DaysConversionsValue = sumNullable(
      members.map((m) => m.campaign30DaysConversionsValue)
    );
    if (scope === "PORTFOLIO") {
      const distinctBudgets = /* @__PURE__ */ new Map();
      for (const m of members) {
        distinctBudgets.set(`${m.accountId}::${m.budgetResourceName || m.campaignId}`, m.campaignBudgetAmount);
      }
      base.campaignBudgetAmount = sumNullable(Array.from(distinctBudgets.values()));
    }
    const totalCost30d = base.campaign30DaysCost ?? 0;
    const totalConv30d = base.campaign30DaysConversions ?? 0;
    const totalConvVal30d = base.campaign30DaysConversionsValue ?? 0;
    base.campaign30DaysCpa = totalConv30d > 0 ? totalCost30d / totalConv30d : null;
    base.campaign30DaysRoas = totalCost30d > 0 ? totalConvVal30d / totalCost30d : null;
    base.campaign30DaysSearchRankLostImpressionShare = costWeightedAverage(
      members,
      (m) => m.campaign30DaysSearchRankLostImpressionShare
    );
    base.campaign30DaysSearchBudgetLostImpressionShare = costWeightedAverage(
      members,
      (m) => m.campaign30DaysSearchBudgetLostImpressionShare
    );
    populateDerivedRowMetrics(base);
    return base;
  }
  function sumNullable(values) {
    let hasNumber = false;
    let total = 0;
    for (const v of values) {
      if (typeof v === "number" && Number.isFinite(v)) {
        hasNumber = true;
        total += v;
      }
    }
    return hasNumber ? total : null;
  }
  function costWeightedAverage(members, selector) {
    let weightedSum = 0;
    let weightTotal = 0;
    for (const m of members) {
      const metric = selector(m);
      const cost = m.campaign30DaysCost ?? 0;
      if (typeof metric === "number" && Number.isFinite(metric) && cost > 0) {
        weightedSum += cost * metric;
        weightTotal += cost;
      }
    }
    return weightTotal > 0 ? weightedSum / weightTotal : null;
  }

  // src/appscript/processors/ranking.ts
  function assignRecommendationNumbers(rows) {
    const byPartition = /* @__PURE__ */ new Map();
    for (const row of rows) {
      const clone = { ...row };
      const partitionKey = getPartitionKey(clone);
      const list = byPartition.get(partitionKey) ?? [];
      list.push(clone);
      byPartition.set(partitionKey, list);
    }
    const result = [];
    for (const group of byPartition.values()) {
      group.sort((a, b) => {
        const typeCmp = a.recommendationType.localeCompare(b.recommendationType);
        if (typeCmp !== 0) {
          return typeCmp;
        }
        return a.recommendationId.localeCompare(b.recommendationId);
      });
      for (let i = 0; i < group.length; i++) {
        group[i].recommendationNumber = i + 1;
        result.push(group[i]);
      }
    }
    return result;
  }
  function getPartitionKey(row) {
    if (row.recommendationScope === "PORTFOLIO") {
      return `PORTFOLIO::${getPortfolioEntityKey(row)}`;
    }
    if (row.recommendationScope === "SHARED_BUDGET") {
      const budgetKey = row.budgetResourceName || row.campaignBudgetName;
      return `SHARED_BUDGET::${row.accountId}::${budgetKey}`;
    }
    return `CAMPAIGN::${row.accountId}::${row.campaignId}`;
  }

  // src/appscript/sheets/data-writer.ts
  var DATA_SHEET_NAME = "Data";
  function writeReportRowsToSheet(sheet, rows, workbook) {
    const runWrite = () => {
      const effectiveRows = [
        createDummyFormattingRow(),
        ...rows.filter((r) => !isDummyFormattingRow(r))
      ];
      const numCols = SCHEMA_COLUMN_NAMES.length;
      const totalSheetRows = effectiveRows.length + 1;
      const currentMaxCols = sheet.getMaxColumns();
      if (currentMaxCols < numCols) {
        sheet.insertColumnsAfter(currentMaxCols, numCols - currentMaxCols);
      }
      const currentMaxRows = sheet.getMaxRows();
      if (currentMaxRows < totalSheetRows) {
        sheet.insertRowsAfter(currentMaxRows, totalSheetRows - currentMaxRows);
      }
      const singleRowFormats = SCHEMA_COLUMN_NAMES.map(
        (colKey) => REPORT_SCHEMA_TEMPLATE[colKey].format
      );
      writeFormatsInChunks(
        sheet,
        2,
        1,
        effectiveRows.length,
        singleRowFormats,
        SHEET_WRITE_CHUNK_ROWS
      );
      const valuesMatrix = [
        SCHEMA_COLUMN_NAMES,
        ...toSheetValueMatrix(SCHEMA_COLUMN_NAMES, effectiveRows)
      ];
      writeMatrixInChunks(sheet, 1, 1, valuesMatrix, SHEET_WRITE_CHUNK_ROWS);
      trimSheetGrid(sheet, totalSheetRows, numCols);
    };
    if (workbook) {
      withSheetLock(workbook, "DATA_WRITE", runWrite);
    } else {
      runWrite();
    }
  }
  function readReportRowsFromSheet(sheet, timeZone = resolveDefaultTimeZone()) {
    if (!sheet)
      return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2)
      return [];
    const numCols = sheet.getMaxColumns();
    if (numCols < 1)
      return [];
    const allValues = sheet.getRange(1, 1, lastRow, numCols).getValues();
    const headerRow = allValues[0] ?? [];
    const colIndexByKey = /* @__PURE__ */ new Map();
    for (let c = 0; c < headerRow.length; c++) {
      const headerName = String(headerRow[c] ?? "").trim();
      if (headerName && !colIndexByKey.has(headerName)) {
        colIndexByKey.set(headerName, c);
      }
    }
    const out = [];
    for (let r = 1; r < allValues.length; r++) {
      const rawRow = allValues[r];
      const obj = {};
      for (const key of SCHEMA_COLUMN_NAMES) {
        const colIdx = colIndexByKey.get(key);
        const cellVal = colIdx === void 0 ? void 0 : rawRow[colIdx];
        obj[key] = coerceReportCell(key, cellVal, timeZone);
      }
      const typed = obj;
      if (isDummyFormattingRow(typed)) {
        continue;
      }
      out.push(typed);
    }
    return out;
  }
  function coerceReportCell(key, cellVal, timeZone) {
    const def = REPORT_SCHEMA_TEMPLATE[key];
    if (cellVal === "" || cellVal === null || cellVal === void 0) {
      return def.defaultValue;
    }
    if (typeof def.defaultValue === "boolean") {
      return parseSheetBoolean(cellVal);
    }
    if (isDateFormat(def.format)) {
      return coerceDateCell(cellVal, def.format, def.defaultValue, timeZone);
    }
    if (def.format === "@") {
      if (cellVal instanceof Date) {
        return cellVal.toISOString().replace("T", " ").slice(0, 19);
      }
      return unescapeFormulaString(String(cellVal));
    }
    const num = Number(cellVal);
    return Number.isFinite(num) ? num : def.defaultValue;
  }
  var MIN_EPOCH_MS = 1e11;
  function isDateFormat(format) {
    return format.startsWith("yyyy");
  }
  function coerceDateCell(cellVal, format, defaultValue, timeZone) {
    const pattern = format.includes("hh") ? "yyyy-MM-dd HH:mm:ss" : "yyyy-MM-dd";
    if (cellVal instanceof Date) {
      return formatDateInTimeZone(cellVal, timeZone, pattern) ?? defaultValue;
    }
    if (typeof cellVal === "number") {
      if (Number.isFinite(cellVal) && cellVal > MIN_EPOCH_MS) {
        return formatDateInTimeZone(new Date(cellVal), timeZone, pattern) ?? defaultValue;
      }
      return defaultValue;
    }
    return String(cellVal);
  }
  function formatDateInTimeZone(date, timeZone, pattern) {
    if (Number.isNaN(date.getTime()))
      return null;
    if (typeof Utilities !== "undefined" && typeof Utilities.formatDate === "function") {
      return Utilities.formatDate(date, timeZone || "UTC", pattern);
    }
    const iso = date.toISOString().replace("T", " ");
    return pattern === "yyyy-MM-dd" ? iso.slice(0, 10) : iso.slice(0, 19);
  }
  function resolveDefaultTimeZone() {
    if (typeof Session !== "undefined" && typeof Session.getScriptTimeZone === "function") {
      return Session.getScriptTimeZone() || "UTC";
    }
    return "UTC";
  }
  function writeFormatsInChunks(sheet, startRow, startCol, totalDataRows, singleRowFormats, chunkRows) {
    const numCols = singleRowFormats.length;
    for (let offset = 0; offset < totalDataRows; offset += chunkRows) {
      const count = Math.min(chunkRows, totalDataRows - offset);
      const formatsSlice = Array.from({ length: count }, () => [...singleRowFormats]);
      sheet.getRange(startRow + offset, startCol, count, numCols).setNumberFormats(formatsSlice);
    }
  }

  // src/appscript/publisher.ts
  var MAX_STATUS_ERROR_LENGTH = 80;
  var MAX_LISTED_DUPLICATE_ACCOUNTS = 3;
  function publishFromStagingToData(options = {}) {
    const spreadsheet = options.spreadsheet ?? SpreadsheetApp.getActiveSpreadsheet();
    const dateStr = resolveSpreadsheetDate(spreadsheet);
    const force = Boolean(options.force);
    withSheetLock(spreadsheet, "SHEET_PUBLISH", () => {
      const configSheet = getOrCreateSheet(spreadsheet, CONFIG_SHEET_NAME, { index: 0 });
      const syncInfo = syncConfigFromMccTabs(spreadsheet, configSheet);
      const runStateMap = readRunStateMap(spreadsheet);
      const plan = buildPublishPlan(spreadsheet, configSheet, dateStr, runStateMap, syncInfo);
      if (!plan || shouldSkipPublish(plan, runStateMap, force)) {
        return;
      }
      const dataSheet = getOrCreateSheet(spreadsheet, DATA_SHEET_NAME);
      const existingDataRows = readReportRowsFromSheet(
        dataSheet,
        spreadsheet.getSpreadsheetTimeZone() || void 0
      );
      const assembled = assembleRows(spreadsheet, plan, existingDataRows);
      const publishedRows = convertCurrencies(spreadsheet, plan, assembled.rows);
      writeReportRowsToSheet(dataSheet, publishedRows);
      recordPublish(spreadsheet, plan, publishedRows.length, assembled);
      updateStatusBanner(configSheet, plan, publishedRows.length, assembled);
    });
  }
  function syncConfigFromMccTabs(spreadsheet, configSheet) {
    const syncInfo = syncAllMccSheetsToConfig(spreadsheet, configSheet);
    if (syncInfo.newlyRegisteredMccs > 0 && syncInfo.latestStatusMessages.length > 0) {
      setStatusBannerIfChanged(
        configSheet,
        syncInfo.latestStatusMessages[syncInfo.latestStatusMessages.length - 1]
      );
    }
    return syncInfo;
  }
  function buildPublishPlan(spreadsheet, configSheet, dateStr, runStateMap, syncInfo) {
    const summaries = getMccSelectionSummary(configSheet);
    if (summaries.length === 0) {
      return null;
    }
    const lockedConfig = readLockedDailyConfig(spreadsheet, dateStr, runStateMap);
    const dailyConfig = lockedConfig ?? readGlobalConfig(configSheet);
    const activeMccs = summaries.map((s) => ({
      ...s,
      selectedCids: dailyConfig.selectedCidsByMcc?.[s.mccId] ?? getSelectedCidsForMcc(configSheet, s.mccId)
    })).filter((s) => s.selectedCids.length > 0);
    if (activeMccs.length === 0) {
      return null;
    }
    if (!lockedConfig) {
      getOrLockDailyConfig(spreadsheet, dateStr, () => dailyConfig);
    }
    const stateByMcc = new Map(
      activeMccs.map((a) => [a.mccId, getMccCheckpoint(spreadsheet, dateStr, a.mccId)])
    );
    const readyMccIds = new Set(
      activeMccs.filter((a) => {
        const st = stateByMcc.get(a.mccId);
        return st?.status === "COMPLETED" && st.dateStr === dateStr;
      }).map((a) => a.mccId)
    );
    const existingDataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    const setupWarnings = buildSetupWarnings(syncInfo);
    const signature = JSON.stringify({
      dateStr,
      targetCurrency: dailyConfig.targetCurrency,
      excludeShoppingAndPmax: dailyConfig.excludeShoppingAndPmax,
      mccs: activeMccs.map((a) => {
        const st = stateByMcc.get(a.mccId);
        return {
          mccId: a.mccId,
          selectedCids: a.selectedCids,
          isReadyToday: readyMccIds.has(a.mccId),
          lastCompletedDate: st?.lastCompletedDate ?? null,
          failedCids: Object.keys(st?.failedCids ?? {}).sort()
        };
      }),
      setupWarnings
    });
    return {
      dateStr,
      dailyConfig,
      activeMccs,
      stateByMcc,
      readyMccIds,
      hasExistingData: existingDataSheet !== null && existingDataSheet.getLastRow() > 1,
      setupWarnings,
      signature
    };
  }
  function buildSetupWarnings(syncInfo) {
    const warnings = [];
    const duplicates = syncInfo.duplicateAccounts;
    if (duplicates.length > 0) {
      const listed = duplicates.slice(0, MAX_LISTED_DUPLICATE_ACCOUNTS).map((d) => `${formatCidHyphenated(d.cid)} (MCCs ${d.mccIds.map(formatCidHyphenated).join(", ")})`);
      const more = duplicates.length > MAX_LISTED_DUPLICATE_ACCOUNTS ? ` and ${duplicates.length - MAX_LISTED_DUPLICATE_ACCOUNTS} more` : "";
      warnings.push(
        `\u26A0\uFE0F ${duplicates.length} account(s) are under more than one MCC \u2014 install the Google Ads Script only in top-level MCCs: ${listed.join(", ")}${more}`
      );
    }
    if (syncInfo.outdatedScripts.length > 0) {
      const listed = syncInfo.outdatedScripts.map((o) => `${formatCidHyphenated(o.mccId)} (${o.version ? `v${o.version}` : "older version"})`).join(", ");
      warnings.push(
        `\u26A0\uFE0F Google Ads Script out of date in MCC(s) ${listed} \u2014 re-copy it with \u{1F4CB} Copy Google Ads Script (this Sheet runs v${DLG_VERSION})`
      );
    }
    return warnings;
  }
  function shouldSkipPublish(plan, runStateMap, force) {
    if (force) {
      return false;
    }
    if (plan.readyMccIds.size === 0 && !plan.hasExistingData) {
      return true;
    }
    const previousSignature = runStateMap.get(`PUBLISH_SIGNATURE_${plan.dateStr}`);
    const publishStatus = runStateMap.get(`PUBLISH_STATUS_${plan.dateStr}`);
    const pendingCount = plan.activeMccs.length - plan.readyMccIds.size;
    return previousSignature === plan.signature || publishStatus === "PUBLISHED" && pendingCount === 0 && !previousSignature;
  }
  function assembleRows(spreadsheet, plan, existingDataRows) {
    const out = { rows: [], staleWarnings: [], failedAccountWarnings: [] };
    for (const mcc of plan.activeMccs) {
      if (plan.readyMccIds.has(mcc.mccId)) {
        out.rows.push(...buildFreshMccRows(spreadsheet, plan, mcc));
        out.rows.push(...buildFailedAccountRows(plan, mcc, existingDataRows, out.failedAccountWarnings));
      } else {
        out.staleWarnings.push(`${formatCidHyphenated(mcc.mccId)} (${mcc.mccName})`);
        out.rows.push(...buildStaleMccRows(spreadsheet, plan, mcc, existingDataRows));
      }
    }
    return out;
  }
  function keepRow(plan, row, selectedCids) {
    if (!selectedCids.has(normalizeCid(row.accountId)))
      return false;
    if (plan.dailyConfig.excludeShoppingAndPmax && row.campaignType === "SHOPPING")
      return false;
    return true;
  }
  function buildFreshMccRows(spreadsheet, plan, mcc) {
    const selectedCids = new Set(mcc.selectedCids);
    const failedCids = new Set(Object.keys(plan.stateByMcc.get(mcc.mccId)?.failedCids ?? {}));
    const stagedBatch = readStagingBatchForMcc(spreadsheet, mcc.mccId);
    if (stagedBatch.recommendations.length === 0) {
      return [];
    }
    const joined = joinStagingBatchToReportRows(stagedBatch).filter(
      (r) => keepRow(plan, r, selectedCids) && !failedCids.has(normalizeCid(r.accountId))
    );
    return assignRecommendationNumbers(consolidateRowsByScope(joined)).map((r) => ({
      ...r,
      timestamp: r.timestamp || `${plan.dateStr} 00:00:00`,
      mccSyncStatus: "OK",
      mccLastUpdatedDate: plan.dateStr
    }));
  }
  function buildFailedAccountRows(plan, mcc, existingDataRows, warnings) {
    const selectedCids = new Set(mcc.selectedCids);
    const failedInfo = plan.stateByMcc.get(mcc.mccId)?.failedCids ?? {};
    const rows = [];
    for (const cid of Object.keys(failedInfo).filter((c) => selectedCids.has(c))) {
      const hyphenatedCid = formatCidHyphenated(cid);
      warnings.push(hyphenatedCid);
      const shortError = truncate(failedInfo[cid].error, MAX_STATUS_ERROR_LENGTH);
      for (const r of existingDataRows) {
        if (normalizeCid(r.mccId) !== mcc.mccId || normalizeCid(r.accountId) !== cid)
          continue;
        if (!keepRow(plan, r, selectedCids))
          continue;
        rows.push({
          ...r,
          mccSyncStatus: `WARNING: Account ${hyphenatedCid} failed (${shortError}); showing snapshot from ${r.mccLastUpdatedDate ?? "a previous run"}`
        });
      }
    }
    return rows;
  }
  function buildStaleMccRows(spreadsheet, plan, mcc, existingDataRows) {
    const state = plan.stateByMcc.get(mcc.mccId);
    const selectedCids = new Set(mcc.selectedCids);
    let rawRows = existingDataRows.filter((r) => normalizeCid(r.mccId) === mcc.mccId);
    const hasProgressToday = state.completedCids.length > 0 || Object.keys(state.failedCids ?? {}).length > 0;
    if (rawRows.length === 0 && !hasProgressToday) {
      const stagedBatch = readStagingBatchForMcc(spreadsheet, mcc.mccId);
      if (stagedBatch.recommendations.length > 0) {
        rawRows = assignRecommendationNumbers(
          consolidateRowsByScope(joinStagingBatchToReportRows(stagedBatch))
        );
      }
    }
    const lastDate = state.lastCompletedDate ?? "previous run";
    return rawRows.filter((r) => keepRow(plan, r, selectedCids)).map((r) => ({
      ...r,
      timestamp: r.timestamp || `${state.lastCompletedDate ?? plan.dateStr} 00:00:00`,
      mccSyncStatus: `WARNING: Stale snapshot from ${lastDate}`,
      // Date column: never write the 'previous run' label; leave it blank when the date is unknown
      mccLastUpdatedDate: state.lastCompletedDate ?? null
    }));
  }
  function convertCurrencies(spreadsheet, plan, rows) {
    if (rows.length === 0) {
      return rows;
    }
    const fxRates = resolveDailyFxRates({
      spreadsheet,
      dateStr: plan.dateStr,
      sourceCurrencies: Array.from(new Set(rows.map((r) => r.currencyCode))),
      targetCurrency: plan.dailyConfig.targetCurrency
    });
    return rows.map((r) => applyCurrencyConversionToRow(r, plan.dailyConfig.targetCurrency, fxRates));
  }
  function recordPublish(spreadsheet, plan, publishedCount, assembled) {
    const hasDataWarnings = assembled.staleWarnings.length > 0 || assembled.failedAccountWarnings.length > 0;
    setRunStateValues(
      spreadsheet,
      [
        { key: `PUBLISH_SIGNATURE_${plan.dateStr}`, value: plan.signature, notes: `Signature for ${plan.dateStr}` },
        {
          key: `PUBLISH_STATUS_${plan.dateStr}`,
          value: hasDataWarnings ? "PUBLISHED_WITH_STALE_WARNING" : "PUBLISHED",
          notes: `Published ${publishedCount} rows on ${plan.dateStr}`
        }
      ],
      plan.dateStr
    );
  }
  function updateStatusBanner(configSheet, plan, publishedCount, assembled) {
    const parts = [];
    if (assembled.staleWarnings.length > 0) {
      parts.push(
        `\u26A0\uFE0F WARNING \u2014 Published (${plan.dateStr}) with stale snapshot for pending MCC(s): ${assembled.staleWarnings.join(", ")}`
      );
    }
    if (assembled.failedAccountWarnings.length > 0) {
      parts.push(
        `\u26A0\uFE0F ${assembled.failedAccountWarnings.length} account(s) failed in today's run (showing their previous snapshot): ${assembled.failedAccountWarnings.join(", ")}`
      );
    }
    if (parts.length === 0) {
      parts.push(
        `\u{1F7E2} PUBLISHED (${plan.dateStr}) \u2014 ${publishedCount} recommendations across ${plan.activeMccs.length} active MCC(s).`
      );
    }
    parts.push(...plan.setupWarnings);
    setStatusBannerIfChanged(configSheet, parts.join(" | "));
  }
  function truncate(text, maxLength) {
    return text.length > maxLength ? `${text.slice(0, maxLength - 1)}\u2026` : text;
  }
  function setStatusBannerIfChanged(configSheet, nextMessage) {
    const cell = configSheet.getRange(STATUS_BANNER_CELL);
    const currentMessage = String(cell.getValue() ?? "");
    if (currentMessage !== nextMessage) {
      cell.setValue(nextMessage);
    }
  }

  // src/appscript/demo-seed.ts
  var SYNTHETIC_MCC_ID = "999-888-7777";
  var SYNTHETIC_NORM_MCC_ID = "9998887777";
  var SYNTHETIC_MCC_NAME = "[Demo] Global Retail MCC";
  var SYNTHETIC_SECOND_MCC_ID = "888-777-6666";
  var SYNTHETIC_SECOND_NORM_MCC_ID = "8887776666";
  var SYNTHETIC_SECOND_MCC_NAME = "[Demo] DACH & Nordics Aviation MCC";
  var SYNTHETIC_CHILD_ACCOUNTS = [
    {
      customerId: "101-202-3030",
      name: "[Demo] US Search & AI Max",
      currencyCode: "USD"
    },
    {
      customerId: "202-303-4040",
      name: "[Demo] EMEA Retail & Shared Budgets",
      currencyCode: "EUR"
    },
    {
      customerId: "303-404-5050",
      name: "[Demo] UK Performance & Portfolios",
      currencyCode: "GBP"
    }
  ];
  var SYNTHETIC_SECOND_MCC_CHILD_ACCOUNTS = [
    {
      customerId: "404-505-6060",
      name: "[Demo] DE Airlines Search & PMax",
      currencyCode: "EUR"
    },
    {
      customerId: "505-606-7070",
      name: "[Demo] CH Alpine Routes & Portfolios",
      currencyCode: "CHF"
    }
  ];
  function buildSyntheticDemoBatch(dateStr) {
    const timestamp = `${dateStr} 00:00:00`;
    const accounts = SYNTHETIC_CHILD_ACCOUNTS.map((acc) => ({
      mccId: SYNTHETIC_MCC_ID,
      mccName: SYNTHETIC_MCC_NAME,
      accountId: acc.customerId,
      accountName: acc.name,
      currencyCode: acc.currencyCode,
      timestamp
    }));
    const budgets = [
      // Account 1 (USD): 101-202-3030
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        budgetResourceName: "customers/1012023030/campaignBudgets/1001",
        campaignBudgetName: "US - Brand Search Daily Budget",
        campaignBudgetAmount: 500,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        budgetResourceName: "customers/1012023030/campaignBudgets/1002",
        campaignBudgetName: "US - Q4 Peak Demand PMax Budget",
        campaignBudgetAmount: 1200,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        budgetResourceName: "customers/1012023030/campaignBudgets/1003",
        campaignBudgetName: "US - Non-Brand High Intent Budget",
        campaignBudgetAmount: 400,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        budgetResourceName: "customers/1012023030/campaignBudgets/1099",
        campaignBudgetName: "US - Legacy Display Shared Budget (Donor)",
        campaignBudgetAmount: 650,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: true
      },
      // Account 2 (EUR): 202-303-4040
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        budgetResourceName: "customers/2023034040/campaignBudgets/2001",
        campaignBudgetName: "EMEA - Core Generic Shared Budget",
        campaignBudgetAmount: 900,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: true
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        budgetResourceName: "customers/2023034040/campaignBudgets/2002",
        campaignBudgetName: "EMEA - Standard Shopping Budget",
        campaignBudgetAmount: 750,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      // Account 3 (GBP): 303-404-5050
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        budgetResourceName: "customers/3034045050/campaignBudgets/3001",
        campaignBudgetName: "UK - London Search Budget",
        campaignBudgetAmount: 350,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        budgetResourceName: "customers/3034045050/campaignBudgets/3002",
        campaignBudgetName: "UK - Midlands Search Budget",
        campaignBudgetAmount: 280,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        budgetResourceName: "customers/3034045050/campaignBudgets/3003",
        campaignBudgetName: "UK - Demand Gen Video Budget",
        campaignBudgetAmount: 420,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      }
    ];
    const campaigns = [
      // 1. US Brand Search with AI Max enabled (CAMPAIGN_BUDGET)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        campaignResourceName: "customers/1012023030/campaigns/11001",
        campaignId: "11001",
        campaignName: "US - Brand Search - AI Max",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: true,
        campaignAiMaxTextCustomizationEnabled: true,
        campaignAiMaxFinalUrlExpansionEnabled: true,
        budgetResourceName: "customers/1012023030/campaignBudgets/1001",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "MAXIMIZE_CONVERSIONS",
        campaignBiddingTargetRoas: null,
        campaignBiddingTargetCpa: 18.5,
        campaignStatsYesterdayCost: 495,
        campaignStats7DaysCost: 3420,
        campaign30DaysCost: 14600,
        campaign30DaysConversionsValue: 58400,
        campaign30DaysConversions: 810,
        campaign30DaysSearchRankLostImpressionShare: 0.08,
        campaign30DaysSearchBudgetLostImpressionShare: 0.34
      },
      // 2. US Non-Feed PMax (FORECASTING_CAMPAIGN_BUDGET)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        campaignResourceName: "customers/1012023030/campaigns/11002",
        campaignId: "11002",
        campaignName: "US - Q4 Peak Demand PMax (Lead Gen)",
        campaignType: "PERFORMANCE_MAX",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/1012023030/campaignBudgets/1002",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "MAXIMIZE_CONVERSION_VALUE",
        campaignBiddingTargetRoas: 3.6,
        campaignBiddingTargetCpa: null,
        campaignStatsYesterdayCost: 1190,
        campaignStats7DaysCost: 8250,
        campaign30DaysCost: 35100,
        campaign30DaysConversionsValue: 133380,
        campaign30DaysConversions: 1420,
        campaign30DaysSearchRankLostImpressionShare: null,
        campaign30DaysSearchBudgetLostImpressionShare: null
      },
      // 3. US Non-Brand Search (MOVE_UNUSED_BUDGET)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        campaignResourceName: "customers/1012023030/campaigns/11003",
        campaignId: "11003",
        campaignName: "US - Non-Brand High Intent Search",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: true,
        campaignAiMaxTextCustomizationEnabled: true,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/1012023030/campaignBudgets/1003",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "TARGET_CPA",
        campaignBiddingTargetRoas: null,
        campaignBiddingTargetCpa: 26,
        campaignStatsYesterdayCost: 398,
        campaignStats7DaysCost: 2760,
        campaign30DaysCost: 11850,
        campaign30DaysConversionsValue: 39105,
        campaign30DaysConversions: 465,
        campaign30DaysSearchRankLostImpressionShare: 0.14,
        campaign30DaysSearchBudgetLostImpressionShare: 0.41
      },
      // 4a & 4b. EMEA Shared Budget across 2 campaigns (MARGINAL_ROI_CAMPAIGN_BUDGET -> SHARED_BUDGET scope)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        campaignResourceName: "customers/2023034040/campaigns/22001",
        campaignId: "22001",
        campaignName: "DE - Generic Search - Core",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/2023034040/campaignBudgets/2001",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "MAXIMIZE_CONVERSIONS",
        campaignBiddingTargetRoas: null,
        campaignBiddingTargetCpa: 22,
        campaignStatsYesterdayCost: 540,
        campaignStats7DaysCost: 3750,
        campaign30DaysCost: 16200,
        campaign30DaysConversionsValue: 64800,
        campaign30DaysConversions: 750,
        campaign30DaysSearchRankLostImpressionShare: 0.11,
        campaign30DaysSearchBudgetLostImpressionShare: 0.29
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        campaignResourceName: "customers/2023034040/campaigns/22002",
        campaignId: "22002",
        campaignName: "FR - Generic Search - Core",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/2023034040/campaignBudgets/2001",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "MAXIMIZE_CONVERSIONS",
        campaignBiddingTargetRoas: null,
        campaignBiddingTargetCpa: 24,
        campaignStatsYesterdayCost: 355,
        campaignStats7DaysCost: 2480,
        campaign30DaysCost: 10600,
        campaign30DaysConversionsValue: 41340,
        campaign30DaysConversions: 445,
        campaign30DaysSearchRankLostImpressionShare: 0.13,
        campaign30DaysSearchBudgetLostImpressionShare: 0.27
      },
      // 5. EMEA Standard Shopping (LOWER_TARGET_ROAS -> tests Exclude Shopping toggle)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        campaignResourceName: "customers/2023034040/campaigns/22003",
        campaignId: "22003",
        campaignName: "EMEA - Merchant Center Standard Shopping",
        campaignType: "SHOPPING",
        campaignSubType: "SHOPPING_SMART_ADS",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/2023034040/campaignBudgets/2002",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "TARGET_ROAS",
        campaignBiddingTargetRoas: 4.5,
        campaignBiddingTargetCpa: null,
        campaignStatsYesterdayCost: 690,
        campaignStats7DaysCost: 4820,
        campaign30DaysCost: 20500,
        campaign30DaysConversionsValue: 96350,
        campaign30DaysConversions: 920,
        campaign30DaysSearchRankLostImpressionShare: 0.22,
        campaign30DaysSearchBudgetLostImpressionShare: 0.09
      },
      // 6a & 6b. UK Locally-Owned Portfolio across 2 campaigns (RAISE_TARGET_CPA -> PORTFOLIO scope)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        campaignResourceName: "customers/3034045050/campaigns/33001",
        campaignId: "33001",
        campaignName: "UK - Search - London & South East",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: true,
        campaignAiMaxTextCustomizationEnabled: true,
        campaignAiMaxFinalUrlExpansionEnabled: true,
        budgetResourceName: "customers/3034045050/campaignBudgets/3001",
        campaignIsPortfolioBiddingStrategy: true,
        campaignPortfolioBiddingStrategyName: "UK - Target CPA Growth Portfolio",
        portfolioStrategyResourceName: "customers/3034045050/biddingStrategies/9001",
        portfolioOwnerId: "3034045050",
        isLocallyOwnedPortfolio: true,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "TARGET_CPA",
        campaignBiddingTargetRoas: null,
        campaignBiddingTargetCpa: 24,
        campaignStatsYesterdayCost: 320,
        campaignStats7DaysCost: 2210,
        campaign30DaysCost: 9450,
        campaign30DaysConversionsValue: 37800,
        campaign30DaysConversions: 410,
        campaign30DaysSearchRankLostImpressionShare: 0.19,
        campaign30DaysSearchBudgetLostImpressionShare: 0.05
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        campaignResourceName: "customers/3034045050/campaigns/33002",
        campaignId: "33002",
        campaignName: "UK - Search - Midlands & North",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/3034045050/campaignBudgets/3002",
        campaignIsPortfolioBiddingStrategy: true,
        campaignPortfolioBiddingStrategyName: "UK - Target CPA Growth Portfolio",
        portfolioStrategyResourceName: "customers/3034045050/biddingStrategies/9001",
        portfolioOwnerId: "3034045050",
        isLocallyOwnedPortfolio: true,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "TARGET_CPA",
        campaignBiddingTargetRoas: null,
        campaignBiddingTargetCpa: 24,
        campaignStatsYesterdayCost: 250,
        campaignStats7DaysCost: 1740,
        campaign30DaysCost: 7400,
        campaign30DaysConversionsValue: 29600,
        campaign30DaysConversions: 320,
        campaign30DaysSearchRankLostImpressionShare: 0.21,
        campaign30DaysSearchBudgetLostImpressionShare: 0.04
      },
      // 7. UK Manager-Owned MCC Portfolio (LOWER_TARGET_ROAS -> PORTFOLIO scope)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        campaignResourceName: "customers/3034045050/campaigns/33003",
        campaignId: "33003",
        campaignName: "UK - Demand Gen Video & Discovery",
        campaignType: "DEMAND_GEN",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/3034045050/campaignBudgets/3003",
        campaignIsPortfolioBiddingStrategy: true,
        campaignPortfolioBiddingStrategyName: "MCC Global ROAS Portfolio",
        portfolioStrategyResourceName: "customers/9998887777/biddingStrategies/9900",
        portfolioOwnerId: "9998887777",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: true,
        campaignBiddingStrategyType: "TARGET_ROAS",
        campaignBiddingTargetRoas: 5.2,
        campaignBiddingTargetCpa: null,
        campaignStatsYesterdayCost: 390,
        campaignStats7DaysCost: 2710,
        campaign30DaysCost: 11600,
        campaign30DaysConversionsValue: 61480,
        campaign30DaysConversions: 580,
        campaign30DaysSearchRankLostImpressionShare: null,
        campaign30DaysSearchBudgetLostImpressionShare: null
      }
    ];
    const recommendations = [
      // Rec 1: CAMPAIGN_BUDGET (US Brand Search)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        recommendationId: "500001",
        recommendationType: "CAMPAIGN_BUDGET",
        campaignResourceName: "customers/1012023030/campaigns/11001",
        recommendationCurrentBudgetAmount: 500,
        recommendationNewBudgetAmount: 850,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 3500,
        recommendationPotentialCost: 5950,
        recommendationBaseClicks: 2800,
        recommendationPotentialClicks: 4650,
        recommendationBaseConversions: 195,
        recommendationPotentialConversions: 315,
        recommendationBaseConversionsValue: 14040,
        recommendationPotentialConversionsValue: 22680,
        recommendationBaseImpressions: 42e3,
        recommendationPotentialImpressions: 68500,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 2: FORECASTING_CAMPAIGN_BUDGET (US Q4 Peak Demand PMax)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        recommendationId: "500002",
        recommendationType: "FORECASTING_CAMPAIGN_BUDGET",
        campaignResourceName: "customers/1012023030/campaigns/11002",
        recommendationCurrentBudgetAmount: 1200,
        recommendationNewBudgetAmount: 1800,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 8400,
        recommendationPotentialCost: 12600,
        recommendationBaseClicks: 6100,
        recommendationPotentialClicks: 8950,
        recommendationBaseConversions: 340,
        recommendationPotentialConversions: 495,
        recommendationBaseConversionsValue: 31960,
        recommendationPotentialConversionsValue: 45540,
        recommendationBaseImpressions: 185e3,
        recommendationPotentialImpressions: 27e4,
        recommendationBaseVideoViews: 12400,
        recommendationPotentialVideoViews: 18900,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 3: MOVE_UNUSED_BUDGET (US Non-Brand Search)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        recommendationId: "500003",
        recommendationType: "MOVE_UNUSED_BUDGET",
        campaignResourceName: "customers/1012023030/campaigns/11003",
        recommendationCurrentBudgetAmount: 400,
        recommendationNewBudgetAmount: 650,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 2800,
        recommendationPotentialCost: 4550,
        recommendationBaseClicks: 1900,
        recommendationPotentialClicks: 3050,
        recommendationBaseConversions: 108,
        recommendationPotentialConversions: 172,
        recommendationBaseConversionsValue: 9180,
        recommendationPotentialConversionsValue: 14620,
        recommendationBaseImpressions: 31e3,
        recommendationPotentialImpressions: 49e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "customers/1012023030/campaignBudgets/1099",
        moveBudgetSourceCampaigns: "US - Upper Funnel Display | US - Remarketing Display",
        moveBudgetSourceTargetCpa: 32,
        moveBudgetSourceTargetRoas: 2.4
      },
      // Rec 4a & 4b: MARGINAL_ROI_CAMPAIGN_BUDGET across 2 campaigns sharing EMEA Core Generic Shared Budget
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        recommendationId: "500004",
        recommendationType: "MARGINAL_ROI_CAMPAIGN_BUDGET",
        campaignResourceName: "customers/2023034040/campaigns/22001",
        recommendationCurrentBudgetAmount: 900,
        recommendationNewBudgetAmount: 1350,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 6300,
        recommendationPotentialCost: 9450,
        recommendationBaseClicks: 4900,
        recommendationPotentialClicks: 7200,
        recommendationBaseConversions: 285,
        recommendationPotentialConversions: 415,
        recommendationBaseConversionsValue: 25650,
        recommendationPotentialConversionsValue: 37350,
        recommendationBaseImpressions: 78e3,
        recommendationPotentialImpressions: 114e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        recommendationId: "500004",
        recommendationType: "MARGINAL_ROI_CAMPAIGN_BUDGET",
        campaignResourceName: "customers/2023034040/campaigns/22002",
        recommendationCurrentBudgetAmount: 900,
        recommendationNewBudgetAmount: 1350,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 6300,
        recommendationPotentialCost: 9450,
        recommendationBaseClicks: 4900,
        recommendationPotentialClicks: 7200,
        recommendationBaseConversions: 285,
        recommendationPotentialConversions: 415,
        recommendationBaseConversionsValue: 25650,
        recommendationPotentialConversionsValue: 37350,
        recommendationBaseImpressions: 78e3,
        recommendationPotentialImpressions: 114e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 5: LOWER_TARGET_ROAS on EMEA Standard Shopping
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "202-303-4040",
        recommendationId: "500005",
        recommendationType: "LOWER_TARGET_ROAS",
        campaignResourceName: "customers/2023034040/campaigns/22003",
        recommendationCurrentBudgetAmount: 750,
        recommendationNewBudgetAmount: 750,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: 3.8,
        recommendationBaseCost: 4800,
        recommendationPotentialCost: 6650,
        recommendationBaseClicks: 5400,
        recommendationPotentialClicks: 7600,
        recommendationBaseConversions: 215,
        recommendationPotentialConversions: 295,
        recommendationBaseConversionsValue: 21600,
        recommendationPotentialConversionsValue: 28800,
        recommendationBaseImpressions: 14e4,
        recommendationPotentialImpressions: 195e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 6a & 6b: RAISE_TARGET_CPA on UK Locally-Owned Portfolio across 2 campaigns
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        recommendationId: "500006",
        recommendationType: "RAISE_TARGET_CPA",
        campaignResourceName: "customers/3034045050/campaigns/33001",
        recommendationCurrentBudgetAmount: 350,
        recommendationNewBudgetAmount: 350,
        recommendationNewTargetCpa: 29.5,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 3950,
        recommendationPotentialCost: 5400,
        recommendationBaseClicks: 3100,
        recommendationPotentialClicks: 4300,
        recommendationBaseConversions: 170,
        recommendationPotentialConversions: 225,
        recommendationBaseConversionsValue: 15640,
        recommendationPotentialConversionsValue: 20700,
        recommendationBaseImpressions: 52e3,
        recommendationPotentialImpressions: 71e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        recommendationId: "500006",
        recommendationType: "RAISE_TARGET_CPA",
        campaignResourceName: "customers/3034045050/campaigns/33002",
        recommendationCurrentBudgetAmount: 280,
        recommendationNewBudgetAmount: 280,
        recommendationNewTargetCpa: 29.5,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 3950,
        recommendationPotentialCost: 5400,
        recommendationBaseClicks: 3100,
        recommendationPotentialClicks: 4300,
        recommendationBaseConversions: 170,
        recommendationPotentialConversions: 225,
        recommendationBaseConversionsValue: 15640,
        recommendationPotentialConversionsValue: 20700,
        recommendationBaseImpressions: 52e3,
        recommendationPotentialImpressions: 71e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 7: LOWER_TARGET_ROAS on UK Manager-Owned MCC Portfolio (Demand Gen Video)
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        recommendationId: "500007",
        recommendationType: "LOWER_TARGET_ROAS",
        campaignResourceName: "customers/3034045050/campaigns/33003",
        recommendationCurrentBudgetAmount: 420,
        recommendationNewBudgetAmount: 420,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: 4.4,
        recommendationBaseCost: 2700,
        recommendationPotentialCost: 3850,
        recommendationBaseClicks: 4200,
        recommendationPotentialClicks: 5900,
        recommendationBaseConversions: 135,
        recommendationPotentialConversions: 188,
        recommendationBaseConversionsValue: 14310,
        recommendationPotentialConversionsValue: 18920,
        recommendationBaseImpressions: 21e4,
        recommendationPotentialImpressions: 295e3,
        recommendationBaseVideoViews: 45e3,
        recommendationPotentialVideoViews: 62e3,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 8 & 9: CAMPAIGN_BUDGET on each campaign of the UK Locally-Owned Portfolio (Rec 6). A portfolio
      // owns bidding, not budgets, so each budget rec stays a CAMPAIGN-scope row next to the PORTFOLIO row.
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        recommendationId: "500008",
        recommendationType: "CAMPAIGN_BUDGET",
        campaignResourceName: "customers/3034045050/campaigns/33001",
        recommendationCurrentBudgetAmount: 350,
        recommendationNewBudgetAmount: 500,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 2200,
        recommendationPotentialCost: 3150,
        recommendationBaseClicks: 1750,
        recommendationPotentialClicks: 2450,
        recommendationBaseConversions: 96,
        recommendationPotentialConversions: 130,
        recommendationBaseConversionsValue: 8960,
        recommendationPotentialConversionsValue: 12100,
        recommendationBaseImpressions: 29e3,
        recommendationPotentialImpressions: 4e4,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "303-404-5050",
        recommendationId: "500009",
        recommendationType: "CAMPAIGN_BUDGET",
        campaignResourceName: "customers/3034045050/campaigns/33002",
        recommendationCurrentBudgetAmount: 280,
        recommendationNewBudgetAmount: 400,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 1730,
        recommendationPotentialCost: 2470,
        recommendationBaseClicks: 1350,
        recommendationPotentialClicks: 1900,
        recommendationBaseConversions: 75,
        recommendationPotentialConversions: 101,
        recommendationBaseConversionsValue: 6900,
        recommendationPotentialConversionsValue: 9400,
        recommendationBaseImpressions: 23e3,
        recommendationPotentialImpressions: 32e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      // Rec 10: RAISE_TARGET_CPA on US Brand Search, which also has a CAMPAIGN_BUDGET rec (Rec 1), so one
      // standalone campaign shows up on both the Budgets and the Bids dashboard pages.
      {
        mccId: SYNTHETIC_MCC_ID,
        accountId: "101-202-3030",
        recommendationId: "500010",
        recommendationType: "RAISE_TARGET_CPA",
        campaignResourceName: "customers/1012023030/campaigns/11001",
        recommendationCurrentBudgetAmount: 500,
        recommendationNewBudgetAmount: 500,
        recommendationNewTargetCpa: 21.5,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 3420,
        recommendationPotentialCost: 4100,
        recommendationBaseClicks: 2700,
        recommendationPotentialClicks: 3150,
        recommendationBaseConversions: 189,
        recommendationPotentialConversions: 215,
        recommendationBaseConversionsValue: 13600,
        recommendationPotentialConversionsValue: 15480,
        recommendationBaseImpressions: 41e3,
        recommendationPotentialImpressions: 47500,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      }
    ];
    return {
      accounts,
      budgets,
      campaigns,
      recommendations
    };
  }
  function buildSyntheticSecondMccDemoBatch(dateStr) {
    const timestamp = `${dateStr} 00:00:00`;
    const accounts = SYNTHETIC_SECOND_MCC_CHILD_ACCOUNTS.map((acc) => ({
      mccId: SYNTHETIC_SECOND_MCC_ID,
      mccName: SYNTHETIC_SECOND_MCC_NAME,
      accountId: acc.customerId,
      accountName: acc.name,
      currencyCode: acc.currencyCode,
      timestamp
    }));
    const budgets = [
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "404-505-6060",
        budgetResourceName: "customers/4045056060/campaignBudgets/4001",
        campaignBudgetName: "DE - Direct Flights Brand Budget",
        campaignBudgetAmount: 1100,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "404-505-6060",
        budgetResourceName: "customers/4045056060/campaignBudgets/4002",
        campaignBudgetName: "DE - Summer Holiday Routes PMax Budget",
        campaignBudgetAmount: 1600,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "505-606-7070",
        budgetResourceName: "customers/5056067070/campaignBudgets/5001",
        campaignBudgetName: "CH - Zurich & Geneva Hub Search Budget",
        campaignBudgetAmount: 820,
        campaignBudgetTotalAmount: null,
        campaignBudgetDeliveryMethod: "STANDARD",
        campaignBudgetType: "DAILY",
        campaignBudgetIsShared: false
      }
    ];
    const campaigns = [
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "404-505-6060",
        campaignResourceName: "customers/4045056060/campaigns/44001",
        campaignId: "44001",
        campaignName: "DE - Direct Flights Brand Search - AI Max",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: true,
        campaignAiMaxTextCustomizationEnabled: true,
        campaignAiMaxFinalUrlExpansionEnabled: true,
        budgetResourceName: "customers/4045056060/campaignBudgets/4001",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "MAXIMIZE_CONVERSION_VALUE",
        campaignBiddingTargetRoas: 5.8,
        campaignBiddingTargetCpa: null,
        campaignStatsYesterdayCost: 1085,
        campaignStats7DaysCost: 7540,
        campaign30DaysCost: 32100,
        campaign30DaysConversionsValue: 192600,
        campaign30DaysConversions: 1280,
        campaign30DaysSearchRankLostImpressionShare: 0.06,
        campaign30DaysSearchBudgetLostImpressionShare: 0.31
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "404-505-6060",
        campaignResourceName: "customers/4045056060/campaigns/44002",
        campaignId: "44002",
        campaignName: "DE - Summer Holiday Routes PMax (Travel)",
        campaignType: "PERFORMANCE_MAX",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: false,
        campaignAiMaxTextCustomizationEnabled: false,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/4045056060/campaignBudgets/4002",
        campaignIsPortfolioBiddingStrategy: false,
        campaignPortfolioBiddingStrategyName: "N/A",
        portfolioStrategyResourceName: "N/A",
        portfolioOwnerId: "N/A",
        isLocallyOwnedPortfolio: false,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "MAXIMIZE_CONVERSION_VALUE",
        campaignBiddingTargetRoas: 4.2,
        campaignBiddingTargetCpa: null,
        campaignStatsYesterdayCost: 1580,
        campaignStats7DaysCost: 10950,
        campaign30DaysCost: 46800,
        campaign30DaysConversionsValue: 205920,
        campaign30DaysConversions: 1640,
        campaign30DaysSearchRankLostImpressionShare: null,
        campaign30DaysSearchBudgetLostImpressionShare: null
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "505-606-7070",
        campaignResourceName: "customers/5056067070/campaigns/55001",
        campaignId: "55001",
        campaignName: "CH - Zurich & Geneva Hub Search",
        campaignType: "SEARCH",
        campaignSubType: "UNSPECIFIED",
        campaignIsAiMax: true,
        campaignAiMaxTextCustomizationEnabled: true,
        campaignAiMaxFinalUrlExpansionEnabled: false,
        budgetResourceName: "customers/5056067070/campaignBudgets/5001",
        campaignIsPortfolioBiddingStrategy: true,
        campaignPortfolioBiddingStrategyName: "CH Alpine Routes ROAS Portfolio",
        portfolioStrategyResourceName: "customers/5056067070/biddingStrategies/9501",
        portfolioOwnerId: "5056067070",
        isLocallyOwnedPortfolio: true,
        campaignPortfolioIsManagerOwned: false,
        campaignBiddingStrategyType: "TARGET_ROAS",
        campaignBiddingTargetRoas: 4.8,
        campaignBiddingTargetCpa: null,
        campaignStatsYesterdayCost: 790,
        campaignStats7DaysCost: 5460,
        campaign30DaysCost: 23400,
        campaign30DaysConversionsValue: 117e3,
        campaign30DaysConversions: 690,
        campaign30DaysSearchRankLostImpressionShare: 0.17,
        campaign30DaysSearchBudgetLostImpressionShare: 0.07
      }
    ];
    const recommendations = [
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "404-505-6060",
        recommendationId: "500008",
        recommendationType: "CAMPAIGN_BUDGET",
        campaignResourceName: "customers/4045056060/campaigns/44001",
        recommendationCurrentBudgetAmount: 1100,
        recommendationNewBudgetAmount: 1550,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 7700,
        recommendationPotentialCost: 10850,
        recommendationBaseClicks: 6400,
        recommendationPotentialClicks: 8900,
        recommendationBaseConversions: 310,
        recommendationPotentialConversions: 425,
        recommendationBaseConversionsValue: 46500,
        recommendationPotentialConversionsValue: 63750,
        recommendationBaseImpressions: 92e3,
        recommendationPotentialImpressions: 128e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "404-505-6060",
        recommendationId: "500009",
        recommendationType: "FORECASTING_CAMPAIGN_BUDGET",
        campaignResourceName: "customers/4045056060/campaigns/44002",
        recommendationCurrentBudgetAmount: 1600,
        recommendationNewBudgetAmount: 2250,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: null,
        recommendationBaseCost: 11200,
        recommendationPotentialCost: 15750,
        recommendationBaseClicks: 8200,
        recommendationPotentialClicks: 11400,
        recommendationBaseConversions: 390,
        recommendationPotentialConversions: 535,
        recommendationBaseConversionsValue: 49140,
        recommendationPotentialConversionsValue: 66875,
        recommendationBaseImpressions: 215e3,
        recommendationPotentialImpressions: 298e3,
        recommendationBaseVideoViews: 14200,
        recommendationPotentialVideoViews: 19800,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        accountId: "505-606-7070",
        recommendationId: "500010",
        recommendationType: "LOWER_TARGET_ROAS",
        campaignResourceName: "customers/5056067070/campaigns/55001",
        recommendationCurrentBudgetAmount: 820,
        recommendationNewBudgetAmount: 820,
        recommendationNewTargetCpa: null,
        recommendationNewTargetRoas: 4.1,
        recommendationBaseCost: 5500,
        recommendationPotentialCost: 7400,
        recommendationBaseClicks: 4100,
        recommendationPotentialClicks: 5550,
        recommendationBaseConversions: 165,
        recommendationPotentialConversions: 220,
        recommendationBaseConversionsValue: 27225,
        recommendationPotentialConversionsValue: 35200,
        recommendationBaseImpressions: 64e3,
        recommendationPotentialImpressions: 86e3,
        recommendationBaseVideoViews: 0,
        recommendationPotentialVideoViews: 0,
        moveBudgetSourceBudgetResourceName: "",
        moveBudgetSourceCampaigns: "N/A",
        moveBudgetSourceTargetCpa: null,
        moveBudgetSourceTargetRoas: null
      }
    ];
    return {
      accounts,
      budgets,
      campaigns,
      recommendations
    };
  }
  function seedSyntheticDemoData(options = {}) {
    const ss = options.spreadsheet ?? SpreadsheetApp.getActiveSpreadsheet();
    const dateStr = resolveSpreadsheetDate(ss);
    const demoMccs = [
      {
        mccId: SYNTHETIC_MCC_ID,
        normMccId: SYNTHETIC_NORM_MCC_ID,
        mccName: SYNTHETIC_MCC_NAME,
        childAccounts: SYNTHETIC_CHILD_ACCOUNTS,
        batch: buildSyntheticDemoBatch(dateStr),
        recCountLabel: "10 recommendations"
      },
      {
        mccId: SYNTHETIC_SECOND_MCC_ID,
        normMccId: SYNTHETIC_SECOND_NORM_MCC_ID,
        mccName: SYNTHETIC_SECOND_MCC_NAME,
        childAccounts: SYNTHETIC_SECOND_MCC_CHILD_ACCOUNTS,
        batch: buildSyntheticSecondMccDemoBatch(dateStr),
        recCountLabel: "3 recommendations"
      }
    ];
    const demoNormMccSet = new Set(demoMccs.map((d) => d.normMccId));
    withSheetLock(ss, "SHEET_SEED_DEMO", () => {
      const configSheet = getOrCreateSheet(ss, CONFIG_SHEET_NAME, { index: 0 });
      assertConfigLayout(configSheet);
      for (const demo of demoMccs) {
        clearMccFromStaging(ss, demo.normMccId);
        appendStagingBatch(ss, demo.batch, demo.normMccId);
        syncDiscoveredAccountsToConfig(
          configSheet,
          { mccId: demo.mccId, mccName: demo.mccName },
          demo.childAccounts
        );
      }
      const mccCount = MCC_TABLE_END_ROW - MCC_TABLE_START_ROW + 1;
      const mccRows = configSheet.getRange(MCC_TABLE_START_ROW, 1, mccCount, sheetCol(MCC_COL.MCC_ID)).getValues();
      for (let i = 0; i < mccRows.length; i++) {
        if (demoNormMccSet.has(normalizeCid(mccRows[i][MCC_COL.MCC_ID]))) {
          configSheet.getRange(MCC_TABLE_START_ROW + i, sheetCol(MCC_COL.SELECT_ALL)).setValue(true);
        }
      }
      const lastRow = configSheet.getLastRow();
      if (lastRow >= CID_TABLE_START_ROW) {
        const cidCount = lastRow - CID_TABLE_START_ROW + 1;
        const cidRows = configSheet.getRange(CID_TABLE_START_ROW, 1, cidCount, sheetCol(CID_COL.PARENT_MCC_ID)).getValues();
        const nextCidColA = cidRows.map((r) => [
          demoNormMccSet.has(normalizeCid(r[CID_COL.PARENT_MCC_ID])) ? true : r[CID_COL.INCLUDE]
        ]);
        configSheet.getRange(CID_TABLE_START_ROW, sheetCol(CID_COL.INCLUDE), cidCount, 1).setValues(nextCidColA);
      }
      resetTodayPublisherState(ss, dateStr);
      for (const demo of demoMccs) {
        writeMccSheetState(ss, {
          mcc: { mccId: demo.mccId, mccName: demo.mccName },
          childAccounts: demo.childAccounts,
          statusMessage: `\u{1F9EA} Demo MCC ${demo.mccId} seeded with ${demo.childAccounts.length} accounts & ${demo.recCountLabel} (${dateStr}).`,
          checkpoint: {
            mccId: demo.normMccId,
            dateStr,
            status: "COMPLETED",
            completedCids: demo.childAccounts.map((a) => normalizeCid(a.customerId)),
            lastCompletedDate: dateStr
          },
          lastCompletedDate: dateStr
        });
      }
    });
    publishFromStagingToData({
      spreadsheet: ss,
      force: true
    });
  }

  // src/appscript/sheet-ui.ts
  var EMBEDDED_ADS_SCRIPT_TEMPLATE = true ? '/**\n * Copyright 2026 Google LLC\n *\n * Licensed under the Apache License, Version 2.0 (the "License");\n * you may not use this file except in compliance with the License.\n * You may obtain a copy of the License at\n *\n *      https://www.apache.org/licenses/LICENSE-2.0\n *\n * Unless required by applicable law or agreed to in writing, software\n * distributed under the License is distributed on an "AS IS" BASIS,\n * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\n * See the License for the specific language governing permissions and\n * limitations under the License.\n *\n * DLG Dashboard \u2014 Google Ads Script (Auto-generated by scripts/build.mjs)\n */\nvar SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";\nvar SPREADSHEET_TIMEZONE = "YOUR_SPREADSHEET_TIMEZONE_HERE";\n"use strict";\nvar DlgAdsBundle = (() => {\n  var __defProp = Object.defineProperty;\n  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;\n  var __getOwnPropNames = Object.getOwnPropertyNames;\n  var __hasOwnProp = Object.prototype.hasOwnProperty;\n  var __export = (target, all) => {\n    for (var name in all)\n      __defProp(target, name, { get: all[name], enumerable: true });\n  };\n  var __copyProps = (to, from, except, desc) => {\n    if (from && typeof from === "object" || typeof from === "function") {\n      for (let key of __getOwnPropNames(from))\n        if (!__hasOwnProp.call(to, key) && key !== except)\n          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });\n    }\n    return to;\n  };\n  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);\n\n  // src/ads/main.ts\n  var main_exports = {};\n  __export(main_exports, {\n    main: () => main\n  });\n\n  // src/shared/utils/cid.ts\n  var HYPHENATED_CID_RE = /^\\d{3}-\\d{3}-\\d{4}$/;\n  function normalizeCid(rawCid) {\n    if (rawCid === null || rawCid === void 0) {\n      return "";\n    }\n    const trimmed = String(rawCid).trim();\n    const digits = HYPHENATED_CID_RE.test(trimmed) ? trimmed.replace(/-/g, "") : trimmed.replace(/\\s+/g, "");\n    if (!/^\\d{1,10}$/.test(digits) || /^0+$/.test(digits)) {\n      return "";\n    }\n    return digits.padStart(10, "0");\n  }\n  function formatCidHyphenated(rawCid) {\n    const digits = normalizeCid(rawCid);\n    if (digits.length !== 10) {\n      return digits;\n    }\n    return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 10)}`;\n  }\n\n  // src/shared/contract.ts\n  var CONTRACT_VERSION = 1;\n  function isRecord(value) {\n    return typeof value === "object" && value !== null && !Array.isArray(value);\n  }\n  function isStringArray(value) {\n    return Array.isArray(value) && value.every((v) => typeof v === "string");\n  }\n  function parseVersionedRecord(raw, label, log) {\n    let value;\n    try {\n      value = JSON.parse(raw);\n    } catch {\n      log(`[DLG v2] ${label} is not valid JSON; ignoring it.`);\n      return null;\n    }\n    if (!isRecord(value)) {\n      log(`[DLG v2] ${label} is not a JSON object; ignoring it.`);\n      return null;\n    }\n    const version = value["schemaVersion"] ?? 1;\n    if (version !== CONTRACT_VERSION) {\n      log(\n        `[DLG v2] ${label} has schemaVersion ${String(version)}, but this script reads version ${CONTRACT_VERSION}. Re-copy the Google Ads Script from the Google Sheet menu (\\u{1F4CB} Copy Google Ads Script) so both scripts match.`\n      );\n      return null;\n    }\n    return value;\n  }\n  function serializeMccCheckpoint(checkpoint) {\n    return JSON.stringify({ schemaVersion: CONTRACT_VERSION, ...checkpoint });\n  }\n  function parseMccCheckpoint(raw, log = defaultLog) {\n    const v = parseVersionedRecord(raw, "MCC CHECKPOINT", log);\n    if (!v)\n      return null;\n    const status = v["status"];\n    if (typeof v["mccId"] !== "string" || typeof v["dateStr"] !== "string" || status !== "IN_PROGRESS" && status !== "COMPLETED" || !isStringArray(v["completedCids"])) {\n      log("[DLG v2] MCC CHECKPOINT has an unexpected shape; ignoring it.");\n      return null;\n    }\n    const failedCids = parseFailedCids(v["failedCids"]);\n    const lastCompletedDate = v["lastCompletedDate"];\n    return {\n      mccId: v["mccId"],\n      dateStr: v["dateStr"],\n      status,\n      completedCids: v["completedCids"].map(normalizeCid).filter(Boolean),\n      ...failedCids ? { failedCids } : {},\n      ...typeof lastCompletedDate === "string" && lastCompletedDate ? { lastCompletedDate } : {}\n    };\n  }\n  function parseFailedCids(value) {\n    if (!isRecord(value))\n      return void 0;\n    const out = {};\n    for (const [rawCid, info] of Object.entries(value)) {\n      const cid = normalizeCid(rawCid);\n      if (!cid || !isRecord(info))\n        continue;\n      const attempts = Number(info["attempts"]);\n      out[cid] = {\n        error: String(info["error"] ?? ""),\n        attempts: Number.isFinite(attempts) && attempts > 0 ? attempts : 1\n      };\n    }\n    return Object.keys(out).length > 0 ? out : void 0;\n  }\n  function parseLockedDailyConfig(raw, log = defaultLog) {\n    const v = parseVersionedRecord(raw, "DAILY_CONFIG_LOCK", log);\n    if (!v)\n      return null;\n    if (typeof v["targetCurrency"] !== "string" || typeof v["excludeShoppingAndPmax"] !== "boolean") {\n      log("[DLG v2] DAILY_CONFIG_LOCK has an unexpected shape; ignoring it.");\n      return null;\n    }\n    const config = {\n      targetCurrency: v["targetCurrency"],\n      excludeShoppingAndPmax: v["excludeShoppingAndPmax"]\n    };\n    const rawSelected = v["selectedCidsByMcc"];\n    if (rawSelected !== void 0) {\n      if (!isRecord(rawSelected)) {\n        log("[DLG v2] DAILY_CONFIG_LOCK selectedCidsByMcc is not an object; ignoring the lock.");\n        return null;\n      }\n      const selectedCidsByMcc = {};\n      for (const [rawMcc, cids] of Object.entries(rawSelected)) {\n        const mcc = normalizeCid(rawMcc);\n        if (!mcc || !isStringArray(cids))\n          continue;\n        selectedCidsByMcc[mcc] = cids.map(normalizeCid).filter(Boolean);\n      }\n      config.selectedCidsByMcc = selectedCidsByMcc;\n    }\n    return config;\n  }\n  function defaultLog(msg) {\n    if (typeof Logger !== "undefined") {\n      Logger.log(msg);\n    }\n  }\n\n  // src/shared/version.ts\n  var DLG_VERSION = true ? "2.0.0" : "dev";\n\n  // src/shared/sheets/sheet-utils.ts\n  var RUN_STATE_SHEET_NAME = "_RunState";\n  var SHEET_WRITE_CHUNK_ROWS = 1e3;\n  function getOrCreateSheet(spreadsheet, name, options = {}) {\n    let sheet = spreadsheet.getSheetByName(name);\n    if (sheet) {\n      return sheet;\n    }\n    sheet = options.index !== void 0 ? spreadsheet.insertSheet(name, options.index) : spreadsheet.insertSheet(name);\n    if (options.hidden) {\n      sheet.hideSheet();\n    }\n    return sheet;\n  }\n  function readRunStateMap(spreadsheet) {\n    const map = /* @__PURE__ */ new Map();\n    const sheet = spreadsheet.getSheetByName(RUN_STATE_SHEET_NAME);\n    if (!sheet)\n      return map;\n    const lastRow = sheet.getLastRow();\n    if (lastRow < 2)\n      return map;\n    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();\n    for (const row of rows) {\n      const k = String(row[0] ?? "").trim();\n      if (k) {\n        map.set(k, String(row[1] ?? ""));\n      }\n    }\n    return map;\n  }\n  function getRunStateValue(spreadsheet, key, stateMap) {\n    const map = stateMap ?? readRunStateMap(spreadsheet);\n    return map.get(key) ?? null;\n  }\n  function readLockedDailyConfig(spreadsheet, dateStr, stateMap) {\n    const lockedJson = getRunStateValue(spreadsheet, `DAILY_CONFIG_LOCK_${dateStr}`, stateMap);\n    return lockedJson ? parseLockedDailyConfig(lockedJson) : null;\n  }\n  var FORMULA_PREFIX_RE = /^[=+\\-@\\t\\r]/;\n  var ESCAPED_FORMULA_PREFIX_RE = /^\'[=+\\-@\\t\\r]/;\n  function escapeFormulaString(val) {\n    return FORMULA_PREFIX_RE.test(val) ? `\'${val}` : val;\n  }\n  function unescapeFormulaString(val) {\n    return ESCAPED_FORMULA_PREFIX_RE.test(val) ? val.slice(1) : val;\n  }\n  function parseSheetBoolean(value) {\n    return value === true || String(value).toUpperCase() === "TRUE";\n  }\n  function toSheetValueMatrix(columns, items) {\n    return items.map((item) => {\n      const rec = item;\n      return columns.map((k) => {\n        const v = rec[k];\n        if (v === null || v === void 0)\n          return "";\n        return typeof v === "string" ? escapeFormulaString(v) : v;\n      });\n    });\n  }\n  function writeMatrixInChunks(sheet, startRow, startCol, matrix, chunkRows) {\n    if (matrix.length === 0)\n      return;\n    const numCols = matrix[0].length;\n    for (let offset = 0; offset < matrix.length; offset += chunkRows) {\n      const slice = matrix.slice(offset, offset + chunkRows).map((r) => [...r]);\n      sheet.getRange(startRow + offset, startCol, slice.length, numCols).setValues(slice);\n    }\n  }\n  function trimSheetGrid(sheet, targetRows, targetCols) {\n    const updatedMaxRows = sheet.getMaxRows();\n    if (updatedMaxRows > targetRows) {\n      sheet.deleteRows(targetRows + 1, updatedMaxRows - targetRows);\n    }\n    const updatedMaxCols = sheet.getMaxColumns();\n    if (updatedMaxCols > targetCols) {\n      sheet.deleteColumns(targetCols + 1, updatedMaxCols - targetCols);\n    }\n  }\n\n  // src/shared/sheets/mcc-state.ts\n  var MCC_SHEET_PREFIX = "_MCC_";\n  var MAX_ACCOUNT_ATTEMPTS = 3;\n  function getMccSheetName(mccId) {\n    return `${MCC_SHEET_PREFIX}${normalizeCid(mccId)}`;\n  }\n  function writeMccSheetState(spreadsheet, payload) {\n    const normMcc = normalizeCid(payload.mcc.mccId);\n    const formattedMcc = formatCidHyphenated(normMcc);\n    const sheetName = getMccSheetName(normMcc);\n    const sheet = getOrCreateSheet(spreadsheet, sheetName, { hidden: true });\n    const existing = readMccSheetState(spreadsheet, normMcc);\n    const checkpoint = payload.checkpoint ?? existing?.checkpoint;\n    const lastCompletedDate = payload.lastCompletedDate ?? (checkpoint?.status === "COMPLETED" ? checkpoint.dateStr : void 0) ?? existing?.lastCompletedDate ?? "";\n    const statusMessage = payload.statusMessage ?? existing?.statusMessage ?? "";\n    const childAccounts = payload.childAccounts ?? existing?.childAccounts ?? [];\n    const nowIso = (/* @__PURE__ */ new Date()).toISOString();\n    const rows = [\n      ["Key", "Value", "Extra1", "Extra2"],\n      ["MCC_META", formattedMcc, payload.mcc.mccName, nowIso],\n      ["CHECKPOINT", checkpoint ? serializeMccCheckpoint(checkpoint) : "", checkpoint?.dateStr ?? "", nowIso],\n      ["LAST_COMPLETED_DATE", lastCompletedDate, "", nowIso],\n      ["STATUS_MESSAGE", statusMessage, "", nowIso],\n      ["SCRIPT_VERSION", DLG_VERSION, "", nowIso]\n    ];\n    for (const acc of childAccounts) {\n      rows.push([\n        "CHILD_ACCOUNT",\n        formatCidHyphenated(acc.customerId),\n        acc.name,\n        String(acc.currencyCode ?? "USD").trim().toUpperCase()\n      ]);\n    }\n    const prevLastRow = Math.max(1, sheet.getLastRow());\n    const writeRowCount = Math.max(prevLastRow, rows.length);\n    if (writeRowCount > sheet.getMaxRows()) {\n      sheet.insertRowsAfter(sheet.getMaxRows(), writeRowCount - sheet.getMaxRows() + 5);\n    }\n    while (rows.length < writeRowCount) {\n      rows.push(["", "", "", ""]);\n    }\n    const range = sheet.getRange(1, 1, rows.length, 4);\n    range.setNumberFormat("@");\n    range.setValues(rows);\n    if (prevLastRow > rows.length) {\n      sheet.deleteRows(rows.length + 1, prevLastRow - rows.length);\n    }\n  }\n  function readMccSheetState(spreadsheet, mccId) {\n    const normMcc = normalizeCid(mccId);\n    const sheet = spreadsheet.getSheetByName(getMccSheetName(normMcc));\n    if (!sheet)\n      return null;\n    const lastRow = sheet.getLastRow();\n    if (lastRow < 2)\n      return null;\n    const rawRows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();\n    let formattedMcc = formatCidHyphenated(normMcc);\n    let mccName = "MCC";\n    let checkpoint;\n    let lastCompletedDate;\n    let statusMessage;\n    let scriptVersion;\n    const childAccounts = [];\n    for (const r of rawRows) {\n      const key = String(r[0] ?? "").trim();\n      const rawVal = r[1];\n      const val = rawVal instanceof Date ? rawVal.toISOString().slice(0, 10) : String(rawVal ?? "").trim();\n      const extra1 = String(r[2] ?? "").trim();\n      const extra2 = String(r[3] ?? "").trim();\n      if (key === "MCC_META") {\n        if (val)\n          formattedMcc = val;\n        if (extra1)\n          mccName = extra1;\n      } else if (key === "CHECKPOINT" && val) {\n        checkpoint = parseMccCheckpoint(val) ?? void 0;\n      } else if (key === "SCRIPT_VERSION" && val) {\n        scriptVersion = val;\n      } else if (key === "LAST_COMPLETED_DATE" && val) {\n        lastCompletedDate = val;\n      } else if (key === "STATUS_MESSAGE" && val) {\n        statusMessage = val;\n      } else if (key === "CHILD_ACCOUNT" && val) {\n        childAccounts.push({\n          customerId: val,\n          name: extra1 || "Account",\n          currencyCode: extra2 || "USD"\n        });\n      }\n    }\n    return {\n      mcc: { mccId: formattedMcc, mccName },\n      childAccounts,\n      statusMessage,\n      checkpoint,\n      lastCompletedDate,\n      scriptVersion\n    };\n  }\n  function getMccCheckpoint(spreadsheet, dateStr, mccId) {\n    const normMcc = normalizeCid(mccId);\n    const mccPayload = readMccSheetState(spreadsheet, normMcc);\n    const lastCompletedDate = mccPayload?.lastCompletedDate ?? mccPayload?.checkpoint?.lastCompletedDate;\n    if (mccPayload?.checkpoint && mccPayload.checkpoint.dateStr === dateStr) {\n      return {\n        ...mccPayload.checkpoint,\n        mccId: normMcc,\n        lastCompletedDate\n      };\n    }\n    return {\n      mccId: normMcc,\n      dateStr,\n      status: "IN_PROGRESS",\n      completedCids: [],\n      lastCompletedDate\n    };\n  }\n\n  // src/shared/sheets/staging-writer.ts\n  var STAGE_ACCOUNTS_SHEET = "_Stage_Accounts";\n  var STAGE_BUDGETS_SHEET = "_Stage_Budgets";\n  var STAGE_CAMPAIGNS_SHEET = "_Stage_Campaigns";\n  var STAGE_RECS_SHEET = "_Stage_Recs";\n  function getMccStagingSheetNames(mccId) {\n    const norm = normalizeCid(mccId);\n    return {\n      accounts: `${STAGE_ACCOUNTS_SHEET}_${norm}`,\n      budgets: `${STAGE_BUDGETS_SHEET}_${norm}`,\n      campaigns: `${STAGE_CAMPAIGNS_SHEET}_${norm}`,\n      recommendations: `${STAGE_RECS_SHEET}_${norm}`\n    };\n  }\n  var STAGED_ACCOUNT_SCHEMA = {\n    mccId: "STRING",\n    mccName: "STRING",\n    accountId: "STRING",\n    accountName: "STRING",\n    currencyCode: "STRING",\n    timestamp: "STRING"\n  };\n  var STAGED_BUDGET_SCHEMA = {\n    mccId: "STRING",\n    accountId: "STRING",\n    budgetResourceName: "STRING",\n    campaignBudgetName: "STRING",\n    campaignBudgetAmount: "NUMBER",\n    campaignBudgetTotalAmount: "NUMBER",\n    campaignBudgetDeliveryMethod: "STRING",\n    campaignBudgetType: "STRING",\n    campaignBudgetIsShared: "BOOLEAN"\n  };\n  var STAGED_CAMPAIGN_SCHEMA = {\n    mccId: "STRING",\n    accountId: "STRING",\n    campaignResourceName: "STRING",\n    campaignId: "STRING",\n    campaignName: "STRING",\n    campaignType: "STRING",\n    campaignSubType: "STRING",\n    campaignIsAiMax: "BOOLEAN",\n    campaignAiMaxTextCustomizationEnabled: "BOOLEAN",\n    campaignAiMaxFinalUrlExpansionEnabled: "BOOLEAN",\n    budgetResourceName: "STRING",\n    campaignIsPortfolioBiddingStrategy: "BOOLEAN",\n    campaignPortfolioBiddingStrategyName: "STRING",\n    portfolioStrategyResourceName: "STRING",\n    portfolioOwnerId: "STRING",\n    isLocallyOwnedPortfolio: "BOOLEAN",\n    campaignPortfolioIsManagerOwned: "BOOLEAN",\n    campaignBiddingStrategyType: "STRING",\n    campaignBiddingTargetRoas: "NUMBER",\n    campaignBiddingTargetCpa: "NUMBER",\n    campaignStatsYesterdayCost: "NUMBER",\n    campaignStats7DaysCost: "NUMBER",\n    campaign30DaysCost: "NUMBER",\n    campaign30DaysConversionsValue: "NUMBER",\n    campaign30DaysConversions: "NUMBER",\n    campaign30DaysSearchRankLostImpressionShare: "NUMBER",\n    campaign30DaysSearchBudgetLostImpressionShare: "NUMBER"\n  };\n  var STAGED_REC_SCHEMA = {\n    mccId: "STRING",\n    accountId: "STRING",\n    recommendationId: "STRING",\n    recommendationType: "STRING",\n    campaignResourceName: "STRING",\n    recommendationCurrentBudgetAmount: "NUMBER",\n    recommendationNewBudgetAmount: "NUMBER",\n    recommendationNewTargetCpa: "NUMBER",\n    recommendationNewTargetRoas: "NUMBER",\n    recommendationBaseCost: "NUMBER",\n    recommendationPotentialCost: "NUMBER",\n    recommendationBaseClicks: "NUMBER",\n    recommendationPotentialClicks: "NUMBER",\n    recommendationBaseConversions: "NUMBER",\n    recommendationPotentialConversions: "NUMBER",\n    recommendationBaseConversionsValue: "NUMBER",\n    recommendationPotentialConversionsValue: "NUMBER",\n    recommendationBaseImpressions: "NUMBER",\n    recommendationPotentialImpressions: "NUMBER",\n    recommendationBaseVideoViews: "NUMBER",\n    recommendationPotentialVideoViews: "NUMBER",\n    moveBudgetSourceBudgetResourceName: "STRING",\n    moveBudgetSourceCampaigns: "STRING",\n    moveBudgetSourceTargetCpa: "NUMBER",\n    moveBudgetSourceTargetRoas: "NUMBER"\n  };\n  var STAGED_ACCOUNT_KEYS = Object.keys(STAGED_ACCOUNT_SCHEMA);\n  var STAGED_BUDGET_KEYS = Object.keys(STAGED_BUDGET_SCHEMA);\n  var STAGED_CAMPAIGN_KEYS = Object.keys(STAGED_CAMPAIGN_SCHEMA);\n  var STAGED_REC_KEYS = Object.keys(STAGED_REC_SCHEMA);\n  function appendStagingBatch(workbook, batch, mccId) {\n    const names = getMccStagingSheetNames(mccId);\n    upsertStagingTable(\n      getOrCreateHiddenSheet(workbook, names.accounts),\n      STAGED_ACCOUNT_KEYS,\n      STAGED_ACCOUNT_SCHEMA,\n      batch.accounts,\n      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}`\n    );\n    upsertStagingTable(\n      getOrCreateHiddenSheet(workbook, names.budgets),\n      STAGED_BUDGET_KEYS,\n      STAGED_BUDGET_SCHEMA,\n      batch.budgets,\n      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.budgetResourceName}`\n    );\n    upsertStagingTable(\n      getOrCreateHiddenSheet(workbook, names.campaigns),\n      STAGED_CAMPAIGN_KEYS,\n      STAGED_CAMPAIGN_SCHEMA,\n      batch.campaigns,\n      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.campaignResourceName}`\n    );\n    upsertStagingTable(\n      getOrCreateHiddenSheet(workbook, names.recommendations),\n      STAGED_REC_KEYS,\n      STAGED_REC_SCHEMA,\n      batch.recommendations,\n      (r) => `${normalizeCid(r.mccId)}::${normalizeCid(r.accountId)}::${r.recommendationId}::${r.campaignResourceName}`\n    );\n  }\n  function createEmptyStagingBatch() {\n    return { accounts: [], budgets: [], campaigns: [], recommendations: [] };\n  }\n  function clearMccFromStaging(workbook, mccId) {\n    const normMcc = normalizeCid(mccId);\n    const perMccNames = getMccStagingSheetNames(normMcc);\n    const perMccSpecs = [\n      [perMccNames.accounts, STAGED_ACCOUNT_KEYS, STAGED_ACCOUNT_SCHEMA],\n      [perMccNames.budgets, STAGED_BUDGET_KEYS, STAGED_BUDGET_SCHEMA],\n      [perMccNames.campaigns, STAGED_CAMPAIGN_KEYS, STAGED_CAMPAIGN_SCHEMA],\n      [perMccNames.recommendations, STAGED_REC_KEYS, STAGED_REC_SCHEMA]\n    ];\n    for (const [name, keys, schema] of perMccSpecs) {\n      const sheet = workbook.getSheetByName(name);\n      if (sheet) {\n        writeTableRows(sheet, keys, schema, []);\n      }\n    }\n  }\n  function getOrCreateHiddenSheet(workbook, name) {\n    return getOrCreateSheet(workbook, name, { hidden: true });\n  }\n  function upsertStagingTable(sheet, keys, schema, newItems, keyFn) {\n    const lastRow = sheet.getLastRow();\n    if (newItems.length === 0 && lastRow > 0)\n      return;\n    const existing = readTableRows(sheet, keys, schema);\n    const existingKeys = new Set(existing.map(keyFn));\n    const newKeys = /* @__PURE__ */ new Set();\n    let collides = false;\n    for (const item of newItems) {\n      const k = keyFn(item);\n      if (existingKeys.has(k) || newKeys.has(k)) {\n        collides = true;\n        break;\n      }\n      newKeys.add(k);\n    }\n    if (lastRow > 0 && !collides && headerMatchesKeys(sheet, keys)) {\n      appendTableRows(sheet, keys, schema, newItems, lastRow + 1);\n      return;\n    }\n    const mergedMap = /* @__PURE__ */ new Map();\n    for (const item of existing) {\n      mergedMap.set(keyFn(item), item);\n    }\n    for (const item of newItems) {\n      mergedMap.set(keyFn(item), item);\n    }\n    writeTableRows(sheet, keys, schema, Array.from(mergedMap.values()));\n  }\n  function headerMatchesKeys(sheet, keys) {\n    const maxCols = sheet.getMaxColumns();\n    if (maxCols < keys.length)\n      return false;\n    const header = sheet.getRange(1, 1, 1, maxCols).getValues()[0] ?? [];\n    for (let c = 0; c < maxCols; c++) {\n      const cell = String(header[c] ?? "").trim();\n      if (cell !== (c < keys.length ? keys[c] : ""))\n        return false;\n    }\n    return true;\n  }\n  function formatForFieldType(fieldType) {\n    switch (fieldType) {\n      case "NUMBER":\n        return "0.########";\n      case "BOOLEAN":\n        return "General";\n      default:\n        return "@";\n    }\n  }\n  function appendTableRows(sheet, keys, schema, items, startRow) {\n    if (items.length === 0)\n      return;\n    const numCols = keys.length;\n    const lastNeededRow = startRow + items.length - 1;\n    if (sheet.getMaxRows() < lastNeededRow) {\n      sheet.insertRowsAfter(sheet.getMaxRows(), lastNeededRow - sheet.getMaxRows());\n    }\n    const rowFormats = keys.map((k) => formatForFieldType(schema[k]));\n    for (let offset = 0; offset < items.length; offset += SHEET_WRITE_CHUNK_ROWS) {\n      const count = Math.min(SHEET_WRITE_CHUNK_ROWS, items.length - offset);\n      const fmtMatrix = Array.from({ length: count }, () => [...rowFormats]);\n      sheet.getRange(startRow + offset, 1, count, numCols).setNumberFormats(fmtMatrix);\n    }\n    writeMatrixInChunks(sheet, startRow, 1, toSheetValueMatrix(keys, items), SHEET_WRITE_CHUNK_ROWS);\n  }\n  function writeTableRows(sheet, keys, schema, items) {\n    const numCols = keys.length;\n    const totalRows = items.length + 1;\n    if (sheet.getMaxColumns() < numCols) {\n      sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());\n    }\n    if (sheet.getMaxRows() < totalRows) {\n      sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());\n    }\n    const rowFormats = keys.map((k) => formatForFieldType(schema[k]));\n    for (let offset = 0; offset < totalRows; offset += SHEET_WRITE_CHUNK_ROWS) {\n      const count = Math.min(SHEET_WRITE_CHUNK_ROWS, totalRows - offset);\n      const fmtMatrix = Array.from(\n        { length: count },\n        () => offset === 0 ? keys.map(() => "@") : [...rowFormats]\n      );\n      if (offset === 0 && count > 1) {\n        for (let r = 1; r < count; r++) {\n          fmtMatrix[r] = [...rowFormats];\n        }\n      }\n      sheet.getRange(1 + offset, 1, count, numCols).setNumberFormats(fmtMatrix);\n    }\n    const matrix = [[...keys], ...toSheetValueMatrix(keys, items)];\n    writeMatrixInChunks(sheet, 1, 1, matrix, SHEET_WRITE_CHUNK_ROWS);\n    trimSheetGrid(sheet, totalRows, numCols);\n  }\n  function readTableRows(sheet, keys, schema) {\n    if (!sheet)\n      return [];\n    const lastRow = sheet.getLastRow();\n    if (lastRow < 2)\n      return [];\n    const numCols = sheet.getMaxColumns();\n    if (numCols < 1)\n      return [];\n    const allValues = sheet.getRange(1, 1, lastRow, numCols).getValues();\n    const headerRow = allValues[0] ?? [];\n    const colIndexByKey = /* @__PURE__ */ new Map();\n    for (let c = 0; c < headerRow.length; c++) {\n      const headerName = String(headerRow[c] ?? "").trim();\n      if (headerName && !colIndexByKey.has(headerName)) {\n        colIndexByKey.set(headerName, c);\n      }\n    }\n    const out = [];\n    for (let r = 1; r < allValues.length; r++) {\n      const rawRow = allValues[r];\n      const obj = {};\n      let hasNonEmptyCell = false;\n      for (const key of keys) {\n        const k = String(key);\n        const colIdx = colIndexByKey.get(k);\n        const cell = colIdx === void 0 ? void 0 : rawRow[colIdx];\n        if (cell !== "" && cell !== null && cell !== void 0) {\n          hasNonEmptyCell = true;\n        }\n        obj[k] = coerceStagingCell(cell, schema[k] ?? "STRING");\n      }\n      if (hasNonEmptyCell) {\n        out.push(obj);\n      }\n    }\n    return out;\n  }\n  function coerceStagingCell(cell, fieldType) {\n    if (fieldType === "BOOLEAN") {\n      return parseSheetBoolean(cell);\n    }\n    if (fieldType === "NUMBER") {\n      if (cell === "" || cell === null || cell === void 0)\n        return null;\n      const num = Number(cell);\n      return Number.isFinite(num) ? num : null;\n    }\n    if (fieldType === "NULLABLE_STRING") {\n      if (cell === "" || cell === null || cell === void 0)\n        return null;\n      if (cell instanceof Date) {\n        return cell.toISOString().replace("T", " ").slice(0, 19);\n      }\n      return unescapeFormulaString(String(cell));\n    }\n    if (cell === null || cell === void 0)\n      return "";\n    if (cell instanceof Date) {\n      return cell.toISOString().replace("T", " ").slice(0, 19);\n    }\n    return unescapeFormulaString(String(cell));\n  }\n\n  // src/ads/gaql/queries.ts\n  var GAQL_IN_CLAUSE_BATCH_SIZE = 500;\n  var SUPPORTED_RECOMMENDATION_TYPES = [\n    "CAMPAIGN_BUDGET",\n    "FORECASTING_CAMPAIGN_BUDGET",\n    "MARGINAL_ROI_CAMPAIGN_BUDGET",\n    "MOVE_UNUSED_BUDGET",\n    "RAISE_TARGET_CPA",\n    "LOWER_TARGET_ROAS"\n  ];\n  function chunkArray(items, batchSize = GAQL_IN_CLAUSE_BATCH_SIZE) {\n    if (items.length === 0)\n      return [];\n    const chunks = [];\n    for (let i = 0; i < items.length; i += batchSize) {\n      chunks.push(items.slice(i, i + batchSize));\n    }\n    return chunks;\n  }\n  function formatInClause(resourceNames, builderName) {\n    if (resourceNames.length === 0) {\n      throw new Error(`${builderName}: at least one resource name is required (got an empty list).`);\n    }\n    return resourceNames.map((name) => `\'${name}\'`).join(",");\n  }\n  function buildRecommendationDetailsQuery() {\n    const typeList = SUPPORTED_RECOMMENDATION_TYPES.map((t) => `\'${t}\'`).join(",");\n    return `\n    SELECT\n      recommendation.resource_name,\n      recommendation.type,\n      recommendation.campaign,\n      recommendation.campaigns,\n      recommendation.campaign_budget_recommendation,\n      recommendation.forecasting_campaign_budget_recommendation,\n      recommendation.marginal_roi_campaign_budget_recommendation,\n      recommendation.move_unused_budget_recommendation,\n      recommendation.raise_target_cpa_recommendation,\n      recommendation.lower_target_roas_recommendation,\n      recommendation.impact\n    FROM recommendation\n    WHERE recommendation.type IN (${typeList})\n  `.trim();\n  }\n  function buildCampaignMetadataQuery(campaignResourceNames, excludeShoppingAndPmax = false) {\n    const shoppingFilter = excludeShoppingAndPmax ? "\\n      AND campaign.advertising_channel_type != \'SHOPPING\'" : "";\n    return `\n    SELECT\n      customer.id,\n      campaign.resource_name,\n      campaign.name,\n      campaign.status,\n      campaign.advertising_channel_type,\n      campaign.advertising_channel_sub_type,\n      campaign.shopping_setting.merchant_id,\n      campaign.ai_max_setting.enable_ai_max,\n      campaign.asset_automation_settings,\n      campaign.bidding_strategy_type,\n      campaign.bidding_strategy,\n      campaign.maximize_conversions.target_cpa_micros,\n      campaign.maximize_conversion_value.target_roas,\n      campaign.target_cpa.target_cpa_micros,\n      campaign.target_roas.target_roas,\n      campaign_budget.resource_name,\n      campaign_budget.amount_micros,\n      campaign_budget.total_amount_micros,\n      campaign_budget.delivery_method,\n      campaign_budget.name,\n      campaign_budget.period,\n      campaign_budget.explicitly_shared,\n      accessible_bidding_strategy.name,\n      accessible_bidding_strategy.owner_customer_id,\n      accessible_bidding_strategy.id,\n      accessible_bidding_strategy.maximize_conversions.target_cpa_micros,\n      accessible_bidding_strategy.maximize_conversion_value.target_roas,\n      accessible_bidding_strategy.target_cpa.target_cpa_micros,\n      accessible_bidding_strategy.target_roas.target_roas,\n      accessible_bidding_strategy.type\n    FROM campaign\n    WHERE campaign.resource_name IN (${formatInClause(campaignResourceNames, "buildCampaignMetadataQuery")})\n      AND campaign.status = \'ENABLED\'${shoppingFilter}\n  `.trim();\n  }\n  function buildCampaignMetrics30DaysQuery(campaignResourceNames) {\n    return `\n    SELECT\n      campaign.resource_name,\n      metrics.cost_micros,\n      metrics.conversions_value,\n      metrics.conversions,\n      metrics.search_rank_lost_impression_share,\n      metrics.search_budget_lost_impression_share\n    FROM campaign\n    WHERE campaign.resource_name IN (${formatInClause(campaignResourceNames, "buildCampaignMetrics30DaysQuery")})\n      AND segments.date DURING LAST_30_DAYS\n  `.trim();\n  }\n  function buildCampaignCostQuery(campaignResourceNames, dateCondition) {\n    return `\n    SELECT\n      campaign.resource_name,\n      metrics.cost_micros\n    FROM campaign\n    WHERE campaign.resource_name IN (${formatInClause(campaignResourceNames, "buildCampaignCostQuery")})\n      AND segments.date ${dateCondition}\n  `.trim();\n  }\n  function buildBudgetDetailsQuery(budgetResourceNames) {\n    return `\n    SELECT\n      campaign_budget.resource_name,\n      campaign_budget.amount_micros,\n      campaign_budget.total_amount_micros,\n      campaign_budget.name,\n      campaign_budget.period,\n      campaign_budget.delivery_method,\n      campaign_budget.explicitly_shared\n    FROM campaign_budget\n    WHERE campaign_budget.resource_name IN (${formatInClause(budgetResourceNames, "buildBudgetDetailsQuery")})\n  `.trim();\n  }\n  function buildCampaignsByBudgetsQuery(budgetResourceNames, excludeShoppingAndPmax = false) {\n    const shoppingFilter = excludeShoppingAndPmax ? "\\n      AND campaign.advertising_channel_type != \'SHOPPING\'" : "";\n    return `\n    SELECT\n      campaign.resource_name,\n      campaign.name,\n      campaign.status,\n      campaign.advertising_channel_type,\n      campaign.shopping_setting.merchant_id,\n      campaign.bidding_strategy,\n      campaign.maximize_conversions.target_cpa_micros,\n      campaign.maximize_conversion_value.target_roas,\n      campaign.target_cpa.target_cpa_micros,\n      campaign.target_roas.target_roas,\n      accessible_bidding_strategy.maximize_conversions.target_cpa_micros,\n      accessible_bidding_strategy.maximize_conversion_value.target_roas,\n      accessible_bidding_strategy.target_cpa.target_cpa_micros,\n      accessible_bidding_strategy.target_roas.target_roas,\n      campaign_budget.resource_name\n    FROM campaign\n    WHERE campaign_budget.resource_name IN (${formatInClause(budgetResourceNames, "buildCampaignsByBudgetsQuery")})\n      AND campaign.status = \'ENABLED\'${shoppingFilter}\n  `.trim();\n  }\n\n  // src/ads/gaql/fetcher.ts\n  function fromMicros(micros) {\n    if (micros === null || micros === void 0 || micros === "") {\n      return null;\n    }\n    const num = Number(micros);\n    return Number.isFinite(num) ? num / 1e6 : null;\n  }\n  function isCampaignExcludedByEcRemedy(campaign, excludeShoppingAndPmax) {\n    if (!excludeShoppingAndPmax) {\n      return false;\n    }\n    const type = String(campaign.campaignType ?? "").toUpperCase();\n    if (type === "SHOPPING") {\n      return true;\n    }\n    if (type === "PERFORMANCE_MAX") {\n      const merchantIdStr = String(campaign.merchantId ?? "").trim();\n      return merchantIdStr.length > 0;\n    }\n    return false;\n  }\n  function executeGaqlSearch(query) {\n    const iterator = AdsApp.search(query);\n    const rows = [];\n    while (iterator.hasNext()) {\n      rows.push(iterator.next());\n    }\n    return rows;\n  }\n  function fetchAllRecommendationDetails() {\n    const detailsMap = /* @__PURE__ */ new Map();\n    const rows = executeGaqlSearch(buildRecommendationDetailsQuery());\n    for (const row of rows) {\n      if (row.recommendation?.resourceName) {\n        detailsMap.set(row.recommendation.resourceName, row.recommendation);\n      }\n    }\n    return detailsMap;\n  }\n  function fetchAllCampaignData(campaignResourceNames, excludeShoppingAndPmax = false) {\n    const campaignMap = /* @__PURE__ */ new Map();\n    if (campaignResourceNames.length === 0) {\n      return campaignMap;\n    }\n    for (const batch of chunkArray(campaignResourceNames)) {\n      const metaRows = executeGaqlSearch(buildCampaignMetadataQuery(batch, excludeShoppingAndPmax));\n      for (const row of metaRows) {\n        const c = row.campaign ?? {};\n        const resName = String(c.resourceName ?? "");\n        if (!resName)\n          continue;\n        const campaignType = c.advertisingChannelType ?? "N/A";\n        const merchantId = extractMerchantId(c);\n        if (isCampaignExcludedByEcRemedy({ campaignType, merchantId }, excludeShoppingAndPmax)) {\n          continue;\n        }\n        const campaignId = resName.split("/")[3] ?? "N/A";\n        const { targetCpa, targetRoas, isPortfolio, portfolioName, portfolioRes, ownerId, isLocal, isManager } = extractBiddingDetails(row);\n        const { isAiMax, textOptIn, urlOptIn } = extractAiMaxDetails(c);\n        const budget = extractBudgetDetails(row.campaignBudget);\n        campaignMap.set(resName, {\n          campaignResourceName: resName,\n          campaignId,\n          campaignName: c.name ?? "N/A",\n          campaignType,\n          campaignSubType: c.advertisingChannelSubType ?? "N/A",\n          merchantId,\n          campaignIsAiMax: isAiMax,\n          campaignAiMaxTextCustomizationEnabled: textOptIn,\n          campaignAiMaxFinalUrlExpansionEnabled: urlOptIn,\n          campaignBudgetResourceName: row.campaignBudget?.resourceName ?? null,\n          campaignBudgetName: budget.campaignBudgetName,\n          campaignBudgetAmount: budget.campaignBudgetAmount,\n          campaignBudgetTotalAmount: budget.campaignBudgetTotalAmount,\n          campaignBudgetDeliveryMethod: budget.campaignBudgetDeliveryMethod,\n          campaignBudgetType: budget.campaignBudgetType,\n          campaignBudgetIsShared: budget.campaignBudgetIsShared,\n          campaignIsPortfolioBiddingStrategy: isPortfolio,\n          campaignPortfolioBiddingStrategyName: portfolioName,\n          portfolioStrategyResourceName: portfolioRes,\n          portfolioOwnerId: ownerId,\n          isLocallyOwnedPortfolio: isLocal,\n          campaignPortfolioIsManagerOwned: isManager,\n          campaignBiddingStrategyType: c.biddingStrategyType ?? "N/A",\n          campaignBiddingTargetCpa: targetCpa,\n          campaignBiddingTargetRoas: targetRoas,\n          campaignStatsYesterdayCost: null,\n          campaignStats7DaysCost: null,\n          campaign30DaysCost: null,\n          campaign30DaysConversionsValue: null,\n          campaign30DaysConversions: null,\n          campaign30DaysRoas: null,\n          campaign30DaysCpa: null,\n          campaign30DaysSearchRankLostImpressionShare: null,\n          campaign30DaysSearchBudgetLostImpressionShare: null\n        });\n      }\n      const enabledBatchNames = batch.filter((name) => campaignMap.has(name));\n      if (enabledBatchNames.length === 0)\n        continue;\n      for (const row of executeGaqlSearch(buildCampaignMetrics30DaysQuery(enabledBatchNames))) {\n        const entry = campaignMap.get(row.campaign?.resourceName ?? "");\n        if (!entry || !row.metrics)\n          continue;\n        const cost = fromMicros(row.metrics.costMicros ?? 0) ?? 0;\n        const convValue = Number(row.metrics.conversionsValue ?? 0);\n        const convs = Number(row.metrics.conversions ?? 0);\n        entry.campaign30DaysCost = cost;\n        entry.campaign30DaysConversionsValue = convValue;\n        entry.campaign30DaysConversions = convs;\n        entry.campaign30DaysRoas = cost > 0 ? convValue / cost : null;\n        entry.campaign30DaysCpa = convs > 0 ? cost / convs : null;\n        entry.campaign30DaysSearchRankLostImpressionShare = row.metrics.searchRankLostImpressionShare ?? null;\n        entry.campaign30DaysSearchBudgetLostImpressionShare = row.metrics.searchBudgetLostImpressionShare ?? null;\n      }\n      for (const row of executeGaqlSearch(buildCampaignCostQuery(enabledBatchNames, "DURING YESTERDAY"))) {\n        const entry = campaignMap.get(row.campaign?.resourceName ?? "");\n        if (entry && row.metrics) {\n          entry.campaignStatsYesterdayCost = fromMicros(row.metrics.costMicros ?? 0);\n        }\n      }\n      for (const row of executeGaqlSearch(buildCampaignCostQuery(enabledBatchNames, "DURING LAST_7_DAYS"))) {\n        const entry = campaignMap.get(row.campaign?.resourceName ?? "");\n        if (entry && row.metrics) {\n          entry.campaignStats7DaysCost = fromMicros(row.metrics.costMicros ?? 0);\n        }\n      }\n    }\n    return campaignMap;\n  }\n  function fetchAllBudgetDetails(budgetResourceNames) {\n    const budgetMap = /* @__PURE__ */ new Map();\n    for (const batch of chunkArray(budgetResourceNames)) {\n      for (const row of executeGaqlSearch(buildBudgetDetailsQuery(batch))) {\n        const resName = row.campaignBudget?.resourceName;\n        if (resName) {\n          budgetMap.set(resName, extractBudgetDetails(row.campaignBudget));\n        }\n      }\n    }\n    return budgetMap;\n  }\n  function fetchCampaignsByBudgets(budgetResourceNames, excludeShoppingAndPmax = false) {\n    const byBudget = /* @__PURE__ */ new Map();\n    for (const batch of chunkArray(budgetResourceNames)) {\n      for (const row of executeGaqlSearch(\n        buildCampaignsByBudgetsQuery(batch, excludeShoppingAndPmax)\n      )) {\n        const budgetRes = row.campaignBudget?.resourceName;\n        const c = row.campaign;\n        if (!budgetRes || !c)\n          continue;\n        const campaignType = c.advertisingChannelType ?? "N/A";\n        const merchantId = extractMerchantId(c);\n        if (isCampaignExcludedByEcRemedy({ campaignType, merchantId }, excludeShoppingAndPmax)) {\n          continue;\n        }\n        const { targetCpa, targetRoas } = extractBiddingDetails(row);\n        const list = byBudget.get(budgetRes) ?? [];\n        list.push({\n          campaignResourceName: c.resourceName ?? "",\n          campaignName: c.name ?? "N/A",\n          campaignType,\n          merchantId,\n          targetCpa,\n          targetRoas\n        });\n        byBudget.set(budgetRes, list);\n      }\n    }\n    return byBudget;\n  }\n  function extractMerchantId(c) {\n    const id = c.shoppingSetting?.merchantId;\n    return id === void 0 || id === null ? null : String(id);\n  }\n  function extractBudgetDetails(b) {\n    if (!b) {\n      return {\n        budgetResourceName: "",\n        campaignBudgetName: "N/A",\n        campaignBudgetAmount: null,\n        campaignBudgetTotalAmount: null,\n        campaignBudgetDeliveryMethod: "N/A",\n        campaignBudgetType: "N/A",\n        campaignBudgetIsShared: false\n      };\n    }\n    return {\n      budgetResourceName: b.resourceName ?? "",\n      campaignBudgetName: b.name || "N/A",\n      campaignBudgetAmount: fromMicros(b.amountMicros),\n      campaignBudgetTotalAmount: fromMicros(b.totalAmountMicros),\n      campaignBudgetDeliveryMethod: b.deliveryMethod || "N/A",\n      // Budget period (DAILY / CUSTOM_PERIOD); campaign_budget.type is not selected\n      campaignBudgetType: b.period || "N/A",\n      campaignBudgetIsShared: Boolean(b.explicitlyShared)\n    };\n  }\n  function extractBiddingDetails(row) {\n    const c = row.campaign ?? {};\n    const abs = row.accessibleBiddingStrategy;\n    const isPortfolio = Boolean(c.biddingStrategy);\n    const portfolioRes = c.biddingStrategy ?? "N/A";\n    const portfolioName = abs?.name ?? "N/A";\n    const ownerCid = normalizeCid(abs?.ownerCustomerId);\n    const customerCid = normalizeCid(row.customer?.id);\n    const isLocal = Boolean(isPortfolio && ownerCid && customerCid && ownerCid === customerCid);\n    const isManager = Boolean(isPortfolio && ownerCid && customerCid && ownerCid !== customerCid);\n    const rawCpaMicros = isPortfolio ? abs?.targetCpa?.targetCpaMicros ?? abs?.maximizeConversions?.targetCpaMicros : c.targetCpa?.targetCpaMicros ?? c.maximizeConversions?.targetCpaMicros;\n    const rawRoas = isPortfolio ? abs?.targetRoas?.targetRoas ?? abs?.maximizeConversionValue?.targetRoas : c.targetRoas?.targetRoas ?? c.maximizeConversionValue?.targetRoas;\n    return {\n      isPortfolio,\n      portfolioRes,\n      portfolioName,\n      ownerId: abs?.ownerCustomerId ? String(abs.ownerCustomerId) : "N/A",\n      isLocal,\n      isManager,\n      targetCpa: fromMicros(rawCpaMicros),\n      targetRoas: typeof rawRoas === "number" && Number.isFinite(rawRoas) ? rawRoas : null\n    };\n  }\n  function extractAiMaxDetails(campaign) {\n    const isAiMax = Boolean(campaign.aiMaxSetting?.enableAiMax);\n    const rawSettings = campaign.assetAutomationSettings ?? [];\n    const settings = Array.isArray(rawSettings) ? rawSettings : typeof rawSettings === "object" && rawSettings !== null ? Object.values(rawSettings) : [];\n    let textOptIn = false;\n    let urlOptIn = false;\n    for (const s of settings) {\n      if (s.assetAutomationType === "TEXT_ASSET_AUTOMATION") {\n        textOptIn = s.assetAutomationStatus === "OPTED_IN";\n      } else if (s.assetAutomationType === "FINAL_URL_EXPANSION_TEXT_ASSET_AUTOMATION") {\n        urlOptIn = s.assetAutomationStatus === "OPTED_IN";\n      }\n    }\n    return { isAiMax, textOptIn, urlOptIn };\n  }\n\n  // src/ads/processors/recommendations.ts\n  function buildAccountStagingBatch(ctx) {\n    const budgetsByKey = /* @__PURE__ */ new Map();\n    const campaignsByKey = /* @__PURE__ */ new Map();\n    const stagedRecs = [];\n    for (const [recResourceName, recDetails] of ctx.recommendationDetailsMap.entries()) {\n      const type = String(recDetails.type ?? "");\n      const targetCampaignNames = extractTargetCampaignResourceNames(recDetails);\n      if (targetCampaignNames.length === 0)\n        continue;\n      const targetCampaigns = targetCampaignNames.map((name) => ctx.campaignDataMap.get(name)).filter((c) => c !== void 0);\n      if (targetCampaigns.length === 0)\n        continue;\n      if (ctx.excludeShoppingAndPmax && targetCampaigns.length < targetCampaignNames.length) {\n        continue;\n      }\n      if (targetCampaigns.some((c) => isCampaignExcludedByEcRemedy(c, ctx.excludeShoppingAndPmax))) {\n        continue;\n      }\n      const excessBudgetRes = String(\n        recDetails.moveUnusedBudgetRecommendation?.excessCampaignBudget ?? ""\n      );\n      const sourceCampaigns = excessBudgetRes ? ctx.budgetToCampaignsMap.get(excessBudgetRes) ?? [] : [];\n      if (type === "MOVE_UNUSED_BUDGET") {\n        if (ctx.excludeShoppingAndPmax && excessBudgetRes && sourceCampaigns.length === 0) {\n          continue;\n        }\n        if (sourceCampaigns.some((sc) => isCampaignExcludedByEcRemedy(sc, ctx.excludeShoppingAndPmax))) {\n          continue;\n        }\n      }\n      const impactStats = extractNativeImpactMetrics(recDetails);\n      if (type === "MOVE_UNUSED_BUDGET" && excessBudgetRes) {\n        const sourceBudget = ctx.budgetDataMap.get(excessBudgetRes);\n        if (sourceBudget && !budgetsByKey.has(excessBudgetRes)) {\n          budgetsByKey.set(excessBudgetRes, {\n            mccId: ctx.mccId,\n            accountId: ctx.accountId,\n            budgetResourceName: excessBudgetRes,\n            campaignBudgetName: sourceBudget.campaignBudgetName,\n            campaignBudgetAmount: sourceBudget.campaignBudgetAmount,\n            campaignBudgetTotalAmount: sourceBudget.campaignBudgetTotalAmount,\n            campaignBudgetDeliveryMethod: sourceBudget.campaignBudgetDeliveryMethod,\n            campaignBudgetType: sourceBudget.campaignBudgetType,\n            campaignBudgetIsShared: sourceBudget.campaignBudgetIsShared\n          });\n        }\n      }\n      for (const campaign of targetCampaigns) {\n        const budgetResName = campaign.campaignBudgetResourceName || `customers/${ctx.accountId}/campaignBudgets/camp_${campaign.campaignId}`;\n        if (!budgetsByKey.has(budgetResName)) {\n          budgetsByKey.set(budgetResName, {\n            mccId: ctx.mccId,\n            accountId: ctx.accountId,\n            budgetResourceName: budgetResName,\n            campaignBudgetName: campaign.campaignBudgetName,\n            campaignBudgetAmount: campaign.campaignBudgetAmount,\n            campaignBudgetTotalAmount: campaign.campaignBudgetTotalAmount,\n            campaignBudgetDeliveryMethod: campaign.campaignBudgetDeliveryMethod,\n            campaignBudgetType: campaign.campaignBudgetType,\n            campaignBudgetIsShared: campaign.campaignBudgetIsShared\n          });\n        }\n        if (!campaignsByKey.has(campaign.campaignResourceName)) {\n          campaignsByKey.set(campaign.campaignResourceName, {\n            mccId: ctx.mccId,\n            accountId: ctx.accountId,\n            campaignResourceName: campaign.campaignResourceName,\n            campaignId: campaign.campaignId,\n            campaignName: campaign.campaignName,\n            campaignType: campaign.campaignType,\n            campaignSubType: campaign.campaignSubType,\n            campaignIsAiMax: campaign.campaignIsAiMax,\n            campaignAiMaxTextCustomizationEnabled: campaign.campaignAiMaxTextCustomizationEnabled,\n            campaignAiMaxFinalUrlExpansionEnabled: campaign.campaignAiMaxFinalUrlExpansionEnabled,\n            budgetResourceName: budgetResName,\n            campaignIsPortfolioBiddingStrategy: campaign.campaignIsPortfolioBiddingStrategy,\n            campaignPortfolioBiddingStrategyName: campaign.campaignPortfolioBiddingStrategyName,\n            portfolioStrategyResourceName: campaign.portfolioStrategyResourceName,\n            portfolioOwnerId: campaign.portfolioOwnerId,\n            isLocallyOwnedPortfolio: campaign.isLocallyOwnedPortfolio,\n            campaignPortfolioIsManagerOwned: campaign.campaignPortfolioIsManagerOwned,\n            campaignBiddingStrategyType: campaign.campaignBiddingStrategyType,\n            campaignBiddingTargetRoas: campaign.campaignBiddingTargetRoas,\n            campaignBiddingTargetCpa: campaign.campaignBiddingTargetCpa,\n            campaignStatsYesterdayCost: campaign.campaignStatsYesterdayCost,\n            campaignStats7DaysCost: campaign.campaignStats7DaysCost,\n            campaign30DaysCost: campaign.campaign30DaysCost,\n            campaign30DaysConversionsValue: campaign.campaign30DaysConversionsValue,\n            campaign30DaysConversions: campaign.campaign30DaysConversions,\n            campaign30DaysSearchRankLostImpressionShare: campaign.campaign30DaysSearchRankLostImpressionShare,\n            campaign30DaysSearchBudgetLostImpressionShare: campaign.campaign30DaysSearchBudgetLostImpressionShare\n          });\n        }\n        const typeDetails = extractTypeSpecificStagingDetails(\n          type,\n          recDetails,\n          campaign,\n          sourceCampaigns\n        );\n        stagedRecs.push({\n          mccId: ctx.mccId,\n          accountId: ctx.accountId,\n          recommendationId: recResourceName,\n          recommendationType: type,\n          campaignResourceName: campaign.campaignResourceName,\n          recommendationCurrentBudgetAmount: typeDetails.recommendationCurrentBudgetAmount,\n          recommendationNewBudgetAmount: typeDetails.recommendationNewBudgetAmount,\n          recommendationNewTargetCpa: typeDetails.recommendationNewTargetCpa,\n          recommendationNewTargetRoas: typeDetails.recommendationNewTargetRoas,\n          recommendationBaseCost: impactStats.recommendationBaseCost,\n          recommendationPotentialCost: impactStats.recommendationPotentialCost,\n          recommendationBaseClicks: impactStats.recommendationBaseClicks,\n          recommendationPotentialClicks: impactStats.recommendationPotentialClicks,\n          recommendationBaseConversions: impactStats.recommendationBaseConversions,\n          recommendationPotentialConversions: impactStats.recommendationPotentialConversions,\n          recommendationBaseConversionsValue: impactStats.recommendationBaseConversionsValue,\n          recommendationPotentialConversionsValue: impactStats.recommendationPotentialConversionsValue,\n          recommendationBaseImpressions: impactStats.recommendationBaseImpressions,\n          recommendationPotentialImpressions: impactStats.recommendationPotentialImpressions,\n          recommendationBaseVideoViews: impactStats.recommendationBaseVideoViews,\n          recommendationPotentialVideoViews: impactStats.recommendationPotentialVideoViews,\n          moveBudgetSourceBudgetResourceName: typeDetails.moveBudgetSourceBudgetResourceName,\n          moveBudgetSourceCampaigns: typeDetails.moveBudgetSourceCampaigns,\n          moveBudgetSourceTargetCpa: typeDetails.moveBudgetSourceTargetCpa,\n          moveBudgetSourceTargetRoas: typeDetails.moveBudgetSourceTargetRoas\n        });\n      }\n    }\n    const accounts = stagedRecs.length > 0 ? [\n      {\n        mccId: ctx.mccId,\n        mccName: ctx.mccName,\n        accountId: ctx.accountId,\n        accountName: ctx.accountName,\n        currencyCode: ctx.currencyCode,\n        timestamp: ctx.timestamp\n      }\n    ] : [];\n    return {\n      accounts,\n      budgets: Array.from(budgetsByKey.values()),\n      campaigns: Array.from(campaignsByKey.values()),\n      recommendations: stagedRecs\n    };\n  }\n  function extractTargetCampaignResourceNames(recDetails) {\n    if (Array.isArray(recDetails.campaigns) && recDetails.campaigns.length > 0) {\n      return recDetails.campaigns.map(String);\n    }\n    if (recDetails.campaign) {\n      return [String(recDetails.campaign)];\n    }\n    return [];\n  }\n  function getBudgetRecommendation(recDetails) {\n    return recDetails.campaignBudgetRecommendation ?? recDetails.forecastingCampaignBudgetRecommendation ?? recDetails.marginalRoiCampaignBudgetRecommendation ?? recDetails.moveUnusedBudgetRecommendation?.budgetRecommendation;\n  }\n  function extractNativeImpactMetrics(recDetails) {\n    const budgetSubRec = getBudgetRecommendation(recDetails);\n    let impactPair = null;\n    if (budgetSubRec && Array.isArray(budgetSubRec.budgetOptions)) {\n      const matchedOption = budgetSubRec.budgetOptions.find(\n        (opt) => opt.budgetAmountMicros !== void 0 && opt.budgetAmountMicros !== null && budgetSubRec.recommendedBudgetAmountMicros !== void 0 && budgetSubRec.recommendedBudgetAmountMicros !== null && Number(opt.budgetAmountMicros) === Number(budgetSubRec.recommendedBudgetAmountMicros)\n      );\n      if (matchedOption?.impact) {\n        impactPair = matchedOption.impact;\n      }\n    }\n    if (!impactPair && recDetails.impact) {\n      impactPair = recDetails.impact;\n    }\n    const base = impactPair?.baseMetrics;\n    const potential = impactPair?.potentialMetrics;\n    return {\n      recommendationBaseCost: fromMicros(base?.costMicros),\n      recommendationPotentialCost: fromMicros(potential?.costMicros),\n      recommendationBaseClicks: toNullableNumber(base?.clicks),\n      recommendationPotentialClicks: toNullableNumber(potential?.clicks),\n      recommendationBaseConversions: toNullableNumber(base?.conversions),\n      recommendationPotentialConversions: toNullableNumber(potential?.conversions),\n      recommendationBaseConversionsValue: toNullableNumber(base?.conversionsValue),\n      recommendationPotentialConversionsValue: toNullableNumber(potential?.conversionsValue),\n      recommendationBaseImpressions: toNullableNumber(base?.impressions),\n      recommendationPotentialImpressions: toNullableNumber(potential?.impressions),\n      recommendationBaseVideoViews: toNullableNumber(base?.videoViews),\n      recommendationPotentialVideoViews: toNullableNumber(potential?.videoViews)\n    };\n  }\n  function extractTypeSpecificStagingDetails(type, recDetails, campaign, sourceCampaigns) {\n    let recommendationCurrentBudgetAmount = null;\n    let recommendationNewBudgetAmount = null;\n    let recommendationNewTargetCpa = null;\n    let recommendationNewTargetRoas = null;\n    let moveBudgetSourceBudgetResourceName = "";\n    let moveBudgetSourceCampaigns = "N/A";\n    let moveBudgetSourceTargetCpa = null;\n    let moveBudgetSourceTargetRoas = null;\n    if (type === "CAMPAIGN_BUDGET" || type === "FORECASTING_CAMPAIGN_BUDGET" || type === "MARGINAL_ROI_CAMPAIGN_BUDGET") {\n      const budgetRec = getBudgetRecommendation(recDetails);\n      if (budgetRec) {\n        recommendationCurrentBudgetAmount = fromMicros(budgetRec.currentBudgetAmountMicros);\n        recommendationNewBudgetAmount = fromMicros(budgetRec.recommendedBudgetAmountMicros);\n      }\n    } else if (type === "MOVE_UNUSED_BUDGET") {\n      const moveRec = recDetails.moveUnusedBudgetRecommendation;\n      const budgetRec = getBudgetRecommendation(recDetails);\n      if (budgetRec) {\n        recommendationCurrentBudgetAmount = fromMicros(budgetRec.currentBudgetAmountMicros);\n        recommendationNewBudgetAmount = fromMicros(budgetRec.recommendedBudgetAmountMicros);\n      }\n      moveBudgetSourceBudgetResourceName = String(moveRec?.excessCampaignBudget ?? "");\n      if (sourceCampaigns.length > 0) {\n        moveBudgetSourceCampaigns = sourceCampaigns.map((sc) => sc.campaignName).join(", ");\n        const { targetCpa, targetRoas } = resolveSourceBudgetTargets(sourceCampaigns);\n        moveBudgetSourceTargetCpa = targetCpa;\n        moveBudgetSourceTargetRoas = targetRoas;\n      }\n    } else if (type === "RAISE_TARGET_CPA") {\n      const cpaAdj = recDetails.raiseTargetCpaRecommendation?.targetAdjustment;\n      const mult = toNullableNumber(\n        cpaAdj?.recommendedTargetMultiplier ?? recDetails.raiseTargetCpaRecommendation?.targetCpaMultiplier\n      );\n      const baseCpa = fromMicros(cpaAdj?.currentAverageTargetMicros) ?? campaign.campaignBiddingTargetCpa;\n      if (mult !== null && baseCpa !== null) {\n        recommendationNewTargetCpa = baseCpa * mult;\n      }\n    } else if (type === "LOWER_TARGET_ROAS") {\n      const roasAdj = recDetails.lowerTargetRoasRecommendation?.targetAdjustment;\n      const mult = toNullableNumber(\n        roasAdj?.recommendedTargetMultiplier ?? recDetails.lowerTargetRoasRecommendation?.targetRoasMultiplier\n      );\n      const baseRoas = fromMicros(roasAdj?.currentAverageTargetMicros) ?? campaign.campaignBiddingTargetRoas;\n      if (mult !== null && baseRoas !== null) {\n        recommendationNewTargetRoas = baseRoas * mult;\n      }\n    }\n    return {\n      recommendationCurrentBudgetAmount,\n      recommendationNewBudgetAmount,\n      recommendationNewTargetCpa,\n      recommendationNewTargetRoas,\n      moveBudgetSourceBudgetResourceName,\n      moveBudgetSourceCampaigns,\n      moveBudgetSourceTargetCpa,\n      moveBudgetSourceTargetRoas\n    };\n  }\n  function resolveSourceBudgetTargets(sourceCampaigns) {\n    const cpas = sourceCampaigns.map((c) => c.targetCpa).filter((v) => typeof v === "number" && Number.isFinite(v));\n    const roases = sourceCampaigns.map((c) => c.targetRoas).filter((v) => typeof v === "number" && Number.isFinite(v));\n    const uniqueCpas = Array.from(new Set(cpas));\n    const uniqueRoases = Array.from(new Set(roases));\n    return {\n      targetCpa: uniqueCpas.length === 1 ? uniqueCpas[0] : null,\n      targetRoas: uniqueRoases.length === 1 ? uniqueRoases[0] : null\n    };\n  }\n  function toNullableNumber(val) {\n    if (val === null || val === void 0 || val === "")\n      return null;\n    const num = Number(val);\n    return Number.isFinite(num) ? num : null;\n  }\n\n  // src/ads/orchestrator.ts\n  var MIN_REMAINING_TIME_SECONDS = 240;\n  var FLUSH_EVERY_N_ACCOUNTS = 25;\n  var FLUSH_EVERY_N_RECS = 500;\n  var NOT_A_MANAGER_ACCOUNT_ERROR = "DLG Dashboard: this Google Ads Script must run in a Google Ads manager (MCC) account.";\n  var MAX_STORED_ERROR_LENGTH = 200;\n  function runAdsScriptOrchestrator(options) {\n    const env = createRunEnvironment(options);\n    const lock = readDailyLockForMcc(env);\n    let checkpoint;\n    if (!env.isPreview && lock) {\n      checkpoint = getMccCheckpoint(env.spreadsheet, env.dateStr, env.normMccId);\n      if (checkpoint.status === "COMPLETED" && checkpoint.dateStr === env.dateStr) {\n        Logger.log(\n          `[DLG v2] MCC ${env.mcc.mccId} already marked COMPLETED for ${env.dateStr}. Nothing to do (use "\\u{1F680} DLG Dashboard v2 -> Apply Config Changes Now" in the Google Sheet to re-run).`\n        );\n        return;\n      }\n    }\n    const { childAccounts, accountEntityMap } = discoverChildAccountsUnderCurrentMcc();\n    const ctx = {\n      spreadsheet: env.spreadsheet,\n      mcc: env.mcc,\n      normMccId: env.normMccId,\n      dateStr: env.dateStr,\n      childAccounts\n    };\n    if (env.isPreview || !lock || !checkpoint) {\n      registerDiscovery(ctx, env.isPreview, lock !== null);\n      return;\n    }\n    const selectedCids = lock.lockedCids.filter((cid) => accountEntityMap.has(cid));\n    const progress = startOrResumeRun(ctx, checkpoint, selectedCids.length);\n    const outcome = processSelectedAccounts(\n      ctx,\n      progress,\n      selectedCids,\n      accountEntityMap,\n      lock.config.excludeShoppingAndPmax\n    );\n    if (outcome.status === "PAUSED_FOR_TIME")\n      return;\n    finalizeRun(ctx, progress, selectedCids, outcome.pendingBatch);\n  }\n  function createRunEnvironment(options) {\n    const mccAccount = AdsApp.currentAccount();\n    const mccId = mccAccount.getCustomerId();\n    return {\n      spreadsheet: SpreadsheetApp.openByUrl(options.spreadsheetUrl),\n      mcc: { mccId, mccName: mccAccount.getName() || "MCC" },\n      normMccId: normalizeCid(mccId),\n      dateStr: Utilities.formatDate(/* @__PURE__ */ new Date(), options.timeZone || "UTC", "yyyy-MM-dd"),\n      isPreview: AdsApp.getExecutionInfo().isPreview()\n    };\n  }\n  function readDailyLockForMcc(env) {\n    const config = readLockedDailyConfig(env.spreadsheet, env.dateStr);\n    const lockedCids = config?.selectedCidsByMcc?.[env.normMccId];\n    return config && Array.isArray(lockedCids) ? { config, lockedCids } : null;\n  }\n  function registerDiscovery(ctx, isPreview, isLockedInRunState) {\n    const { mccId, mccName } = ctx.mcc;\n    const msg = `\\u{1F7E2} MCC ${mccId} (${mccName}) registered (${ctx.childAccounts.length} accounts). Review checkboxes & Target Currency in Config; full run will execute on next trigger.`;\n    Logger.log(\n      `[DLG v2] Discovery / waiting for daily lock (isPreview=${isPreview}, isLockedInRunState=${isLockedInRunState}): discovered ${ctx.childAccounts.length} child accounts under MCC ${mccId} (${mccName}).`\n    );\n    writeMccSheetState(ctx.spreadsheet, {\n      mcc: ctx.mcc,\n      childAccounts: ctx.childAccounts,\n      statusMessage: msg\n    });\n  }\n  function startOrResumeRun(ctx, checkpoint, selectedCount) {\n    const progress = {\n      completed: new Set(checkpoint.completedCids),\n      failed: { ...checkpoint.failedCids }\n    };\n    Logger.log(\n      `[DLG v2] MCC ${ctx.mcc.mccId} (${ctx.mcc.mccName}) on ${ctx.dateStr}: ${selectedCount} selected account(s) (of ${ctx.childAccounts.length} discovered). Checkpoint status: ${checkpoint.status} (${checkpoint.completedCids.length} completed, ${Object.keys(progress.failed).length} failed).`\n    );\n    if (checkpoint.completedCids.length === 0 && Object.keys(progress.failed).length === 0) {\n      clearMccFromStaging(ctx.spreadsheet, ctx.normMccId);\n    }\n    return progress;\n  }\n  function processSelectedAccounts(ctx, progress, selectedCids, accountEntityMap, excludeShoppingAndPmax) {\n    let batch = createEmptyStagingBatch();\n    let unFlushedAccountCount = 0;\n    for (const cid of selectedCids) {\n      if (progress.completed.has(cid))\n        continue;\n      const remainingSeconds = Number(AdsApp.getExecutionInfo().getRemainingTime());\n      if (remainingSeconds < MIN_REMAINING_TIME_SECONDS) {\n        saveProgress(\n          ctx,\n          batch,\n          progress,\n          "IN_PROGRESS",\n          `\\u23F3 IN PROGRESS \\u2014 MCC ${ctx.mcc.mccId}: ${progress.completed.size}/${selectedCids.length} accounts processed. Will resume on next hourly trigger.`\n        );\n        return { status: "PAUSED_FOR_TIME" };\n      }\n      processOneAccount(ctx, progress, cid, accountEntityMap.get(cid), excludeShoppingAndPmax, batch);\n      unFlushedAccountCount++;\n      if (unFlushedAccountCount >= FLUSH_EVERY_N_ACCOUNTS || batch.recommendations.length >= FLUSH_EVERY_N_RECS) {\n        saveProgress(\n          ctx,\n          batch,\n          progress,\n          "IN_PROGRESS",\n          `\\u23F3 IN PROGRESS \\u2014 MCC ${ctx.mcc.mccId}: ${progress.completed.size}/${selectedCids.length} accounts processed.`\n        );\n        batch = createEmptyStagingBatch();\n        unFlushedAccountCount = 0;\n      }\n    }\n    return { status: "FINISHED", pendingBatch: batch };\n  }\n  function processOneAccount(ctx, progress, cid, accountEntity, excludeShoppingAndPmax, batch) {\n    try {\n      const accountBatch = processSingleChildAccountToBatch(\n        accountEntity,\n        ctx.mcc.mccId,\n        ctx.mcc.mccName,\n        `${ctx.dateStr} 00:00:00`,\n        excludeShoppingAndPmax\n      );\n      batch.accounts.push(...accountBatch.accounts);\n      batch.budgets.push(...accountBatch.budgets);\n      batch.campaigns.push(...accountBatch.campaigns);\n      batch.recommendations.push(...accountBatch.recommendations);\n      delete progress.failed[cid];\n      progress.completed.add(cid);\n      Logger.log(\n        `[DLG v2] Processed account ${formatCidHyphenated(cid)}: ${accountBatch.recommendations.length} recommendation(s) staged.`\n      );\n    } catch (err) {\n      const errMsg = truncate(err instanceof Error ? err.message : String(err), MAX_STORED_ERROR_LENGTH);\n      const attempts = (progress.failed[cid]?.attempts ?? 0) + 1;\n      progress.failed[cid] = { error: errMsg, attempts };\n      if (attempts >= MAX_ACCOUNT_ATTEMPTS) {\n        progress.completed.add(cid);\n      }\n      Logger.log(\n        `Error processing child account ${cid} under MCC ${ctx.mcc.mccId} (attempt ${attempts}/${MAX_ACCOUNT_ATTEMPTS}): ${errMsg}`\n      );\n    }\n  }\n  function finalizeRun(ctx, progress, selectedCids, pendingBatch) {\n    const retryCount = selectedCids.filter((cid) => !progress.completed.has(cid)).length;\n    if (retryCount > 0) {\n      saveProgress(\n        ctx,\n        pendingBatch,\n        progress,\n        "IN_PROGRESS",\n        `\\u23F3 IN PROGRESS \\u2014 MCC ${ctx.mcc.mccId}: ${progress.completed.size}/${selectedCids.length} accounts processed; retrying ${retryCount} failed account(s) on the next hourly trigger.`\n      );\n      return;\n    }\n    const givenUpCount = Object.keys(progress.failed).length;\n    saveProgress(\n      ctx,\n      pendingBatch,\n      progress,\n      "COMPLETED",\n      `\\u23F3 STAGED FOR PUBLISH (${ctx.dateStr}) \\u2014 MCC ${ctx.mcc.mccId} completed (${progress.completed.size} accounts${givenUpCount > 0 ? `, ${givenUpCount} failed` : ""}).`\n    );\n  }\n  function truncate(text, maxLength) {\n    return text.length > maxLength ? `${text.slice(0, maxLength - 1)}\\u2026` : text;\n  }\n  function discoverChildAccountsUnderCurrentMcc() {\n    const childAccounts = [];\n    const accountEntityMap = /* @__PURE__ */ new Map();\n    if (typeof AdsManagerApp === "undefined") {\n      throw new Error(NOT_A_MANAGER_ACCOUNT_ERROR);\n    }\n    const iterator = AdsManagerApp.accounts().get();\n    while (iterator.hasNext()) {\n      const acc = iterator.next();\n      const rawCid = String(acc.getCustomerId() ?? "");\n      const normCid = normalizeCid(rawCid);\n      if (!normCid)\n        continue;\n      accountEntityMap.set(normCid, acc);\n      childAccounts.push({\n        customerId: rawCid,\n        name: String(acc.getName() || "Account"),\n        currencyCode: String(acc.getCurrencyCode() || "USD")\n      });\n    }\n    return { childAccounts, accountEntityMap };\n  }\n  function processSingleChildAccountToBatch(accountEntity, mccId, mccName, timestamp, excludeShoppingAndPmax) {\n    AdsManagerApp.select(accountEntity);\n    const accountId = accountEntity.getCustomerId();\n    const accountName = accountEntity.getName() || "Account";\n    const currencyCode = accountEntity.getCurrencyCode() || "USD";\n    const recommendationDetailsMap = fetchAllRecommendationDetails();\n    if (recommendationDetailsMap.size === 0) {\n      return createEmptyStagingBatch();\n    }\n    const campaignResourceNames = /* @__PURE__ */ new Set();\n    const budgetResourceNames = /* @__PURE__ */ new Set();\n    for (const details of recommendationDetailsMap.values()) {\n      for (const cName of extractTargetCampaignResourceNames(details)) {\n        campaignResourceNames.add(cName);\n      }\n      const excessBudget = details.moveUnusedBudgetRecommendation?.excessCampaignBudget;\n      if (excessBudget) {\n        budgetResourceNames.add(String(excessBudget));\n      }\n    }\n    const campaignDataMap = fetchAllCampaignData(\n      Array.from(campaignResourceNames),\n      excludeShoppingAndPmax\n    );\n    const budgetDataMap = fetchAllBudgetDetails(Array.from(budgetResourceNames));\n    const budgetToCampaignsMap = fetchCampaignsByBudgets(\n      Array.from(budgetResourceNames),\n      excludeShoppingAndPmax\n    );\n    return buildAccountStagingBatch({\n      mccId,\n      mccName,\n      accountId,\n      accountName,\n      currencyCode,\n      timestamp,\n      excludeShoppingAndPmax,\n      recommendationDetailsMap,\n      campaignDataMap,\n      budgetDataMap,\n      budgetToCampaignsMap\n    });\n  }\n  function saveProgress(ctx, batch, progress, status, statusMessage) {\n    appendStagingBatch(ctx.spreadsheet, batch, ctx.normMccId);\n    const lastCompletedDate = status === "COMPLETED" ? ctx.dateStr : void 0;\n    const checkpoint = {\n      mccId: ctx.normMccId,\n      dateStr: ctx.dateStr,\n      status,\n      completedCids: Array.from(progress.completed),\n      ...Object.keys(progress.failed).length > 0 ? { failedCids: progress.failed } : {},\n      ...lastCompletedDate ? { lastCompletedDate } : {}\n    };\n    writeMccSheetState(ctx.spreadsheet, {\n      mcc: ctx.mcc,\n      childAccounts: ctx.childAccounts,\n      statusMessage,\n      checkpoint,\n      ...lastCompletedDate ? { lastCompletedDate } : {}\n    });\n  }\n\n  // src/ads/main.ts\n  function main() {\n    const url = typeof SPREADSHEET_URL !== "undefined" ? SPREADSHEET_URL : "";\n    if (!url || url === "YOUR_SPREADSHEET_URL_HERE") {\n      throw new Error(\n        \'SPREADSHEET_URL is not configured. Use "\\u{1F4CB} Copy Google Ads Script" in your copy of the Master Template Google Sheet.\'\n      );\n    }\n    const timeZone = typeof SPREADSHEET_TIMEZONE !== "undefined" ? SPREADSHEET_TIMEZONE : "";\n    if (!timeZone || timeZone === "YOUR_SPREADSHEET_TIMEZONE_HERE") {\n      throw new Error(\n        \'SPREADSHEET_TIMEZONE is not configured. Use "\\u{1F4CB} Copy Google Ads Script" in your copy of the Master Template Google Sheet.\'\n      );\n    }\n    runAdsScriptOrchestrator({ spreadsheetUrl: url, timeZone });\n  }\n  return __toCommonJS(main_exports);\n})();\nfunction main() { return DlgAdsBundle.main(); }\n' : "";
  var SPREADSHEET_URL_PLACEHOLDER = 'var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";';
  var SPREADSHEET_TIMEZONE_PLACEHOLDER = 'var SPREADSHEET_TIMEZONE = "YOUR_SPREADSHEET_TIMEZONE_HERE";';
  var DATA_STUDIO_TEMPLATE_REPORT_ID = "8efcb962-4ce8-4df1-bca0-bd354d946b2d";
  var DEMO_TOOLS_PROPERTY = "DLG_ENABLE_DEMO_TOOLS";
  var PUBLISHER_HEARTBEAT_PROPERTY = "DLG_PUBLISHER_HEARTBEAT_MS";
  var PUBLISHER_HEARTBEAT_STALE_MS = 45 * 60 * 1e3;
  function generateAdsScriptForCurrentSheet(spreadsheetUrl, timeZone = "UTC", template = EMBEDDED_ADS_SCRIPT_TEMPLATE || globalThis.__EMBEDDED_ADS_SCRIPT__ || "") {
    if (!template.includes(SPREADSHEET_URL_PLACEHOLDER)) {
      throw new Error(
        "Embedded Ads Script template is missing the SPREADSHEET_URL placeholder."
      );
    }
    if (!template.includes(SPREADSHEET_TIMEZONE_PLACEHOLDER)) {
      throw new Error(
        "Embedded Ads Script template is missing the SPREADSHEET_TIMEZONE placeholder."
      );
    }
    return template.replace(
      SPREADSHEET_URL_PLACEHOLDER,
      `var SPREADSHEET_URL = ${JSON.stringify(spreadsheetUrl)};`
    ).replace(
      SPREADSHEET_TIMEZONE_PLACEHOLDER,
      `var SPREADSHEET_TIMEZONE = ${JSON.stringify(timeZone || "UTC")};`
    );
  }
  function buildDataStudioCreationUrlForActiveSheet(spreadsheet) {
    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    const gid = dataSheet?.getSheetId() ?? 100;
    const params = [
      "reporting/create?",
      "c.mode=edit",
      `&c.reportId=${DATA_STUDIO_TEMPLATE_REPORT_ID}`,
      "&r.reportName=DLG_Dashboard_v2",
      "&ds.ds0.connector=googleSheets",
      `&ds.ds0.spreadsheetId=${encodeURIComponent(spreadsheet.getId())}`,
      `&ds.ds0.worksheetId=${encodeURIComponent(String(gid))}`,
      "&ds.ds0.includeHiddenCells=true",
      "&ds.ds0.includeFilteredCells=true",
      "&ds.ds0.refreshFields=false"
    ].join("");
    return `https://datastudio.google.com/${params}`;
  }
  function onOpen() {
    if (typeof SpreadsheetApp === "undefined" || typeof SpreadsheetApp.getUi !== "function") {
      return;
    }
    const ui = SpreadsheetApp.getUi();
    if (!ui)
      return;
    const menu = ui.createMenu("\u{1F680} DLG Dashboard v2").addItem("\u{1F4CB} Copy Google Ads Script", "showCopyAdsScriptModal").addItem("\u{1F4CA} Create New Data Studio Dashboard", "openDataStudioCreator").addItem("\u{1F680} Open Saved Dashboard", "openSavedDashboard").addSeparator().addItem("\u{1F504} Apply Config Changes Now (Reset Today's Lock)", "applyConfigChangesNow");
    if (isDemoToolsEnabled()) {
      menu.addSeparator().addItem("\u{1F9EA} Seed Synthetic Demo Data (For Testing)", "seedSyntheticDemoDataFromMenu");
    }
    menu.addToUi();
  }
  function isDemoToolsEnabled() {
    return typeof PropertiesService !== "undefined" && PropertiesService.getScriptProperties().getProperty(DEMO_TOOLS_PROPERTY) === "true";
  }
  function seedSyntheticDemoDataFromMenu() {
    if (!isDemoToolsEnabled()) {
      throw new Error(
        `Demo tools are disabled. Set the Script Property ${DEMO_TOOLS_PROPERTY}=true to enable them.`
      );
    }
    const ui = SpreadsheetApp.getUi();
    const answer = ui.alert(
      "Seed demo data?",
      "Adds 2 demo MCCs (using 2 of the 8 MCC slots) and sample recommendations to Config and Data. Real MCCs are not reset.",
      ui.ButtonSet.YES_NO
    );
    if (answer === ui.Button.YES) {
      seedSyntheticDemoData();
    }
  }
  function onEdit(e) {
    const range = e?.range;
    if (!range)
      return;
    const colStart = range.columnStart ?? range.getColumn();
    if (colStart !== 1)
      return;
    const startRow = range.rowStart ?? range.getRow();
    const endRow = range.rowEnd ?? startRow + (range.getNumRows?.() ?? 1) - 1;
    if (endRow < GLOBAL_SELECT_ALL_ROW || startRow > MCC_TABLE_END_ROW && endRow < CID_TABLE_START_ROW) {
      return;
    }
    const sheet = range.getSheet();
    if (sheet.getName() !== CONFIG_SHEET_NAME) {
      return;
    }
    const lastRow = Math.max(sheet.getLastRow(), MCC_TABLE_END_ROW);
    const blockRowsCount = lastRow - GLOBAL_SELECT_ALL_ROW + 1;
    const blockWidth = sheetCol(CID_COL.PARENT_MCC_ID);
    const block = sheet.getRange(GLOBAL_SELECT_ALL_ROW, 1, blockRowsCount, blockWidth).getValues();
    const getBlockRow = (sheetRow) => block[sheetRow - GLOBAL_SELECT_ALL_ROW] ?? new Array(blockWidth).fill("");
    const mccCount = MCC_TABLE_END_ROW - MCC_TABLE_START_ROW + 1;
    const hasCids = lastRow >= CID_TABLE_START_ROW;
    const cidCount = hasCids ? lastRow - CID_TABLE_START_ROW + 1 : 0;
    if (startRow <= GLOBAL_SELECT_ALL_ROW && endRow >= GLOBAL_SELECT_ALL_ROW) {
      const isChecked = e?.value !== void 0 ? parseSheetBoolean(e.value) : parseSheetBoolean(getBlockRow(GLOBAL_SELECT_ALL_ROW)[0]);
      const nextMccColA = [];
      let mccChanged = false;
      for (let r = MCC_TABLE_START_ROW; r <= MCC_TABLE_END_ROW; r++) {
        const row = getBlockRow(r);
        const hasMcc = Boolean(normalizeCid(row[MCC_COL.MCC_ID]));
        const targetVal = hasMcc ? isChecked : false;
        if (parseSheetBoolean(row[MCC_COL.SELECT_ALL]) !== targetVal || row[MCC_COL.SELECT_ALL] !== false && !hasMcc) {
          mccChanged = true;
        }
        nextMccColA.push([targetVal]);
      }
      if (mccChanged) {
        sheet.getRange(MCC_TABLE_START_ROW, 1, mccCount, 1).setValues(nextMccColA);
      }
      if (hasCids) {
        const nextCidColA = [];
        let cidChanged = false;
        for (let r = CID_TABLE_START_ROW; r <= lastRow; r++) {
          const row = getBlockRow(r);
          const hasCid = Boolean(normalizeCid(row[CID_COL.CID]));
          const targetVal = hasCid ? isChecked : false;
          if (parseSheetBoolean(row[CID_COL.INCLUDE]) !== targetVal || row[CID_COL.INCLUDE] !== false && !hasCid) {
            cidChanged = true;
          }
          nextCidColA.push([targetVal]);
        }
        if (cidChanged) {
          sheet.getRange(CID_TABLE_START_ROW, 1, cidCount, 1).setValues(nextCidColA);
        }
      }
      return;
    }
    if (startRow >= MCC_TABLE_START_ROW && startRow <= MCC_TABLE_END_ROW) {
      const mccEnd = Math.min(endRow, MCC_TABLE_END_ROW);
      const toggledMccStates = /* @__PURE__ */ new Map();
      for (let r = startRow; r <= mccEnd; r++) {
        const row = getBlockRow(r);
        const mccId = normalizeCid(row[MCC_COL.MCC_ID]);
        if (mccId) {
          toggledMccStates.set(mccId, parseSheetBoolean(row[MCC_COL.SELECT_ALL]));
        }
      }
      if (hasCids && toggledMccStates.size > 0) {
        const nextCidColA = [];
        let cidChanged = false;
        let totalValidCids = 0;
        let checkedValidCids = 0;
        for (let r = CID_TABLE_START_ROW; r <= lastRow; r++) {
          const row = getBlockRow(r);
          const rowCid = normalizeCid(row[CID_COL.CID]);
          if (!rowCid) {
            if (row[CID_COL.INCLUDE] !== false)
              cidChanged = true;
            nextCidColA.push([false]);
            continue;
          }
          totalValidCids++;
          const rowMcc = normalizeCid(row[CID_COL.PARENT_MCC_ID]);
          const currentChecked = parseSheetBoolean(row[CID_COL.INCLUDE]);
          const targetChecked = toggledMccStates.has(rowMcc) ? toggledMccStates.get(rowMcc) : currentChecked;
          if (currentChecked !== targetChecked) {
            cidChanged = true;
          }
          if (targetChecked) {
            checkedValidCids++;
          }
          nextCidColA.push([targetChecked]);
        }
        if (cidChanged) {
          sheet.getRange(CID_TABLE_START_ROW, 1, cidCount, 1).setValues(nextCidColA);
        }
        if (totalValidCids > 0) {
          const nextGlobal = checkedValidCids === totalValidCids;
          if (parseSheetBoolean(getBlockRow(GLOBAL_SELECT_ALL_ROW)[0]) !== nextGlobal) {
            sheet.getRange(GLOBAL_SELECT_ALL_CELL).setValue(nextGlobal);
          }
        }
      }
      return;
    }
    if (endRow >= CID_TABLE_START_ROW && hasCids) {
      const cidStart = Math.max(startRow, CID_TABLE_START_ROW);
      const cidEnd = Math.min(endRow, lastRow);
      const affectedMccs = /* @__PURE__ */ new Set();
      for (let r = cidStart; r <= cidEnd; r++) {
        const parentMccId = normalizeCid(getBlockRow(r)[3]);
        if (parentMccId) {
          affectedMccs.add(parentMccId);
        }
      }
      const mccTotals = /* @__PURE__ */ new Map();
      let totalValidCids = 0;
      let checkedValidCids = 0;
      for (let r = CID_TABLE_START_ROW; r <= lastRow; r++) {
        const row = getBlockRow(r);
        const cid = normalizeCid(row[CID_COL.CID]);
        if (!cid)
          continue;
        totalValidCids++;
        const isRowChecked = parseSheetBoolean(row[CID_COL.INCLUDE]);
        if (isRowChecked)
          checkedValidCids++;
        const mcc = normalizeCid(row[CID_COL.PARENT_MCC_ID]);
        if (mcc && affectedMccs.has(mcc)) {
          const entry = mccTotals.get(mcc) ?? { total: 0, checked: 0 };
          entry.total++;
          if (isRowChecked)
            entry.checked++;
          mccTotals.set(mcc, entry);
        }
      }
      for (let r = MCC_TABLE_START_ROW; r <= MCC_TABLE_END_ROW; r++) {
        const mccRow = getBlockRow(r);
        const mccId = normalizeCid(mccRow[MCC_COL.MCC_ID]);
        if (!mccId || !affectedMccs.has(mccId))
          continue;
        const stats = mccTotals.get(mccId);
        const allChecked = Boolean(stats && stats.total > 0 && stats.checked === stats.total);
        if (parseSheetBoolean(mccRow[MCC_COL.SELECT_ALL]) !== allChecked) {
          sheet.getRange(r, 1).setValue(allChecked);
        }
      }
      if (totalValidCids > 0) {
        const nextGlobal = checkedValidCids === totalValidCids;
        if (parseSheetBoolean(getBlockRow(GLOBAL_SELECT_ALL_ROW)[0]) !== nextGlobal) {
          sheet.getRange(GLOBAL_SELECT_ALL_CELL).setValue(nextGlobal);
        }
      }
    }
  }
  function ensureAutoPublishTriggerInstalled() {
    if (typeof ScriptApp === "undefined" || typeof ScriptApp.getProjectTriggers !== "function" || typeof ScriptApp.newTrigger !== "function") {
      return;
    }
    const existing = ScriptApp.getProjectTriggers();
    if (existing.some((t) => t.getHandlerFunction() === "runScheduledPublish")) {
      return;
    }
    const lastHeartbeat = readPublisherHeartbeat();
    if (lastHeartbeat !== null && Date.now() - lastHeartbeat < PUBLISHER_HEARTBEAT_STALE_MS) {
      return;
    }
    ScriptApp.newTrigger("runScheduledPublish").timeBased().everyMinutes(15).create();
    recordPublisherHeartbeat();
  }
  function runScheduledPublish(options) {
    recordPublisherHeartbeat();
    publishFromStagingToData(options);
  }
  function readPublisherHeartbeat() {
    if (typeof PropertiesService === "undefined")
      return null;
    const raw = PropertiesService.getDocumentProperties().getProperty(PUBLISHER_HEARTBEAT_PROPERTY);
    const ms = Number(raw);
    return raw && Number.isFinite(ms) ? ms : null;
  }
  function recordPublisherHeartbeat() {
    if (typeof PropertiesService === "undefined")
      return;
    PropertiesService.getDocumentProperties().setProperty(
      PUBLISHER_HEARTBEAT_PROPERTY,
      String(Date.now())
    );
  }
  function showCopyAdsScriptModal() {
    ensureAutoPublishTriggerInstalled();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scriptCode = generateAdsScriptForCurrentSheet(
      ss.getUrl(),
      ss.getSpreadsheetTimeZone() || "UTC"
    );
    const escaped = escapeHtml(scriptCode);
    const html = `
    <div style="font-family: Arial, sans-serif; padding: 4px 12px 12px; box-sizing: border-box;">
      <p style="margin-top:0; margin-bottom:6px;"><b>Setup Steps (repeat for each MCC):</b></p>
      <ol style="margin-top:0; margin-bottom:12px; padding-left:20px; font-size:12px; line-height:1.45;">
        <li>Click <b>\u{1F4CB} Copy to Clipboard</b> below.</li>
        <li>In Google Ads (<b>Tools &amp; settings \u2192 Bulk actions \u2192 Scripts</b>), create a new script in your MCC, paste the code, click <b>Save</b>, then <b>Authorize</b> and <b>Preview</b> once to register its accounts on the <b>Config</b> tab.</li>
        <li>Set the script <b>Frequency = Hourly</b> (required for automatic daily updates and 30-min timeout checkpoint recovery). <i>This Sheet's 15-minute background publisher trigger is installed automatically.</i></li>
        <li>Return to the <b>Config</b> tab in this Sheet and check the accounts you want to include.</li>
      </ol>
      <div style="margin-bottom:8px;">
        <button id="copyBtn" onclick="copyAdsScript(this)" style="padding:8px 16px; background:#1a73e8; color:#fff; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">\u{1F4CB} Copy to Clipboard</button>
      </div>
      <textarea id="code" readonly onclick="this.select();" style="width:100%; height:200px; box-sizing:border-box; font-family:monospace; font-size:11px;">${escaped}</textarea>
      <script>
        (function() {
          var el = document.getElementById('code');
          if (el) { el.focus(); el.select(); }
        })();
        function copyAdsScript(btn) {
          var el = document.getElementById('code');
          el.focus();
          el.select();
          function fallbackCopy() {
            try {
              if (document.execCommand && document.execCommand('copy')) {
                btn.innerText = '\u2705 Copied!';
                return;
              }
            } catch (e) {}
            btn.innerText = '\u26A0\uFE0F Select & press Ctrl+C / Cmd+C';
          }
          if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(el.value).then(function() {
              btn.innerText = '\u2705 Copied!';
            }).catch(fallbackCopy);
          } else {
            fallbackCopy();
          }
        }
      <\/script>
    </div>
  `;
    if (typeof HtmlService !== "undefined" && typeof HtmlService.createHtmlOutput === "function") {
      const output = HtmlService.createHtmlOutput(html).setWidth(680).setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(output, "\u{1F4CB} Copy DLG v2 Google Ads Script");
    }
  }
  function openDataStudioCreator() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureDataSheetInitialized(ss);
    const url = buildDataStudioCreationUrlForActiveSheet(ss);
    openUrlModal(url, "\u{1F4CA} Create New Data Studio Dashboard");
  }
  function ensureDataSheetInitialized(ss) {
    withSheetLock(ss, "SHEET_INIT_DATA", () => {
      const dataSheet = getOrCreateSheet(ss, DATA_SHEET_NAME);
      if (dataSheet.getLastRow() === 0) {
        writeReportRowsToSheet(dataSheet, [], ss);
      }
    });
  }
  function openSavedDashboard() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = ss.getSheetByName(SETUP_SHEET_NAME);
    const savedUrl = String(setupSheet?.getRange(SAVED_DASHBOARD_URL_CELL)?.getValue() ?? "").trim();
    if (!isAllowedDashboardUrl(savedUrl)) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "No Saved Dashboard URL Found",
        `Paste your saved Data Studio report URL (starting with https://lookerstudio.google.com/ or https://datastudio.google.com/) into cell ${SAVED_DASHBOARD_URL_CELL} on the ${SETUP_SHEET_NAME} tab first.`,
        ui.ButtonSet.OK
      );
      return;
    }
    openUrlModal(savedUrl, "\u{1F680} Open Saved Data Studio Dashboard");
  }
  function applyConfigChangesNow() {
    ensureAutoPublishTriggerInstalled();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dateStr = resolveSpreadsheetDate(ss);
    withSheetLock(ss, "SHEET_RESET", () => {
      const configSheet = getOrCreateSheet(ss, CONFIG_SHEET_NAME, { index: 0 });
      syncAllMccSheetsToConfig(ss, configSheet);
      const lockedConfig = readLockedDailyConfig(ss, dateStr);
      const liveConfig = readGlobalConfig(configSheet);
      const shoppingFilterChanged = lockedConfig !== null && Boolean(lockedConfig.excludeShoppingAndPmax) !== Boolean(liveConfig.excludeShoppingAndPmax);
      const resetMccIds = [];
      for (const mcc of getMccSelectionSummary(configSheet).filter((m) => m.isActive)) {
        const checkpoint = getMccCheckpoint(ss, dateStr, mcc.mccId);
        const hasProgressToday = checkpoint.dateStr === dateStr && (checkpoint.status === "COMPLETED" || checkpoint.completedCids.length > 0);
        if (!hasProgressToday) {
          continue;
        }
        const completed = new Set(checkpoint.completedCids);
        const hasNewlyCheckedCids = checkpoint.status === "COMPLETED" && getSelectedCidsForMcc(configSheet, mcc.mccId).some((cid) => !completed.has(cid));
        if (shoppingFilterChanged || hasNewlyCheckedCids) {
          resetMccIds.push(mcc.mccId);
        }
      }
      resetTodayPublisherState(ss, dateStr);
      for (const mccId of resetMccIds) {
        clearMccSheetCheckpoint(ss, mccId);
      }
      const resetNote = resetMccIds.length > 0 ? ` Re-collecting ${resetMccIds.length} MCC(s) on their next hourly Google Ads Script run${shoppingFilterChanged ? " (Exclude Shopping & Feed-PMax changed)" : " (newly checked accounts)"}.` : " Republishing from the data already collected today.";
      configSheet.getRange(STATUS_BANNER_CELL).setValue(`\u{1F504} Config changes applied (${dateStr}).${resetNote}`);
    });
    publishFromStagingToData({
      spreadsheet: ss,
      force: true
    });
  }
  var DASHBOARD_URL_RE = /^https:\/\/(lookerstudio|datastudio)\.google\.com\/[^\s"'<>\\]*$/i;
  function isAllowedDashboardUrl(url) {
    return DASHBOARD_URL_RE.test(url);
  }
  function escapeHtml(raw) {
    return raw.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function openUrlModal(url, title) {
    if (!isAllowedDashboardUrl(url)) {
      throw new Error("Only Data Studio report links (lookerstudio.google.com / datastudio.google.com) can be opened.");
    }
    const safeHref = escapeHtml(url);
    const safeTitle = escapeHtml(title);
    const html = `
    <div style="font-family: Arial, sans-serif; padding: 16px;">
      <p>Click below if the tab did not open automatically:</p>
      <p><a id="target" href="${safeHref}" target="_blank" rel="noopener" style="font-weight:bold; color:#1a73e8;">${safeTitle}</a></p>
      <script>window.open(document.getElementById('target').href, '_blank');<\/script>
    </div>
  `;
    if (typeof HtmlService !== "undefined" && typeof HtmlService.createHtmlOutput === "function") {
      const output = HtmlService.createHtmlOutput(html).setWidth(480).setHeight(180);
      SpreadsheetApp.getUi().showModalDialog(output, title);
    }
  }
  return __toCommonJS(sheet_ui_exports);
})();
var EMBEDDED_ADS_SCRIPT_TEMPLATE = DlgSheetBundle.EMBEDDED_ADS_SCRIPT_TEMPLATE;
function onOpen() { return DlgSheetBundle.onOpen(); }
function onEdit(e) { return DlgSheetBundle.onEdit(e); }
function showCopyAdsScriptModal() { return DlgSheetBundle.showCopyAdsScriptModal(); }
function openDataStudioCreator() { return DlgSheetBundle.openDataStudioCreator(); }
function openSavedDashboard() { return DlgSheetBundle.openSavedDashboard(); }
function applyConfigChangesNow() { return DlgSheetBundle.applyConfigChangesNow(); }
function runScheduledPublish() { return DlgSheetBundle.runScheduledPublish(); }
function seedSyntheticDemoDataFromMenu() { return DlgSheetBundle.seedSyntheticDemoDataFromMenu(); }
