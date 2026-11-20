# Development guide

This guide shows how to make the most common kinds of change to the DLG Dashboard. It assumes you've set up the project as described in [CONTRIBUTING.md](../CONTRIBUTING.md) and read [architecture.md](architecture.md), which explains how the Google Ads script and the Google Sheet's script work together.

## Contents

* [Where things live](#where-things-live)
* [Write tests](#write-tests)
* [Add a column to the Data tab](#add-a-column-to-the-data-tab)
* [Collect a new field from Google Ads](#collect-a-new-field-from-google-ads)
* [Add a recommendation type](#add-a-recommendation-type)
* [Add a setting to the Config tab](#add-a-setting-to-the-config-tab)
* [Change the template Google Sheet layout](#change-the-template-google-sheet-layout)
* [Add a menu item](#add-a-menu-item)
* [Use a new Apps Script or Google Ads service](#use-a-new-apps-script-or-google-ads-service)
* [Change the data the two scripts exchange](#change-the-data-the-two-scripts-exchange)
* [Use demo data](#use-demo-data)

## Where things live

| You want to change… | Look in |
|---|---|
| The GAQL queries | `src/ads/gaql/queries.ts` |
| How query results become campaign and budget data | `src/ads/gaql/fetcher.ts` |
| How recommendations become staged rows (forecasts, targets, Shopping/PMax filter) | `src/ads/processors/recommendations.ts` |
| The Google Ads script's daily run (discovery, checkpoints, retries) | `src/ads/orchestrator.ts` |
| The `Data` tab's columns, formats, and order | `src/shared/schema.ts` and `ReportRow` in `src/shared/types.ts` |
| The hidden staging tabs (`_Stage_*`) | `src/shared/sheets/staging-writer.ts` and the `Staged*` types in `src/shared/types.ts` |
| How staged rows are joined and derived metrics are calculated | `src/appscript/processors/joiner.ts` |
| Portfolio and shared-budget consolidation | `src/appscript/processors/consolidator.ts` and `entity-keys.ts` |
| `recommendationNumber` | `src/appscript/processors/ranking.ts` |
| Google Ads links | `src/appscript/processors/ads-ui-urls.ts` |
| Currency conversion | `src/appscript/fx-rates.ts` |
| Publishing to the `Data` tab | `src/appscript/publisher.ts` |
| The `Setup` and `Config` tabs | `src/appscript/config-sheet.ts` |
| The menu, dialogs, checkboxes, and triggers | `src/appscript/sheet-ui.ts` |
| The JSON both scripts exchange through the Sheet | `src/shared/contract.ts` |
| The build and the Sheet script's permissions | `scripts/build.mjs` |

## Write tests

Every change needs a test that fails without it. Run one file while you work with `npx vitest run test/<file>.test.ts`, and the whole suite with `npm test` before you send a pull request.

**Pure functions** (query builders, the joiner, consolidation, ranking, URL builders) need no mocks. Call them with literal inputs and compare the result with literal expected values.

**Code that reads or writes the Sheet, or calls Google Ads,** uses the fakes in `test/mocks/google-apps-script.mock.ts`:

```ts
import {afterEach, expect, it, vi} from 'vitest';
import {buildMockAccount, setTestDate} from './helpers/factories';
import {
  createMockAdsApp,
  createMockSpreadsheetApp,
  installAppsScriptGlobals,
} from './mocks/google-apps-script.mock';

afterEach(() => {
  vi.useRealTimers();
});

it('…', () => {
  setTestDate('2026-10-14');
  const spreadsheetApp = createMockSpreadsheetApp('https://docs.google.com/spreadsheets/d/test-sheet/edit');
  const adsApp = createMockAdsApp({
    mccId: '111-222-3333',
    childAccounts: [buildMockAccount({customerId: '100-000-0001', currencyCode: 'GBP'})],
    searchHandler: (query, accountId) => [/* GAQL result rows for this query and account */],
  });
  installAppsScriptGlobals({spreadsheetApp, adsApp, lockService: null});

  // Act: call the code under test, for example runAdsScriptOrchestrator(…) or publishFromStagingToData(…).
  // Assert: read the mock spreadsheet's tabs, for example with readReportRowsFromSheet(…).
});
```

* `installAppsScriptGlobals` replaces every global on each call, so tests never share state. Pass `lockService: null` to run code the way the Google Ads script runs it (it has no document lock); leave it out for Sheet-script code.
* `createMockAdsApp` options simulate the situations the Google Ads script must handle: `remainingTime` (seconds left before the 30-minute limit), `isPreview`, and `searchHandler`, which returns GAQL rows per query and account and can throw to simulate an account error.
* `setTestDate` freezes the clock at a date (and optional UTC hour); always restore real timers afterwards.
* If your code calls an Apps Script method the fakes don't have, add it to the mock file.

## Add a column to the Data tab

> [!IMPORTANT]
> **Add columns at the end only.** Data Studio matches the Sheet's columns by position, so inserting, reordering, or renaming a column breaks every dashboard built on the Sheet.

1. **Define the column.** Add it at the **end** of `REPORT_SCHEMA_TEMPLATE` in `src/shared/schema.ts`, with a `defaultValue`, a Sheets number `format`, and a `dummyValueForFormatting` of the right type. The schema dummy row in row 2 uses that value so Data Studio detects the column's type even on days when no row has a value.
2. **Add the field** to `ReportRow` in `src/shared/types.ts`.
3. **Fill it in.**
   * A value that comes straight from staged data: set it in `joinStagingBatchToReportRows` in `joiner.ts`.
   * A value calculated from other columns: set it in `populateDerivedRowMetrics` in `joiner.ts`. Consolidation calls that function again after it merges rows, so the value stays correct on portfolio and shared-budget rows.
4. **Decide how consolidation treats it.** `mergeGroupRows` in `consolidator.ts` keeps the first member's value unless you add a rule. Historical campaign figures are summed across member campaigns; budget and forecast values are kept once, because Google Ads already reports them for the whole portfolio or shared budget.
5. **Money columns:** values stay in the account's currency. If the column needs a converted copy, add a separate `…Global` column and convert it in `syncGlobalColumnsWithRate` in `fx-rates.ts`.
6. **Update the tests.** `test/schema.test.ts` checks the column count and which column is last; update both. Leave the list of original columns alone: it guards their positions. Add a test for how the value is calculated.
7. **Update the docs.** Add the column to the `Data` tab reference in [data-reference.md](data-reference.md), and add a CHANGELOG line.
8. **Update the template dashboard.** **📊 Create New Data Studio Dashboard** copies the template dashboard's field list, so new reports only see the column after a maintainer refreshes the template's data source (**Resource → Manage added data sources → Edit → Refresh fields → Apply**). Existing reports need the same refresh.

Older `Data` tabs are read by header name, so rows published before the change simply get the column's default value.

## Collect a new field from Google Ads

A field travels from a GAQL query, through a staging tab, into the `Data` tab. For a campaign field:

1. **Query it.** Add the field to the matching query in `src/ads/gaql/queries.ts`, for example `buildCampaignMetadataQuery`. Check that the field can be selected together with the others in the [Google Ads API query builder](https://developers.google.com/google-ads/api/fields/latest/overview_query_builder).
2. **Read it.** In `src/ads/gaql/fetcher.ts`, add the field to the response type (for example `GaqlCampaignProto`) and to `EnrichedCampaignData`, and set it in `fetchAllCampaignData`. Convert money from micros with `fromMicros`, and use `null`, not `0`, for missing values.
3. **Stage it.** Add the field to `StagedCampaign` in `src/shared/types.ts` and to `STAGED_CAMPAIGN_SCHEMA` in `src/shared/sheets/staging-writer.ts`, with its `FieldType`. TypeScript reports an error until both match. Copy it into the staged campaign in `buildAccountStagingBatch` in `src/ads/processors/recommendations.ts`.
4. **Publish it.** Follow [Add a column to the Data tab](#add-a-column-to-the-data-tab), setting the value from the staged campaign in the joiner.

Staging tabs are read by header name, so tabs written before the change still load; the new field gets its default value until the next run rewrites them.

This changes the Google Ads script (`dist/ads/Code.gs`). Each MCC keeps running the copy that was pasted into it, so the new field only appears once the script is copied again from the Sheet's menu.

## Add a recommendation type

1. **Allow the type.** Add it to `RecommendationType` in `src/shared/types.ts` and to `SUPPORTED_RECOMMENDATION_TYPES` in `src/ads/gaql/queries.ts`.
2. **Query its details.** Add its sub-message (for example `recommendation.<type>_recommendation`) to `buildRecommendationDetailsQuery`, and its shape to `GaqlRecommendationRow` in `src/ads/gaql/fetcher.ts`.
3. **Extract its values** in `src/ads/processors/recommendations.ts`:
   * If it's a budget recommendation with budget options, add it to `getBudgetRecommendation`. The forecast is then taken from the option that matches the recommended budget.
   * Otherwise, the forecast comes from `recommendation.impact`.
   * Add a branch to `extractTypeSpecificStagingDetails` for its current and recommended values.
4. **Decide its scope** in `consolidateRowsByScope` in `consolidator.ts`: should it merge into one row per portfolio, per shared budget, or stay one row per campaign? Ranking needs no change: `recommendationNumber` orders types alphabetically.
5. **Test it.** Extend the type test in `test/gaql.test.ts` and add extraction and consolidation tests to `test/processors.test.ts`.
6. **Add it to the demo data** in `src/appscript/demo-seed.ts`, so the dashboard can be checked without live recommendations.
7. **Update the docs and the dashboard:** the recommendation types in [data-reference.md](data-reference.md), the CHANGELOG, and any charts or filters in the template dashboard that list types.

This changes the Google Ads script, so it reaches each MCC only after the script is copied again.

## Add a setting to the Config tab

Settings are read from the `Config` tab once a day and saved in the hidden `_RunState` tab as the day's **settings lock**. The Google Ads script reads only the lock, never the `Config` tab.

1. **Add the cell to the template Google Sheet** next to the existing settings (`B5:B6`), keeping the tables below it where they are. See [Change the template Google Sheet layout](#change-the-template-google-sheet-layout).
2. **Read it.** Add a cell constant in `src/appscript/config-sheet.ts`, add the field to `GlobalConfig` in `src/shared/types.ts`, and read it in `readGlobalConfig`.
3. **Save it in the lock.** `parseLockedDailyConfig` in `src/shared/contract.ts` copies only the fields it knows, so add the new field there and to `test/contract.test.ts`; otherwise the Google Ads script never sees it.
4. **Republish when it changes.** Add the field to the publish signature in `buildPublishPlan` in `publisher.ts`. Otherwise the 15-minute trigger thinks nothing changed.
5. **Decide what Apply Config Changes Now resets.** If the setting only affects publishing (like Target Currency), nothing else is needed. If it changes what the Google Ads script collects (like Exclude Shopping & Feed PMax), reset the checkpoints of the affected MCCs in `applyConfigChangesNow` in `sheet-ui.ts`, as the Shopping setting does.
6. **Use it** where it applies: in the publisher, or in the Google Ads script through `lock.config` in `orchestrator.ts`.

## Change the template Google Sheet layout

The template Google Sheet (its tabs, text, formatting, and cell positions) isn't stored in this repository. Maintainers edit it directly, and the code finds things by fixed positions defined in `src/appscript/config-sheet.ts`:

* `Setup!D13`: the saved dashboard URL.
* `Config!B2`, `B5`, `B6`, and `A9`: the status, Target Currency, Exclude Shopping & Feed PMax, and the select-all checkbox.
* `Config` rows 10–18: the MCC table (8 slots), with its header in row 10.
* `Config` row 21 onwards: the accounts table, with its header in row 21.

If you move anything, update the constants in the same change. `assertConfigLayout` stops the scripts when the table headers aren't where it expects them, which protects copies of the template made before the change. Those copies keep the old layout, so only new instances get the new one.

## Add a menu item

1. Add the item to the menu in `onOpen` in `src/appscript/sheet-ui.ts`, and export the function it calls.
2. Wrap any Sheet writes in `withSheetLock`, so the item can't run at the same time as the 15-minute trigger.
3. **Expose the function to Apps Script.** The build wraps the code in a function, so Apps Script can only call functions listed in the footer of the Sheet bundle in `scripts/build.mjs`. Add a line such as `function myAction() { return DlgSheetBundle.myAction(); }`.
4. Test the function directly with the Sheet mocks; the mock records menu items, alerts, and dialogs.

## Use a new Apps Script or Google Ads service

* **The Google Sheet's script** declares its permissions explicitly in `SHEET_OAUTH_SCOPES` in `scripts/build.mjs`, and `test/build.test.ts` checks that list. Add a scope only when there's no alternative: users are asked to authorize again, and narrow permissions are a feature of the tool.
* **The Google Ads script** gets its permissions from what the code uses. A new service (for example `UrlFetchApp`) means users must re-authorize the script in every MCC.
* Add the service, or the methods you use, to `test/mocks/google-apps-script.mock.ts`.

## Change the data the two scripts exchange

The two scripts exchange JSON through Sheet cells: each MCC's checkpoint in its `_MCC_<id>` tab, and the daily settings lock in `_RunState`. The format is defined in `src/shared/contract.ts`. An MCC can run an older copy of the Google Ads script than the Sheet, so:

* **Adding an optional field** is safe. Older readers ignore it; make new readers handle it being missing.
* **Removing a field, renaming it, or changing its meaning** is a breaking change. Increase `CONTRACT_VERSION`. Readers ignore data with an unknown version and log that the script must be copied again, so MCCs on the old copy stop collecting until they're updated.

Update `test/contract.test.ts` either way.

## Use demo data

The Sheet's script can fill your copy with made-up data for testing the publish pipeline and the dashboard without live recommendations:

1. In your copy, open **Extensions → Apps Script → Project Settings → Script Properties** and add `DLG_ENABLE_DEMO_TOOLS` with the value `true`.
2. Reload the Sheet and click **🚀 DLG Dashboard v2 → 🧪 Seed Synthetic Demo Data (For Testing)**.

This adds 2 demo MCCs (using 2 of the 8 MCC slots), 5 accounts in 4 currencies, and recommendations covering every type and scope, then publishes the `Data` tab. It doesn't reset your real MCCs. Demo recommendation IDs can't be decoded, so their Google Ads links fall back to links built from the account ID.

When you add a recommendation type or a field that the dashboard shows, add it to `src/appscript/demo-seed.ts` too.
