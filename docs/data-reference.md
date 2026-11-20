# Data reference

This document describes the `Data` tab that the DLG Dashboard publishes and Data Studio reads: what each row and column means, where every value comes from, and the rules that decide which recommendations appear and how they're combined. It's for technical operators who build or debug dashboards, and for developers who change the data.

For how the data is collected and published, see [architecture.md](architecture.md).

## Contents

1. [How to read the Data tab](#1-how-to-read-the-data-tab)
2. [Recommendation types](#2-recommendation-types)
3. [Forecasts](#3-forecasts)
4. [Scopes and consolidation](#4-scopes-and-consolidation)
5. [recommendationNumber](#5-recommendationnumber)
6. [Which recommendations are included](#6-which-recommendations-are-included)
7. [Currencies](#7-currencies)
8. [Freshness and status](#8-freshness-and-status)
9. [Links to Google Ads](#9-links-to-google-ads)
10. [The schema dummy row](#10-the-schema-dummy-row)
11. [Column reference](#11-column-reference)

## 1. How to read the Data tab

* **Each row is one recommendation for one entity**: a campaign, a portfolio bid strategy, or a shared budget. The `recommendationScope` column says which.
* **Row 1** holds the column names, and **row 2** is always the [schema dummy row](#10-the-schema-dummy-row). Real data starts in row 3.
* **The columns never move.** Data Studio reads Google Sheets columns by position, so new columns are only ever added at the end. The current layout has 82 columns.
* **Values describe the day the data was collected.** Each MCC's rows are replaced when its daily run finishes; until then they keep the previous day's values, marked stale (see [Freshness and status](#8-freshness-and-status)).
* **Empty cells mean "not available" or "doesn't apply"**, never zero. For example, a target CPA column is empty on a budget recommendation, and a CPA is empty when there were no conversions. Text columns show `N/A` instead of being empty.

## 2. Recommendation types

The tool collects six Google Ads recommendation types. They fill different columns:

| `recommendationType` | What Google Ads recommends | Columns it fills |
|---|---|---|
| `CAMPAIGN_BUDGET` | Raise a budget that limits the campaign | Current and recommended budget |
| `FORECASTING_CAMPAIGN_BUDGET` | Raise a budget for demand forecast in the coming days | Current and recommended budget |
| `MARGINAL_ROI_CAMPAIGN_BUDGET` | Raise a budget based on the expected return of extra spend | Current and recommended budget |
| `MOVE_UNUSED_BUDGET` | Move unused budget from another budget (the **source**) to this campaign's budget | Current and recommended budget of the receiving budget, the amount moved, and the `moveBudget*` details of the source budget |
| `RAISE_TARGET_CPA` | Raise the target CPA | Recommended target CPA |
| `LOWER_TARGET_ROAS` | Lower the target ROAS | Recommended target ROAS |

Every type also fills the [7-day forecast](#3-forecasts) and the campaign's own details and recent performance.

**How the recommended targets are calculated.** Google Ads returns a multiplier rather than the new target. The tool calculates `recommendationNewTargetCpa` and `recommendationNewTargetRoas` as the current average target that Google Ads reports, times the recommended multiplier. If Google Ads doesn't report a current average target, the campaign's (or portfolio's) own target is used instead.

## 3. Forecasts

Each recommendation comes with Google Ads' forecast for the next 7 days, in two versions:

* **Base** (`recommendationBase…`): what Google Ads expects if you change nothing.
* **Potential** (`recommendationPotential…`): what it expects if you apply the recommendation.

The forecast covers cost, clicks, conversions, conversion value, impressions, and video views.

**Where the forecast comes from:**

* For budget recommendations, Google Ads offers several budget options, each with its own forecast. The tool uses the forecast of the option that matches the recommended budget.
* For the other types, or when no option matches, it uses the recommendation's overall forecast.

**Calculated from the forecast:**

| Column | Formula |
|---|---|
| `weeklyCostIncrease` | Potential cost − base cost |
| `newWeeklyConversions` | Potential conversions − base conversions |
| `recommendationBaseCpa`, `recommendationPotentialCpa` | Cost ÷ conversions (empty with no conversions) |
| `recommendationBaseRoas`, `recommendationPotentialRoas` | Conversion value ÷ cost (empty with no cost) |
| `recommendationCpaDelta` | Base CPA − potential CPA. **Negative when applying the recommendation raises CPA**, which is expected for `RAISE_TARGET_CPA` |
| `recommendationRoasDelta` | Base ROAS − potential ROAS. **Positive when applying the recommendation lowers ROAS**, which is expected for `LOWER_TARGET_ROAS` |

## 4. Scopes and consolidation

When several campaigns share a portfolio bid strategy or a shared budget, Google Ads attaches the same recommendation to every one of them, each with the **whole** portfolio's or budget's budget and forecast. If those copies stayed as separate rows, a dashboard total would count the same budget and forecast once per campaign. So the tool merges them into one row.

**Scopes:**

| `recommendationScope` | When | The row represents |
|---|---|---|
| `PORTFOLIO` | A `RAISE_TARGET_CPA` or `LOWER_TARGET_ROAS` recommendation for campaigns that use a portfolio bid strategy | The portfolio bid strategy |
| `SHARED_BUDGET` | A `CAMPAIGN_BUDGET`, `FORECASTING_CAMPAIGN_BUDGET`, or `MARGINAL_ROI_CAMPAIGN_BUDGET` recommendation for campaigns with a shared budget | The shared budget |
| `CAMPAIGN` | Everything else, including every `MOVE_UNUSED_BUDGET` recommendation, budget recommendations for campaigns that use a portfolio bid strategy, and target recommendations for campaigns with a shared budget | One campaign |

**What changes when rows are merged:**

| Columns | Rule |
|---|---|
| `campaignName`, `campaignId`, `campaignUrl` | The portfolio's or shared budget's name, ID, and Google Ads report link |
| `portfolioCampaignsCount`, `portfolioMemberCampaignNames` | The number and names of the member campaigns the recommendation covers |
| Yesterday, 7-day, and 30-day cost, conversions, and conversion value | **Summed** across member campaigns, since each campaign spent separately |
| `campaign30DaysCpa`, `campaign30DaysRoas` | Recalculated from the summed values |
| Search lost impression share (rank and budget) | Averaged across member campaigns, weighted by 30-day cost |
| `campaignBudgetAmount` | `PORTFOLIO`: the sum of the member campaigns' distinct budgets. `SHARED_BUDGET`: the shared budget's amount |
| Recommendation budgets, targets, and forecasts | **Kept once**, not summed, because Google Ads already reports them for the whole portfolio or budget |
| Calculated columns | Recalculated from the merged values |
| Account, recommendation ID, and the other columns | Taken from one member campaign |

**Portfolios across accounts.** A portfolio bid strategy owned by a manager account can be used by campaigns in several child accounts, and Google Ads creates a copy of the recommendation in each one. The tool merges those copies into one row, identified by the strategy's owner and ID. Copies in different currencies stay separate rows, so amounts in different currencies are never added together.

## 5. recommendationNumber

An entity can have several recommendations on the same day, for example a campaign with both a `CAMPAIGN_BUDGET` and a `RAISE_TARGET_CPA` recommendation. Each row repeats the campaign's recent spend, so summing spend over all rows would count that campaign twice.

`recommendationNumber` numbers each entity's recommendations `1, 2, 3, …`:

* **Entities** are each portfolio bid strategy (by owner and ID), each shared budget, and each campaign, following the row's `recommendationScope`.
* **Order** is alphabetical by `recommendationType`, then by recommendation ID. **`1` is not "the best" or "the most important" recommendation**; it's just the first in that order.

To count each entity's recent spend once, filter to `recommendationNumber = 1`.

> [!NOTE]
> A campaign can belong to two entities at once. For example, a budget recommendation for a campaign that uses a portfolio bid strategy is a `CAMPAIGN` row, while a target recommendation for the same portfolio is a `PORTFOLIO` row that includes the campaign's spend. Both rows have `recommendationNumber = 1`, so filter by `recommendationScope` too if you need each campaign's spend exactly once.

## 6. Which recommendations are included

A recommendation appears in `Data` when all of these are true:

* Its account is checked on the `Config` tab.
* It's one of the [six supported types](#2-recommendation-types).
* At least one of its campaigns is **enabled**. Rows are only created for enabled campaigns, so paused or removed campaigns never appear.
* It passes the **Exclude Shopping & Feed PMax** setting, when that's turned on:
  * Shopping campaigns are excluded.
  * Performance Max campaigns connected to a Merchant Center feed are excluded. Performance Max campaigns without a feed stay.
  * A recommendation is left out if **any** of its campaigns is excluded, or isn't enabled.
  * A `MOVE_UNUSED_BUDGET` recommendation is also left out if **no** enabled campaign using the source budget remains after the filter. Excluded campaigns are dropped from `moveBudgetSourceCampaigns`.

The setting is applied while the data is collected. Rows kept from an earlier day (stale rows) can only be checked for Shopping campaigns, not for feed-connected Performance Max campaigns, until the MCC's next run.

## 7. Currencies

**Every money column is in the account's own currency** (`currencyCode`), except the six `…Global` columns, which are converted to the **Target Currency** chosen on the `Config` tab (`currencyCodeGlobal`):

| Converted column | Converted from |
|---|---|
| `recommendationCurrentBudgetAmountGlobal` | `recommendationCurrentBudgetAmount` |
| `recommendationNewBudgetAmountGlobal` | `recommendationNewBudgetAmount` |
| `recommendationBaseConversionsValueGlobal` | `recommendationBaseConversionsValue` |
| `recommendationPotentialConversionsValueGlobal` | `recommendationPotentialConversionsValue` |
| `weeklyCostIncreaseGlobal` | `weeklyCostIncrease` |
| `moveBudgetAmountGlobal` | `moveBudgetAmount` |

* **Only add up `…Global` columns across accounts in different currencies.** Summing a column in the account's currency (for example `campaign30DaysCost`) over accounts with different currencies gives a meaningless total.
* **Rates are fixed for the day.** The first publish of the day looks up each rate with `GOOGLEFINANCE` and reuses it for the rest of the day.
* **If no rate is available,** the tool uses the last rate it found on an earlier day. If it has never found one, the `…Global` columns are left empty rather than guessing.
* Converted values are rounded to 2 decimal places.

## 8. Freshness and status

| Column | Values |
|---|---|
| `mccSyncStatus` | `OK`: collected today. `WARNING: Stale snapshot from <date>`: the MCC hasn't finished today's run, so the row shows an earlier day's data. `WARNING: Account <id> failed (<error>); showing snapshot from <date>`: the account failed on 3 runs today, so the row shows its previous data |
| `mccLastUpdatedDate` | The date the row's data was collected. Empty if unknown |
| `timestamp` | The date of the run that collected the row, at `00:00:00` in the Google Sheet's time zone. It's a date, not the exact time of collection |

## 9. Links to Google Ads

| Column | Opens |
|---|---|
| `recommendationsDetailsUrl` | The recommendation type's page in the Google Ads recommendations view, for the recommendation's account |
| `campaignUrl` | `CAMPAIGN`: the campaign's overview. `PORTFOLIO`: the portfolio bid strategy's report. `SHARED_BUDGET`: the shared budget's report |

Google Ads addresses accounts in its web interface by an internal ID that differs from the 10-digit customer ID. The tool reads that ID, and the recommendation type, from the recommendation ID. When it can't (for example with demo data), the links fall back to the 10-digit customer ID, and `recommendationsDetailsUrl` opens the account's main recommendations page.

For a portfolio bid strategy owned by a manager account, the internal ID of the owner isn't available, so its `campaignUrl` uses the owner's customer ID and may not open the right page.

## 10. The schema dummy row

Row 2 of `Data` is always a placeholder row with a correctly typed value in every column: a number in number columns, a date in date columns, and `FALSE` in true/false columns. Data Studio detects each column's type from its values, so without this row a column that's empty in every real row that day (for example the target CPA columns on a day without `RAISE_TARGET_CPA` recommendations) would be treated as text, and charts using it would break.

* Its `accountName` is `!!! SCHEMA DEFINITION (FILTER ME OUT)`.
* **Exclude it in every report** with a filter on that `accountName`. It has `recommendationNumber = 1`, so a `recommendationNumber = 1` filter doesn't remove it.
* Dashboards created with **📊 Create New Data Studio Dashboard** copy the template's field types instead of detecting them again.

## 11. Column reference

**Type** is the Data Studio type the column should have. **Currency** is *Account* (the account's own currency, `currencyCode`) or *Target* (the Target Currency, `currencyCodeGlobal`).

### Rank, scope, and MCC (columns 1–6)

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 1 | `recommendationNumber` | Number | — | The recommendation's position among its entity's recommendations. See [recommendationNumber](#5-recommendationnumber) |
| 2 | `recommendationScope` | Text | — | `CAMPAIGN`, `PORTFOLIO`, or `SHARED_BUDGET`. See [Scopes](#4-scopes-and-consolidation) |
| 3 | `mccId` | Text | — | The manager account's customer ID (`XXX-XXX-XXXX`) |
| 4 | `mccName` | Text | — | The manager account's name |
| 5 | `mccSyncStatus` | Text | — | Whether the row is from today. See [Freshness](#8-freshness-and-status) |
| 6 | `mccLastUpdatedDate` | Date | — | The date the row's data was collected |

### Account and recommendation (columns 7–12)

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 7 | `accountId` | Text | — | The account's customer ID (`XXX-XXX-XXXX`) |
| 8 | `accountName` | Text | — | The account's name. `!!! SCHEMA DEFINITION (FILTER ME OUT)` on the [dummy row](#10-the-schema-dummy-row) |
| 9 | `timestamp` | Date & time | — | The run date at `00:00:00` |
| 10 | `recommendationId` | Text | — | The recommendation's Google Ads resource name (`customers/<id>/recommendations/<id>`) |
| 11 | `recommendationType` | Text | — | One of the [six types](#2-recommendation-types) |
| 12 | `recommendationsDetailsUrl` | Text (URL) | — | Link to the recommendation type in Google Ads. See [Links](#9-links-to-google-ads) |

### Campaign and currency (columns 13–22)

On `PORTFOLIO` and `SHARED_BUDGET` rows, columns 13–15 describe the portfolio or shared budget instead of a campaign.

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 13 | `campaignUrl` | Text (URL) | — | Link to the campaign, portfolio, or shared budget in Google Ads |
| 14 | `campaignId` | Text | — | Campaign ID, portfolio bid strategy ID, or shared budget ID |
| 15 | `campaignName` | Text | — | Campaign name, portfolio bid strategy name, or shared budget name |
| 16 | `campaignType` | Text | — | Google Ads campaign type, such as `SEARCH`, `PERFORMANCE_MAX`, or `DEMAND_GEN` |
| 17 | `campaignSubType` | Text | — | Google Ads campaign subtype, or `N/A` |
| 18 | `campaignIsAiMax` | True/false | — | Whether AI Max for Search is turned on |
| 19 | `campaignAiMaxTextCustomizationEnabled` | True/false | — | Whether AI Max text customization is turned on |
| 20 | `campaignAiMaxFinalUrlExpansionEnabled` | True/false | — | Whether AI Max final URL expansion is turned on |
| 21 | `currencyCode` | Text | — | The account's currency (ISO code such as `USD`) |
| 22 | `currencyCodeGlobal` | Text | — | The Target Currency of the `…Global` columns |

### Budget (columns 23–26)

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 23 | `campaignBudgetAmount` | Number | Account | The campaign's budget amount. On merged rows, see [consolidation](#4-scopes-and-consolidation) |
| 24 | `campaignBudgetTotalAmount` | Number | Account | The total amount of a budget set for a custom period. Empty for daily budgets |
| 25 | `campaignBudgetDeliveryMethod` | Text | — | `STANDARD` or `ACCELERATED` |
| 26 | `campaignBudgetType` | Text | — | The budget's period: `DAILY` or `CUSTOM_PERIOD` |

### Bidding and members (columns 27–31)

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 27 | `portfolioCampaignsCount` | Number | — | On `PORTFOLIO` and `SHARED_BUDGET` rows, the number of member campaigns the recommendation covers. Empty on `CAMPAIGN` rows |
| 28 | `portfolioMemberCampaignNames` | Text | — | Those campaigns' names, separated by commas. `N/A` on `CAMPAIGN` rows |
| 29 | `campaignBiddingStrategyType` | Text | — | Google Ads bidding strategy type, such as `TARGET_CPA`, `MAXIMIZE_CONVERSION_VALUE`, or `MANUAL_CPC` |
| 30 | `campaignBiddingTargetRoas` | Number | — | The current target ROAS as a ratio (`4.0` is 400%), from the campaign or its portfolio bid strategy |
| 31 | `campaignBiddingTargetCpa` | Number | Account | The current target CPA, from the campaign or its portfolio bid strategy |

### Recent performance (columns 32–40)

On merged rows these are combined across member campaigns, as described in [consolidation](#4-scopes-and-consolidation).

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 32 | `campaignStatsYesterdayCost` | Number | Account | Cost yesterday |
| 33 | `campaignStats7DaysCost` | Number | Account | Cost over the last 7 days |
| 34 | `campaign30DaysCost` | Number | Account | Cost over the last 30 days |
| 35 | `campaign30DaysConversionsValue` | Number | Account | Conversion value over the last 30 days |
| 36 | `campaign30DaysConversions` | Number | — | Conversions over the last 30 days. Can be fractional |
| 37 | `campaign30DaysRoas` | Number | — | 30-day conversion value ÷ 30-day cost. Empty with no cost |
| 38 | `campaign30DaysCpa` | Number | Account | 30-day cost ÷ 30-day conversions. Empty with no conversions |
| 39 | `campaign30DaysSearchRankLostImpressionShare` | Percent | — | Search impression share lost to ad rank over the last 30 days. Empty for non-Search campaigns |
| 40 | `campaign30DaysSearchBudgetLostImpressionShare` | Percent | — | Search impression share lost to budget over the last 30 days. Empty for non-Search campaigns |

### Recommendation and 7-day forecast (columns 41–64)

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 41 | `recommendationCurrentBudgetAmount` | Number | Account | Budget recommendations: the current budget, as reported by the recommendation. For `MOVE_UNUSED_BUDGET`, the receiving budget |
| 42 | `recommendationCurrentBudgetAmountGlobal` | Number | Target | Column 41, converted |
| 43 | `recommendationNewBudgetAmount` | Number | Account | Budget recommendations: the recommended budget |
| 44 | `recommendationNewBudgetAmountGlobal` | Number | Target | Column 43, converted |
| 45 | `recommendationNewTargetCpa` | Number | Account | `RAISE_TARGET_CPA`: the recommended target CPA, [calculated](#2-recommendation-types) by the tool |
| 46 | `recommendationNewTargetRoas` | Number | — | `LOWER_TARGET_ROAS`: the recommended target ROAS as a ratio, [calculated](#2-recommendation-types) by the tool |
| 47 | `recommendationBaseCost` | Number | Account | Forecast 7-day cost without the change |
| 48 | `recommendationPotentialCost` | Number | Account | Forecast 7-day cost with the change |
| 49 | `recommendationBaseClicks` | Number | — | Forecast 7-day clicks without the change |
| 50 | `recommendationPotentialClicks` | Number | — | Forecast 7-day clicks with the change |
| 51 | `recommendationBaseConversions` | Number | — | Forecast 7-day conversions without the change |
| 52 | `recommendationPotentialConversions` | Number | — | Forecast 7-day conversions with the change |
| 53 | `recommendationBaseConversionsValue` | Number | Account | Forecast 7-day conversion value without the change |
| 54 | `recommendationBaseConversionsValueGlobal` | Number | Target | Column 53, converted |
| 55 | `recommendationPotentialConversionsValue` | Number | Account | Forecast 7-day conversion value with the change |
| 56 | `recommendationPotentialConversionsValueGlobal` | Number | Target | Column 55, converted |
| 57 | `recommendationBaseCpa` | Number | Account | Column 47 ÷ column 51 |
| 58 | `recommendationPotentialCpa` | Number | Account | Column 48 ÷ column 52 |
| 59 | `recommendationBaseRoas` | Number | — | Column 53 ÷ column 47 |
| 60 | `recommendationPotentialRoas` | Number | — | Column 55 ÷ column 48 |
| 61 | `recommendationBaseImpressions` | Number | — | Forecast 7-day impressions without the change |
| 62 | `recommendationPotentialImpressions` | Number | — | Forecast 7-day impressions with the change |
| 63 | `recommendationBaseVideoViews` | Number | — | Forecast 7-day video views without the change |
| 64 | `recommendationPotentialVideoViews` | Number | — | Forecast 7-day video views with the change |

### Calculated (columns 65–72)

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 65 | `campaignPercentageOfBudgetUsedYesterday` | Percent | — | Yesterday's cost ÷ the current budget (column 41, or column 23 if 41 is empty). Empty without a budget |
| 66 | `weeklyCostIncrease` | Number | Account | Column 48 − column 47 |
| 67 | `weeklyCostIncreaseGlobal` | Number | Target | Column 66, converted |
| 68 | `newWeeklyConversions` | Number | — | Column 52 − column 51 |
| 69 | `dailyBudget30Days` | Number | Account | Column 23 × 30: what the budget allows over 30 days. Only meaningful for daily budgets |
| 70 | `dailyBudgetDelta` | Number | Account | Column 69 − column 34: budget left unspent over the last 30 days |
| 71 | `recommendationCpaDelta` | Number | Account | Column 57 − column 58. See [Forecasts](#3-forecasts) for the sign |
| 72 | `recommendationRoasDelta` | Number | — | Column 59 − column 60. See [Forecasts](#3-forecasts) for the sign |

### Moving unused budget (columns 73–82)

These columns are only filled on `MOVE_UNUSED_BUDGET` rows.

| # | Column | Type | Currency | Description |
|:-:|---|---|---|---|
| 73 | `moveBudgetAmount` | Number | Account | How much the recommendation adds to the receiving budget: column 43 − column 41 |
| 74 | `moveBudgetSourceCampaigns` | Text | — | The names of the enabled campaigns that use the source budget, separated by commas |
| 75 | `moveBudgetSourceBudgetName` | Text | — | The source budget's name |
| 76 | `moveBudgetSourceBudgetType` | Text | — | The source budget's period: `DAILY` or `CUSTOM_PERIOD` |
| 77 | `moveBudgetSourceBudgetDeliveryMethod` | Text | — | The source budget's delivery method |
| 78 | `moveBudgetSourceBudgetIsShared` | True/false | — | Whether the source budget is a shared budget |
| 79 | `moveBudgetSourceBudgetAmount` | Number | Account | The source budget's current amount |
| 80 | `moveBudgetSourceTargetCpa` | Number | Account | The target CPA of the source campaigns, when they all have the same one. Empty otherwise |
| 81 | `moveBudgetSourceTargetRoas` | Number | — | The target ROAS of the source campaigns as a ratio, when they all have the same one. Empty otherwise |
| 82 | `moveBudgetAmountGlobal` | Number | Target | Column 73, converted. Added after the other columns, so it's the last column in the tab |
