# Changelog

All notable changes to the **Demand-Led Growth (DLG) Dashboard — Ads Script edition** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project uses [Semantic Versioning](https://semver.org/).

---

## [2.0.0] - 2026-11-20

Version 2 is a rebuild. Instead of a single Google Ads Script that creates its own Google Sheet, you now copy a template Google Sheet that holds all the settings, combines data from several MCCs, and creates your Data Studio dashboard.

### Breaking changes

* **New setup.** Make a copy of the template Google Sheet, then install the Google Ads Script from the Sheet's menu. The script no longer creates a Google Sheet or sends a setup email. To move from version 1, set up version 2 as a new instance by following the [README](README.md).
* **New `Data` tab layout.** The columns and their order have changed, so dashboards built on version 1 don't work with version 2 data. Create a new dashboard from the Sheet's menu.
* **Hourly schedule.** Schedule the Google Ads Script to run **hourly** instead of daily. It finishes the day's data collection once, then exits immediately on later runs.
* **Enabled campaigns only.** Recommendations for paused or removed campaigns are no longer included.

### Added

* **Template Google Sheet and menu.** A **🚀 DLG Dashboard v2** menu that gives you a Google Ads Script already set up for your Sheet, creates and opens your Data Studio dashboard, and applies setting changes.
* **Account selection.** Choose which accounts to include on the `Config` tab: all accounts, everything under one MCC, or individual accounts. Accounts that are added, removed, or renamed in Google Ads are synced automatically.
* **Multiple MCCs.** Combine up to 8 MCCs in one Google Sheet. Each MCC's data is published as soon as its daily run finishes; an MCC that hasn't finished yet keeps its previous data, marked as stale.
* **Large MCC support.** Progress is saved during each run and resumed on the next hourly run, so MCCs with many accounts are no longer limited by the 30-minute Google Ads Scripts time limit. An account that fails is retried on up to 3 runs.
* **Two new recommendation types:** `FORECASTING_CAMPAIGN_BUDGET` (budget increases for forecasted demand) and `MARGINAL_ROI_CAMPAIGN_BUDGET` (budget increases based on marginal return).
* **Portfolio and shared-budget consolidation.** A recommendation for a portfolio bid strategy or shared budget appears as one row, with the number and names of its member campaigns and a link to the strategy or budget report, so budgets and forecasts are no longer counted once per campaign. New columns: `recommendationScope`, `recommendationNumber`, `portfolioCampaignsCount`, and `portfolioMemberCampaignNames`.
* **Currency conversion.** Choose a **Target Currency** to report accounts in different currencies together. The `*Global` columns are converted with a daily exchange rate; the other money columns stay in each account's currency.
* **Exclude Shopping & Feed PMax** option. Leaves out Shopping campaigns and Performance Max campaigns connected to a Merchant Center feed.
* **Source targets for `MOVE_UNUSED_BUDGET`.** The target CPA and target ROAS of the campaigns that the budget would move from (`moveBudgetSourceTargetCpa`, `moveBudgetSourceTargetRoas`).
* **MCC columns:** `mccId`, `mccName`, `mccSyncStatus`, and `mccLastUpdatedDate`, which show which MCC each row comes from and how up to date it is.
* **Status and warnings** on the `Config` tab: MCCs with stale data, accounts that failed, accounts listed under more than one MCC, and MCCs running an older copy of the Google Ads Script.
* **Apply Config Changes Now.** Settings are fixed for the day so that every MCC uses the same ones; use this menu item to apply changes straight away instead of waiting for the next day.

### Changed

* **Settings** moved from the `CONFIG` block at the top of the script to the `Config` tab. You don't need to edit the script.
* **Ratios that can't be calculated are left blank** instead of showing `0`, for example CPA with no conversions or ROAS with no cost.
* **Fewer permissions.** The Google Ads Script no longer needs access to Google Drive or Gmail, and the Google Sheet's script can only access the Sheet it belongs to.
* **Source code** is now written in TypeScript with automated tests. See [CONTRIBUTING.md](CONTRIBUTING.md).

---

## [1.0.0]

* Single Google Ads Script for one MCC. It collected `CAMPAIGN_BUDGET`, `MOVE_UNUSED_BUDGET`, `RAISE_TARGET_CPA`, and `LOWER_TARGET_ROAS` recommendations from every child account into a Google Sheet that it created, and emailed a link to create the dashboard.
