# Budget Allocation Dashboard - Ads Script

This directory contains the Google Ads Script (`Code.js`) that powers the Budget
Allocation Dashboard.

**Last Updated:** 2025-10-30

## Overview

This script is the data-gathering engine for the Budget Allocation Dashboard
(aka DLG Dashboard), a tool that helps users identify budget reallocation
opportunities within their Google Ads accounts.

### Key Functions

1.  **Processes All Sub-Accounts:** Iterates through all accounts under the MCC.
2.  **Fetches Budget Recommendations:** Identifies active `CAMPAIGN_BUDGET` and
    `MOVE_UNUSED_BUDGET` recommendations.
3.  **Performs Bulk Data Extraction:** Efficiently gathers detailed campaign
    performance metrics (cost, conversions, ROAS, CPA), budget settings,
    impression share data, and historical bidding targets using a minimal number
    of API calls.
4.  **Handles Portfolio Strategies:** Accurately fetches historical average
    tCPA/tROAS for all strategy types:
    *   **Standard Strategies:** Fetched directly from campaign metrics.
    *   **Client-Owned Portfolios:** Fetched from the local bidding strategy.
    *   **Manager-Owned Portfolios:** The script automatically **switches
        context** to the owning Manager Account to fetch historical metrics that
        are otherwise inaccessible to the client account.
5.  **Automated Dashboard Setup:** Automatically creates the data source Google
    Sheet and emails the user a "magic link" to instantly generate a configured
    Looker Studio dashboard.

## Architecture & Performance

The script is designed to be both performant and maintainable:

*   **High Performance:** It uses a **bulk-fetch and cache** architecture to
    minimize API calls, ensuring it can run efficiently on large MCCs without
    timing out.
*   **Extensible:** It uses the **Strategy Pattern** to handle different
    recommendation types. Each handler encapsulates the logic for a specific
    recommendation, including declaring its own data dependencies (e.g.,
    required budget resources). This makes it easy to add support for new
    recommendation types in the future without altering the core logic.

For more details, see the `architecture.md` file.

## Setup

1.  **Create a New Ads Script:** In your Google Ads MCC, create a new script.
2.  **Copy & Paste Code:** Copy the code from `Code.js` and paste it into the
    Google Ads script editor.
3.  **Run the Script:** Run the `main` function (preview or run).
4.  **Check Your Email:** The script will automatically create a new Google
    Sheet and send you an email with a "Create Report in Looker Studio" button.
5.  **Finalize Setup:** Click the button in the email to generate your Looker
    Studio dashboard with the data pre-connected.
6.  **Schedule the Script:** Set up a schedule for the script to run daily to
    keep the data fresh.

## Configuration

The script contains a `CONFIG` object at the top of the file that you can modify
if needed:

*   `debug_mode`: Set to `true` to enable detailed logging in the "Execution
    logs" tab.
*   `force_setup_email`: Set to `true` to force the script to resend the setup
    email (useful if you lost the original email).
