# Script Architecture and Design

This document explains the core architectural patterns used in the Budget
Allocation Dashboard script (`Code.js`).

**Last Updated:** 2025-10-30

---

## 1. Core Principle: Performance via Bulk Operations

The script's primary design principle is performance, achieved by minimizing API
calls. To avoid timeouts in large Manager Accounts (MCCs), the script uses a
**bulk-fetch and cache** pattern.

The main `processAccountRecommendations` function operates in distinct phases:

1.  **ID Gathering:** The script first iterates through all recommendations to
    compile unique lists of all campaign, budget, and recommendation resources
    that need to be analyzed.
2.  **Bulk Data Fetching:** It then uses a minimal number of bulk GAQL queries
    to fetch all the necessary data for the identified resources at once. This
    data is stored in-memory in `Map` objects, which act as a temporary cache.
3.  **Processing from Cache:** The script processes the recommendations,
    retrieving the pre-fetched data instantly from the in-memory cache instead
    of making new API calls.
4.  **Cross-Account Enrichment:** In a final pass, the script identifies
    manager-owned portfolio strategies and performs a targeted context switch to
    the owning manager account to fetch historical metrics that are otherwise
    inaccessible.

This pattern reduces the number of API calls from potentially hundreds down to a
handful, completely avoiding the "N+1 query problem."

---

## 2. Unified Schema: Structure and Presentation

To ensure data integrity and consistent presentation, the script uses a single,
unified schema object: `REPORT_SCHEMA_TEMPLATE`.

This object is the single source of truth for the report's structure and defines
three key properties for each field:

*   **`defaultValue`**: The value to use when creating a new report row.
*   **`format`**: The Google Sheets display format to be applied to the column
    (e.g., `'#,##0.00'` for currency, `'0.00%'` for percentages, `'@'` for plain
    text IDs).

The `writeDataToSheet` function uses this schema to programmatically apply the
correct formatting to each column after the data has been written, preventing
Google Sheets from misinterpreting data types.

---

## 3. The "Strategy" Pattern: Handling Recommendation Types

To handle different recommendation types (e.g., `CAMPAIGN_BUDGET` vs.
`MOVE_UNUSED_BUDGET`), the script uses the **Strategy Pattern**.

The `RECOMMENDATION_HANDLERS` object contains a "strategy" for each supported
recommendation type. Each strategy encapsulates the specific logic required to
process that type, such as extracting target campaigns, parsing impact
metrics, or declaring its own data dependencies (e.g., required budget
resources). This makes the script easily extensible—adding support for a new
recommendation type only requires adding a new handler, with no changes to the
core processing loop.

---

## 4. The "Read-Only" Principle

This is the script's most important safety principle. Its only purpose is to
**read** data from Google Ads and **write** it to a Google Sheet. It **NEVER**
makes any changes to the Ads account.
