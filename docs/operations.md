# Operations and troubleshooting

This guide is for the people who set up and run the DLG Dashboard for one or more manager accounts (MCCs): who needs which access, how the schedules work, what every control in the Google Sheet does, how fresh the data is, and how to find and fix problems.

To set up the dashboard for the first time, follow [Set up the dashboard](../README.md#set-up-the-dashboard) in the README. For what each column in the `Data` tab means, see the [data reference](data-reference.md).

## Contents

1. [Access and permissions](#1-access-and-permissions)
2. [Several manager accounts](#2-several-manager-accounts)
3. [Schedules](#3-schedules)
4. [The Setup tab](#4-the-setup-tab)
5. [The Config tab](#5-the-config-tab)
6. [The daily cycle and data freshness](#6-the-daily-cycle-and-data-freshness)
7. [Changing settings during the day](#7-changing-settings-during-the-day)
8. [Status messages](#8-status-messages)
9. [Troubleshooting](#9-troubleshooting)
10. [Following the data through the hidden tabs](#10-following-the-data-through-the-hidden-tabs)

## 1. Access and permissions

| Who | Needs | Why |
|---|---|---|
| Whoever sets up the dashboard | Access to the template Google Sheet and dashboard, through your Google point of contact | The template is shared with an access group |
| Whoever authorizes the Google Ads script in an MCC | Permission to create and authorize scripts in that MCC, **and edit access to the Google Sheet** | The script writes its data to the Sheet as that person |
| Editors of the Google Sheet | Edit access, and to authorize the Sheet's script the first time they use its menu | The menu actions write to the Sheet |
| People who only view the dashboard | View access to the Data Studio report | They don't need access to the Sheet if the report's data source uses the owner's credentials (the default) |

**Permissions each script asks for:**

* **The Google Ads script:** access to the Google Ads accounts under the MCC, and to Google Sheets. It only reads Google Ads; it never changes anything.
* **The Sheet's script:** access to this Google Sheet only, permission to show its menu and dialogs, and permission to install its own 15-minute trigger.

## 2. Several manager accounts

One Google Sheet can combine up to **8 MCCs**. To add an MCC, paste the same script (from **🚀 DLG Dashboard v2 → 📋 Copy Google Ads Script**) into it, run **Preview** once, and set its frequency to **Hourly**.

* **Install the script only in top-level MCCs.** Each account should be reachable from exactly one MCC in the Sheet. If you install the script in an MCC and in a sub-manager account under it, the same accounts are discovered twice: only the first MCC processes them, and the `Config` tab shows a warning.
* **Each MCC runs on its own schedule.** MCCs don't wait for each other: each MCC's data is published as soon as its own daily run finishes.
* **Removing an MCC** isn't supported yet. To stop including an MCC, uncheck all its accounts on the `Config` tab and remove the script from the MCC. The MCC then shows `⏸️ INACTIVE`, its rows leave the `Data` tab on the next publish, but it keeps its slot in the MCC table.

## 3. Schedules

Two schedules keep the dashboard up to date:

| What | Runs | Set up by |
|---|---|---|
| **The Google Ads script** in each MCC | Every hour | You, in Google Ads: set the script's frequency to **Hourly** |
| **The Sheet's publisher** (`runScheduledPublish`) | Every 15 minutes | Automatically, the first time an editor opens **📋 Copy Google Ads Script** or clicks **🔄 Apply Config Changes Now** |

**Why hourly?** Each MCC collects its data once a day. The hourly schedule lets a large MCC finish over several runs (each run stops before Google Ads' 30-minute limit and the next one continues), and retries accounts that failed. Once the day's run is complete, later runs exit after a few seconds.

**Who owns the publisher trigger.** Apps Script triggers belong to the person who installed them and run with that person's access.

* To see the trigger, open **Extensions → Apps Script → Triggers** in the Sheet. You only see your own triggers, so if another editor installed it, it appears only in their list.
* The Sheet installs only one publisher trigger. Each run records the time, and a new trigger is installed only when the current editor has none and no publisher has run for 45 minutes.
* **If the trigger's owner loses access to the Sheet,** publishing stops. After 45 minutes, any editor can reinstall it by clicking **🔄 Apply Config Changes Now**.

## 4. The Setup tab

The `Setup` tab (**DLG Dashboard v2 - Setup Guide**) lists the setup steps in order. Every action is in the **🚀 DLG Dashboard v2** menu. Yellow cells are yours to edit; everything else is filled in automatically.

| Row | What it says | Notes |
|---|---|---|
| **System Status** | The current status | Shows the same message as `Config!B2`. See [Status messages](#8-status-messages) |
| **Step 1: Install the Google Ads Script** | 1.1 Copy the script with **📋 Copy Google Ads Script** (authorize the Sheet's script the first time). 1.2 Create a script in each MCC in Google Ads and paste the code. 1.3 Authorize it and run Preview once. 1.4 Set its frequency to **Hourly**. 1.5 Click **🔄 Apply Config Changes Now** to fill in the MCC and account tables | Opening the copy dialog also installs the 15-minute publisher trigger. See [Schedules](#3-schedules) |
| **Step 2: Choose Accounts & Settings** | 2.1 Set the Target Currency and the other settings on the `Config` tab. 2.2 Check the MCCs or accounts to include | Data is collected on the script's next hourly run. To collect straight away, click **🔄 Apply Config Changes Now** and then run the script in Google Ads |
| **Step 3: Create Dashboard** | 3.1 **📊 Create New Data Studio Dashboard** opens a copy of the template dashboard, already connected to this Sheet. 3.2 In Data Studio, click **Edit and share → Acknowledge and save** | You can do this before the first data is published |
| **Step 4: Save Dashboard URL** | Paste the saved report's URL into the yellow cell (`D13`) | Only `lookerstudio.google.com` and `datastudio.google.com` links are accepted |
| **Step 5: Open Dashboard** | **🚀 Open Saved Dashboard**, or the link next to it | Opens the URL saved in `D13` |

## 5. The Config tab

| Item | Cell | What it does |
|---|---|---|
| **System Status** | `B2` | The latest status message. See [Status messages](#8-status-messages) |
| **Target Currency** | `B5` | The currency of the `…Global` columns. The list contains the currencies of the accounts in the Sheet; until you choose one, it's set to the first one registered |
| **Exclude Shopping & Feed PMax** | `B6` | When checked, leaves out Shopping campaigns and Performance Max campaigns connected to a Merchant Center feed. See the [data reference](data-reference.md#6-which-recommendations-are-included) |
| **Select / Deselect All Accounts** | `A9` | Checks or unchecks every account in every MCC |
| **MCC table** | Rows 11–18 | One row per registered MCC, added automatically. **Select All in MCC** checks or unchecks all its accounts; **Discovered Accounts** and **Selected Accounts** count them; **Sync Status** shows `🟢 ACTIVE (n Selected)` or `⏸️ INACTIVE (0 Selected — Skipped)` |
| **Accounts table** | From row 22 | One row per account: **Include**, **Account CID**, **Account Name**, **Parent MCC ID**, **Parent MCC Name**, and **Currency** |

**How the tables stay up to date.** Each MCC's Google Ads script reports its accounts to the Sheet, and the publisher syncs them into the tables:

* **New accounts** are added at the bottom. They're checked if their MCC's **Select All in MCC** box is checked, and unchecked otherwise.
* **Accounts removed** from the MCC in Google Ads are removed from the table.
* **Renamed accounts and MCCs** are updated in place; your checkboxes are kept.

The checkboxes stay in step with each other: checking **Select All in MCC** checks that MCC's accounts, and the top-level box is checked only when every account is.

> [!IMPORTANT]
> **Don't insert or delete rows or columns above or inside the tables.** The scripts find the settings and tables by position, and stop with a "Config layout changed" error if the table headers move.

## 6. The daily cycle and data freshness

A typical day for one MCC:

1. **First publish of the day:** the Sheet's publisher fixes the day's settings (accounts, Target Currency, and the Shopping setting) as soon as at least one MCC has accounts selected. Changes made in `Config` later that day don't apply until the next day, or until you click **🔄 Apply Config Changes Now**.
2. **Next hourly run:** the MCC's Google Ads script clears yesterday's data and starts collecting. A large MCC can take several hourly runs.
3. **Run complete:** within 15 minutes, the publisher replaces the MCC's rows in `Data` with today's data, and the status shows **🟢 PUBLISHED**.
4. **Rest of the day:** both scripts find nothing to do and exit straight away.

**Until an MCC finishes its run for the day,** its rows keep the previous day's data, with `mccSyncStatus` set to `WARNING: Stale snapshot from <date>`. Dashboards keep working during the run.

**Dates and time zones.** "Today" is based on the Google Sheet's time zone (**File → Settings**). The Google Ads script takes that time zone when it's copied, so if you change the Sheet's time zone, copy the script into each MCC again.

**What Data Studio shows.** Data Studio caches data from Google Sheets. If the dashboard looks older than the `Data` tab, use **Refresh data** in the report's menu.

## 7. Changing settings during the day

Edit the `Config` tab, then click **🚀 DLG Dashboard v2 → 🔄 Apply Config Changes Now (Reset Today's Lock)**. It applies your changes straight away, and only collects data again where it has to:

| You changed | What happens |
|---|---|
| **Target Currency** | Rates are looked up again and `Data` is republished immediately from the data already collected |
| **Unchecked accounts** | `Data` is republished immediately without them |
| **Checked accounts that an MCC already finished without** | That MCC collects again on its next hourly run, and shows as stale until then |
| **Exclude Shopping & Feed PMax** | Every active MCC collects again on its next hourly run, because the filter is applied while collecting. They show as stale until then |
| **Nothing** | `Data` is republished from the data already collected. Use this to publish straight away after a run instead of waiting up to 15 minutes |

## 8. Status messages

**On the `Config` tab (`B2`):**

| Message | Meaning |
|---|---|
| `🟢 MCC <id> (<name>) registered (<n> accounts). …` | A new MCC was registered. Choose its accounts, then click **🔄 Apply Config Changes Now** |
| `🟢 PUBLISHED (<date>) — <n> recommendations across <m> active MCC(s).` | All active MCCs are up to date |
| `⚠️ WARNING — Published (<date>) with stale snapshot for pending MCC(s): <MCCs>` | The listed MCCs haven't finished today's run yet; their rows show earlier data |
| `⚠️ <n> account(s) failed in today's run (showing their previous snapshot): <accounts>` | These accounts failed on 3 runs today and show their previous data |
| `⚠️ <n> account(s) are under more than one MCC — …` | The script is installed in an MCC and in a sub-manager account under it. See [Troubleshooting](#9-troubleshooting) |
| `⚠️ Google Ads Script out of date in MCC(s) <MCCs> — …` | Those MCCs run a different version of the script than the Sheet |
| `🔄 Config changes applied (<date>). …` | **🔄 Apply Config Changes Now** ran; the rest of the message says whether any MCC collects again |

Warnings can appear together, separated by `|`.

**In each MCC's hidden `_MCC_<id>` tab** (`STATUS_MESSAGE` row), the Google Ads script records its own progress: `⏳ IN PROGRESS — MCC <id>: <done>/<total> accounts processed`, or `⏳ STAGED FOR PUBLISH (<date>) — MCC <id> completed …` when its run is finished and waiting for the publisher.

## 9. Troubleshooting

| Symptom | Likely cause | What to do |
|---|---|---|
| **An MCC doesn't appear on the `Config` tab after Preview** | The publisher hasn't run yet, or the script couldn't write to the Sheet | Click **🔄 Apply Config Changes Now**. If the MCC still doesn't appear, check the script's logs in Google Ads for an error, and make sure the person who authorized the script can edit the Sheet |
| **The script's log says `SPREADSHEET_URL is not configured`** | The script wasn't copied from the Sheet's menu (for example, it was copied from this repository) | Copy it again with **📋 Copy Google Ads Script** |
| **The script's log says it must run in a Google Ads manager (MCC) account** | The script was installed in a regular account | Install it in a manager account |
| **The script runs but only registers accounts; no data arrives** | Today's settings didn't include this MCC when it ran: accounts were checked after the day's settings were fixed, or the MCC was registered later | Click **🔄 Apply Config Changes Now**, then run the script again (or wait for the next hourly run) |
| **Changes on the `Config` tab have no effect** | Settings are fixed for the day | Click **🔄 Apply Config Changes Now**. See [Changing settings](#7-changing-settings-during-the-day) |
| **An MCC stays stale all day** | Its run hasn't finished: the script isn't scheduled hourly, it fails, or the MCC is very large | In Google Ads, check that the script's frequency is **Hourly** and look at its recent runs and logs. Check progress in the MCC's `_MCC_<id>` tab ([see below](#10-following-the-data-through-the-hidden-tabs)) |
| **The status message and the `Data` tab stop updating** | The publisher trigger isn't running, for example because its owner lost access to the Sheet | Click **🔄 Apply Config Changes Now** to publish and, after 45 minutes without a publisher run, install a new trigger. Check **Extensions → Apps Script → Triggers** |
| **Accounts failed in today's run** | Those accounts returned errors on 3 runs, for example because of a permission or quota problem | Look at the script's logs in that MCC. Failed accounts are retried the next day. To retry sooner, fix the cause, click **🔄 Apply Config Changes Now**, and run the script |
| **Accounts are under more than one MCC** | The script is installed in an MCC and in a sub-manager account under it | Remove the script from the sub-manager account. If the warning stays, delete that sub-manager's `_MCC_<id>` tab |
| **Google Ads script out of date** | An MCC runs a copy of the script from a different version than the Sheet | Copy the script into that MCC again with **📋 Copy Google Ads Script** |
| **A new account doesn't appear** | Accounts linked after an MCC has finished its run for the day are discovered the next day | Wait for the next day's run |
| **`…Global` columns are empty for one currency** | No exchange rate was available from `GOOGLEFINANCE`, and no earlier rate was saved | Usually temporary. Click **🔄 Apply Config Changes Now** later to look up rates again |
| **The dashboard shows a `!!! SCHEMA DEFINITION (FILTER ME OUT)` row or an extra count** | The report doesn't filter out the [schema dummy row](data-reference.md#10-the-schema-dummy-row) | Add a report-level filter that excludes that `accountName` |
| **The dashboard shows older data than the `Data` tab** | Data Studio's cache | Refresh the report's data in Data Studio |
| **An error says `Config layout changed: expected "…" in …`** | Rows or columns were inserted or deleted on the `Config` tab | Undo the change so the table headers are back in rows 10 and 21 |
| **An error says `Maximum of 8 MCCs already registered in Config`** | The Sheet already has 8 MCCs | Use a second Google Sheet for more MCCs |
| **A log says a checkpoint or settings lock `has schemaVersion …`** | An MCC runs a much older or newer copy of the script than the Sheet | Copy the script into that MCC again |
| **Dates look a day off** | The Sheet's time zone changed after the script was copied | Copy the script into each MCC again |

**Finding the Google Ads script's logs:** in the MCC, go to **Tools and settings → Bulk actions → Scripts**, open **Script history**, and open the logs of a run. Most of the script's messages start with `[DLG v2]`.

## 10. Following the data through the hidden tabs

To see the hidden tabs, open **View → Hidden sheets** in the Google Sheet. **Look but don't edit:** the scripts rely on their exact contents.

1. **`_MCC_<id>`** (one per MCC) shows what the MCC's script last reported:
   * `CHILD_ACCOUNT` rows: the accounts it discovered.
   * `CHECKPOINT`: today's progress as JSON. `status` is `IN_PROGRESS` or `COMPLETED`, `completedCids` lists the finished accounts, and `failedCids` lists failed accounts with their error and number of attempts.
   * `LAST_COMPLETED_DATE`, `STATUS_MESSAGE`, and `SCRIPT_VERSION`.
2. **`_Stage_Accounts_<id>`, `_Stage_Budgets_<id>`, `_Stage_Campaigns_<id>`, `_Stage_Recs_<id>`** hold the data the MCC collected today. If `_Stage_Recs_<id>` is empty after a completed run, Google Ads had no supported recommendations for the selected accounts (or the Shopping setting excluded them).
3. **`_RunState`** holds the publisher's state:
   * `DAILY_CONFIG_LOCK_<date>`: today's fixed settings, including the selected accounts of each MCC.
   * `PUBLISH_STATUS_<date>` and `PUBLISH_SIGNATURE_<date>`: the last publish's result.
   * `FX_RATES_<date>_<currency>` and `FX_LAST_GOOD_<from>_<to>`: exchange rates.
4. **`Data`** is the result. The `mccSyncStatus` column says whether each row is from today.

**Example:** an MCC is stale. Its `_MCC_<id>` tab shows `"status":"IN_PROGRESS"` with 40 of 120 accounts in `completedCids` and no `failedCids`: the MCC is large and will finish in later hourly runs. If `failedCids` lists accounts with errors, those accounts are being retried; check the errors and the script's logs.
