# How to Contribute

We would love to accept your patches and contributions to this project.

## Before you begin

### Sign our Contributor License Agreement

Contributions to this project must be accompanied by a
[Contributor License Agreement](https://cla.developers.google.com/about) (CLA).
You (or your employer) retain the copyright to your contribution; this simply
gives us permission to use and redistribute your contributions as part of the
project.

If you or your current employer have already signed the Google CLA (even if it
was for a different project), you probably don't need to do it again.

Visit <https://cla.developers.google.com/> to see your current agreements or to
sign a new one.

### Review our Community Guidelines

This project follows [Google's Open Source Community
Guidelines](https://opensource.google/conduct/).

## Contribution process

### Code Reviews

All submissions, including submissions by project members, require review. We
use [GitHub pull requests](https://docs.github.com/articles/about-pull-requests)
for this purpose.

For anything larger than a small fix, please [open an issue](https://github.com/google-marketing-solutions/dlg-tool/issues) first so we can agree on the approach before you write the code.

### Before you send a pull request

1. Add or update tests for every change in `src/`.
2. Run `npm test` and `npm run build`. Both must pass with no errors.
3. Commit the rebuilt `dist/` files together with your source changes (see [Generated files](#generated-files)).
4. If you changed behaviour that users or operators can see, update the matching docs and add a line to [CHANGELOG.md](CHANGELOG.md).

## Set up your development environment

You need [Node.js](https://nodejs.org/) 20 or later, which includes npm.

```bash
git clone https://github.com/google-marketing-solutions/dlg-tool.git
cd dlg-tool
npm install
```

| Command | What it does |
|---|---|
| `npm test` | Runs the unit and integration tests with [Vitest](https://vitest.dev/) |
| `npx vitest run test/<file>.test.ts` | Runs one test file |
| `npm run typecheck` | Type-checks `src/`, `test/`, and `scripts/` without building |
| `npm run build` | Type-checks, then builds both Apps Script bundles into `dist/` |
| `npm run push:sheet` | Builds, then uploads the Google Sheet's script with [clasp](https://github.com/google/clasp) (see [Try your changes in a Google Sheet](#try-your-changes-in-a-google-sheet)) |

## How the project is organized

The tool is made of two Apps Script programs that share one Google Sheet:

* **The Google Ads script** (`src/ads/`, built into `dist/ads/Code.gs`) runs hourly in each manager account and collects recommendations.
* **The Google Sheet's script** (`src/appscript/`, built into `dist/sheet/Code.gs`) adds the Sheet's menu, keeps the `Config` tab up to date, and publishes the `Data` tab.
* **Shared code** (`src/shared/`) holds the types, the `Data` tab schema, and the helpers both programs use to read and write the Sheet.

[docs/architecture.md](docs/architecture.md) explains how the two programs work together. [docs/development-guide.md](docs/development-guide.md) walks through common changes step by step.

## Generated files

`scripts/build.mjs` builds two files that users copy or install, so they're committed to the repository:

* `dist/ads/Code.gs`: the Google Ads script. The build embeds it inside the Sheet's script, which serves it from the **📋 Copy Google Ads Script** menu with the user's Sheet URL filled in.
* `dist/sheet/Code.gs` and `dist/sheet/appsscript.json`: the Google Sheet's script and its manifest.

**Never edit files in `dist/` by hand.** Change the TypeScript in `src/`, run `npm run build`, and commit the result.

## Coding guidelines

* Follow the [Google TypeScript Style Guide](https://google.github.io/styleguide/tsguide.html). The project compiles in TypeScript's strict mode.
* Start every new source file, including tests and scripts, with the Apache 2.0 license header used in the existing files in `src/`.
* Apps Script has no module system, so don't add runtime dependencies from npm. The build bundles `src/` into self-contained files.
* **The `Data` tab is append-only.** Add new columns at the end of `REPORT_SCHEMA_TEMPLATE` in `src/shared/schema.ts`, and never insert, reorder, or rename existing columns. Data Studio matches the Sheet's columns by position, so moving one breaks every dashboard built on it. See [Add a column to the Data tab](docs/development-guide.md#add-a-column-to-the-data-tab).
* Keep the scripts read-only towards Google Ads: they collect data and never change accounts.

## Tests

Tests live in `test/` and run in Node with Vitest. Apps Script services don't exist in Node, so `test/mocks/google-apps-script.mock.ts` provides in-memory fakes for the parts of `SpreadsheetApp`, `AdsApp`, `AdsManagerApp`, `LockService`, `PropertiesService`, `HtmlService`, `Utilities`, and `Logger` that the code uses.

* Every change in `src/` needs a test that fails without the change.
* If your code calls an Apps Script method the mocks don't support yet, add it to the mock file in the same pull request.
* See [Write tests](docs/development-guide.md#write-tests) for the test helpers and patterns.

## Try your changes in a Google Sheet

Tests cover most logic, but changes to menus, dialogs, or Sheet formatting are worth checking in a real Google Sheet:

1. Make your own copy of the template Google Sheet (see the [README](README.md#set-up-the-dashboard)).
2. In your copy, open **Extensions → Apps Script → Project Settings** and copy the **Script ID**.
3. Copy `.clasp.json.example` to `.clasp.json`, replace `YOUR_SCRIPT_ID` with your Script ID, and run `npx clasp login` once.
4. Run `npm run push:sheet` to upload your build to your copy. The command refuses to run unless `.clasp.json` points at `dist/sheet`.
5. To test the Google Ads script, use **📋 Copy Google Ads Script** in your copy and run it with **Preview** in a test manager account. If you have no recommendations to test with, use the [demo data](docs/development-guide.md#use-demo-data).

`.clasp.json` is ignored by Git, so your Script ID stays local.

## Commit messages

Use [Conventional Commits](https://www.conventionalcommits.org/): `type(scope): summary`, for example `feat(publisher): …`, `fix(ads): …`, or `docs: …`. Keep the subject line to 50 characters or fewer and wrap the body at 72 characters.

## Releasing (maintainers)

1. Update `version` in `package.json` and add a dated entry to [CHANGELOG.md](CHANGELOG.md).
2. Run `npm test` and `npm run build`, and commit the rebuilt `dist/` files.
3. Point `.clasp.json` at the template Google Sheet's script and run `npm run push:sheet`, so new copies of the template get the release.

Each Google Ads script records the version it was built with, and the Sheet warns when an MCC runs a different version than the Sheet. Copies of the template made before a release keep their old version; users who want a new release set up a new instance.
