# Contributing to FetchXML Studio

Thanks for helping. This guide covers setup, the branch and PR workflow, and what a PR
needs before it can merge. Releases are described in [RELEASING.md](RELEASING.md).

## Prerequisites

- **Node.js 22.22.2+, 24.15.0+ or 26+** (`engines` in `package.json`, the range supported by jsdom, the Vitest environment; CI uses Node 24) and npm.
- **Power Platform ToolBox 1.2.7 or later** to run the tool for real (`minAPI` 1.2.7).
- For the browser dev harness: the **Azure CLI** (`az login`) and a Dataverse dev org.

## Setup

```bash
npm ci                 # exact versions from package-lock.json
npm run dev            # Vite dev server (no Dataverse)
npm run dev:harness    # the real UI in a browser against a live org; see dev/README.md
```

Before opening a PR, run what CI runs:

```bash
npm run typecheck
npm run lint           # zero warnings
npm test
npm run validate       # PPTB manifest rules
npm run build
```

CI also fails on any **high or critical** `npm audit --omit=dev` finding and checks the
package contents (`.github/scripts/check-pack.mjs`).

To try a build inside PPTB, load the repository folder as a local tool after
`npm run build`.

## Branches and PRs

- Branch from `dev`: `feat/<issue>-<slug>`, `fix/<issue>-<slug>`, `chore/<issue>-<slug>`
  or `docs/<issue>-<slug>`.
- Open the PR into `dev`. A fix for the release being stabilized goes into `release/X.Y`,
  and a fix for the current stable release goes into `main` as a `hotfix/*` branch. See
  [RELEASING.md](RELEASING.md) → *Where a fix goes*.
- Link the issue in the description: `Closes #123`, or `Part of #123` while manual checks
  are pending.
- The `check` status must pass. PRs into `dev` are squash-merged.
- Commit messages follow [Conventional Commits](https://www.conventionalcommits.org/):
  `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`.

## What every PR includes

- **CHANGELOG:** user-facing changes go under `## [Unreleased]` in `CHANGELOG.md`.
- **No version change.** The version, the README version badge and the dated CHANGELOG
  section change only in release-prep PRs, so PRs stay independent.
- **README:** update the feature docs when behaviour changes, and the
  "What this tool changes" section when an operation writes data.
- **Tests:** unit tests for changes to the FetchXML model, parser, serializer or rules.

## Product rules

- **Ground behaviour in Microsoft Learn, then verify it live.** Cite the Learn page in the
  code or PR. When the docs and Dataverse disagree, Dataverse's actual behaviour wins, and
  the PR says what was verified and how.
- **Never refuse supported FetchXML.** Block only what is known to be invalid; warn only
  when the outcome can't be determined. A Dataverse error is an acceptable outcome.
- **Data-changing operations** (delete, bulk delete, state changes, workflows, saving
  views) confirm with the **environment** and the **scope** (entity and record count), per
  the PPTB Marketplace policy.

## UI changes

- **Agree on the design first.** Share a prototype or screenshots in the issue before
  building.
- **Fluent UI v9 only** (`@fluentui/react-components`), following the PPTB theme in light
  and dark. PR screenshots show both themes.
- **Wording:** follow the [style table](#ui-wording-style) below: sentence case, "…" for
  ellipses, and "Entity" / "Attribute" as the nouns. The same wording is shared with Command
  Bar Studio and Dataverse Request Studio.

### UI wording style

Every string a user sees (labels, menu items, tooltips, titles, placeholders, `aria-label`s,
messages, empty and progress text) follows these rules. They come from the
[Fluent 2 content design](https://fluent2.microsoft.design/content-design) guidance and the
[Microsoft Writing Style Guide](https://learn.microsoft.com/style-guide/) it builds on. Where a
rule comes from model-driven apps, Microsoft Learn or this project instead, the table says so.
`npm run check:wording` ([`.github/scripts/check-wording.mjs`](.github/scripts/check-wording.mjs))
checks the first two rules; it isn't enforced in CI yet (#231).

| Rule | Use | Not | Source |
|---|---|---|---|
| Sentence case | "Export as Excel", "Does not equal". Capitals only for the first word and proper nouns (FetchXML, Dataverse, Excel, Power Automate, Advanced Find) | "Export (Server)", "Less Than" | Fluent 2 [Content design → Capitalization](https://fluent2.microsoft.design/content-design) (sentence case on Windows and web); [Style Guide: Capitalization](https://learn.microsoft.com/style-guide/capitalization) (UI labels, titles and headings) |
| Ellipsis on commands | "…" (U+2026) at the end of a command that asks for more input or confirmation: "Save as…", "Delete…". Nothing on a command that acts at once: "Run", "Copy link". Not on an item that only opens a window or a submenu | "...", or "…" on "Run" | [Windows menus: Using ellipses](https://learn.microsoft.com/windows/win32/uxguide/cmd-menus#using-ellipses). Fluent 2 doesn't cover ellipses |
| Ellipsis on progress | "…" ends a progress label: "Loading workflows…", "Running query…" | "Loading...", "… in progress" | Project convention (Fluent 2 and the Style Guide don't cover progress text) |
| Run | Run, Running query…, Run failed | Execute, Executing, Execution failed | Project convention: plain, short verbs ([Style Guide: Use simple words, concise sentences](https://learn.microsoft.com/style-guide/word-choice/use-simple-words-concise-sentences); Fluent 2 [Content design → Plain language](https://fluent2.microsoft.design/content-design)) |
| Progress, indeterminate | `<Verb>ing <object>…` in a spinner label: "Loading workflows…", "Checking the status options…" | "Reading…", "Processing batch" | Project convention; Windows [Writing style → Lead with what's important](https://learn.microsoft.com/windows/apps/design/style/writing-style) |
| Progress, determinate | Title "Deleting records…", detail `Batch 2 of 5`, stats `6 of 12 (6 succeeded, 0 failed)` | "Processing batch 2 of 5", "(batch 2 of 5)" | Project convention |
| No exclamation points | "Deleted 3 records." | "Deleted!" | [Style Guide: Exclamation points](https://learn.microsoft.com/style-guide/punctuation/exclamation-points); Fluent 2 [Content design → Punctuation](https://fluent2.microsoft.design/content-design) (avoid exclamation points) |
| No "Please" | "Check your connection and run the query again" | "Please try again" | [Style Guide: please](https://learn.microsoft.com/style-guide/a-z-word-list-term-collections/p/please) (avoid it; the app isn't asking for anything inconvenient) |
| "Select", not "Click" | "Select **Run**" | "Click Run" | [Style Guide: click](https://learn.microsoft.com/style-guide/a-z-word-list-term-collections/c/click) (mouse-specific; use "select") |
| No "record(s)"; real plurals | "1 record", "3 records" | "record(s)" | Plain, concise language: Fluent 2 [Content design → Plain language](https://fluent2.microsoft.design/content-design) |
| No full stop on labels | "Copy link" | "Copy link." | Fluent 2 [Content design → Punctuation](https://fluent2.microsoft.design/content-design); Windows [Writing style → Periods](https://learn.microsoft.com/windows/apps/design/style/writing-style) (full stops end full sentences in tooltips, errors and dialogs only) |
| Nouns | **Entity** and **Attribute**. They are common nouns in sentence case: "All entities", "Show entities from", "Entity" only as a first word. "Column" only for results-grid columns ("Edit columns") | table, column (for the builder) | Project decision (shared with Command Bar Studio and Dataverse Request Studio), not Fluent 2 |
| Operator labels | Two rules. 1: the model-driven app filter editor's words, in Fluent 2 sentence case: "Is greater than", "On or before", "Next 7 days", "Does not contain data", "Equals current user". 2: the app's lowercase placeholder "x": "Last x days", "Older than x hours". Operators the app doesn't offer get a sentence-case label based on the FetchXML operator | "Not Equals", "Is Greater than", "On Or Before", "Last X Hours" | Model-driven apps: [Create or edit view filters](https://learn.microsoft.com/power-apps/maker/model-driven-apps/create-edit-view-filters); the full list read live from the filter editor: [#231 comment](https://github.com/mohsinonxrm/pptb-mxrm-fetchxml-studio/issues/231#issuecomment-6713341114); sentence case from Fluent 2 [Content design](https://fluent2.microsoft.design/content-design) |
| Power Automate field names | The List rows action's names: "Row count", "Select columns", "Filter rows", "Fetch Xml Query" (as the connector shows it) | "Top Count", "Table name" | Microsoft Learn: [List rows](https://learn.microsoft.com/power-automate/dataverse/list-rows). These quote another product, so the check's allowlist covers them |

The check's proper-noun list is `PROPER_NOUNS` in `check-wording.mjs`; a justified exception
(for example a quoted name from another product) goes in
`.github/scripts/wording-allowlist.json` as `{ file, text, reason }`. Out of scope: code
comments, log messages, FetchXML and OData text, and record data such as job names.

## Dependencies

- Prefer the platform and existing dependencies. New dependencies need a reason in the PR.
- **Only versions published at least 14 days ago.** Dependabot's cooldown applies the same
  rule to its version updates.
- After changing dependencies, run `npm install` and commit `package-lock.json`.

## Security

Report vulnerabilities privately; see [SECURITY.md](SECURITY.md). Never commit tokens,
org URLs with credentials, or `.env.local`.
