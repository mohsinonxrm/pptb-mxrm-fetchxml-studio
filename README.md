# FetchXML Studio for Power Platform ToolBox

A powerful, modern FetchXML query builder and data explorer for [Power Platform ToolBox](https://github.com/PowerPlatformToolBox/desktop-app). Inspired by the XrmToolBox FetchXML Builder, reimagined with React 18, Fluent UI v9, and seamless Dataverse integration.

![Version](https://img.shields.io/badge/version-1.4.5-blue)
[![npm](https://img.shields.io/npm/v/@mohsinonxrm/pptb-fetchxml-studio?label=npm)](https://www.npmjs.com/package/@mohsinonxrm/pptb-fetchxml-studio)
![React](https://img.shields.io/badge/React-18-61DAFB?logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178C6?logo=typescript)
![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft)
![PPTB Types](https://img.shields.io/badge/%40pptb%2Ftypes-1.2.3-blue)

## ✨ Features

### 🔧 Visual Query Builder
- **Tree-based hierarchy** — Build FetchXML queries visually with entities, attributes, filters, orders, and link-entities
- **Property panel for every node** — Fetch options, entity, attributes, all-attributes, sorts, filters, conditions and link-entities each have an editor covering the full FetchXML reference. Every field explains itself in an info popover with a link to Microsoft Learn.
  - **Metadata-driven** — Attribute pickers, operators, value pickers, aggregate functions, date grouping, sortable attributes and compare-to-attribute choices follow each attribute's metadata, so the panel only offers what Dataverse accepts.
  - **Shows problems where they are** — Rule errors appear on the field they're about and as an **Error** badge in the tree; settings Dataverse ignores are listed with a **Remove** button.
  - **Tree badges** — Aggregate, Distinct, Retained and Hints on the query; Group by, date grouping and aggregate functions on attributes; Union and Quick find on filters.
  - **Safe mode changes** — Turning off aggregate, or switching a link-entity to a type that returns no attributes, lists what has to change and asks first.
- **Full FetchXML support** — Aggregate queries, grouping, distinct, top/count, paging, and advanced hints
- **Nested filters** — Create complex AND/OR filter groups with unlimited nesting
- **Link-entity relationships** — Browse and add 1:N, N:1, and N:N relationships with inner/outer joins; N:N relationships join through their intersect entity, as documented; intersect entities (such as `systemuserroles` or `teammembership`) can be joined to each side of their N:N relationship

### 📊 Results Grid
- **Virtualized DataGrid** — Handle large result sets with smooth scrolling (powered by react-window)
- **Copy results** — Right-click a cell to copy its value, its raw value (GUID, number, ISO date or option value), the row, the selected rows or the column. **Ctrl+C** copies the selected rows, or the focused cell. The command bar's **Copy** menu copies the selected or all loaded rows, as shown or as raw values. Rows copy as tab-separated values with a header row, so they paste into Excel as cells
- **Paging and Retrieve All** — Scrolling to the end loads the next page; **Retrieve All** loads every page, with **Stop** to keep what's loaded so far. Pages use the documented paging cookie and the query's own `count` and starting `page`. Queries Dataverse returns no cookie for (sorted by a link-entity attribute, or distinct) page by page number, with a note. Views follow `@odata.nextLink`. A failed page shows how many rows loaded, with **Retry**
- **Rich cell rendering** — Power Apps-style display for lookups, option sets, dates, currency, and more
- **Column header menu** — Click a header for **Sort** (A to Z, smallest to largest or oldest to newest, by type), **Filter by…** and **Clear filter**. Shift+click a header still adds it to a multi-column sort.
- **Column filters** — **Filter by…** adds a `<condition>` to the query and runs it again, so the filter covers every record, not just the loaded ones. Operators and value pickers follow the attribute's metadata, as in the property panel. Column filters are kept in their own `<filter type="and">` inside the root filter, apart from the conditions you built. An `or` root filter is wrapped so the column filters narrow the results; it's unwrapped again when the last column filter goes. Link-entity columns use `entityname`, which filters after the join. Active filters show as removable tags above the grid, and filtered columns show a filter icon
- **Quick search** — The search box above the grid filters the loaded rows as you type: any column, values as shown, with matches highlighted. The info bar shows how many loaded rows match. Rows it hides are deselected, and scrolling doesn't load more pages while searching
- **Multi-column sorting** — Shift+click headers for multi-column sort. Each click writes a valid `<order>`: by alias in aggregate queries, with `entityname` for link-entity columns, and not at all for attribute types Dataverse can't sort (multi-select option sets)
- **Edit / Add columns** — Remove, reorder and add columns from the root entity or related entities (N:1 and 1:N); the tree and FetchXML change to match. Removing a column removes its attribute (aliased or not, root or link-entity), any sort on its alias, and an outer link-entity left empty. At least one column must remain (a query without attributes returns every attribute). Add columns isn't available in aggregate queries
- **Link-entities without an alias** — Their columns show in the grid: Dataverse names them `{name}{N}.attribute`, and the grid uses the same names
- **Row selection** — Select single or multiple records for bulk operations. Record actions use the table's primary key, which Dataverse returns automatically for regular queries; for aggregate and distinct queries (which return no record IDs) record actions are disabled with an explanation
- **Resizable & reorderable columns** — Customize your view with drag-and-drop columns
- **Formatted values** — Display OData formatted values or raw values (configurable)
- **Works with every entity** — Queries run against the entity set name from Dataverse metadata, so entities with irregular names (for example `feedback`, `msdyn_conversationdata`) run correctly. The Power Platform ToolBox host's own name guess is wrong for about 4% of entities.

### 📝 FetchXML Editor
- **Monaco editor** — Full-featured XML editor with XML syntax highlighting and line numbers
- **Bi-directional editing** — Edit XML directly and parse back to the visual builder
- **Intellisense from the reference** — Element, attribute and value suggestions follow the FetchXML reference (every documented attribute, including `aggregatelimit`, `datasource`, `options`, `rowaggregate`, order `entityname` and the quick find overrides). Undocumented attributes are suggested only where Dataverse was verified to accept them, and say so. Value suggestions work for hyphenated attributes such as `link-type`
- **Power Automate–safe values** — Apostrophes stay literal in the generated FetchXML (attribute values use double quotes), so expressions such as `@{body('x')}` paste into Power Automate unchanged; `"`, `&`, `<` and `>` are escaped everywhere, including names, aliases and `from` / `to`
- **Editor mode actions run what you see** — In editor mode, **Execute**, **Save** and **Send to Tool** first load your FetchXML into the tree, so results, saved views and sent queries always match the editor. Execute is available as soon as the editor has content, even before an entity is chosen. If the tree would change anything, a confirmation lists the changes first, and the editor then shows the cleaned FetchXML that ran. Invalid FetchXML is reported and nothing runs.
- **What you build is what runs** — The generated FetchXML contains exactly what is in the tree; nothing is injected, so what you copy, save as a view, or send to another tool matches what executes
- **Copy to clipboard** — One-click copy of generated FetchXML, available in both read-only and editor modes
- **Alias validation** — Invalid alias characters are sanitized in real time in the property editors; Parse to Tree corrects invalid aliases, lists each correction, and updates every reference to a corrected alias (condition and order `entityname`, `valueof`, aggregate order aliases) so the query keeps working
- **Accurate validation** — Everything in the FetchXML reference is supported, plus undocumented behaviour verified to work (e.g. `link-type="natural"`). Every rule is checked against a live Dataverse environment:
  - **errors** block execution — only for FetchXML Dataverse is known to reject;
  - **warnings** offer **Execute anyway** — only where it can't be determined in the client (e.g. a link-entity without `from`/`to`, which is valid only for system many-to-many relationships);
  - **notes** never block (e.g. `xmlns:*` from other tools, which Dataverse ignores).
- **Query-level checks** — Validate and Execute catch the mistakes Dataverse rejects, each verified on a live environment:
  - aggregate queries where an attribute has no Group by / aggregate or no alias;
  - an aggregate order using an alias from another entity;
  - `aggregatelimit` combined with an order;
  - duplicate aliases;
  - an invalid union hint;
  - misuse of `rowaggregate`;
  - `top` together with `count`, `page` or `returntotalrecordcount`, and `page` below 1;
  - comparing to another attribute (`valueof`) with an operator other than eq, ne, gt, ge, lt or le, or through a link-entity without an alias;
  - a link-entity inside a filter whose type isn't any / not any / all / not all;
  - attributes, all-attributes or orders in an exists, in, any, not any, all or not all link-entity.

  Settings Dataverse accepts but ignores (such as `aggregatelimit` on a non-aggregate query, `aggregate` on a condition, which FetchXML can't use to filter on aggregate values, or `visible` on a link-entity) are shown as notes only.
- **Boolean spellings** — `"1"`/`"0"` and other spellings Dataverse accepts are read correctly (e.g. `usertimezone="1"`, `descending="1"`, `enableprefiltering="1"`), so Parse to Tree never flips or drops them.
- **Lossless Parse to Tree** — Every attribute in the FetchXML reference is kept, including `datasource`, `aggregatelimit`, `useraworderby`, query hints (`options`), `latematerialize`, `<order alias>`, `distinct` counts, `rowaggregate` and quick find flags. So are unknown, namespaced and legacy attributes (e.g. `xmlns:generator` from SQL 4 CDS, `version`/`mapping` from Advanced Find, `uiname`/`uitype` from saved views, on conditions and on each `<value>`). Values are kept exactly as written: every decimal digit, spaces in `<value>` text, and line breaks or tabs in attribute values. Every FetchXML example in the Microsoft Learn FetchXML articles round-trips without loss.
- **No silent changes** — If Parse to Tree must correct something Dataverse would reject (for example an invalid alias), a confirmation dialog lists every change first
- **LayoutXML preview** — See the column layout configuration

### 📥 Load & Save Views
- **Load system/personal views** — Browse and load existing Dataverse views
- **Optimized view execution** — Uses SavedQuery/UserQuery APIs for better performance
- **Save to Dataverse** — Save your queries as new personal views or update existing ones
- **Solution-aware** — Add views to solutions during save

### 📤 Export & Data Operations
- **Local Excel export** — Native `.xlsx` generation via ExcelJS with proper data types (numbers, dates, currencies, booleans)
- **Dataverse Excel export** — Server-side export via the `ExportToExcel` action (requires a saved view)
- **Native save dialog** — When running inside PPTB, both export paths use `fileSystem.saveFile` for a native OS save-file dialog; falls back to browser blob download in standalone/dev mode
- **Formatted value options** — Export formatted values, raw values, or both side-by-side columns
- **Record deletion** — Delete selected records with confirmation dialog
- **Bulk delete jobs** — Submit async `BulkDelete` jobs to Dataverse with job-tracking URL; any number of selected records (the query for selected records is built in the tool, with no URL length limit)
- **Delete all deletes what you see** — with no selection, Delete all and Bulk delete (job) submit the query that produced the results on screen; after editing the query in the left pane, run it again first ("Delete all unavailable" explains why). Adding, removing or sorting columns in the grid doesn't count as a change
- **Batch delete** — Parallel per-record delete for small sets (up to 100 records) with progress and ETA
- **Run workflows** — Execute on-demand workflows on selected records with batch execution and progress tracking

### 📋 Bulk Attribute Selection
- **Select Attributes dialog** — Multi-select attributes from a sortable, resizable grid (display name, logical name, type). Long names end with an ellipsis and show in full on hover. Search by display or logical name, filter by type, show selected only, and see how many are selected. Opens from the entity and link-entity Add menus
- **Smart updates** — Adds new selections and removes deselected attributes while preserving existing order and properties
- **Search & filter** — Real-time filtering across logical name, display name, and data type columns

### 🔀 Code Generation
- **Code tab** — New tab in the right pane that generates ready-to-use code from the current query in seven formats
- **C# QueryExpression** — Full SDK object graph (`QueryExpression`, `ColumnSet`, `FilterExpression`, `ConditionExpression`, `LinkEntity`, `OrderExpression`) via the Dataverse `FetchXmlToQueryExpression` API; supports aggregate queries with `XrmAggregateType` and `XrmDateTimeGrouping`
- **C# FetchExpression** — Simple `new FetchExpression(@"...")` wrapper using a verbatim string literal
- **JavaScript** — `Xrm.WebApi.retrieveMultipleRecords` call for model-driven apps and PCF components
- **pac CLI** — `pac env fetch --xml '...'` for the Power Platform CLI, as a bash / zsh line and a PowerShell line. The XML is single-quoted, so the shell never expands values such as `$100` or `$(...)`
- **Power Automate** — Field-by-field form mirroring the Dataverse "List rows" connector UI; table name resolved from `EntityDefinitions` metadata (not a heuristic), each field individually copyable; Top Count is taken only from `top` (`count` is a page size)
- **Web API** — Full OData URL with request headers (including `Prefer: odata.include-annotations="*"` for formatted values and paging cookies), the documented paging-cookie steps, a cURL snippet and a PowerShell `Invoke-RestMethod` example; entity set name resolved from metadata
- **SQL** — Dataverse-generated T-SQL via `FetchXMLToSQL` (undocumented API, labelled Preview)

### 🎨 User Experience
- **Dark/Light themes** — Every surface (tree, properties, editors, results grid, panels and dialogs) uses the Fluent UI v9 light or dark theme and follows the Power Platform ToolBox theme live: changing it in ToolBox settings, including **System**, and changing the OS theme while ToolBox is on System, update the tool without a reload
- **Lazy metadata loading** — Loads only what's needed, when it's needed
- **Intelligent caching** — In-memory cache prevents duplicate API calls per session
- **Resizable panes** — Adjust split-pane layout to your preference
- **Keyboard shortcuts** — `Ctrl+Enter` to execute query, copy XML to clipboard
- **Display settings** — Toggle logical names vs display names in column headers; choose formatted, raw, or both value modes in the grid and exports
- **Query scope settings** — Control which entities appear in the entity picker: *Publisher + Solution* (full filter), *Solution Only*, or *All Entities* (no filter). Persisted across sessions.
- **Advanced Find Only toggle** — When on (default), entities, attributes and relationships are limited to the ones Advanced Find shows (`IsValidForAdvancedFind`), in every entity scope. Turn it off to build queries Advanced Find can't: system and customization entities (saved views, forms, plug-in steps, solution components), intersect entities (team membership, user roles), and hidden attributes and relationships. Attributes that can't be read (`IsValidForRead = false`) are never listed. Hidden label attributes (such as `statuscodename`) can be filtered and sorted. Toggling is instant: metadata is fetched once and filtered locally. Attributes already in a query keep their operators and value inputs either way.

### 🔗 Tool-to-Tool (T2T) Integration
- **Send to Tool button** — Launch another PPTB tool directly from FetchXML Studio, pre-loading it with the current FetchXML query and active Dataverse connection (one-way handoff).
- **Automatic tool discovery** — Targets are discovered via the PPTB capability registry — installed tools that declare the `fetchxml` capability. The Send to Tool control always renders as a dropdown listing the discovered tools (even a single match). No configuration needed.
- **Active connection forwarding** — The active Dataverse connection is forwarded automatically so the target tool opens against the same environment.
- **Inbound prefill** — FetchXML Studio also accepts incoming T2T invocations from other tools (see [Callee Contract](#-callee-contract-pptbconfigjson) below).

### 🔒 Privilege-Aware
- **Export privilege check** — Only shows Dataverse export option if user has access
- **Delete privilege check** — Validates entity-specific delete permissions before enabling delete actions
- **Bulk delete privilege check** — Validates `prvBulkDelete` before surfacing bulk delete
- **Workflow privilege check** — Validates `prvReadWorkflow` + `prvWorkflowExecution` before showing workflow picker
- **View save privilege check** — Validates `prvWriteQuery` / `prvWriteCustomization` / `prvPublishCustomization` / `prvWriteUserQuery` before save operations

## 🖼️ Interface Overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│  [Entity Selector ▼]  [Load View ▼]  [Save View] [Send to Tool ▼]  [⚙] │
├──────────────────────┬──────────────────────────────────────────────────┤
│                      │  [FetchXML]  [LayoutXML]  [Results]  [Code]  [▶ Execute]│
│   Query Tree         ├──────────────────────────────────────────────────┤
│                      │                                                   │
│   📁 fetch           │   Results Grid / Monaco Editor                   │
│   └─📁 entity        │                                                   │
│      ├─📋 attribute  │   ┌────────────┬────────────┬──────────────┐     │
│      ├─📋 attribute  │   │ Column 1   │ Column 2   │ Column 3     │     │
│      ├─🔽 order      │   ├────────────┼────────────┼──────────────┤     │
│      ├─🔍 filter     │   │ Value      │ Value      │ Value        │     │
│      │  └─ condition │   │ Value      │ Value      │ Value        │     │
│      └─🔗 link-entity│   └────────────┴────────────┴──────────────┘     │
│                      │                                                   │
├──────────────────────┤   [◀ Prev] Page 1 of 10 [Next ▶]  100 records    │
│                      │                                                   │
│   Properties Panel   │   [+ Add Columns] [✎ Edit Columns] [⬇ Export]   │
│   (context-aware)    │   [🗑 Delete] [Bulk Delete] [⚡ Run Workflow]    │
│                      │                                                   │
└──────────────────────┴──────────────────────────────────────────────────┘
```

## 🚀 Getting Started

### Prerequisites
- [Power Platform ToolBox](https://github.com/PowerPlatformToolBox/desktop-app) desktop application
- Connection to a Dataverse environment

### Installation
FetchXML Studio is available as a tool in Power Platform ToolBox. Install it from the tool gallery or load as a custom tool. The package is published on npm as [`@mohsinonxrm/pptb-fetchxml-studio`](https://www.npmjs.com/package/@mohsinonxrm/pptb-fetchxml-studio).

## 🛠️ Development

### Setup

```bash
# Clone the repository
git clone https://github.com/mohsinonxrm/pptb-mxrm-fetchxml-studio.git
cd pptb-mxrm-fetchxml-studio

# Install the exact dependency versions from package-lock.json
npm ci

# Start development server
npm run dev
```

### Building

```bash
# Production build
npm run build

# Preview production build
npm run preview

# Validate package manifest against PPTB registry rules
npm run validate
```

**Releases.** CI (`.github/workflows/ci.yml`) runs typecheck, lint, tests, PPTB validation, the build and a package-contents check on every pull request. Pushing a `vX.Y.Z` tag on `main` publishes to npm with provenance and creates the GitHub Release; see [RELEASING.md](RELEASING.md).

**Tool icon.** `package.json` points `icon` at `icons/icon.svg` under `dist/`, where Power Platform ToolBox looks for it. The build bundles `icons/icon-insider.svg` (gold stamp) for prerelease versions such as `1.4.2-beta.7`, and `icons/icon.svg` for releases.

**Dependencies.** `package-lock.json` is committed and pins every dependency with an integrity hash. Use `npm ci` for a clean install. After changing dependencies, run `npm install` and commit the updated lockfile.

### Testing

```bash
# Unit tests (Vitest + jsdom): parser, serializer, rules, round-trip over the Microsoft Learn examples
npm test

# Watch mode
npm run test:watch

# Type check and lint
npm run typecheck
npm run lint

# Formatting (Prettier)
npm run format:check
```

### Debug Logging

Enable debug logging in the browser console:

```javascript
// Enable specific category
enableDebug("metadataAPI");       // API calls and responses
enableDebug("treeExpansion");     // Tree node operations
enableDebug("relationshipPicker"); // Relationship loading
enableDebug("linkEntityEditor");  // Link-entity configuration
enableDebug("propertiesPanel");   // Property editor routing

// Enable all
enableAllDebug();

// Disable
disableDebug("metadataAPI");
disableAllDebug();
```

## 📦 Tech Stack

| Technology | Version | Purpose |
|------------|---------|--------|
| **React** | 18.3 | UI framework |
| **TypeScript** | 5.9 | Type-safe development |
| **Vite** | 7 | Build tooling and HMR |
| **Fluent UI v9** | 9.74 | Microsoft design system (Tree, DataGrid, Tabs, Drawer, etc.) |
| **Monaco Editor** | 0.54 | VS Code XML editor for FetchXML authoring |
| **react-window** | 2 | Virtualized list rendering for large datasets |
| **ExcelJS** | 4.4 | Native `.xlsx` generation with typed cells |
| **@pptb/types** | 1.2.3 | Power Platform ToolBox host API types (`window.dataverseAPI`, `window.toolboxAPI`) |
| **Vitest** | 5 | Unit tests (jsdom) |

## 📁 Project Structure

```
src/
├── app/
│   └── AppShell.tsx              # Main layout with resizable split panes
├── features/fetchxml/
│   ├── api/
│   │   ├── pptbClient.ts         # window.dataverseAPI wrapper (all Dataverse ops)
│   │   ├── dataverseMetadata.ts  # Lazy metadata loading with cache + dedup
│   │   ├── excelExport.ts        # Local ExcelJS export with native types
│   │   ├── formattedValues.ts    # OData @FormattedValue annotation helpers
│   │   └── invocation.ts         # T2T caller API (isT2TSupported, sendFetchXmlToTool)
│   ├── model/
│   │   ├── nodes.ts              # FetchXML node TypeScript definitions
│   │   ├── fetchxml.ts           # FetchXML XML generation
│   │   ├── fetchxmlParser.ts     # FetchXML XML → node tree parser (severity-classified findings)
│   │   ├── fetchxmlRules.ts      # Query-level rules shared by the parser and editors
│   │   ├── attributeCapabilities.ts # What each attribute type supports (sort, group by, aggregates, valueof)
│   │   ├── aggregateToggle.ts    # What turning off an aggregate query changes
│   │   ├── fetchxmlIntellisense.ts # Intellisense / autocomplete helpers
│   │   ├── layoutxml.ts          # LayoutXML generation
│   │   ├── operators.ts          # Operator definitions by attribute type
│   │   ├── displaySettings.ts    # Display settings types and defaults
│   │   ├── treeUtils.ts          # Tree traversal utilities
│   │   └── __fixtures__/         # FetchXML examples from Microsoft Learn (MIT) for round-trip tests
│   ├── engine/
│   │   ├── queryExpressionTypes.ts   # QE JSON schema interfaces (number | string enums)
│   │   ├── queryExpressionCodegen.ts # C# SDK code emitter (resolveEnum, unwrapODataValue)
│   │   └── fetchxmlCodeGenerators.ts # Sync generators: FetchExpr, JS, pac, PA, WebAPI
│   ├── state/
│   │   ├── builderStore.tsx      # React context + reducer state management
│   │   └── cache.ts              # Per-session in-memory metadata cache
│   └── ui/
│       ├── LeftPane/             # Tree view + context-aware properties panel
│       │   └── PropertiesPanel/
│       │       ├── editors/      # One editor per node type
│       │       ├── fields.tsx    # InfoField, InfoCheckbox, NumberField, LeftoverSetting
│       │       ├── docLinks.ts   # Microsoft Learn links used by the info popovers
│       │       ├── ruleFindings.ts # Rule findings → field validation
│       │       └── useEditorMetadata.ts # Attribute / entity metadata hooks
│       ├── RightPane/            # Monaco editor, LayoutXML viewer, results grid,
│       │                         # code generation panel, Power Automate pane
│       ├── Toolbar/              # Entity selector, load view picker, save button,
│       │                         # send-to-tool button

│       ├── Dialogs/              # Save view, select attributes, delete, bulk
│       │                         # delete, workflow picker, solution picker
│       └── Settings/             # Settings drawer (display preferences)
└── shared/
    ├── components/               # Reusable value pickers (option set, boolean,
    │                             # date, numeric, multi-value, relationship, etc.)
    ├── contexts/                 # ThemeContext
    ├── hooks/                    # usePptbContext, useLazyMetadata, useAccessMode,
    │                             # usePublisherFilter, useSolutionFilter
    └── utils/                    # Debug logging utilities
```

## 🔧 FetchXML Features Support

| Feature | Status |
|---------|--------|
| Basic queries | ✅ |
| Attributes (select columns) | ✅ |
| All-attributes | ✅ |
| Filters (and/or) | ✅ |
| Nested filters | ✅ |
| All condition operators (offered per attribute from the documented data types and a live operator × attribute-type matrix; value inputs match each operator: whole numbers for "x" operators, period + year for fiscal period operators) | ✅ |
| Link-entity (joins) | ✅ |
| Inner/outer joins | ✅ |
| N:N relationships (through the intersect entity) | ✅ |
| Orders (sorting) | ✅ |
| Multi-column sort | ✅ |
| Aggregate queries | ✅ |
| Groupby | ✅ |
| Distinct | ✅ |
| Top/Count | ✅ |
| Paging with cookies (`fetchxmlpagingcookie`), page-number paging when Dataverse returns no cookie, views via `@odata.nextLink` | ✅ |
| Value-of conditions | ✅ |
| Entity name on conditions | ✅ |
| Filter link-entity (any/all) | ✅ |
| Query hints | ✅ |
| Order by alias (aggregate) | ✅ |
| Quick find, union hint | ✅ |
| Row aggregate (hierarchy child count) | ✅ |
| Long-term retention (`datasource`) | ✅ |
| Validation against Dataverse rules | ✅ |

## 🔗 Callee Contract (`pptb.config.json`)

FetchXML Studio declares a PPTB [Inter-Tool Invocation](https://github.com/PowerPlatformToolBox/desktop-app) callee contract in `pptb.config.json` at the repository root, and declares the `fetchxml` capability so callers can discover it. Any other PPTB tool can launch FetchXML Studio and pre-populate it by passing a prefill payload that matches the following schema:

```json
{
  "fetchXml": "<fetch><entity name=\"account\">...</entity></fetch>"
}
```

| Field | Type | Required | Description |
|---|---|---|---|
| `fetchXml` | `string` | One of these three | Full serialized FetchXML query string. The root entity is parsed from this. |
| `viewRef` | `object` | One of these three | EntityReference to an existing view. The FetchXML is retrieved from the view record. |
| `entityLogicalName` | `string` | One of these three | Root entity logical name (e.g. `"account"`). Starts a fresh query on that entity. **Ignored when `fetchXml` is present** — the root entity is then derived from the FetchXML, since a mismatched value would break this metadata-driven tool's lookups. |
| `viewRef.id` | `string` (uuid) | Required if viewRef | GUID of the view record |
| `viewRef.entityLogicalName` | `"savedquery"` \| `"userquery"` | Required if viewRef | Dataverse entity logical name — `savedquery` = system/public view, `userquery` = personal view |

Resolution priority at runtime: **`fetchXml` → `viewRef` → `entityLogicalName`**. With no prefill (a standalone launch) the tool opens empty.

### Returning a query to the caller

When FetchXML Studio is launched by another tool, a **Return FetchXML** button appears in the FetchXML toolbar. Clicking it returns the current query to the caller:

```json
{ "fetchXml": "<fetch>...</fetch>" }
```

The caller receives this as the resolved value of `launchTool(...)`. This is distinct from the host-injected "Return to [Caller]" banner, which simply navigates back and resolves the caller's promise with `null` (no data).

### Sending a query from another PPTB tool

To launch FetchXML Studio from your own PPTB tool with a pre-loaded query:

```typescript
const result = await window.toolboxAPI.invocation.launchTool(
  "@mohsinonxrm/pptb-fetchxml-studio",
  { fetchXml: "<fetch><entity name=\"account\"><attribute name=\"name\"/></entity></fetch>" },
  { primaryConnectionId: connection?.id ?? null },
);

// If the user clicks "Return FetchXML", result is { fetchXml }.
// If they close the window or click the host's back banner, result is null.
const editedFetchXml = (result as { fetchXml?: string } | null)?.fetchXml;
```

### Sending a query to another tool (Send to Tool)

FetchXML Studio discovers send targets automatically — no configuration needed. It calls the host capability registry (`toolboxAPI.invocation.findToolsByCapability("fetchxml")`) to find installed tools that declare the `fetchxml` capability in their own `pptb.config.json`, and lists them in a **Send to Tool** button (a single button for one match, a dropdown picker for several). The button is hidden when no other `fetchxml`-capable tools are installed, or on hosts that don't support capability discovery.

The launch is one-way (`noReturn: true`) — FetchXML Studio hands the current query off and does not wait for data back. Any tool that wants to appear here only needs to declare:

```json
{ "invocation": { "version": "1.0.0", "capabilities": ["fetchxml"], "prefill": { "properties": { "fetchXml": { "type": "string" } } } } }
```

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'feat: add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Commit Convention
This project follows [Conventional Commits](https://www.conventionalcommits.org/):
- `feat:` New feature
- `fix:` Bug fix
- `refactor:` Code refactoring
- `docs:` Documentation
- `chore:` Maintenance

## 📋 Roadmap

- [ ] Import/export query definitions (JSON)
- [ ] Quick query templates
- [ ] Query history
- [ ] Query performance insights
- [ ] Syntax validation in Monaco

## 📄 License

Licensed under the GNU Affero General Public License v3.0 (AGPL-3.0-only). See [LICENSE](LICENSE) for details.

## 🙏 Acknowledgments

- [XrmToolBox FetchXML Builder](https://github.com/rappen/FetchXMLBuilder) - Original inspiration
- [Power Platform ToolBox](https://github.com/PowerPlatformToolBox/desktop-app) - Host platform
- [Fluent UI](https://react.fluentui.dev/) - UI component library
- [Monaco Editor](https://microsoft.github.io/monaco-editor/) - Code editor

---

Built with ❤️ for the Power Platform community
