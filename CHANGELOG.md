# Changelog

All notable changes to FetchXML Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.2.3-beta.2] - 2026-07-01

### 🐛 Fixed

- **Send to Tool no longer spins forever.** The launch is now fire-and-forget — previously the button awaited `launchTool`, whose promise doesn't resolve until the callee window closes, so the spinner stayed up for the whole session (reported when sending FXS → Bulk Data Studio). This is purely a caller-side fix; a callee does **not** need to declare a `returnTopic` to avoid it.
- **Always shows a "Send to Tool" dropdown.** Previously a single discovered tool collapsed the toolbar control into a button labelled with that tool's name; it now consistently reads "Send to Tool" and lists the target(s) in the menu.
- **Each target shows its tool id on its own line** beneath the display name in the menu.
- **FetchXML Studio no longer lists itself** as a target. Self-exclusion now matches on the runtime tool id from `getToolContext()` (with the npm package id as a fallback), since discovery reports the runtime id (e.g. `npm-mohsinonxrm-pptb-fetchxml-studio`) rather than the npm package name.

## [1.2.3-beta.1] - 2026-06-17

### ✨ Added

#### Tool-to-Tool (T2T) — capability-based tool discovery

The "Send to Tool" target list is now populated automatically via the PPTB capability registry instead of a manually maintained list. Requires **`@pptb/types` ≥ 1.2.3-beta.0** and a host that exposes capability discovery.

- **Automatic discovery** — FetchXML Studio calls `toolboxAPI.invocation.findToolsByCapability("fetchxml")` to find installed tools that declare the `fetchxml` capability, excludes itself, and lists them in the Send to Tool button (single button for one match, dropdown for several). Hidden when no matches or when the host lacks discovery — feature-detected at runtime, so older hosts degrade gracefully.
- **One-way send** — Launches now pass `noReturn: true`, the dedicated flag for the "Send To" pattern, which suppresses the callee's "Return to [Caller]" banner since FXS isn't waiting for data back.

### 🔧 Changed

- **Removed the manual "Tool Integration" settings.** The Settings → Tool Integration section and the `targetTools` display setting are gone; discovery replaces them entirely.

### 🏗️ Technical

- Bumped `@pptb/types` `1.2.2-beta.1` → `1.2.3-beta.0`, which now types `invocation.findToolsByCapability` / `getKnownCapabilityTags`, the `capabilities` field on `InvocationConfig`, the `CapabilityTag`/`KnownCapabilityTag` types, and `launchTool`'s `noReturn` option. Removed the local `InvocationAPI` interface in `invocation.ts` in favor of the now-published host types.
- `src/features/fetchxml/api/invocation.ts` — Added `DiscoveredTool` + `discoverFetchXmlTools()` (feature-detected; filters out our own id; defaults `name` to `id`); added `noReturn: true` to `sendFetchXmlToTool`.
- `src/shared/hooks/useSendToTools.ts` — New hook; discovers send targets once on mount.
- `src/app/AppShell.tsx` — Uses `useSendToTools`; Send to Tool slot now gated on `isT2TSupported() && sendToTools.length > 0`.
- `src/features/fetchxml/ui/Toolbar/SendToToolButton.tsx` — Prop changed from `targetTools: string[]` to `tools: DiscoveredTool[]`; labels use the discovered tool name.
- `src/features/fetchxml/ui/Settings/SettingsDrawer.tsx` — Removed the Tool Integration section, its handlers/state/styles, and now-unused imports.
- `src/features/fetchxml/model/displaySettings.ts` — Removed `targetTools` from `DisplaySettings` and the defaults.

## [1.2.2-beta.2] - 2026-06-16

### ✨ Added

#### Tool-to-Tool (T2T) Invocation — Callee (prefill + return)

FetchXML Studio can now be *launched by* another PPTB tool, pre-populated with an inbound query, and hand a query back. Completes the round-trip alongside the existing caller support.

- **Inbound prefill** — On launch via T2T, FetchXML Studio reads its launch context and pre-populates the builder. Resolution priority is `fetchXml` → `viewRef` → `entityLogicalName`: a `fetchXml` string is loaded directly; a `viewRef` has its FetchXML retrieved from the referenced `savedquery`/`userquery`; and a bare `entityLogicalName` starts a fresh query on that entity. When `fetchXml` is present the root entity is **derived from the FetchXML itself** and any `entityLogicalName` hint is ignored, because a mismatched value would break this metadata-driven tool's lookups.
- **Return FetchXML button** — When launched as a callee, a "Return FetchXML" button appears in the FetchXML toolbar. It returns `{ fetchXml }` (the current editor-or-tree query) to the caller via `returnData()`, which the host resolves as the caller's `launchTool(...)` result. This is intentionally separate from the host-injected "Return to [Caller]" banner, which only navigates back and resolves the caller's promise with `null`.
- **`pptb.config.json`** — Added the `fetchxml` capability tag (for discovery via `findToolsByCapability`) and a `returnTopic` describing the `{ fetchXml }` return shape. Relaxed the prefill contract so `entityLogicalName` is optional/hint-only.

### 🏗️ Technical

- `src/features/fetchxml/api/invocation.ts` — Added callee helpers: `getLaunchContext()`, `resolveLaunchPrefill(context)` (returns a `{ kind: "fetchxml" | "entity" }` action following the `fetchXml` → `viewRef` → `entityLogicalName` priority; resolves `viewRef` by reusing the `dataverseAPI.queryData` view-read route filtered to a single record), `returnDataToInvokingTool(data)`, and `returnFetchXmlToInvokingTool(fetchXml)`. Added `FetchXmlStudioLaunchPrefill` type and `LaunchPrefillAction` union.
- `src/shared/hooks/useLaunchContext.ts` — New hook. Calls `getLaunchContext()` once on mount and exposes `{ loading, isCallee, context }` so detection happens in one place.
- `src/app/AppShell.tsx` — Consumes the prefill exactly once on launch (loads incoming FetchXML via `builder.loadFetchXml`), and passes `isCallee` to `PreviewTabs`.
- `src/features/fetchxml/ui/RightPane/PreviewTabs.tsx` — Added `isCallee?` prop and a `ReturnFetchXmlButton` (gated on `isCallee`) rendered next to Execute; uses a `getCurrentXml()` helper that returns the live editor buffer when in editor mode.

### 🔧 Changed

- **"Send to Tool" button visibility** — The caller button is now hidden entirely unless at least one target tool is configured under Settings → Tool Integration (previously it rendered as a disabled button with a tooltip). The two T2T toolbar actions are now strictly capability-gated and never change meaning by mode: **Return FetchXML** shows only when launched as a callee (`isCallee`); **Send to Tool** shows only when targets are configured.

## [1.2.2-beta.1] - 2026-05-25

### ✨ Added

#### Tool-to-Tool (T2T) Invocation — Caller

FetchXML Studio can now send the active query to any other PPTB tool that accepts a FetchXML prefill payload. Requires **PPTB host ≥ 1.2.2**.

- **Send to Tool button** — A new "Send to Tool" button appears in the toolbar (alongside Save View) whenever the PPTB host exposes the `toolboxAPI.invocation` API. Three interaction modes depending on how many target tools are configured:
  - *No tools configured* — Button is disabled with a tooltip pointing to Settings → Tool Integration.
  - *One tool configured* — Single button; click launches the tool directly.
  - *Multiple tools configured* — Dropdown menu listing all configured tools; click an item to launch.
- **Settings → Tool Integration section** — New section in the Settings drawer to manage the list of target tool npm package IDs (e.g. `@linked365/pptb-bulk-data-studio`). Supports add (text input + Enter/Add button) and remove (× button per item). The list is persisted via `toolboxAPI.settings` alongside other display preferences.
- **Active connection forwarding** — The caller automatically retrieves the currently active Dataverse connection and forwards its ID as `primaryConnectionId` when launching the target tool, so the callee opens against the same environment.
- **`pptb.config.json`** — Added PPTB callee contract file at the repository root. Declares the prefill schema this tool accepts when *receiving* a T2T invocation from another tool (see [Callee Contract](#-callee-contract-pptbconfigjson) in the README). `viewRef` is modelled as a Dataverse EntityReference: required fields `id` (UUID string) and `entityLogicalName` (enum: `"savedquery"` | `"userquery"`).

### 🏗️ Technical

- `src/features/fetchxml/api/invocation.ts` — New module. Defines a local `InvocationAPI` interface (mirrors the PPTB host API not yet in `@pptb/types`) and exposes `isT2TSupported()` and `sendFetchXmlToTool(targetToolId, { fetchXml, entityLogicalName })`.
- `src/features/fetchxml/ui/Toolbar/SendToToolButton.tsx` — New toolbar component. Handles the disabled / single / multi-tool rendering variants; calls `sendFetchXmlToTool`; surfaces errors via `toolboxAPI.utils.showNotification`.
- `src/features/fetchxml/model/displaySettings.ts` — Added `targetTools: string[]` (default `[]`) to `DisplaySettings` interface and `defaultDisplaySettings`.
- `src/features/fetchxml/ui/Settings/SettingsDrawer.tsx` — Added Tool Integration section with add/remove UI. Added `Input`, `Field`, `Add20Regular`, `Dismiss16Regular`, `PlugConnected20Regular` to imports.
- `src/features/fetchxml/ui/RightPane/PreviewTabs.tsx` — Added `sendToToolButton?: ReactNode` prop; rendered unconditionally (all tabs) next to `saveViewButton`.
- `src/app/AppShell.tsx` — Imports `SendToToolButton` + `isT2TSupported`; passes `sendToToolButton` prop to `PreviewTabs` gated on `isT2TSupported()`.
- `package.json` — `features.minAPI` bumped from `1.2.0` → `1.2.2` to reflect the new PPTB host requirement. `scheduler` added to `devDependencies` (required peer dep of `@fluentui/react-context-selector` that was missing from the install tree).

#### Packaging fixes (resolves #33)

- **`icon` path corrected** — `"icon"` field in `package.json` was missing the `icons/` directory prefix (`"icon-insider.svg"` → `"icons/icon-insider.svg"`). Icon was not resolvable by PPTB at install time.
- **`files` array expanded** — Added `"icons"`, `"index.html"`, `"LICENSE"`, `"README.md"`, `"CHANGELOG.md"`, `"SECURITY.md"` alongside the existing `"dist"` and `"npm-shrinkwrap.json"` entries. Previously the `icons/` folder (referenced by the `"icon"` field) and all root-level docs were absent from the published package.
- **Production source maps disabled** — Added `sourcemap: false` to `vite.config.ts` `build` options. Monaco worker source maps were silently adding ~17 MB to the package; disabling them brings the published size down to ~3–4 MB.

---

## [1.2.1] - 2026-04-07

### ✨ Added

- **Copy button in editor mode** — The FetchXML editor toolbar now includes a Copy button when in editor mode, making it consistent with the read-only toolbar. Previously Copy was only available when the editor was not in edit mode.

### 🔧 Fixed

- **Results grid column overlap** — Column widths are now estimated dynamically from header text (capped at 400 px) so adjacent columns no longer overlap on first render. All cells use `TableCellLayout truncate` for clean overflow.
- **Results grid header/row misalignment** — Removed a stray CSS class from `DataGridBody` that was causing header and row widths to drift out of sync when the grid had overflowing content.
- **Alias validation in property editors** — Typing an invalid alias character (space, hyphen, special character) in `AttributeEditor` or `LinkEntityEditor` now sanitizes the value in real time: invalid characters are stripped, and a leading digit is prefixed with `_`.
- **Alias validation in parser** — When parsing FetchXML (both *Parse to Tree* and *Execute*), invalid aliases are now automatically sanitized. A parse warning describes the correction (e.g. *"Alias 'this is id' had invalid characters and was corrected to 'thisisid'."*) instead of silently forwarding the invalid alias to Dataverse where it would cause an API error.
- **Execute blocked when alias warnings are present** — Clicking Execute in editor mode now validates the XML completely before sending it to Dataverse. If parse warnings exist (e.g. an invalid alias that wasn't auto-corrected via *Parse to Tree*), execution is blocked and the warning is surfaced in a message bar above the editor with a hint to use *Parse to Tree* to auto-correct. Previously the bad alias was silently sent to Dataverse and the error appeared in the Results tab.
- **Editor state preserved across tab switches** — Switching from the FetchXML tab to Results or LayoutXML and back no longer resets editor mode or discards the edited XML. The editor component is now always mounted (hidden via `display: none` when not active) so Monaco state — editor mode toggle, cursor position, edit buffer — is fully preserved.
- **Validate contradictory messages** — *Validate* no longer shows both a "Valid" success message and warnings simultaneously. When warnings are present, validate now shows each warning plus an "Action required" notice and stops — the "Valid" success message is only shown when there are zero warnings.
- **Parse to Tree corrections shown as informational** — Auto-corrected alias warnings from *Parse to Tree* are now shown with `intent="info"` (blue) rather than `intent="warning"` (yellow), correctly communicating that the tool has already resolved the problem.
- **Execute API errors surfaced** — Dataverse API errors that occur during query execution are now captured and displayed in a message bar in the Results tab. Previously they were only visible in the browser console.

### 🏗️ Technical

- Added `src/features/fetchxml/model/aliasUtils.ts` — `isValidAlias()` and `sanitizeAlias()` shared utilities enforcing the Dataverse alias constraint (`[A-Za-z_][A-Za-z0-9_]*`).
- `fetchxmlParser.ts` — `parseAttribute` and `parseLinkEntity` now import `sanitizeAlias`; invalid aliases are corrected in place and described in the parse warning rather than propagated as-is.
- `PreviewTabs.tsx` — `FetchXmlEditor` is always mounted; `handleExecute` blocks on parse warnings; navigating back to the FetchXML tab clears any stale editor validation error banner.
- `AppShell.tsx` — `handleExecute` accepts an optional `xmlOverride` (used when editor mode is active); `executeError` state threads Dataverse API errors back to the `PreviewTabs` message bar.

---

## [1.2.0] - 2026-04-03

### ✨ Added

#### Query Scope Settings
- **Entity Scope Mode** — New setting (persisted across sessions) to control which entities appear in the entity dropdown:
  - **Publisher + Solution** — Select a publisher, then narrow to specific solutions; only entities belonging to those solutions are shown
  - **Solution Only** — Skip the publisher step and pick solutions directly
  - **All Entities** — Show every entity in the environment (no solution filter)
- **Advanced Find Only toggle** — When enabled (default), entities and attributes are limited to those marked `IsValidForAdvancedFind = true`, matching the classic Advanced Find behavior. Disabling shows all entities and attributes — including system and developer-only tables — useful for building admin or integration queries. The filter is applied locally so toggling is instant with no additional API call.
- Scope mode and Advanced Find preference are persisted individually via `toolboxAPI.settings` and survive across sessions.

#### Monaco Editor — Local Bundling
- Monaco is now fully bundled in `dist/` instead of being fetched from a CDN at runtime.
- Eliminates CSP errors and ensures line numbers, syntax highlighting, and all editor features work correctly in the PPTB host environment without network access to external CDNs.

### 🔧 Fixed

- **Internal solutions in Solution-only mode** — Solutions from `isreadonly = true` publishers (Microsoft, system publishers) are now excluded from the solution picker, consistent with what Publisher + Solution mode shows. Previously some Microsoft-owned visible solutions appeared in the list.
- **Parse-to-tree no longer resets the entity** — Pasting FetchXML into the editor and clicking "Parse to Tree" no longer wipes the selected entity. The validation that auto-clears the entity selection now only fires when there is an active solution filter in place (i.e. specific solutions have been selected) and the entity has dropped out of scope — not when the entity was set programmatically via the XML editor.
- **Entity list cache race condition** — `loadAllEntities` no longer accepts an `advancedFindOnly` parameter. It always fetches all entities from the server (no `$filter=IsValidForAdvancedFind eq true`). The Advanced Find filter is now applied locally in the EntitySelector `memo`. This eliminates a bug where an earlier `advancedFindOnly=true` load could poison the cache and cause subsequent `advancedFindOnly=false` calls to silently return the wrong (filtered) set.
- **Privilege check console noise** — `checkPrivilegeByName` errors for privileges that don't exist in the environment (e.g. `prvDeleteGitorganization` for system tables) are now logged as `console.warn` instead of `console.error`. The UI behavior is unchanged.
- **CSP exception** — Removed the invalid `*.crm*.dynamics.com` wildcard from `cspExceptions` (browsers reject multi-depth subdomain wildcards). The only `connect-src` exception kept is `https://*.dynamics.com`.

### 🏗️ Technical

- `useAccessMode` preloads only the AF-valid entity metadata cache (`allEntityMetadataCache`) on startup — used by publisher-solution and solution-only modes. The all-entities cache is populated lazily the first time "All Entities" scope mode is used.
- `loadAllEntities` / `useLazyMetadata.loadEntities` signatures simplified (no `advancedFindOnly` param — always fetches everything).
- `vite.config.ts` — `manualChunks` splits Monaco into its own chunk (`monaco-editor-*.js`, ~3.7 MB; `monaco-editor-*.css`, ~145 kB).
- TypeScript strict compile: **zero errors** (`tsc -b --noEmit`).

---

## [1.1.0] - 2026-03-23

### ⬆️ Upgraded

- **`@pptb/types` 1.0.7 → 1.2.0** — now the authoritative source for all PPTB host API types (`window.dataverseAPI`, `window.toolboxAPI`). Added to `tsconfig.app.json` `types` array so global augmentations apply project-wide without per-file triple-slash references.
- **`features.minAPI`** set to `"1.2.0"` in `package.json` to enforce the correct PPTB host version at install time.

### ✨ Added

#### Native File Save Dialog
- Both export paths (local ExcelJS export and Dataverse `ExportToExcel` action) now use `window.toolboxAPI.fileSystem.saveFile` when running inside the PPTB desktop host, presenting a native OS save-file dialog instead of a forced browser blob download.
- Graceful fallback to browser blob/anchor download when running in standalone or dev context (i.e. when `window.toolboxAPI` is not available).

#### Connection Environment Awareness
- `PptbContext` now exposes `environment?: "Dev" | "Test" | "UAT" | "Production"` sourced from the active `DataverseConnection`.
- Subscribes to `connection:updated`, `connection:created`, and `connection:deleted` PPTB events to keep the field in sync dynamically.

#### Package Manifest Hardening
- Added top-level `"icon"` field pointing to `icons/icon.svg` (replacing deprecated `configurations.iconUrl`).
- `"features"` block added: `{ "multiConnection": "none", "minAPI": "1.2.0" }`.
- `cspExceptions` migrated from plain-string arrays to the required object format with `domain` + `exceptionReason` fields (+ `optional: true` on style-src).
- Added `"validate": "pptb-validate"` npm script for pre-publish manifest validation.
- Added `"finalize-package": "npm run build && npm shrinkwrap"` convenience script.

### ♻️ Refactored

#### `pptbClient.ts`
- Removed the entire hand-written `declare global { interface Window { dataverseAPI?: { ... } } }` block (~60 lines). The `window.dataverseAPI` global is now typed entirely by `@pptb/types`.
- Added local `DataverseFetchXmlResponse` interface to capture OData response annotations beyond what `@pptb/types` `FetchXmlResult` declares (`@Microsoft.Dynamics.CRM.totalrecordcount`, `morerecords`, `pagingcookie`, `totalrecordcountlimitexceeded`). These are OData wire annotations returned directly by Dataverse — not fabricated by PPTB — so the cast is safe at runtime.
- `downloadBase64File()` changed from `void` to `async Promise<void>`; uses `fileSystem.saveFile` when available.

#### `usePptbContext.ts`
- Removed the entire hand-written `declare global { interface Window { toolboxAPI?: { ... } } }` block. The `window.toolboxAPI` global is now typed entirely by `@pptb/types`.
- `PptbContext` interface: removed legacy `organizationId?: string` (not present in `DataverseConnection` v1.2.0 types); replaced with `environment?: "Dev" | "Test" | "UAT" | "Production"`.

#### `excelExport.ts`
- `downloadExcelFile()` changed from `void` to `async Promise<void>`; uses `fileSystem.saveFile` when available.

#### `AppShell.tsx`
- Both download call sites updated with `await` to match the new async signatures.

### 🔧 Technical

- TypeScript strict compile: **zero errors** (`tsc -b --noEmit`).
- `pptb-validate`: **✔ Validation passed**.

---

## [1.0.6] - 2026-01-14

### ✨ Added

#### Select Attributes Dialog (#12)
- **Bulk Attribute Selection**: New "Select Attributes" modal dialog for efficient multi-attribute selection
  - Accessible from entity and link-entity context menus in the tree view
  - Works on root entity, link-entities, and nested link-entities at any level
- **Smart Selection Grid**: DataGrid showing all available entity attributes
  - Three sortable columns: Logical Name, Display Name, and Data Type
  - Pre-selects currently selected attributes with checkboxes
  - Multi-select support for bulk add/remove operations
- **Search & Filter**: Integrated SearchBox for quick attribute filtering
  - Filters across all columns (logical name, display name, data type)
  - Real-time filtering as you type
- **Fixed Headers**: Column headers remain visible while scrolling through attributes
- **Confirmation Required**: Changes only applied when user clicks Apply button
  - Cancel button discards changes without affecting the tree
- **Smart Updates**: Intelligently adds new attributes and removes deselected ones
  - Preserves existing attributes that remain selected
  - Maintains attribute order and properties
- **Performance**: Leverages metadata cache for fast attribute loading
  - Falls back to API loading if attributes not cached
  - Shows only attributes valid for Advanced Find

### 🔧 Technical

- Added `SelectAttributesDialog` component with FluentUI DataGrid
- New `UPDATE_ATTRIBUTES` action in builderStore for bulk attribute operations
- Enhanced TreeView with dialog state management and attribute loading
- Integrated with existing metadata caching system

## [1.0.0] - 2024-12-15

### 🎉 Initial Release

FetchXML Studio v1.0.0 - A powerful FetchXML query builder for Power Platform ToolBox.

### ✨ Features

#### Query Builder
- Tree-based visual query builder with hierarchical FetchXML structure
- Full entity, attribute, filter, order, and link-entity support
- Context-aware properties panel for each node type
- Nested filter groups with AND/OR logic
- Support for all FetchXML operators (60+ operators)
- Aggregate queries with groupby and distinct
- Query hints for performance optimization

#### Relationships & Joins
- Relationship picker with 1:N, N:1, N:N categorization
- Inner and outer join support
- Filter link-entity for any/all/not any/not all scenarios
- Value-of conditions for cross-entity comparisons
- Entity name conditions for outer join filtering

#### Results Grid
- Virtualized DataGrid for large datasets (react-window)
- Rich cell rendering for lookups, option sets, dates, currency
- Multi-column sorting (click + Shift+click)
- Row selection (single and multi-select)
- Resizable and reorderable columns
- Formatted values with raw value toggle

#### FetchXML Editor
- Monaco editor with XML syntax highlighting
- Bi-directional editing (visual ↔ XML)
- Parse FetchXML to visual builder
- Copy to clipboard
- LayoutXML preview tab

#### Views
- Load system and personal views
- Optimized view execution via SavedQuery/UserQuery APIs
- Save queries as personal views
- Update existing views
- Solution-aware save (add to solution)

#### Data Operations
- Export to Excel with native data types
- Delete selected records
- Bulk delete jobs (async)
- Run on-demand workflows on selection
- Batch progress tracking with ETA

#### User Experience
- Dark/Light theme support (follows PPTB theme)
- Lazy metadata loading
- Intelligent caching (per-session)
- Resizable split panes
- Keyboard shortcuts (Ctrl+Enter to execute)

#### Security
- Privilege-aware operations
- Export/Delete privilege checks
- Bulk delete privilege validation
- Workflow execution privilege checks

### 🏗️ Technical

- Built with React 19 and TypeScript 5.9
- Fluent UI v9 components
- Monaco Editor integration
- ExcelJS for native Excel export
- Vite build tooling
- Power Platform ToolBox integration via @pptb/types

### 📚 Documentation

- Comprehensive README with feature overview
- Debug logging system documentation
- Project structure guide
- Contributing guidelines

---

## Development History

### Pre-release commits (dev branch)

- `feat: Beta release enhancements - Settings, Export, and UX improvements`
- `feat(preview): add LayoutXML tab to preview pane`
- `feat(commands): add record action commands - delete, bulk delete, workflow`
- `feat(columns): Add/Edit Columns panels with related entity support`
- `feat(export): add Export to Excel with status feedback and privilege check`
- `feat(paging): fix paging with cookies and page size support`
- `feat(grid): multi-entity metadata & column display name fixes`
- `feat(builder): Sprint 1 - Foundation & Layout improvements`
- `feat(save): implement Save View functionality`
- `feat(sort): Add multi-column sort with FetchXML integration`
- `feat(results): Add Column from Results Grid`
- `feat(views): Load view LayoutXML for column configuration`
- `feat(layout): Add column resize and reorder`
- `feat(layout): Add LayoutXML foundation`
- `feat(filter): Add filter link-entity support for any/not any/all scenarios`
- `feat(condition): Add valueof cross-entity comparison`
- `feat(condition): Add valueof same-row column comparison`
- `feat(condition): Add entityname support for outer join filtering`
- `feat(views): Add Load View picker and optimized view execution`
- `feat(parser): Add FetchXML parser with Monaco editor dialog`
- `feat: Implement Phases 1-4 - DataGrid UI polish`
- `feat: add filter consistency, cascading validation, and global metadata caching`
- `feat(entity-selector): implement advanced filtering with publisher/solution multiselect`
- `feat(ui): add resizable panes with visual grip indicators and command bar`
- `Initial commit - setup main branch`
