# Changelog

All notable changes to FetchXML Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
