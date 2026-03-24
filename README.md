# FetchXML Studio for Power Platform ToolBox

A powerful, modern FetchXML query builder and data explorer for [Power Platform ToolBox](https://github.com/PowerPlatformToolBox/desktop-app). Inspired by the XrmToolBox FetchXML Builder, reimagined with React 18, Fluent UI v9, and seamless Dataverse integration.

![Version](https://img.shields.io/badge/version-1.1.0-blue)
![React](https://img.shields.io/badge/React-18-61DAFB?logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178C6?logo=typescript)
![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft)
![PPTB Types](https://img.shields.io/badge/%40pptb%2Ftypes-1.2.0-orange)

## ✨ Features

### 🔧 Visual Query Builder
- **Tree-based hierarchy** — Build FetchXML queries visually with entities, attributes, filters, orders, and link-entities
- **Smart property editors** — Context-aware panels for each node type (fetch options, conditions, relationships)
- **Full FetchXML support** — Aggregate queries, grouping, distinct, top/count, paging, and advanced hints
- **Nested filters** — Create complex AND/OR filter groups with unlimited nesting
- **Link-entity relationships** — Browse and add 1:N, N:1, and N:N relationships with inner/outer joins

### 📊 Results Grid
- **Virtualized DataGrid** — Handle large result sets with smooth scrolling (powered by react-window)
- **Rich cell rendering** — Power Apps-style display for lookups, option sets, dates, currency, and more
- **Multi-column sorting** — Click headers to sort; Shift+click for multi-column sort
- **Row selection** — Select single or multiple records for bulk operations
- **Resizable & reorderable columns** — Customize your view with drag-and-drop columns
- **Formatted values** — Display OData formatted values or raw values (configurable)

### 📝 FetchXML Editor
- **Monaco editor** — Full-featured XML editor with syntax highlighting and line numbers
- **Bi-directional editing** — Edit XML directly and parse back to the visual builder
- **Copy to clipboard** — One-click copy of generated FetchXML
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
- **Bulk delete jobs** — Submit async `BulkDelete` jobs to Dataverse with job-tracking URL
- **Batch delete** — Parallel per-record delete for small sets (up to 100 records) with progress and ETA
- **Run workflows** — Execute on-demand workflows on selected records with batch execution and progress tracking

### 📋 Bulk Attribute Selection
- **Select Attributes dialog** — Multi-select attributes from a searchable DataGrid; accessible from entity and link-entity context menus
- **Smart updates** — Adds new selections and removes deselected attributes while preserving existing order and properties
- **Search & filter** — Real-time filtering across logical name, display name, and data type columns

### 🎨 User Experience
- **Dark/Light themes** — Follows Power Platform ToolBox theme with Fluent UI tokens
- **Lazy metadata loading** — Loads only what's needed, when it's needed
- **Intelligent caching** — In-memory cache prevents duplicate API calls per session
- **Resizable panes** — Adjust split-pane layout to your preference
- **Keyboard shortcuts** — `Ctrl+Enter` to execute query, copy XML to clipboard
- **Display settings** — Toggle logical names vs display names in column headers; choose formatted, raw, or both value modes in the grid and exports

### 🔒 Privilege-Aware
- **Security checks** — Validates user privileges before every destructive or restricted operation
- **Export privilege check** — Only shows Dataverse export option if user has access
- **Delete privilege check** — Validates entity-specific delete permissions before enabling delete actions
- **Bulk delete privilege check** — Validates `prvBulkDelete` before surfacing bulk delete
- **Workflow privilege check** — Validates `prvReadWorkflow` + `prvWorkflowExecution` before showing workflow picker
- **View save privilege check** — Validates `prvWriteQuery` / `prvWriteCustomization` / `prvPublishCustomization` / `prvWriteUserQuery` before save operations

## 🖼️ Interface Overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│  [Entity Selector ▼]  [Load View ▼]  [Save View]           [⚙ Settings]│
├──────────────────────┬──────────────────────────────────────────────────┤
│                      │  [FetchXML]  [LayoutXML]  [Results]   [▶ Execute]│
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
FetchXML Studio is available as a tool in Power Platform ToolBox. Install it from the tool gallery or load as a custom tool.

## 🛠️ Development

### Setup

```bash
# Clone the repository
git clone https://github.com/mohsinonxrm/pptb-fetchxml-studio.git
cd pptb-fetchxml-studio

# Install dependencies
npm install

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

# Finalize for publishing (build + shrinkwrap)
npm run finalize-package
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
| **Fluent UI v9** | 9.72 | Microsoft design system (Tree, DataGrid, Tabs, Drawer, etc.) |
| **Monaco Editor** | 0.54 | VS Code XML editor for FetchXML authoring |
| **react-window** | 2 | Virtualized list rendering for large datasets |
| **ExcelJS** | 4.4 | Native `.xlsx` generation with typed cells |
| **@pptb/types** | 1.2.0 | Power Platform ToolBox host API types (`window.dataverseAPI`, `window.toolboxAPI`) |

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
│   │   └── formattedValues.ts    # OData @FormattedValue annotation helpers
│   ├── model/
│   │   ├── nodes.ts              # FetchXML node TypeScript definitions
│   │   ├── fetchxml.ts           # FetchXML XML generation
│   │   ├── fetchxmlParser.ts     # FetchXML XML → node tree parser
│   │   ├── fetchxmlIntellisense.ts # Intellisense / autocomplete helpers
│   │   ├── layoutxml.ts          # LayoutXML generation
│   │   ├── operators.ts          # Operator definitions by attribute type
│   │   ├── displaySettings.ts    # Display settings types and defaults
│   │   └── treeUtils.ts          # Tree traversal utilities
│   ├── state/
│   │   ├── builderStore.tsx      # React context + reducer state management
│   │   └── cache.ts              # Per-session in-memory metadata cache
│   └── ui/
│       ├── LeftPane/             # Tree view + context-aware properties panel
│       │   └── PropertiesPanel/
│       │       └── editors/      # Node-specific property editors
│       ├── RightPane/            # Monaco editor, LayoutXML viewer, results grid
│       ├── Toolbar/              # Entity selector, load view picker, save button
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
| All condition operators | ✅ |
| Link-entity (joins) | ✅ |
| Inner/outer joins | ✅ |
| N:N relationships | ✅ |
| Orders (sorting) | ✅ |
| Multi-column sort | ✅ |
| Aggregate queries | ✅ |
| Groupby | ✅ |
| Distinct | ✅ |
| Top/Count | ✅ |
| Paging with cookies | ✅ |
| Value-of conditions | ✅ |
| Entity name on conditions | ✅ |
| Filter link-entity (any/all) | ✅ |
| Query hints | ✅ |

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
- [ ] Web API code generation
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
