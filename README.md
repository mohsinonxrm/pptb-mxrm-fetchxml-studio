# FetchXML Studio for PowerPlatform Toolbox

A powerful, modern FetchXML query builder and data explorer for [PowerPlatform Toolbox](https://github.com/PowerPlatformToolBox/desktop-app). Inspired by the XrmToolBox FetchXML Builder, reimagined with React 19, Fluent UI v9, and seamless Dataverse integration.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![React](https://img.shields.io/badge/React-19-61DAFB?logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178C6?logo=typescript)
![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft)

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
- **Export to Excel** — Native Excel export with proper data types (numbers, dates, currencies)
- **Formatted value options** — Export formatted, raw, or both value types
- **Record deletion** — Delete selected records with confirmation
- **Bulk delete jobs** — Submit async bulk delete operations to Dataverse
- **Run workflows** — Execute on-demand workflows on selected records

### 🎨 User Experience
- **Dark/Light themes** — Follows PowerPlatform Toolbox theme with Fluent UI tokens
- **Lazy metadata loading** — Loads only what's needed, when it's needed
- **Intelligent caching** — In-memory cache prevents duplicate API calls
- **Resizable panes** — Adjust the layout to your preference
- **Keyboard shortcuts** — Execute queries, copy XML, and more

### 🔒 Privilege-Aware
- **Security checks** — Validates user privileges before operations
- **Export privilege check** — Only shows export option if user has access
- **Delete privilege check** — Validates delete permissions per entity

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
- [PowerPlatform Toolbox](https://github.com/PowerPlatformToolBox/desktop-app) desktop application
- Connection to a Dataverse environment

### Installation
FetchXML Studio is available as a tool in PowerPlatform Toolbox. Install it from the tool gallery or load as a custom tool.

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

| Technology | Purpose |
|------------|---------|
| **React 19** | UI framework with latest features (transitions, actions) |
| **TypeScript 5.9** | Type-safe development |
| **Vite** | Fast build tooling and HMR |
| **Fluent UI v9** | Microsoft's design system (Tree, DataGrid, Tabs, etc.) |
| **Monaco Editor** | VS Code's editor for XML editing |
| **react-window** | Virtualized list rendering for large datasets |
| **exceljs** | Native Excel file generation |
| **@pptb/types** | PowerPlatform Toolbox API types |

## 📁 Project Structure

```
src/
├── app/
│   └── AppShell.tsx              # Main layout with split panes
├── features/fetchxml/
│   ├── api/
│   │   ├── pptbClient.ts         # Dataverse API wrapper
│   │   ├── dataverseMetadata.ts  # Metadata fetching
│   │   ├── excelExport.ts        # Excel export logic
│   │   └── formattedValues.ts    # OData formatted value handling
│   ├── model/
│   │   ├── nodes.ts              # TypeScript node definitions
│   │   ├── fetchxml.ts           # FetchXML generation
│   │   ├── fetchxmlParser.ts     # FetchXML parsing
│   │   ├── layoutxml.ts          # LayoutXML generation
│   │   └── operators.ts          # Operator definitions by type
│   ├── state/
│   │   ├── builderStore.tsx      # React context state management
│   │   └── cache.ts              # Metadata caching
│   └── ui/
│       ├── LeftPane/             # Tree view and properties
│       ├── RightPane/            # Tabs, editor, grid
│       ├── Toolbar/              # Entity selector, view picker
│       ├── Dialogs/              # Delete, bulk delete, workflow
│       └── Settings/             # Display preferences
└── shared/
    ├── components/               # Reusable pickers
    ├── hooks/                    # Custom React hooks
    └── utils/                    # Debug utilities
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
- [PowerPlatform Toolbox](https://github.com/PowerPlatformToolBox/desktop-app) - Host platform
- [Fluent UI](https://react.fluentui.dev/) - UI component library
- [Monaco Editor](https://microsoft.github.io/monaco-editor/) - Code editor

---

Built with ❤️ for the Power Platform community
