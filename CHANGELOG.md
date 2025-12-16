# Changelog

All notable changes to FetchXML Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-12-15

### 🎉 Initial Release

FetchXML Studio v1.0.0 - A powerful FetchXML query builder for PowerPlatform Toolbox.

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
- PowerPlatform Toolbox integration via @pptb/types

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
