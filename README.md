# PowerPlatform Toolbox - FetchXML Builder

A modern FetchXML query builder for PowerPlatform Toolbox, built with React 19, TypeScript, and Fluent UI v9.

## Features

- **Tree-based query builder** with hierarchical FetchXML structure
- **Relationship picker** with categorized display (1:N, N:1, N:N)
- **Virtualized data grid** for large result sets
- **Live FetchXML preview** with Monaco editor
- **Lazy metadata loading** with intelligent caching
- **Dark theme support** aligned with PowerPlatform Toolbox

## Development

### Setup

```bash
npm install
npm run dev
```

### Building

```bash
npm run build
```

### Debug Logging

The app includes a feature-flag-based debug logging system. To enable specific feature logging, open the browser console and use:

```javascript
// Enable tree expansion debugging
enableDebug("treeExpansion");

// Enable relationship picker debugging
enableDebug("relationshipPicker");

// Enable link entity editor debugging
enableDebug("linkEntityEditor");

// Enable properties panel debugging
enableDebug("propertiesPanel");

// Enable all debug logging
enableAllDebug();

// Disable specific category
disableDebug("treeExpansion");

// Disable all debug logging
disableAllDebug();
```

Debug categories:

- `treeExpansion` - Tree node expansion and path finding
- `relationshipPicker` - Relationship loading and selection
- `linkEntityEditor` - Link-entity configuration
- `propertiesPanel` - Properties panel routing and context

## Tech Stack

- **React 19** with TypeScript
- **Vite** for build tooling
- **Fluent UI v9** for all UI components
- **Monaco Editor** for XML preview
- **@pptb/types** for PowerPlatform Toolbox integration

## Architecture

```
src/
  app/              # App shell and layout
  features/         # Feature modules (fetchxml builder)
  shared/           # Shared components and utilities
```

See `.github/copilot-instructions.md` for detailed architecture and implementation guidance.

---

# React + TypeScript + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Babel](https://babeljs.io/) (or [oxc](https://oxc.rs) when used in [rolldown-vite](https://vite.dev/guide/rolldown)) for Fast Refresh
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/) for Fast Refresh

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
	globalIgnores(["dist"]),
	{
		files: ["**/*.{ts,tsx}"],
		extends: [
			// Other configs...

			// Remove tseslint.configs.recommended and replace with this
			tseslint.configs.recommendedTypeChecked,
			// Alternatively, use this for stricter rules
			tseslint.configs.strictTypeChecked,
			// Optionally, add this for stylistic rules
			tseslint.configs.stylisticTypeChecked,

			// Other configs...
		],
		languageOptions: {
			parserOptions: {
				project: ["./tsconfig.node.json", "./tsconfig.app.json"],
				tsconfigRootDir: import.meta.dirname,
			},
			// other options...
		},
	},
]);
```

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from "eslint-plugin-react-x";
import reactDom from "eslint-plugin-react-dom";

export default defineConfig([
	globalIgnores(["dist"]),
	{
		files: ["**/*.{ts,tsx}"],
		extends: [
			// Other configs...
			// Enable lint rules for React
			reactX.configs["recommended-typescript"],
			// Enable lint rules for React DOM
			reactDom.configs.recommended,
		],
		languageOptions: {
			parserOptions: {
				project: ["./tsconfig.node.json", "./tsconfig.app.json"],
				tsconfigRootDir: import.meta.dirname,
			},
			// other options...
		},
	},
]);
```
