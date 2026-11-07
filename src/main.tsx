import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { AppShell } from "./app/AppShell";
// Import debug utility to expose functions to window
import "./shared/utils/debug";
// Configure Monaco Editor to load from the local bundle instead of CDN.
// This avoids any CSP issues with cdn.jsdelivr.net and ensures the editor
// stylesheet, syntax highlighting and line numbers always render correctly.
//
// Monaco's root entry ("monaco-editor") loads every language and, since 0.56, bundles the CSS,
// HTML, JSON and TypeScript worker scripts (~10 MB). FXS only shows XML, C#, JavaScript, shell,
// SQL (the Code tab's FetchXML to SQL) and plain text, so it uses the tree-shakeable entry points
// Monaco 0.56 added instead: the editor core, all editor features, and just the language
// definitions it needs (#145).
import * as monaco from "monaco-editor/editor";
import "monaco-editor/features/register.all";
import "monaco-editor/languages/definitions/xml/register";
import "monaco-editor/languages/definitions/csharp/register";
import "monaco-editor/languages/definitions/javascript/register";
import "monaco-editor/languages/definitions/shell/register";
import "monaco-editor/languages/definitions/sql/register";
import { loader } from "@monaco-editor/react";

// Run Monaco's editor worker on the main thread. PPTB's CSP has no `worker-src`, so `blob:`
// workers are blocked there anyway. Monaco ≤ 0.55 found no worker location and fell back with one
// console warning. Since 0.56 it first tries a bundler URL, which Vite doesn't resolve (a 404 in
// dev, a `data:` URL with broken relative imports in the build), and logs errors before falling
// back. Throwing here takes the same fallback path as before, with the same single warning.
self.MonacoEnvironment = {
	getWorker(_workerId, label) {
		throw new Error(`FXS runs Monaco's "${label}" worker on the main thread.`);
	},
};
loader.config({ monaco });

createRoot(document.getElementById("root")!).render(
	<StrictMode>
		<AppShell />
	</StrictMode>,
);
