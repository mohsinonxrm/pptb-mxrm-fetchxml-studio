import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { AppShell } from "./app/AppShell";
// Import debug utility to expose functions to window
import "./shared/utils/debug";
// Configure Monaco Editor to load from the local bundle instead of CDN.
// This avoids any CSP issues with cdn.jsdelivr.net and ensures the editor
// stylesheet, syntax highlighting and line numbers always render correctly.
import * as monaco from "monaco-editor";
import { loader } from "@monaco-editor/react";
loader.config({ monaco });

createRoot(document.getElementById("root")!).render(
	<StrictMode>
		<AppShell />
	</StrictMode>,
);
