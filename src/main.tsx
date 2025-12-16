import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { AppShell } from "./app/AppShell";
// Import debug utility to expose functions to window
import "./shared/utils/debug";

createRoot(document.getElementById("root")!).render(
	<StrictMode>
		<AppShell />
	</StrictMode>
);
