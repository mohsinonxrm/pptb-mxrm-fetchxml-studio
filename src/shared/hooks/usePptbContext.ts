/**
 * Hook to access PPTB host context (connection, theme, etc.)
 * Subscribes to PPTB events for dynamic updates
 */

import { useState, useEffect } from "react";
import { getPrimaryConnection } from "../../features/fetchxml/api/host";
import { connectionKey } from "./connectionIdentity";
import { debugLog } from "../utils/debug";

export interface PptbContext {
	theme: "light" | "dark";
	connected: boolean;
	environmentUrl?: string;
	environment?: "Dev" | "Test" | "UAT" | "Production";
	/** Display name of the tool's connection */
	connectionName?: string;
	/**
	 * Identity of the tool's connection (connectionIdentity.connectionKey): undefined until the
	 * host has answered, "" with no connection. Changes only when the connection itself changes
	 * (#66), not on token refreshes or edits to other connections.
	 */
	connectionKey?: string;
}

// PPTB event payload type (compatible with @pptb/types ToolBoxEventPayload)
type ToolBoxEventPayload = {
	event: string;
	data: unknown;
	timestamp?: string;
};

// NOTE: window.toolboxAPI global types are provided by @pptb/types (see tsconfig.app.json)

/**
 * Get tool context from PPTB host
 * Returns theme and connection status
 * Subscribes to PPTB events for dynamic updates (theme, connection changes)
 */
/** The OS colour scheme as the webview sees it. */
const systemTheme = (): "light" | "dark" =>
	typeof window !== "undefined" && window.matchMedia?.("(prefers-color-scheme: dark)").matches ? "dark" : "light";

export function usePptbContext(): PptbContext {
	const [context, setContext] = useState<PptbContext>({
		theme: typeof window !== "undefined" && window.toolboxAPI ? "light" : systemTheme(),
		connected: false,
	});

	useEffect(() => {
		const colorScheme = window.matchMedia?.("(prefers-color-scheme: dark)");

		// Standalone (dev server): follow the OS colour scheme
		if (typeof window === "undefined" || !window.toolboxAPI) {
			const onSchemeChange = () => setContext((prev) => ({ ...prev, theme: systemTheme() }));
			colorScheme?.addEventListener("change", onSchemeChange);
			return () => colorScheme?.removeEventListener("change", onSchemeChange);
		}

		const toolboxAPI = window.toolboxAPI;

		// ============ THEME ============
		// Always ask the host: it resolves the "system" setting to light or dark (#67). Only the latest
		// request applies, so a slow earlier answer can't overwrite a newer one.
		let themeRequest = 0;
		const updateTheme = async () => {
			const request = ++themeRequest;
			try {
				const hostTheme = await toolboxAPI.utils?.getCurrentTheme?.();
				if (request !== themeRequest) return;
				const theme = hostTheme && String(hostTheme).toLowerCase() === "dark" ? "dark" : "light";
				setContext((prev) => (prev.theme === theme ? prev : { ...prev, theme }));
			} catch (err) {
				console.warn("Failed to get the theme from the host:", err);
			}
		};
		// The host sends no event when the OS theme changes while PPTB is set to System; the webview's
		// colour scheme follows the OS, so ask the host again when it changes.
		const onSchemeChange = () => void updateTheme();
		colorScheme?.addEventListener("change", onSchemeChange);

		// ============ CONNECTION ============
		// The tool's primary connection (getConnection(0) on PPTB ≥ 1.2.7). Only the latest request
		// applies; an unchanged answer (token refresh, another connection edited) keeps the same
		// state object, so nothing re-renders (#66).
		let connectionRequest = 0;
		const updateConnection = async () => {
			const request = ++connectionRequest;
			try {
				const connection = await getPrimaryConnection();
				if (request !== connectionRequest) return;
				const next = {
					connected: !!connection,
					environmentUrl: connection?.url,
					environment: connection?.environment,
					connectionName: connection?.name,
					connectionKey: connectionKey(connection),
				};
				setContext((prev) =>
					(Object.keys(next) as (keyof typeof next)[]).every((key) => prev[key] === next[key]) ? prev : { ...prev, ...next },
				);
			} catch (err) {
				console.warn("❌ Failed to get the tool's connection:", err);
			}
		};

		// ============ EVENT HANDLER ============
		// PPTB calls this with (event, payload) where payload has { event, data, timestamp }
		const handlePptbEvent = (_event: unknown, payload: ToolBoxEventPayload) => {
			// The event name only: connection payloads carry environment details (#71)
			debugLog("app", `🔔 PPTB event: ${payload.event}`);

			switch (payload.event) {
				case "settings:updated":
					// The payload holds only the changed settings, and theme can be "system": never read it
					// directly, ask the host for the resolved theme (#67)
					void updateTheme();
					break;

				case "connection:updated":
				case "connection:created":
				case "connection:deleted":
					// Broadcast for any connection, not only this tool's (#66): re-read the tool's own
					// connection; the state changes only if it actually changed
					void updateConnection();
					break;

				default:
					debugLog("app", `📢 Unhandled PPTB event: ${payload.event}`);
			}
		};

		// ============ INITIALIZATION ============
		// Get initial theme and connection
		updateTheme();
		updateConnection();

		// Subscribe to events
		if (toolboxAPI.events?.on) {
			debugLog("app", "✅ Subscribing to PPTB events");
			toolboxAPI.events.on(handlePptbEvent);
		} else {
			console.warn("⚠️ toolboxAPI.events.on not available");
		}

		// ============ CLEANUP ============
		return () => {
			colorScheme?.removeEventListener("change", onSchemeChange);
			if (toolboxAPI.events?.off) {
				debugLog("app", "🧹 Unsubscribing from PPTB events");
				toolboxAPI.events.off(handlePptbEvent);
			}
		};
	}, []);

	return context;
}
