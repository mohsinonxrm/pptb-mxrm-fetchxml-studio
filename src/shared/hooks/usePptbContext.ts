/**
 * Hook to access PPTB host context (connection, theme, etc.)
 * Subscribes to PPTB events for dynamic updates
 */

import { useState, useEffect } from "react";

export interface PptbContext {
	theme: "light" | "dark";
	connected: boolean;
	environmentUrl?: string;
	environment?: "Dev" | "Test" | "UAT" | "Production";
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
		const updateConnection = async () => {
			try {
				const connection = await toolboxAPI.connections?.getActiveConnection?.();
				console.log("🔗 getActiveConnection() returned:", connection);
				setContext((prev) => ({
					...prev,
					connected: !!connection,
					environmentUrl: connection?.url,
					environment: connection?.environment,
				}));
			} catch (err) {
				console.warn("❌ Failed to get active connection:", err);
			}
		};

		// ============ EVENT HANDLER ============
		// PPTB calls this with (event, payload) where payload has { event, data, timestamp }
		const handlePptbEvent = (_event: unknown, payload: ToolBoxEventPayload) => {
			console.log("🔔 PPTB Event received:", payload.event, payload.data);

			switch (payload.event) {
				case "settings:updated":
					// The payload holds only the changed settings, and theme can be "system": never read it
					// directly, ask the host for the resolved theme (#67)
					void updateTheme();
					break;

				case "connection:updated":
				case "connection:created":
					// Connection changed - refetch connection info
					console.log("🔗 Connection changed, refreshing...");
					updateConnection();
					break;

				case "connection:deleted":
					// Connection was deleted
					console.log("🔗 Connection deleted");
					setContext((prev) => ({
						...prev,
						connected: false,
						environmentUrl: undefined,
						environment: undefined,
					}));
					break;

				default:
					// Log other events for debugging
					console.log("📢 Unhandled PPTB event:", payload.event);
			}
		};

		// ============ INITIALIZATION ============
		// Get initial theme and connection
		updateTheme();
		updateConnection();

		// Subscribe to events
		if (toolboxAPI.events?.on) {
			console.log("✅ Subscribing to PPTB events...");
			toolboxAPI.events.on(handlePptbEvent);
		} else {
			console.warn("⚠️ toolboxAPI.events.on not available");
		}

		// ============ CLEANUP ============
		return () => {
			colorScheme?.removeEventListener("change", onSchemeChange);
			if (toolboxAPI.events?.off) {
				console.log("🧹 Unsubscribing from PPTB events...");
				toolboxAPI.events.off(handlePptbEvent);
			}
		};
	}, []);

	return context;
}
