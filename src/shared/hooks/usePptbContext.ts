/**
 * Hook to access PPTB host context (connection, theme, etc.)
 * Subscribes to PPTB events for dynamic updates
 */

import { useState, useEffect } from "react";

export interface PptbContext {
	theme: "light" | "dark";
	connected: boolean;
	environmentUrl?: string;
	organizationId?: string;
}

// PPTB Event payload structure
interface ToolBoxEventPayload {
	event: string;
	data: unknown;
	timestamp?: string;
}

// Extend window interface for PPTB API
declare global {
	interface Window {
		toolboxAPI?: {
			utils?: {
				getCurrentTheme?: () => Promise<"light" | "dark">;
			};
			connections?: {
				getActiveConnection?: () => Promise<{
					url?: string;
					organizationId?: string;
					name?: string;
					environment?: string;
				} | null>;
			};
			events?: {
				// PPTB uses a single callback that receives all events
				on?: (callback: (event: unknown, payload: ToolBoxEventPayload) => void) => void;
				off?: (callback: (event: unknown, payload: ToolBoxEventPayload) => void) => void;
				getHistory?: (count: number) => Promise<ToolBoxEventPayload[]>;
			};
		};
	}
}

/**
 * Get tool context from PPTB host
 * Returns theme and connection status
 * Subscribes to PPTB events for dynamic updates (theme, connection changes)
 */
export function usePptbContext(): PptbContext {
	const [context, setContext] = useState<PptbContext>({
		theme: "light", // Default to light theme
		connected: false,
	});

	useEffect(() => {
		console.log("🔄 usePptbContext: Initializing...");

		// Check if running in PPTB host
		if (typeof window === "undefined" || !window.toolboxAPI) {
			console.log("💡 Running standalone - using default light theme");
			return;
		}

		const toolboxAPI = window.toolboxAPI;

		// ============ THEME ============
		const updateTheme = async () => {
			try {
				const hostTheme = await toolboxAPI.utils?.getCurrentTheme?.();
				console.log("🎨 getCurrentTheme() returned:", hostTheme);
				const normalizedTheme =
					hostTheme && String(hostTheme).toLowerCase() === "dark" ? "dark" : "light";
				setContext((prev) => ({
					...prev,
					theme: normalizedTheme,
				}));
			} catch (err) {
				console.warn("❌ Failed to get theme from host:", err);
			}
		};

		// ============ CONNECTION ============
		const updateConnection = async () => {
			try {
				const connection = await toolboxAPI.connections?.getActiveConnection?.();
				console.log("🔗 getActiveConnection() returned:", connection);
				setContext((prev) => ({
					...prev,
					connected: !!connection,
					environmentUrl: connection?.url,
					organizationId: connection?.organizationId,
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
					// Settings changed - check if theme was updated
					if (payload.data && typeof payload.data === "object" && "theme" in payload.data) {
						const themeValue = (payload.data as { theme: string }).theme;
						console.log("🎨 Theme changed via settings:updated:", themeValue);
						const normalizedTheme = String(themeValue).toLowerCase() === "dark" ? "dark" : "light";
						setContext((prev) => ({
							...prev,
							theme: normalizedTheme,
						}));
					} else {
						// Theme might have changed, refetch it
						updateTheme();
					}
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
						organizationId: undefined,
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
			if (toolboxAPI.events?.off) {
				console.log("🧹 Unsubscribing from PPTB events...");
				toolboxAPI.events.off(handlePptbEvent);
			}
		};
	}, []);

	return context;
}
