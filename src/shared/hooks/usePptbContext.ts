/**
 * Hook to access PPTB host context (connection, theme, etc.)
 */

import { useState, useEffect } from "react";

export interface PptbContext {
	theme: "light" | "dark";
	connected: boolean;
	environmentUrl?: string;
	organizationId?: string;
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
				} | null>;
			};
			events?: {
				on?: (event: string, callback: (data: unknown) => void) => void;
				off?: (event: string, callback: (data: unknown) => void) => void;
			};
		};
	}
}

/**
 * Get tool context from PPTB host
 * Returns theme and connection status
 */
export function usePptbContext(): PptbContext {
	const [context, setContext] = useState<PptbContext>({
		theme: "light", // Default to light theme
		connected: false,
	});

	useEffect(() => {
		console.log("🔄 usePptbContext useEffect running - component mounted or remounted");
		
		const updateTheme = async () => {
			if (typeof window !== "undefined" && window.toolboxAPI) {
				try {
					const hostTheme = await window.toolboxAPI.utils?.getCurrentTheme?.();
					console.log("✅ getCurrentTheme() returned (raw):", hostTheme, typeof hostTheme);
					// Normalize theme value to handle case sensitivity and ensure it's valid
					const normalizedTheme = 
						hostTheme && String(hostTheme).toLowerCase() === "dark" ? "dark" : "light";
					console.log("🎨 Normalized theme:", normalizedTheme);
					setContext((prev) => ({
						...prev,
						theme: normalizedTheme,
					}));
				} catch (err) {
					console.warn("❌ Failed to get theme from host:", err);
					// Default to light on error
					setContext((prev) => ({ ...prev, theme: "light" }));
				}
			} else {
				// Standalone mode - default to light theme
				console.log("💡 Running standalone - using light theme");
				setContext((prev) => ({ ...prev, theme: "light" }));
			}
		};

		// Get theme from PPTB host
		if (typeof window !== "undefined" && window.toolboxAPI) {
			const toolboxAPI = window.toolboxAPI;

			// Initial theme load
			updateTheme();

			// Listen for settings updates (including theme changes)
			const handleSettingsUpdate = (eventData: unknown) => {
				console.log("🔄 Settings updated event received (raw):", eventData);
				console.log("🔍 Event type:", typeof eventData);
				console.log("🔍 Event keys:", eventData && typeof eventData === "object" ? Object.keys(eventData) : "not an object");
				
				if (
					eventData &&
					typeof eventData === "object" &&
					"data" in eventData &&
					eventData.data &&
					typeof eventData.data === "object" &&
					"theme" in eventData.data
				) {
					const themeValue = (eventData.data as { theme: string }).theme;
					console.log("🎨 Theme changed to:", themeValue);
					// Update theme directly from event data
					const normalizedTheme = String(themeValue).toLowerCase() === "dark" ? "dark" : "light";
					console.log("✅ Setting normalized theme:", normalizedTheme);
					setContext((prev) => ({
						...prev,
						theme: normalizedTheme,
					}));
				} else {
					console.warn("⚠️ Event data structure unexpected:", eventData);
				}
			};

			// Register event listener using the CORRECT callback signature
			// PPTB might be calling with just the data object, not wrapped
			const handleSettingsUpdateWrapper = (data: unknown) => {
				console.log("🔔 Event listener called with data:", data);
				
				// Check if data is already the event object with event/data/timestamp
				if (data && typeof data === "object" && "event" in data && "data" in data) {
					handleSettingsUpdate(data);
				} else if (data && typeof data === "object" && "theme" in data) {
					// Data is just the settings object directly
					const themeValue = (data as { theme: string }).theme;
					console.log("🎨 Theme changed to (direct):", themeValue);
					const normalizedTheme = String(themeValue).toLowerCase() === "dark" ? "dark" : "light";
					console.log("✅ Setting normalized theme:", normalizedTheme);
					setContext((prev) => ({
						...prev,
						theme: normalizedTheme,
					}));
				} else {
					console.warn("⚠️ Unknown event data structure:", data);
				}
			};

			// Register event listener
			if (toolboxAPI.events?.on) {
				console.log("🔧 Attempting to register listener...");
				console.log("🔧 toolboxAPI.events:", toolboxAPI.events);
				console.log("🔧 toolboxAPI.events.on type:", typeof toolboxAPI.events.on);
				
				// Try registering with the exact event name
				try {
					toolboxAPI.events.on("settings:updated", handleSettingsUpdateWrapper);
					console.log("✅ Registered listener for 'settings:updated' event");
				} catch (err) {
					console.error("❌ Failed to register event listener:", err);
				}

				return () => {
					// Cleanup event listener
					if (toolboxAPI.events?.off) {
						toolboxAPI.events.off("settings:updated", handleSettingsUpdateWrapper);
						console.log("🧹 Unregistered listener for settings:updated event");
					}
				};
			} else {
				console.warn("⚠️ toolboxAPI.events or toolboxAPI.events.on not available");
			}

			// Get connection
			toolboxAPI.connections
				?.getActiveConnection?.()
				.then((connection) => {
					console.log("✅ getActiveConnection() returned:", connection);
					setContext((prev) => ({
						...prev,
						connected: !!connection,
						environmentUrl: connection?.url,
						organizationId: connection?.organizationId,
					}));
				})
				.catch((err: unknown) => {
					console.warn("❌ Failed to get active connection:", err);
				});

			return () => {
				// Cleanup event listener
				if (toolboxAPI.events?.off) {
					toolboxAPI.events.off("settings:updated", handleSettingsUpdate);
					console.log("🧹 Unregistered listener for settings:updated event");
				}
			};
		} else {
			// Standalone mode - default to light theme
			console.log("💡 Running standalone - using light theme");
			setContext((prev) => ({ ...prev, theme: "light" }));
		}
	}, []);

	return context;
}
