/**
 * Hook for persisting DisplaySettings via window.toolboxAPI.settings.
 *
 * When running inside the PPTB host, settings are loaded from and saved to
 * toolboxAPI.settings (key: "displaySettings") so they survive across sessions.
 *
 * When running standalone/dev (no toolboxAPI), settings live only in memory.
 */

import { useState, useEffect, useCallback, useRef } from "react";
import {
	defaultDisplaySettings,
	type DisplaySettings,
} from "../../features/fetchxml/model/displaySettings";

const SETTINGS_KEY = "displaySettings";

/**
 * Merge persisted value with defaults so any newly added fields always get
 * their default value for users who have an older saved settings object.
 */
function mergeWithDefaults(persisted: unknown): DisplaySettings {
	if (!persisted || typeof persisted !== "object") {
		return { ...defaultDisplaySettings };
	}
	return {
		...defaultDisplaySettings,
		...(persisted as Partial<DisplaySettings>),
	};
}

export function usePersistedSettings(): [DisplaySettings, (settings: DisplaySettings) => void] {
	const [settings, setSettingsState] = useState<DisplaySettings>(defaultDisplaySettings);

	// Track whether the initial load from toolboxAPI has completed.
	// Prevents the save effect from writing back before we've read the value.
	const loadedRef = useRef(false);

	// Load settings from toolboxAPI on mount
	useEffect(() => {
		let mounted = true;

		async function loadSettings() {
			try {
				if (typeof window !== "undefined" && window.toolboxAPI?.settings?.get) {
					const persisted = await window.toolboxAPI.settings.get(SETTINGS_KEY);
					if (mounted) {
						const merged = mergeWithDefaults(persisted);
						setSettingsState(merged);
						console.log("⚙️ Settings loaded from toolboxAPI:", merged);
					}
				} else {
					console.log("⚙️ toolboxAPI.settings not available — using in-memory defaults");
				}
			} catch (err) {
				console.warn("⚙️ Failed to load settings from toolboxAPI:", err);
			} finally {
				if (mounted) {
					loadedRef.current = true;
				}
			}
		}

		loadSettings();

		return () => {
			mounted = false;
		};
	}, []);

	// Persist settings to toolboxAPI whenever they change (after initial load)
	const updateSettings = useCallback((newSettings: DisplaySettings) => {
		setSettingsState(newSettings);

		// Fire-and-forget save; only after the initial load has completed
		if (loadedRef.current) {
			if (typeof window !== "undefined" && window.toolboxAPI?.settings?.set) {
				window.toolboxAPI.settings
					.set(SETTINGS_KEY, newSettings)
					.then(() => {
						console.log("⚙️ Settings saved to toolboxAPI:", newSettings);
					})
					.catch((err: unknown) => {
						console.warn("⚙️ Failed to save settings to toolboxAPI:", err);
					});
			}
		}
	}, []);

	return [settings, updateSettings];
}
