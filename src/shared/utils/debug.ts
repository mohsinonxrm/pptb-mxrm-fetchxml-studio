/**
 * Debug utility for conditional logging
 * Set flags to true to enable logging for specific features
 */

// Feature flags for debugging
const DEBUG_FLAGS = {
	treeExpansion: false, // Tree auto-expansion when nodes are selected
	relationshipPicker: false, // Relationship picker loading and selection
	linkEntityEditor: false, // Link entity editor behavior
	propertiesPanel: false, // Properties panel parent entity resolution
	metadataAPI: true, // Metadata API calls and responses
	fetchXmlAPI: true, // FetchXML query execution and results
	publisherAPI: true, // Publisher API calls
	solutionAPI: true, // Solution API calls
	solutionComponentAPI: true, // Solution component API calls
	viewAPI: true, // View (savedquery/userquery) API calls
} as const;

type DebugCategory = keyof typeof DEBUG_FLAGS;

/**
 * Conditional logger that only logs when the feature flag is enabled
 */
export function debugLog(category: DebugCategory, message: string, ...args: unknown[]): void {
	if (DEBUG_FLAGS[category]) {
		console.log(`[DEBUG:${category}] ${message}`, ...args);
	}
}

/**
 * Conditional warning logger
 */
export function debugWarn(category: DebugCategory, message: string, ...args: unknown[]): void {
	if (DEBUG_FLAGS[category]) {
		console.warn(`[DEBUG:${category}] ${message}`, ...args);
	}
}

/**
 * Enable debugging for a specific category
 * Usage: enableDebug('treeExpansion')
 */
export function enableDebug(category: DebugCategory): void {
	(DEBUG_FLAGS as Record<string, boolean>)[category] = true;
	console.log(`✅ Debug enabled for: ${category}`);
}

/**
 * Disable debugging for a specific category
 */
export function disableDebug(category: DebugCategory): void {
	(DEBUG_FLAGS as Record<string, boolean>)[category] = false;
	console.log(`❌ Debug disabled for: ${category}`);
}

/**
 * Enable all debugging
 * Usage (in browser console): window.enableAllDebug()
 */
export function enableAllDebug(): void {
	Object.keys(DEBUG_FLAGS).forEach((key) => {
		(DEBUG_FLAGS as Record<string, boolean>)[key] = true;
	});
	console.log("✅ All debugging enabled");
}

/**
 * Disable all debugging
 */
export function disableAllDebug(): void {
	Object.keys(DEBUG_FLAGS).forEach((key) => {
		(DEBUG_FLAGS as Record<string, boolean>)[key] = false;
	});
	console.log("❌ All debugging disabled");
}

// Expose to window for easy access in browser console
if (typeof window !== "undefined") {
	(window as unknown as Record<string, unknown>).enableDebug = enableDebug;
	(window as unknown as Record<string, unknown>).disableDebug = disableDebug;
	(window as unknown as Record<string, unknown>).enableAllDebug = enableAllDebug;
	(window as unknown as Record<string, unknown>).disableAllDebug = disableAllDebug;
}
