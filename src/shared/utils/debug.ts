/**
 * Debug utility for conditional logging
 * Set flags to true to enable logging for specific features
 *
 * All categories are off by default (#71): some log query payloads or record data, which shouldn't
 * reach DevTools unless someone turns them on for support, e.g. window.enableDebug("fetchXmlAPI")
 * or window.enableAllDebug() in the console.
 */

// Feature flags for debugging
const DEBUG_FLAGS = {
	treeExpansion: false, // Tree auto-expansion when nodes are selected
	relationshipPicker: false, // Relationship picker loading and selection
	linkEntityEditor: false, // Link entity editor behavior
	propertiesPanel: false, // Properties panel parent entity resolution
	metadataAPI: false, // Metadata API calls and responses
	fetchXmlAPI: false, // FetchXML query execution and results
	publisherAPI: false, // Publisher API calls
	solutionAPI: false, // Solution API calls
	solutionComponentAPI: false, // Solution component API calls
	viewAPI: false, // View (savedquery/userquery) API calls
	exportAPI: false, // Export to Excel API calls
	recordAPI: false, // Record CRUD operations
	workflowAPI: false, // Workflow execution
	accessAPI: false, // WhoAmI and privilege checks (#71)
	app: false, // Startup checks, executions, paging, record actions (#71)
	settings: false, // Settings load/save through toolboxAPI (#71)
	resultsGrid: false, // Results grid sizing and infinite scroll (#71)
	entityFilters: false, // Entity picker publisher/solution filtering (#71)
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
