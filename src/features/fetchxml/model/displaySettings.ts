/**
 * Display settings types and defaults
 */

export type ValueDisplayMode = "formatted" | "raw" | "both";

export interface DisplaySettings {
	/** Show logical names in column headers instead of display names */
	useLogicalNames: boolean;
	/** How to display cell values: formatted, raw, or both */
	valueDisplayMode: ValueDisplayMode;
}

/** Default display settings */
export const defaultDisplaySettings: DisplaySettings = {
	useLogicalNames: false,
	valueDisplayMode: "formatted",
};
