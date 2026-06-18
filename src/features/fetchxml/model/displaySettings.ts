/**
 * Display settings types and defaults
 */

export type ValueDisplayMode = "formatted" | "raw" | "both";

/**
 * Entity scope mode — controls how the entity list is sourced.
 *
 * "publisher-solution" — Publisher → Solution cascade (requires prvReadPublisher + prvReadSolution).
 *                        Silently demotes to "solution-only" when publisher privilege is absent.
 * "solution-only"      — Solution picker only (requires prvReadSolution).
 *                        Silently demotes to "all" when solution privilege is absent.
 * "all"                — All entities loaded directly, no publisher/solution filter.
 */
export type EntityScopeMode = "publisher-solution" | "solution-only" | "all";

export interface DisplaySettings {
	/** Show logical names in column headers instead of display names */
	useLogicalNames: boolean;
	/** How to display cell values: formatted, raw, or both */
	valueDisplayMode: ValueDisplayMode;
	/**
	 * How to scope the entity list in the toolbar.
	 * May be silently demoted by the privilege ceiling (see EntityScopeMode docs).
	 */
	entityScopeMode: EntityScopeMode;
	/**
	 * When true, entities and attributes are limited to those marked IsValidForAdvancedFind.
	 * When false, all entities/attributes are shown (power-user / developer mode).
	 */
	advancedFindOnly: boolean;
}

/** Default display settings */
export const defaultDisplaySettings: DisplaySettings = {
	useLogicalNames: false,
	valueDisplayMode: "formatted",
	entityScopeMode: "publisher-solution",
	advancedFindOnly: true,
};
