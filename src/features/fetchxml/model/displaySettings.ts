/**
 * Display settings types and defaults
 */

import type { TreeIconStyle } from "./treeIcons";

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

/**
 * Density of the results grid and the query tree (#228): Fluent's `DataGrid size` values. One
 * setting drives both; Fluent's Tree has only medium and small, so Small and Extra small both use
 * the small tree.
 */
export type GridDensity = "medium" | "small" | "extra-small";

/** The densities in menu order, with their labels (the same as Dataverse Request Studio's) */
export const GRID_DENSITY_LABELS: Record<GridDensity, string> = {
	medium: "Medium",
	small: "Small",
	"extra-small": "Extra small",
};

/** A saved value that isn't a known density (an older or edited settings object) falls back to medium */
export function normalizeGridDensity(value: unknown): GridDensity {
	return value === "small" || value === "extra-small" ? value : "medium";
}

/** Fluent `Tree size` for a density: the tree has no extra-small size */
export function treeSizeFor(density: GridDensity): "medium" | "small" {
	return density === "medium" ? "medium" : "small";
}

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
	/** Density of the results grid and the query tree (#228) */
	gridDensity: GridDensity;
	/** Query tree icons: one-colour tinted icons, or Fluent's multicolour Color icons (#230) */
	treeIconStyle: TreeIconStyle;
}

/** Default display settings */
export const defaultDisplaySettings: DisplaySettings = {
	useLogicalNames: false,
	valueDisplayMode: "formatted",
	entityScopeMode: "publisher-solution",
	advancedFindOnly: true,
	gridDensity: "medium",
	treeIconStyle: "tinted",
};
