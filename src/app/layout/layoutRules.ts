/**
 * Pure layout rules for the builder pane (#234). Not used yet: the hook and
 * components that consume these come later.
 *
 * Owner decisions (2026-11-07, #234):
 * - The builder pane always starts docked. Docked/collapsed/peek is NOT
 *   persisted; only the sizes are.
 * - Properties go into the inspector drawer when the window is under 1000 px
 *   OR the docked pane is under 320 px. "Under" means strictly `<`.
 * - Saved sizes are stored raw and clamped on use: a saved width is clamped to
 *   the current window without being written back.
 */

export const BUILDER_WIDTH_DEFAULT = 380;
export const BUILDER_WIDTH_NARROW_DEFAULT = 300;
export const BUILDER_WIDTH_MIN = 240;
const BUILDER_WIDTH_MAX_FLOOR = 320;
const BUILDER_WIDTH_WINDOW_MARGIN = 420;
/** Sanity cap for a persisted width (no real window is wider). */
const BUILDER_WIDTH_PERSIST_CAP = 8192;

export const TREE_HEIGHT_DEFAULT = 240;
export const TREE_HEIGHT_MIN = 120;
export const TREE_HEIGHT_MAX = 460;

export const DRAWER_WINDOW_BREAKPOINT = 1000;
export const DRAWER_PANE_BREAKPOINT = 320;

export type BuilderPaneState = "docked" | "collapsed" | "peek";
export type PropertiesHost = "docked" | "drawer" | "peek";

export interface SizeLimits {
	min: number;
	max: number;
	defaultValue: number;
}

export interface PersistedLayout {
	builderWidth?: number;
	treeHeight?: number;
}

/** Limits for the builder width at a given window width. */
export function builderWidthLimits(windowWidth: number): SizeLimits {
	return {
		min: BUILDER_WIDTH_MIN,
		max: Math.max(BUILDER_WIDTH_MAX_FLOOR, windowWidth - BUILDER_WIDTH_WINDOW_MARGIN),
		defaultValue:
			windowWidth < DRAWER_WINDOW_BREAKPOINT
				? BUILDER_WIDTH_NARROW_DEFAULT
				: BUILDER_WIDTH_DEFAULT,
	};
}

function clamp(value: number, min: number, max: number): number {
	return Math.min(max, Math.max(min, value));
}

export function clampBuilderWidth(width: number, windowWidth: number): number {
	const { min, max } = builderWidthLimits(windowWidth);
	return clamp(width, min, max);
}

export function clampTreeHeight(height: number): number {
	return clamp(height, TREE_HEIGHT_MIN, TREE_HEIGHT_MAX);
}

/** Where the Properties editor lives for the current pane state and sizes. */
export function propertiesHost(input: {
	pane: BuilderPaneState;
	windowWidth: number;
	builderWidth: number;
}): PropertiesHost {
	if (input.pane !== "docked") return "peek";
	if (
		input.windowWidth < DRAWER_WINDOW_BREAKPOINT ||
		input.builderWidth < DRAWER_PANE_BREAKPOINT
	) {
		return "drawer";
	}
	return "docked";
}

function validNumber(value: unknown, min: number, max: number): number | undefined {
	return typeof value === "number" && Number.isFinite(value) && value >= min && value <= max
		? value
		: undefined;
}

/**
 * Reads a saved layout of unknown shape. Bad values are dropped (the default
 * applies); good ones are returned raw, to be clamped on use. Any `pane`
 * field (legacy or otherwise) is ignored: the pane always starts docked.
 */
export function normalizePersistedLayout(raw: unknown): PersistedLayout {
	if (!raw || typeof raw !== "object") return {};
	const source = raw as Record<string, unknown>;
	const result: PersistedLayout = {};
	const builderWidth = validNumber(
		source.builderWidth,
		BUILDER_WIDTH_MIN,
		BUILDER_WIDTH_PERSIST_CAP,
	);
	if (builderWidth !== undefined) result.builderWidth = builderWidth;
	const treeHeight = validNumber(source.treeHeight, TREE_HEIGHT_MIN, TREE_HEIGHT_MAX);
	if (treeHeight !== undefined) result.treeHeight = treeHeight;
	return result;
}

/** What gets saved: sizes only, never the pane state. */
export function layoutToPersist(state: {
	builderWidth: number;
	treeHeight: number;
}): Required<PersistedLayout> {
	return { builderWidth: state.builderWidth, treeHeight: state.treeHeight };
}
