/**
 * The app bar's wording and rules (#278), apart from the components so they stay fast-refreshable.
 */
import type { SavedView } from "../../api/pptbClient";

/** What New and Open view say when they discard the query's changes */
export const discardedToast = (viewName: string | null) =>
	viewName ? `Discarded changes to "${viewName}"` : "Discarded the unsaved query";

/** Why Open view can't be used yet: it needs the query's entity and its metadata (to run a view by ID) */
export function openViewDisabledReason(
	entityLogicalName: string | undefined,
	entitySetName: string | undefined,
): string | undefined {
	if (!entityLogicalName) return "Choose an entity first";
	if (!entitySetName) return "Loading the entity…";
	return undefined;
}

/**
 * Open view's groups (owner decision Q2): "My views" (the user's personal views, and ones shared with
 * them) above "System views", each filtered by name, ignoring case.
 */
export function openViewGroups(
	systemViews: SavedView[],
	personalViews: SavedView[],
	search: string,
): { mine: SavedView[]; system: SavedView[] } {
	const query = search.trim().toLowerCase();
	const matches = (view: SavedView) => !query || view.name.toLowerCase().includes(query);
	return { mine: personalViews.filter(matches), system: systemViews.filter(matches) };
}
