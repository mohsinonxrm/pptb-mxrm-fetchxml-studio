/**
 * The app bar's wording (#278), apart from the component so it stays fast-refreshable.
 */

/** What New says when it discards the query's changes */
export const discardedToast = (viewName: string | null) =>
	viewName ? `Discarded changes to "${viewName}"` : "Discarded the unsaved query";
