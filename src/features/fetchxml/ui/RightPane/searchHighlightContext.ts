/**
 * Quick search over the loaded rows (#23): the search text, for highlighting matches in cells.
 * A context rather than a column prop, so typing doesn't rebuild the grid's columns.
 */

import { createContext } from "react";

export const SearchHighlightContext = createContext("");

/** The ranges of text that match the search, case-insensitively */
export function searchMatches(text: string, search: string): [number, number][] {
	const q = search.trim().toLowerCase();
	if (!q) return [];
	const lower = text.toLowerCase();
	const ranges: [number, number][] = [];
	for (let i = lower.indexOf(q); i >= 0; i = lower.indexOf(q, i + q.length)) ranges.push([i, i + q.length]);
	return ranges;
}
