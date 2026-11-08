import type { LayoutColumn } from "../../model/layoutxml";

const columnKey = (c: LayoutColumn) => `${c.linkEntityAlias ?? ""}|${c.name}`;

/**
 * Reconcile the Edit columns panel's unapplied local list with the query's
 * current columns: keep the local order for columns that still exist, drop
 * removed ones, and append new ones in query order.
 */
export function mergeColumns(local: LayoutColumn[], incoming: LayoutColumn[]): LayoutColumn[] {
	const incomingKeys = new Set(incoming.map(columnKey));
	const seen = new Set<string>();
	const merged: LayoutColumn[] = [];
	for (const c of local) {
		const key = columnKey(c);
		if (incomingKeys.has(key) && !seen.has(key)) {
			seen.add(key);
			merged.push(c);
		}
	}
	for (const c of incoming) {
		const key = columnKey(c);
		if (!seen.has(key)) {
			seen.add(key);
			merged.push(c);
		}
	}
	return merged;
}
