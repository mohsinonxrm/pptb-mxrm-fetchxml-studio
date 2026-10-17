/**
 * Row identity for the results grid.
 *
 * Grid row IDs must be unique and stable for the lifetime of a result. The root table's primary
 * key is the natural ID, but it is not always usable (#59):
 *  - aggregate and distinct results contain no primary key;
 *  - 1:N / N:N joins repeat the same primary key on several rows.
 *
 * Rows get the primary key as their ID when present and not already taken by an earlier row;
 * otherwise a position-based ID. Selected row IDs are translated back to real record IDs
 * (primary key values, de-duplicated) so record actions never receive a synthetic ID.
 */

type Row = Record<string, unknown>;

export interface RowIdentity {
	/** Unique, stable grid row ID for a row of this result. */
	getRowId: (row: Row) => string;
	/** Record IDs (primary key values) for the given grid row IDs, de-duplicated, in order. */
	toRecordIds: (rowIds: Iterable<string | number>) => string[];
}

export function buildRowIdentity(rows: readonly Row[], primaryIdAttribute?: string): RowIdentity {
	const rowIds = new WeakMap<Row, string>();
	const recordIdByRowId = new Map<string, string>();
	const used = new Set<string>();

	rows.forEach((row, index) => {
		const pk = primaryIdAttribute ? row[primaryIdAttribute] : undefined;
		const hasPk = typeof pk === "string" && pk.length > 0;
		const id = hasPk && !used.has(pk) ? pk : `__row_${index}`;
		used.add(id);
		rowIds.set(row, id);
		if (hasPk) recordIdByRowId.set(id, pk);
	});

	return {
		getRowId: (row) => rowIds.get(row) ?? `__row_unknown_${rows.indexOf(row)}`,
		toRecordIds: (selected) => {
			const ids = new Set<string>();
			for (const rowId of selected) {
				const recordId = recordIdByRowId.get(String(rowId));
				if (recordId) ids.add(recordId);
			}
			return Array.from(ids);
		},
	};
}
