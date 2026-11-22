/**
 * The builder's undo history (#192): which builder actions are undo steps, their labels, and how a
 * stored state comes back. It wraps `builderReducer` with the generic core in history.ts.
 *
 * Owner decisions (2026-11-07): 100 steps; edits to one field merge for 1 s; New and Open view
 * clear the history instead of being undoable; column widths and Retrieve all records aren't
 * recorded (undo never reverts them); text fields and Monaco keep their own undo (the keys, PR 2).
 *
 * - **Steps** (labelled in sentence case, as "Undo: {label}" shows them): adding, deleting and
 *   editing tree nodes, Select attributes, Change entity, Parse to tree, Sort by column, column
 *   filters, and column changes other than widths. Edits to one field of one node merge.
 * - **Not steps:** selection, the layout sync each edit triggers (it keeps redo and doesn't end a
 *   merge), column widths, Retrieve all records, and Save updating the loaded view, which is
 *   patched into the stored states so an undo never pairs an older query with the saved view.
 * - **Clears:** New, Open view, and loads marked as a baseline (the query carried over a
 *   connection change, one sent from another tool, a view opened on launch), so the first Ctrl+Z
 *   can't blank the tree.
 * - **Selection after undo or redo:** the node the step was about if it exists (a deleted node comes
 *   back selected; an add selects its new node on redo and the parent on undo), else its nearest
 *   ancestor that does, else the stored selection.
 */
import { generateFetchXml } from "../model/fetchxml";
import type { LayoutXmlConfig } from "../model/layoutxml";
import type { FetchNode, NodeId } from "../model/nodes";
import {
	builderReducer,
	findNodeById,
	type BuilderAction,
	type BuilderState,
	type SelectedNode,
} from "./builderReducer";
import {
	beginTransaction,
	clear,
	commitTransaction,
	createHistory,
	lastEntryId,
	mapSnapshots,
	nearestExisting,
	record,
	redo,
	skip,
	undo,
	type History,
	type RestoreContext,
} from "./history";

export type BuilderHistoryState = History<BuilderState>;

/** What the provider sends: a builder action (stamped with the time), or a history command */
export type BuilderHistoryMessage =
	| { kind: "apply"; action: BuilderAction; at: number; baseline?: boolean }
	| { kind: "undo" }
	| { kind: "redo" }
	| { kind: "undoIfLast"; id: number }
	| { kind: "begin"; label: string; at: number; mergeKey?: string; target?: NodeId }
	| { kind: "commit" };

type Step =
	| { kind: "record"; label: string; target?: NodeId; mergeKey?: string }
	| { kind: "skip"; keepRedo?: boolean; endsMerge?: boolean }
	| { kind: "clear" }
	| { kind: "saved" };

export function createBuilderHistory(initial: BuilderState): BuilderHistoryState {
	return createHistory(initial, { restore: restoreBuilderState });
}

export function builderHistoryReducer(
	h: BuilderHistoryState,
	message: BuilderHistoryMessage,
): BuilderHistoryState {
	switch (message.kind) {
		case "undo":
			return undo(h);
		case "redo":
			return redo(h);
		case "undoIfLast":
			return lastEntryId(h) === message.id ? undo(h) : h;
		case "begin":
			return beginTransaction(h, {
				label: message.label,
				mergeKey: message.mergeKey,
				target: message.target,
				now: message.at,
			});
		case "commit":
			return h.transaction ? commitTransaction(h) : h;
		case "apply": {
			const before = h.present;
			const after = builderReducer(before, message.action);
			if (message.baseline) return clear(h, after);
			const step = stepFor(message.action, before, after);
			switch (step.kind) {
				case "clear":
					return clear(h, after);
				case "skip":
					return skip(h, after, { keepRedo: step.keepRedo, endsMerge: step.endsMerge });
				case "saved":
					return mapSnapshots(skip(h, after, { keepRedo: true }), (snapshot) =>
						patchSavedView(snapshot, before, after),
					);
				case "record":
					// Only a change to the query or its columns is a step (a no-op edit isn't)
					if (
						after.fetchQuery === before.fetchQuery &&
						after.columnConfig === before.columnConfig
					) {
						return skip(h, after, { keepRedo: true, endsMerge: false });
					}
					return record(h, after, {
						label: step.label,
						target: step.target,
						mergeKey: step.mergeKey,
						now: message.at,
					});
			}
		}
	}
}

/** What a builder action is to the history */
export function stepFor(action: BuilderAction, before: BuilderState, after: BuilderState): Step {
	const target = after.selectedNodeId ?? undefined;
	const step = (label: string, t: NodeId | undefined = target): Step => ({
		kind: "record",
		label,
		target: t,
	});
	switch (action.type) {
		case "SELECT_NODE":
		case "REVEAL_NODE":
		case "UPDATE_COLUMN_WIDTH":
		case "SET_RETRIEVE_ALL_RECORDS":
			return { kind: "skip", keepRedo: true };
		case "SYNC_LAYOUT_WITH_FETCHXML":
			// Follows every edit (an effect after the render): neither a step nor the end of a merge
			return { kind: "skip", keepRedo: true, endsMerge: false };
		case "CLEAR_LOADED_VIEW":
			return { kind: "skip" };
		case "SET_LOADED_VIEW":
			return { kind: "saved" };
		case "NEW_QUERY":
		case "LOAD_VIEW":
			return { kind: "clear" };
		case "SET_ENTITY":
			return step("Change entity");
		case "ADD_ATTRIBUTE":
		case "ADD_ATTRIBUTE_BY_NAME":
			return step("Add attribute");
		case "UPDATE_ATTRIBUTES":
			return step("Select attributes");
		case "ADD_ALL_ATTRIBUTES":
			return step("Add all attributes");
		case "ADD_ORDER":
			return step("Add order");
		case "ADD_FILTER":
			return step("Add filter");
		case "ADD_SUBFILTER":
			return step("Add sub-filter");
		case "ADD_CONDITION":
			return step("Add condition");
		case "ADD_LINK_ENTITY":
		case "ADD_LINK_ENTITY_WITH_CONFIG":
			return step("Add link-entity");
		case "REMOVE_NODE":
			return step(
				`Delete ${nodeNoun(findNodeById(before.fetchQuery, action.nodeId))}`,
				action.nodeId,
			);
		case "UPDATE_NODE": {
			const node = findNodeById(before.fetchQuery, action.nodeId);
			if (node && isNoOpUpdate(node, action.updates)) {
				return { kind: "skip", keepRedo: true, endsMerge: false };
			}
			return {
				kind: "record",
				label: `Edit ${nodeNoun(node)}`,
				target: action.nodeId,
				mergeKey: `${action.nodeId}:${Object.keys(action.updates).sort().join(",")}`,
			};
		}
		case "LOAD_FETCHXML":
			return step("Parse to tree");
		case "SET_SORT":
			return step("Sort by column");
		case "ADD_COLUMN_FILTER":
			return step("Filter column");
		case "REMOVE_COLUMN_FILTER":
			return step("Remove column filter");
		case "CLEAR_COLUMN_FILTERS":
			return step("Clear column filters");
		case "SET_COLUMN_CONFIG":
			return step("Edit columns");
		case "REORDER_COLUMNS":
			return step("Reorder columns");
	}
}

/** The tree's word for a node, as its menus and delete confirmations use it */
function nodeNoun(node: SelectedNode): string {
	switch (node?.type) {
		case "fetch":
			return "query options";
		case "entity":
			return "entity";
		case "attribute":
			return "attribute";
		case "all-attributes":
			return "all attributes";
		case "order":
			return "order";
		case "filter":
			return "filter";
		case "condition":
			return "condition";
		case "link-entity":
			return "link-entity";
		default:
			return "node";
	}
}

/** Whether every update already holds: the fetch node's values are in its options (the reducer merges there) */
function isNoOpUpdate(node: NonNullable<SelectedNode>, updates: Record<string, unknown>): boolean {
	const current = (node.type === "fetch" ? node.options : node) as unknown as Record<
		string,
		unknown
	>;
	return Object.keys(updates).every((key) => Object.is(current[key], updates[key]));
}

/**
 * The state shown after an undo or redo: the stored one, with what the history doesn't own carried
 * from the present (column widths, Retrieve all records), the selection picked from the step's
 * target, and the tree asked to reveal it.
 */
export function restoreBuilderState(
	snapshot: BuilderState,
	present: BuilderState,
	{ target }: RestoreContext,
): BuilderState {
	const query = snapshot.fetchQuery;
	const exists = (id: NodeId) => !!findNodeById(query, id);
	const path = target ? (pathTo(query, target) ?? pathTo(present.fetchQuery, target) ?? []) : [];
	const selectedNodeId =
		nearestExisting(path, exists) ??
		(snapshot.selectedNodeId && exists(snapshot.selectedNodeId)
			? snapshot.selectedNodeId
			: (query?.entity.id ?? null));
	return {
		...snapshot,
		selectedNodeId,
		selectedNode: selectedNodeId ? findNodeById(query, selectedNodeId) : null,
		revealSeq: present.revealSeq + 1,
		retrieveAllRecords: present.retrieveAllRecords,
		columnConfig: withPresentWidths(snapshot.columnConfig, present.columnConfig),
	};
}

/** The stored columns, with the present width of every column both have */
function withPresentWidths(
	stored: LayoutXmlConfig | null,
	present: LayoutXmlConfig | null,
): LayoutXmlConfig | null {
	if (!stored || !present || stored === present) return stored;
	const widths = new Map(present.columns.map((c) => [c.name, c.width]));
	let changed = false;
	const columns = stored.columns.map((c) => {
		const width = widths.get(c.name);
		if (width === undefined || width === c.width) return c;
		changed = true;
		return { ...c, width };
	});
	return changed ? { ...stored, columns } : stored;
}

/** The node and its ancestors, nearest first; null when the tree hasn't got it */
export function pathTo(query: FetchNode | null, id: NodeId): NodeId[] | null {
	const walk = (node: unknown, trail: NodeId[]): NodeId[] | null => {
		if (!node || typeof node !== "object") return null;
		const n = node as Record<string, unknown>;
		const here = typeof n.id === "string" ? [n.id, ...trail] : trail;
		if (n.id === id) return here;
		for (const key of [
			"attributes",
			"orders",
			"filters",
			"conditions",
			"subfilters",
			"links",
		]) {
			const items = n[key];
			if (!Array.isArray(items)) continue;
			for (const item of items) {
				const found = walk(item, here);
				if (found) return found;
			}
		}
		for (const key of ["entity", "allAttributes"]) {
			const found = walk(n[key], here);
			if (found) return found;
		}
		return null;
	};
	return walk(query, []);
}

/**
 * After Save, a stored state of the same query (the view saved over, or the query saved as a new
 * view, for the same entity) is that view now: its identity is the saved one, and it runs by the
 * view's ID only if its query is what was saved. Other stored states are left as they are.
 */
export function patchSavedView(
	snapshot: BuilderState,
	before: BuilderState,
	after: BuilderState,
): BuilderState {
	if (!after.viewIdentity) return snapshot;
	const sameView =
		(snapshot.viewIdentity?.view.id ?? null) === (before.viewIdentity?.view.id ?? null);
	const sameEntity = snapshot.fetchQuery?.entity.name === after.fetchQuery?.entity.name;
	if (!sameView || !sameEntity || !snapshot.fetchQuery) return snapshot;
	const saved = after.loadedView;
	const unchanged = !!saved && generateFetchXml(snapshot.fetchQuery) === saved.originalFetchXml;
	return {
		...snapshot,
		viewIdentity: after.viewIdentity,
		loadedView: unchanged ? saved : null,
	};
}
