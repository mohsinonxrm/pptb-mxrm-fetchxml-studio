/**
 * The builder's undo history (#192, builderHistory.ts): which actions are steps, merging, the
 * clears, restoring (selection, widths, Retrieve all records) and the patch after Save. The
 * reducer is driven directly with explicit times, so merging is deterministic.
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml } from "../model/fetchxmlParser";
import { generateFetchXml } from "../model/fetchxml";
import type { FetchNode, FilterNode } from "../model/nodes";
import type { NewCondition } from "../model/columnFilters";
import { initialState, type BuilderAction, type BuilderState } from "./builderReducer";
import {
	builderHistoryReducer,
	createBuilderHistory,
	pathTo,
	type BuilderHistoryMessage,
	type BuilderHistoryState,
} from "./builderHistory";
import { canRedo, canUndo, lastEntryId, redoLabel, undoLabel } from "./history";

let clock = 1_000_000;
/** Each apply is 5 s after the last unless `at` says otherwise: no accidental merges */
function apply(
	h: BuilderHistoryState,
	action: BuilderAction,
	options: { at?: number; baseline?: boolean } = {},
) {
	clock = options.at ?? clock + 5000;
	return builderHistoryReducer(h, {
		kind: "apply",
		action,
		at: clock,
		baseline: options.baseline,
	});
}
const send = (h: BuilderHistoryState, message: BuilderHistoryMessage) =>
	builderHistoryReducer(h, message);
const undo = (h: BuilderHistoryState) => send(h, { kind: "undo" });
const redo = (h: BuilderHistoryState) => send(h, { kind: "redo" });

const ACCOUNT =
	"<fetch><entity name='account'><attribute name='name'/><attribute name='revenue'/><filter><condition attribute='name' operator='eq' value='a'/></filter></entity></fetch>";
const parsed = (xml: string): FetchNode => {
	const result = parseFetchXml(xml);
	if (!result.fetchNode) throw new Error("parse");
	return result.fetchNode;
};
/** A history whose present is ACCOUNT loaded as a baseline (nothing to undo) */
function loaded(xml = ACCOUNT): BuilderHistoryState {
	return apply(
		createBuilderHistory(initialState),
		{ type: "LOAD_FETCHXML", fetchNode: parsed(xml) },
		{ baseline: true },
	);
}
const q = (h: BuilderHistoryState) => h.present.fetchQuery!;
const rootFilter = (h: BuilderHistoryState): FilterNode => q(h).entity.filters[0];

describe("undo steps", () => {
	it("a baseline load can't be undone", () => {
		const h = loaded();
		expect(canUndo(h)).toBe(false);
		expect(q(h).entity.name).toBe("account");
	});

	it("an added condition goes on undo and comes back with the same id on redo, selected", () => {
		let h = loaded();
		const filterId = rootFilter(h).id;
		h = apply(h, { type: "ADD_CONDITION", filterId });
		const added = rootFilter(h).conditions[1];
		expect(undoLabel(h)).toBe("Add condition");
		h = undo(h);
		expect(rootFilter(h).conditions).toHaveLength(1);
		// The new node is gone: its parent is selected
		expect(h.present.selectedNodeId).toBe(filterId);
		expect(redoLabel(h)).toBe("Add condition");
		h = redo(h);
		expect(rootFilter(h).conditions[1].id).toBe(added.id);
		expect(h.present.selectedNodeId).toBe(added.id);
	});

	it("undoing a delete brings the node back selected, even if another was selected", () => {
		let h = loaded();
		const condition = rootFilter(h).conditions[0];
		h = apply(h, { type: "REMOVE_NODE", nodeId: condition.id });
		expect(undoLabel(h)).toBe("Delete condition");
		h = apply(h, { type: "SELECT_NODE", nodeId: q(h).entity.id });
		h = undo(h);
		expect(rootFilter(h).conditions[0]).toBe(condition);
		expect(h.present.selectedNodeId).toBe(condition.id);
		expect(h.present.selectedNode).toBe(condition);
		// Redoing the delete selects its nearest ancestor that's left
		h = redo(h);
		expect(rootFilter(h).conditions).toHaveLength(0);
		expect(h.present.selectedNodeId).toBe(rootFilter(h).id);
	});

	it("labels each kind of step as the tree names it", () => {
		let h = loaded();
		const entityId = q(h).entity.id;
		const label = (action: BuilderAction) => {
			h = apply(h, action);
			return undoLabel(h);
		};
		expect(label({ type: "ADD_FILTER", parentId: entityId })).toBe("Add filter");
		expect(label({ type: "ADD_SUBFILTER", filterId: rootFilter(h).id })).toBe("Add sub-filter");
		expect(label({ type: "ADD_ORDER", parentId: entityId, attribute: "name" })).toBe(
			"Add order",
		);
		expect(label({ type: "ADD_LINK_ENTITY", parentId: entityId })).toBe("Add link-entity");
		expect(label({ type: "ADD_ALL_ATTRIBUTES", parentId: entityId })).toBe(
			"Add all attributes",
		);
		expect(
			label({ type: "UPDATE_ATTRIBUTES", parentId: entityId, attributeNames: ["name"] }),
		).toBe("Select attributes");
		expect(label({ type: "UPDATE_NODE", nodeId: q(h).id, updates: { count: 10 } })).toBe(
			"Edit query options",
		);
		expect(label({ type: "REMOVE_NODE", nodeId: rootFilter(h).id })).toBe("Delete filter");
		expect(label({ type: "SET_ENTITY", entityName: "contact" })).toBe("Change entity");
		expect(label({ type: "LOAD_FETCHXML", fetchNode: parsed(ACCOUNT) })).toBe("Parse to tree");
	});

	it("a new step clears redo; 100 steps are kept", () => {
		let h = loaded();
		const entityId = q(h).entity.id;
		h = apply(h, { type: "ADD_FILTER", parentId: entityId });
		h = undo(h);
		expect(canRedo(h)).toBe(true);
		h = apply(h, { type: "ADD_ORDER", parentId: entityId, attribute: "name" });
		expect(canRedo(h)).toBe(false);
		for (let i = 0; i < 120; i++) h = apply(h, { type: "ADD_FILTER", parentId: entityId });
		expect(h.past).toHaveLength(100);
	});
});

describe("typing", () => {
	const typing = (h: BuilderHistoryState, value: string, at: number) =>
		apply(
			h,
			{ type: "UPDATE_NODE", nodeId: rootFilter(h).conditions[0].id, updates: { value } },
			{ at },
		);

	it("edits to one field within 1 s are one step, through the layout sync after each", () => {
		let h = loaded();
		const t = clock + 5000;
		h = typing(h, "ab", t);
		h = apply(h, { type: "SYNC_LAYOUT_WITH_FETCHXML" }, { at: t + 100 });
		h = typing(h, "abc", t + 900);
		h = typing(h, "abcd", t + 1800);
		expect(h.past).toHaveLength(1);
		expect(undoLabel(h)).toBe("Edit condition");
		h = undo(h);
		expect(rootFilter(h).conditions[0].value).toBe("a");
	});

	it("a pause over 1 s, another field or selecting a node starts a new step", () => {
		let h = loaded();
		const t = clock + 5000;
		h = typing(h, "ab", t);
		h = typing(h, "abc", t + 1200);
		expect(h.past).toHaveLength(2);
		h = apply(
			h,
			{
				type: "UPDATE_NODE",
				nodeId: rootFilter(h).conditions[0].id,
				updates: { operator: "ne" },
			},
			{ at: t + 1300 },
		);
		expect(h.past).toHaveLength(3);
		h = apply(h, { type: "SELECT_NODE", nodeId: q(h).entity.id }, { at: t + 1400 });
		h = apply(
			h,
			{
				type: "UPDATE_NODE",
				nodeId: rootFilter(h).conditions[0].id,
				updates: { operator: "eq" },
			},
			{ at: t + 1500 },
		);
		expect(h.past).toHaveLength(4);
	});

	it("clearing a query option is a step; setting the value it has isn't (options live in fetch.options)", () => {
		let h = loaded(
			"<fetch distinct='true' top='5'><entity name='account'><attribute name='name'/></entity></fetch>",
		);
		const fetchId = q(h).id;
		h = apply(h, { type: "UPDATE_NODE", nodeId: fetchId, updates: { top: 5 } });
		expect(canUndo(h)).toBe(false);
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		h = undo(h);
		h = apply(h, { type: "UPDATE_NODE", nodeId: fetchId, updates: { distinct: undefined } });
		expect(undoLabel(h)).toBe("Edit query options");
		// A real edit clears redo, so redo can't bring the option back
		expect(canRedo(h)).toBe(false);
		h = undo(h);
		expect(q(h).options.distinct).toBe(true);
	});

	it("an edit that changes nothing isn't a step", () => {
		let h = loaded();
		h = typing(h, "a", clock + 5000);
		expect(canUndo(h)).toBe(false);
		const name = q(h).entity.attributes[0].name;
		h = apply(h, {
			type: "ADD_ATTRIBUTE_BY_NAME",
			parentId: q(h).entity.id,
			attributeName: name,
		});
		expect(canUndo(h)).toBe(false);
	});
});

describe("not undo steps", () => {
	it("selection, column widths and Retrieve all records aren't steps, keep redo and survive undo", () => {
		let h = loaded();
		const entityId = q(h).entity.id;
		const width = (s: BuilderState) =>
			s.columnConfig?.columns.find((c) => c.name === "name")?.width;
		h = apply(h, { type: "ADD_FILTER", parentId: entityId });
		h = undo(h);
		h = apply(h, { type: "SELECT_NODE", nodeId: entityId });
		h = apply(h, { type: "UPDATE_COLUMN_WIDTH", columnName: "name", width: 333 });
		h = apply(h, { type: "SET_RETRIEVE_ALL_RECORDS", enabled: true });
		expect(h.past).toHaveLength(0);
		expect(canRedo(h)).toBe(true);
		h = redo(h);
		expect(width(h.present)).toBe(333);
		expect(h.present.retrieveAllRecords).toBe(true);
		h = undo(h);
		expect(width(h.present)).toBe(333);
		expect(h.present.retrieveAllRecords).toBe(true);
	});

	it("undo asks the tree to reveal the restored selection", () => {
		let h = loaded();
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		const before = h.present.revealSeq;
		h = undo(h);
		expect(h.present.revealSeq).toBe(before + 1);
	});
});

describe("transactions", () => {
	const begin = (h: BuilderHistoryState, label: string, mergeKey?: string) =>
		send(h, { kind: "begin", label, at: (clock += 10), mergeKey });
	const commit = (h: BuilderHistoryState) => send(h, { kind: "commit" });

	it("group several changes into one step with the transaction's label, nested or not", () => {
		let h = loaded();
		const entityId = q(h).entity.id;
		h = begin(h, "Remove columns");
		h = begin(h, "inner");
		h = apply(h, { type: "REMOVE_NODE", nodeId: q(h).entity.attributes[0].id });
		h = commit(h);
		h = apply(h, { type: "REMOVE_NODE", nodeId: q(h).entity.attributes[0].id });
		h = commit(h);
		expect(h.past).toHaveLength(1);
		expect(undoLabel(h)).toBe("Remove columns");
		h = undo(h);
		expect(q(h).entity.attributes).toHaveLength(2);
		expect(q(h).entity.id).toBe(entityId);
	});

	it("record nothing when nothing changed", () => {
		let h = loaded();
		h = commit(begin(h, "Nothing"));
		expect(canUndo(h)).toBe(false);
	});

	it("merge with typing in the same field (an alias rename that also renames its sorts)", () => {
		let h = loaded();
		const attr = q(h).entity.attributes[0];
		const key = `${attr.id}:alias`;
		h = begin(h, "Edit attribute", key);
		h = apply(
			h,
			{ type: "UPDATE_NODE", nodeId: attr.id, updates: { alias: "n" } },
			{ at: clock },
		);
		h = commit(h);
		h = begin(h, "Edit attribute", key);
		h = apply(
			h,
			{ type: "UPDATE_NODE", nodeId: attr.id, updates: { alias: "nm" } },
			{ at: clock },
		);
		h = commit(h);
		expect(h.past).toHaveLength(1);
		h = undo(h);
		expect(q(h).entity.attributes[0].alias).toBeUndefined();
	});
});

describe("Parse to tree, Change entity and the clears", () => {
	it("Parse to tree can be undone: the old tree, its ids and its column filter come back", () => {
		let h = loaded();
		const condition: NewCondition = {
			attribute: "name",
			operator: "eq",
			value: "x",
		} as NewCondition;
		h = apply(h, { type: "ADD_COLUMN_FILTER", condition });
		expect(undoLabel(h)).toBe("Filter column");
		const filterId = h.present.columnFilterId;
		const tree = q(h);
		h = apply(h, {
			type: "LOAD_FETCHXML",
			fetchNode: parsed("<fetch><entity name='contact'/></fetch>"),
		});
		h = undo(h);
		expect(q(h)).toBe(tree);
		expect(h.present.columnFilterId).toBe(filterId);
	});

	it("undoing Change entity brings back the view the query was", () => {
		let h = loaded();
		h = apply(h, { type: "SET_LOADED_VIEW", viewInfo: savedView(q(h)) });
		expect(h.present.viewIdentity?.name).toBe("FXS TEST view");
		h = apply(h, { type: "SET_ENTITY", entityName: "contact" });
		expect(h.present.viewIdentity).toBeNull();
		h = undo(h);
		expect(h.present.viewIdentity?.name).toBe("FXS TEST view");
	});

	it("New, Open view and baseline loads clear undo and redo", () => {
		let h = loaded();
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		h = undo(h);
		h = apply(h, { type: "NEW_QUERY" });
		expect([canUndo(h), canRedo(h)]).toEqual([false, false]);
		h = loaded();
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		const fetchNode = parsed(ACCOUNT);
		h = apply(h, {
			type: "LOAD_VIEW",
			fetchNode,
			viewInfo: { ...savedView(fetchNode), originalFetchXml: generateFetchXml(fetchNode) },
		});
		expect(canUndo(h)).toBe(false);
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		h = apply(h, { type: "LOAD_FETCHXML", fetchNode: parsed(ACCOUNT) }, { baseline: true });
		expect(canUndo(h)).toBe(false);
	});
});

describe("after Save", () => {
	it("stored states of the saved query take the saved view; only the saved query runs by its id", () => {
		let h = loaded();
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		const saved = savedView(q(h));
		h = apply(h, { type: "SET_LOADED_VIEW", viewInfo: saved });
		// Save isn't a step
		expect(h.past).toHaveLength(1);
		h = undo(h);
		// The older query: the saved view (for Save and Unsaved changes), but not run by its id
		expect(h.present.viewIdentity?.view.id).toBe(saved.id);
		expect(h.present.loadedView).toBeNull();
		h = redo(h);
		expect(h.present.loadedView).toEqual(saved);
	});

	it("a stored state of another entity is left alone", () => {
		let h = loaded();
		h = apply(h, { type: "SET_ENTITY", entityName: "contact" });
		h = apply(h, { type: "SET_LOADED_VIEW", viewInfo: savedView(q(h), "contact") });
		h = undo(h);
		expect(q(h).entity.name).toBe("account");
		expect(h.present.viewIdentity).toBeNull();
	});
});

describe("an Undo toast's step", () => {
	it("is undone only while it's still the last step", () => {
		let h = loaded();
		h = apply(h, { type: "ADD_FILTER", parentId: q(h).entity.id });
		const id = lastEntryId(h)!;
		h = apply(h, { type: "ADD_ORDER", parentId: q(h).entity.id, attribute: "name" });
		expect(send(h, { kind: "undoIfLast", id })).toBe(h);
		h = undo(h);
		h = send(h, { kind: "undoIfLast", id });
		expect(q(h).entity.filters).toHaveLength(1);
	});
});

describe("pathTo", () => {
	it("lists the node and its ancestors, nearest first", () => {
		const h = loaded();
		const condition = rootFilter(h).conditions[0];
		expect(pathTo(q(h), condition.id)).toEqual([
			condition.id,
			rootFilter(h).id,
			q(h).entity.id,
			q(h).id,
		]);
		expect(pathTo(q(h), "nope")).toBeNull();
		expect(pathTo(null, condition.id)).toBeNull();
	});
});

function savedView(query: FetchNode, entity = "account") {
	return {
		id: "11111111-1111-1111-1111-111111111111",
		type: "personal" as const,
		originalFetchXml: generateFetchXml(query),
		entitySetName: entity === "account" ? "accounts" : "contacts",
		name: "FXS TEST view",
	};
}
