/**
 * Builder context and hook, kept separate from BuilderProvider so the
 * provider module only exports components (React Fast Refresh).
 */

import { createContext, useContext } from "react";
import type { NodeId } from "../model/nodes";
import type { LayoutXmlConfig } from "../model/layoutxml";
import type { ParseResult } from "../model/fetchxmlParser";
import type { BuilderState, ViewLoadInfo, ViewSourceInfo } from "./builderStore";
import type { SortPlan } from "../model/queryColumns";
import type { NewCondition } from "../model/columnFilters";

export interface LoadFetchXmlOptions {
	/** Keep the current column layout, merging new attributes into it (editor sync, #102) */
	preserveLayout?: boolean;
	/**
	 * A starting point, not an undo step: the load clears the undo history (the query carried over a
	 * connection change, one sent from another tool), so the first Ctrl+Z can't blank the tree (#192)
	 */
	baseline?: boolean;
}

export interface TransactionOptions {
	mergeKey?: string;
	target?: NodeId;
}

/** The query's undo history (#192, builderHistory.ts). The functions keep their identity. */
export interface BuilderHistoryApi {
	canUndo: boolean;
	canRedo: boolean;
	/** "Delete condition": what Undo would undo, for "Undo: {label}" */
	undoLabel: string | null;
	redoLabel: string | null;
	/** The step Undo would undo, for an Undo toast to check it's still the last one (waves 2–3) */
	lastEntryId: number | null;
	undo: () => void;
	redo: () => void;
	/** Undo only if step `id` is still the last one (a toast's Undo) */
	undoIfLast: (id: number) => void;
	/**
	 * Runs `fn`'s builder changes as one undo step called `label`. `mergeKey` merges it with edits of
	 * the same key (typing); `target` is the node to select after undoing or redoing it.
	 */
	transaction: (label: string, fn: () => void, options?: TransactionOptions) => void;
}

export interface BuilderContextValue extends BuilderState {
	/** Start a query on an entity; `baseline` (a launch with an entity) isn't an undo step (#192) */
	setEntity: (entityName: string, options?: { baseline?: boolean }) => void;
	selectNode: (nodeId: NodeId) => void;
	revealNode: (nodeId: NodeId) => void;
	addAttribute: (parentId: NodeId) => void;
	/** Add a specific attribute by name to an entity or link-entity */
	addAttributeByName: (parentId: NodeId, attributeName: string) => void;
	/** Update attributes for an entity or link-entity (bulk add/remove) */
	updateAttributes: (parentId: NodeId, attributeNames: string[]) => void;
	addAllAttributes: (parentId: NodeId) => void;
	/** Add an order to an entity or link-entity: on `attribute` when given, else a placeholder to edit */
	addOrder: (parentId: NodeId, attribute?: string) => void;
	addFilter: (parentId: NodeId) => void;
	addSubfilter: (filterId: NodeId) => void;
	addCondition: (filterId: NodeId) => void;
	addLinkEntity: (parentId: NodeId) => void;
	/** Add a link-entity with specific configuration, returns the new link-entity's ID */
	addLinkEntityWithConfig: (
		parentId: NodeId,
		name: string,
		from: string,
		to: string,
		linkType: "inner" | "outer",
		relationshipType?: "N1" | "1N" | "NN",
	) => NodeId;
	removeNode: (nodeId: NodeId) => void;
	updateNode: (nodeId: NodeId, updates: Record<string, unknown>) => void;
	newQuery: () => void;
	loadFetchXml: (xmlString: string, options?: LoadFetchXmlOptions) => ParseResult;
	/** Load a view's FetchXML while preserving view info for execution optimization */
	loadView: (xmlString: string, viewInfo: ViewSourceInfo, layoutXml?: string) => ParseResult;
	/** Clear the loaded view info (called when FetchXML is manually edited) */
	clearLoadedView: () => void;
	/** Update loaded view info after save (without re-parsing FetchXML); layoutXml is what was saved */
	setLoadedView: (viewInfo: ViewLoadInfo, layoutXml?: string) => void;
	/** Set the column configuration (from parsed layoutxml or generated) */
	setColumnConfig: (config: LayoutXmlConfig) => void;
	/** Update a single column's width */
	updateColumnWidth: (columnName: string, width: number) => void;
	/** Reorder columns by moving one from fromIndex to toIndex */
	reorderColumns: (fromIndex: number, toIndex: number) => void;
	/** Sync layout with current FetchXML (regenerate/merge as needed) */
	syncLayoutWithFetchXml: (
		attributeTypeMap?: Map<string, string>,
		primaryIdAttribute?: string,
	) => void;
	/**
	 * Sort by a result column (planned by queryColumns.planSortByColumn, #114). With isMultiSort the
	 * order is added (or its direction changed); otherwise it replaces every order in the query.
	 */
	setSort: (plan: SortPlan, descending: boolean, isMultiSort: boolean) => void;
	/**
	 * Enable or disable "Retrieve All Records" mode.
	 * When enabled, execution fetches all pages progressively.
	 */
	setRetrieveAllRecords: (enabled: boolean) => void;
	/** Add a results-grid column filter's condition to the query (#23) */
	addColumnFilter: (condition: NewCondition) => void;
	/** Remove one column filter's condition */
	removeColumnFilter: (conditionId: NodeId) => void;
	/** Remove every column filter */
	clearColumnFilters: () => void;
	/** Undo and redo (#192) */
	history: BuilderHistoryApi;
}

export const BuilderContext = createContext<BuilderContextValue | null>(null);

export function useBuilder() {
	const context = useContext(BuilderContext);
	if (!context) {
		throw new Error("useBuilder must be used within BuilderProvider");
	}
	return context;
}
