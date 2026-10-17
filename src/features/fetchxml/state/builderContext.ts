/**
 * Builder context and hook, kept separate from BuilderProvider so the
 * provider module only exports components (React Fast Refresh).
 */

import { createContext, useContext } from "react";
import type { NodeId } from "../model/nodes";
import type { LayoutXmlConfig } from "../model/layoutxml";
import type { ParseResult } from "../model/fetchxmlParser";
import type { BuilderState, ViewLoadInfo } from "./builderStore";
import type { SortPlan } from "../model/queryColumns";
import type { NewCondition } from "../model/columnFilters";

export interface LoadFetchXmlOptions {
	/** Keep the current column layout, merging new attributes into it (editor sync, #102) */
	preserveLayout?: boolean;
}

export interface BuilderContextValue extends BuilderState {
	setEntity: (entityName: string) => void;
	selectNode: (nodeId: NodeId) => void;
	addAttribute: (parentId: NodeId) => void;
	/** Add a specific attribute by name to an entity or link-entity */
	addAttributeByName: (parentId: NodeId, attributeName: string) => void;
	/** Update attributes for an entity or link-entity (bulk add/remove) */
	updateAttributes: (parentId: NodeId, attributeNames: string[]) => void;
	addAllAttributes: (parentId: NodeId) => void;
	addOrder: (parentId: NodeId) => void;
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
	loadView: (xmlString: string, viewInfo: ViewLoadInfo, layoutXml?: string) => ParseResult;
	/** Clear the loaded view info (called when FetchXML is manually edited) */
	clearLoadedView: () => void;
	/** Update loaded view info after save (without re-parsing FetchXML) */
	setLoadedView: (viewInfo: ViewLoadInfo) => void;
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
}

export const BuilderContext = createContext<BuilderContextValue | null>(null);

export function useBuilder() {
	const context = useContext(BuilderContext);
	if (!context) {
		throw new Error("useBuilder must be used within BuilderProvider");
	}
	return context;
}
