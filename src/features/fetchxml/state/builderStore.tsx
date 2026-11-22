/**
 * Builder state management with React Context and useReducer
 * Manages FetchXML query structure and tree operations
 */

import { useMemo, useReducer, type ReactNode } from "react";
import {
	BuilderContext,
	type BuilderContextValue,
	type LoadFetchXmlOptions,
} from "./builderContext";
import type { NodeId } from "../model/nodes";
import type { LayoutXmlConfig } from "../model/layoutxml";
import { parseFetchXml, type ParseResult } from "../model/fetchxmlParser";
import { generateFetchXml } from "../model/fetchxml";
import type { SortPlan } from "../model/queryColumns";
import type { NewCondition } from "../model/columnFilters";
import {
	builderReducer,
	generateId,
	initialState,
	type ViewLoadInfo,
	type ViewSourceInfo,
} from "./builderReducer";

export type {
	BuilderState,
	LoadedViewState,
	ViewIdentity,
	ViewLoadInfo,
	ViewSourceInfo,
} from "./builderReducer";

export function BuilderProvider({ children }: { children: ReactNode }) {
	const [state, dispatch] = useReducer(builderReducer, initialState);

	// The actions are created once (dispatch never changes) and the context value only changes with the
	// state, so consumers can list builder actions as hook dependencies without re-running every render (#65)
	const actions = useMemo(
		() => ({
			setEntity: (entityName: string) => dispatch({ type: "SET_ENTITY", entityName }),
			selectNode: (nodeId: NodeId) => dispatch({ type: "SELECT_NODE", nodeId }),
			/** Select a node and reveal it in the tree, even when it is already selected */
			revealNode: (nodeId: NodeId) => dispatch({ type: "REVEAL_NODE", nodeId }),
			addAttribute: (parentId: NodeId) => dispatch({ type: "ADD_ATTRIBUTE", parentId }),
			addAttributeByName: (parentId: NodeId, attributeName: string) =>
				dispatch({ type: "ADD_ATTRIBUTE_BY_NAME", parentId, attributeName }),
			updateAttributes: (parentId: NodeId, attributeNames: string[]) =>
				dispatch({ type: "UPDATE_ATTRIBUTES", parentId, attributeNames }),
			addAllAttributes: (parentId: NodeId) =>
				dispatch({ type: "ADD_ALL_ATTRIBUTES", parentId }),
			addOrder: (parentId: NodeId, attribute?: string) =>
				dispatch({ type: "ADD_ORDER", parentId, attribute }),
			addFilter: (parentId: NodeId) => dispatch({ type: "ADD_FILTER", parentId }),
			addSubfilter: (filterId: NodeId) => dispatch({ type: "ADD_SUBFILTER", filterId }),
			addCondition: (filterId: NodeId) => dispatch({ type: "ADD_CONDITION", filterId }),
			addLinkEntity: (parentId: NodeId) => dispatch({ type: "ADD_LINK_ENTITY", parentId }),
			addLinkEntityWithConfig: (
				parentId: NodeId,
				name: string,
				from: string,
				to: string,
				linkType: "inner" | "outer",
				relationshipType?: "N1" | "1N" | "NN",
			): NodeId => {
				const id = generateId();
				dispatch({
					type: "ADD_LINK_ENTITY_WITH_CONFIG",
					id,
					parentId,
					name,
					from,
					to,
					linkType,
					relationshipType,
				});
				return id;
			},
			removeNode: (nodeId: NodeId) => dispatch({ type: "REMOVE_NODE", nodeId }),
			updateNode: (nodeId: NodeId, updates: Record<string, unknown>) =>
				dispatch({ type: "UPDATE_NODE", nodeId, updates }),
			newQuery: () => dispatch({ type: "NEW_QUERY" }),
			loadFetchXml: (xmlString: string, options?: LoadFetchXmlOptions): ParseResult => {
				const result = parseFetchXml(xmlString);
				if (result.success && result.fetchNode) {
					dispatch({
						type: "LOAD_FETCHXML",
						fetchNode: result.fetchNode,
						preserveLayout: options?.preserveLayout,
					});
				}
				return result;
			},
			loadView: (
				xmlString: string,
				viewInfo: ViewSourceInfo,
				layoutXml?: string,
			): ParseResult => {
				const result = parseFetchXml(xmlString);
				if (result.success && result.fetchNode) {
					dispatch({
						type: "LOAD_VIEW",
						fetchNode: result.fetchNode,
						viewInfo: {
							...viewInfo,
							// The stored FetchXML has defaults the generator leaves out (version, mapping,
							// descending="false"...): compare in generated form, or the view never runs by ID (#93)
							originalFetchXml: generateFetchXml(result.fetchNode),
						},
						layoutXml,
					});
				}
				return result;
			},
			clearLoadedView: () => dispatch({ type: "CLEAR_LOADED_VIEW" }),
			setLoadedView: (viewInfo: ViewLoadInfo, layoutXml?: string) =>
				dispatch({
					type: "SET_LOADED_VIEW",
					// originalFetchXml is supplied by the caller, generated from the tree that was saved
					viewInfo: { ...viewInfo },
					layoutXml,
				}),
			setColumnConfig: (config: LayoutXmlConfig) =>
				dispatch({ type: "SET_COLUMN_CONFIG", config }),
			updateColumnWidth: (columnName: string, width: number) =>
				dispatch({ type: "UPDATE_COLUMN_WIDTH", columnName, width }),
			reorderColumns: (fromIndex: number, toIndex: number) =>
				dispatch({ type: "REORDER_COLUMNS", fromIndex, toIndex }),
			syncLayoutWithFetchXml: (
				attributeTypeMap?: Map<string, string>,
				primaryIdAttribute?: string,
			) =>
				dispatch({
					type: "SYNC_LAYOUT_WITH_FETCHXML",
					attributeTypeMap,
					primaryIdAttribute,
				}),
			setSort: (plan: SortPlan, descending: boolean, isMultiSort: boolean) =>
				dispatch({ type: "SET_SORT", plan, descending, isMultiSort }),
			setRetrieveAllRecords: (enabled: boolean) =>
				dispatch({ type: "SET_RETRIEVE_ALL_RECORDS", enabled }),
			addColumnFilter: (condition: NewCondition) =>
				dispatch({ type: "ADD_COLUMN_FILTER", condition }),
			removeColumnFilter: (conditionId: NodeId) =>
				dispatch({ type: "REMOVE_COLUMN_FILTER", conditionId }),
			clearColumnFilters: () => dispatch({ type: "CLEAR_COLUMN_FILTERS" }),
		}),
		[],
	);
	const contextValue = useMemo<BuilderContextValue>(
		() => ({ ...state, ...actions }),
		[state, actions],
	);

	return <BuilderContext.Provider value={contextValue}>{children}</BuilderContext.Provider>;
}
