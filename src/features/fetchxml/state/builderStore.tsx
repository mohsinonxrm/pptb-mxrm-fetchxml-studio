/**
 * Builder state management with React Context and useReducer
 * Manages FetchXML query structure and tree operations
 */

import { createContext, useContext, useReducer, type ReactNode } from "react";
import type {
	FetchNode,
	EntityNode,
	AttributeNode,
	AllAttributesNode,
	OrderNode,
	FilterNode,
	ConditionNode,
	LinkEntityNode,
	NodeId,
} from "../model/nodes";
import type { LayoutXmlConfig } from "../model/layoutxml";
import {
	generateLayoutFromFetchXml,
	mergeLayoutWithFetchXml,
	updateColumnWidth as updateLayoutColumnWidth,
	reorderColumns as reorderLayoutColumns,
	parseLayoutXml,
} from "../model/layoutxml";
import { parseFetchXml, type ParseResult } from "../model/fetchxmlParser";

// Temporary ID generator
let idCounter = 0;
const generateId = (): string => `node_${++idCounter}`;

type SelectedNode =
	| FetchNode
	| EntityNode
	| AttributeNode
	| AllAttributesNode
	| OrderNode
	| FilterNode
	| ConditionNode
	| LinkEntityNode
	| null;

/**
 * Information about a loaded view for execution optimization
 * Used to determine if the view can be executed via savedQuery/userQuery
 */
interface LoadedViewState {
	/** View ID (savedqueryid or userqueryid) */
	id: string;
	/** View type for determining execution method */
	type: "system" | "personal";
	/** Original FetchXML from the view - used for comparison */
	originalFetchXml: string;
	/** Entity set name for the execution URL */
	entitySetName: string;
	/** View name for display purposes */
	name: string;
}

interface BuilderState {
	fetchQuery: FetchNode | null;
	selectedNodeId: NodeId | null;
	selectedNode: SelectedNode;
	/** Loaded view info - null if query was not loaded from a view or has been modified */
	loadedView: LoadedViewState | null;
	/** Column layout configuration - generated from FetchXML or loaded from view */
	columnConfig: LayoutXmlConfig | null;
	/** Track if layout needs regeneration due to FetchXML changes */
	layoutNeedsSync: boolean;
}

type BuilderAction =
	| { type: "SET_ENTITY"; entityName: string }
	| { type: "SELECT_NODE"; nodeId: NodeId }
	| { type: "ADD_ATTRIBUTE"; parentId: NodeId }
	| { type: "ADD_ATTRIBUTE_BY_NAME"; parentId: NodeId; attributeName: string }
	| { type: "ADD_ALL_ATTRIBUTES"; parentId: NodeId }
	| { type: "ADD_ORDER"; parentId: NodeId }
	| { type: "ADD_FILTER"; parentId: NodeId }
	| { type: "ADD_SUBFILTER"; filterId: NodeId }
	| { type: "ADD_CONDITION"; filterId: NodeId }
	| { type: "ADD_LINK_ENTITY"; parentId: NodeId }
	| { type: "REMOVE_NODE"; nodeId: NodeId }
	| { type: "UPDATE_NODE"; nodeId: NodeId; updates: Record<string, unknown> }
	| { type: "NEW_QUERY" }
	| { type: "LOAD_FETCHXML"; fetchNode: FetchNode }
	| { type: "LOAD_VIEW"; fetchNode: FetchNode; viewInfo: LoadedViewState; layoutXml?: string }
	| { type: "CLEAR_LOADED_VIEW" }
	| { type: "SET_COLUMN_CONFIG"; config: LayoutXmlConfig }
	| { type: "UPDATE_COLUMN_WIDTH"; columnName: string; width: number }
	| { type: "REORDER_COLUMNS"; fromIndex: number; toIndex: number }
	| { type: "SYNC_LAYOUT_WITH_FETCHXML"; attributeTypeMap?: Map<string, string> }
	| {
			type: "SET_SORT";
			attribute: string;
			descending: boolean;
			isMultiSort: boolean;
			entityName?: string;
	  };

const initialState: BuilderState = {
	fetchQuery: null,
	selectedNodeId: null,
	selectedNode: null,
	loadedView: null,
	columnConfig: null,
	layoutNeedsSync: false,
};

// Helper to find node by ID in the tree
function findNodeById(node: unknown, targetId: NodeId): SelectedNode {
	if (!node || typeof node !== "object") return null;

	const n = node as Record<string, unknown>;
	if (n.id === targetId) return node as SelectedNode;

	// Search in arrays
	const arrayProps = ["attributes", "orders", "filters", "conditions", "subfilters", "links"];
	for (const prop of arrayProps) {
		if (Array.isArray(n[prop])) {
			for (const item of n[prop]) {
				const found = findNodeById(item, targetId);
				if (found) return found;
			}
		}
	}

	// Search in nested objects
	if (n.entity) {
		const found = findNodeById(n.entity, targetId);
		if (found) return found;
	}

	if (n.allAttributes) {
		const found = findNodeById(n.allAttributes, targetId);
		if (found) return found;
	}

	return null;
}

// Helper to update a node in the tree immutably
function updateNodeInTree<T extends { id: NodeId }>(
	node: T,
	targetId: NodeId,
	updater: (node: unknown) => unknown
): T {
	// If this is the target node, apply the updater
	if (node.id === targetId) {
		return updater(node) as T;
	}

	// Clone the node
	const updated = { ...node } as Record<string, unknown>;
	let changed = false;

	// Update arrays
	const arrayProps = ["attributes", "orders", "filters", "conditions", "subfilters", "links"];
	for (const prop of arrayProps) {
		const value = (node as Record<string, unknown>)[prop];
		if (Array.isArray(value)) {
			const updatedArray = (value as Array<{ id: NodeId }>).map((item) => {
				const result = updateNodeInTree(item, targetId, updater);
				if (result !== item) changed = true;
				return result;
			});
			if (changed) {
				updated[prop] = updatedArray;
			}
		}
	}

	// Update nested objects
	const nodeRecord = node as Record<string, unknown>;
	if (nodeRecord.entity) {
		const updatedEntity = updateNodeInTree(nodeRecord.entity as EntityNode, targetId, updater);
		if (updatedEntity !== nodeRecord.entity) {
			updated.entity = updatedEntity;
			changed = true;
		}
	}

	if (nodeRecord.allAttributes) {
		const updatedAllAttrs = updateNodeInTree(
			nodeRecord.allAttributes as AllAttributesNode,
			targetId,
			updater
		);
		if (updatedAllAttrs !== nodeRecord.allAttributes) {
			updated.allAttributes = updatedAllAttrs;
			changed = true;
		}
	}

	return (changed ? updated : node) as T;
}

// Helper to remove a node from the tree immutably
function removeNodeFromTree<T extends { id: NodeId }>(node: T, targetId: NodeId): T | null {
	const nodeRecord = node as Record<string, unknown>;

	// Can't remove the root fetch node or its entity
	if (node.id === targetId && (nodeRecord.type === "fetch" || nodeRecord.type === "entity")) {
		return node;
	}

	// Clone the node
	const updated = { ...node } as Record<string, unknown>;
	let changed = false;

	// Remove from arrays
	const arrayProps = ["attributes", "orders", "filters", "conditions", "subfilters", "links"];
	for (const prop of arrayProps) {
		const value = nodeRecord[prop];
		if (Array.isArray(value)) {
			const filteredArray = (value as Array<{ id: NodeId }>)
				.filter((item) => item.id !== targetId)
				.map((item) => removeNodeFromTree(item, targetId))
				.filter((item) => item !== null) as Array<{ id: NodeId }>;

			// Check if anything changed (length or nested changes)
			if (
				filteredArray.length !== value.length ||
				filteredArray.some((item, i) => item !== value[i])
			) {
				updated[prop] = filteredArray;
				changed = true;
			}
		}
	}

	// Remove from nested objects
	if (nodeRecord.entity) {
		const updatedEntity = removeNodeFromTree(nodeRecord.entity as EntityNode, targetId);
		if (updatedEntity !== nodeRecord.entity) {
			updated.entity = updatedEntity;
			changed = true;
		}
	}

	// Handle allAttributes removal
	if (nodeRecord.allAttributes && (nodeRecord.allAttributes as AllAttributesNode).id === targetId) {
		delete updated.allAttributes;
		changed = true;
	}

	return (changed ? updated : node) as T;
}

function builderReducer(state: BuilderState, action: BuilderAction): BuilderState {
	switch (action.type) {
		case "NEW_QUERY":
			return initialState;

		case "LOAD_FETCHXML": {
			// Load a pre-parsed FetchNode tree (from parser or editor)
			// This clears loadedView since it's a manual load, not from a saved view
			const fetchNode = action.fetchNode;
			// Generate initial layout from the FetchXML
			const newConfig = generateLayoutFromFetchXml(fetchNode);
			return {
				fetchQuery: fetchNode,
				selectedNodeId: fetchNode.entity.id,
				selectedNode: fetchNode.entity,
				loadedView: null, // Clear view info - this was manual edit
				columnConfig: newConfig,
				layoutNeedsSync: false,
			};
		}

		case "LOAD_VIEW": {
			// Load a pre-parsed FetchNode tree from a saved view
			// Preserves view info for execution optimization
			const fetchNode = action.fetchNode;

			// If layoutXml is provided, parse it; otherwise generate from FetchXML
			let newConfig: LayoutXmlConfig;
			if (action.layoutXml) {
				try {
					newConfig = parseLayoutXml(action.layoutXml);
				} catch (e) {
					console.warn("Failed to parse view layoutxml, generating from FetchXML:", e);
					newConfig = generateLayoutFromFetchXml(fetchNode);
				}
			} else {
				newConfig = generateLayoutFromFetchXml(fetchNode);
			}

			return {
				fetchQuery: fetchNode,
				selectedNodeId: fetchNode.entity.id,
				selectedNode: fetchNode.entity,
				loadedView: action.viewInfo, // Track the loaded view
				columnConfig: newConfig,
				layoutNeedsSync: false,
			};
		}

		case "CLEAR_LOADED_VIEW": {
			// Called when FetchXML is edited in the editor
			return {
				...state,
				loadedView: null,
				layoutNeedsSync: true, // May need to regenerate layout
			};
		}

		case "SET_ENTITY": {
			const entityNode: EntityNode = {
				id: generateId(),
				type: "entity",
				name: action.entityName,
				attributes: [],
				orders: [],
				filters: [],
				links: [],
			};

			const fetchNode: FetchNode = {
				id: generateId(),
				type: "fetch",
				entity: entityNode,
				options: {},
			};

			return {
				fetchQuery: fetchNode,
				selectedNodeId: entityNode.id,
				selectedNode: entityNode,
				loadedView: null, // Clear view info - entity changed
				columnConfig: null, // Reset layout - no attributes yet
				layoutNeedsSync: false,
			};
		}

		case "SELECT_NODE": {
			if (!state.fetchQuery) return state;
			const node = findNodeById(state.fetchQuery, action.nodeId);
			return {
				...state,
				selectedNodeId: action.nodeId,
				selectedNode: node,
			};
		}

		case "ADD_ATTRIBUTE": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity or link-entity)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (!parent || (parent.type !== "entity" && parent.type !== "link-entity")) {
				return state;
			}

			// Create new attribute node
			const newAttribute: AttributeNode = {
				id: generateId(),
				type: "attribute",
				name: "new_attribute",
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const n = node as EntityNode | LinkEntityNode;
				return {
					...n,
					attributes: [...n.attributes, newAttribute],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newAttribute.id,
				selectedNode: newAttribute,
				loadedView: null, // Clear view info - query modified
				layoutNeedsSync: true, // New attribute added - layout needs update
			};
		}

		case "ADD_ATTRIBUTE_BY_NAME": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity or link-entity)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (!parent || (parent.type !== "entity" && parent.type !== "link-entity")) {
				return state;
			}

			// Check if attribute already exists
			const existingAttr = (parent as EntityNode | LinkEntityNode).attributes.find(
				(a) => a.name === action.attributeName
			);
			if (existingAttr) {
				// Attribute already exists, just select it
				return {
					...state,
					selectedNodeId: existingAttr.id,
					selectedNode: existingAttr,
				};
			}

			// Create new attribute node with the specified name
			const newAttrByName: AttributeNode = {
				id: generateId(),
				type: "attribute",
				name: action.attributeName,
			};

			// Update the tree immutably
			const updatedFetchByName = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const n = node as EntityNode | LinkEntityNode;
				return {
					...n,
					attributes: [...n.attributes, newAttrByName],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetchByName,
				selectedNodeId: newAttrByName.id,
				selectedNode: newAttrByName,
				loadedView: null, // Clear view info - query modified
				layoutNeedsSync: true, // New attribute added - layout needs update
			};
		}

		case "ADD_ALL_ATTRIBUTES": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity or link-entity)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (!parent || (parent.type !== "entity" && parent.type !== "link-entity")) {
				return state;
			}

			// Create new all-attributes node
			const newAllAttributes: AllAttributesNode = {
				id: generateId(),
				type: "all-attributes",
				enabled: true,
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const n = node as EntityNode | LinkEntityNode;
				return {
					...n,
					allAttributes: newAllAttributes,
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newAllAttributes.id,
				selectedNode: newAllAttributes,
				loadedView: null, // Clear view info - query modified
				layoutNeedsSync: true, // All attributes changes layout
			};
		}

		case "ADD_ORDER": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity or link-entity)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (!parent || (parent.type !== "entity" && parent.type !== "link-entity")) {
				return state;
			}

			// Create new order node
			const newOrder: OrderNode = {
				id: generateId(),
				type: "order",
				attribute: "new_attribute",
				descending: false,
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const n = node as EntityNode | LinkEntityNode;
				return {
					...n,
					orders: [...n.orders, newOrder],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newOrder.id,
				selectedNode: newOrder,
				loadedView: null, // Clear view info - query modified
			};
		}

		case "ADD_FILTER": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity or link-entity)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (!parent || (parent.type !== "entity" && parent.type !== "link-entity")) {
				return state;
			}

			// Create new filter node
			const newFilter: FilterNode = {
				id: generateId(),
				type: "filter",
				conjunction: "and",
				conditions: [],
				subfilters: [],
				links: [],
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const n = node as EntityNode | LinkEntityNode;
				return {
					...n,
					filters: [...n.filters, newFilter],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newFilter.id,
				selectedNode: newFilter,
				loadedView: null, // Clear view info - query modified
			};
		}

		case "ADD_SUBFILTER": {
			if (!state.fetchQuery) return state;

			// Find parent filter node
			const parent = findNodeById(state.fetchQuery, action.filterId);
			if (!parent || parent.type !== "filter") {
				return state;
			}

			// Create new nested filter node
			const newSubfilter: FilterNode = {
				id: generateId(),
				type: "filter",
				conjunction: "and",
				conditions: [],
				subfilters: [],
				links: [],
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.filterId, (node) => {
				const n = node as FilterNode;
				return {
					...n,
					subfilters: [...n.subfilters, newSubfilter],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newSubfilter.id,
				selectedNode: newSubfilter,
				loadedView: null, // Clear view info - query modified
			};
		}

		case "ADD_CONDITION": {
			if (!state.fetchQuery) return state;

			// Find parent filter node
			const parent = findNodeById(state.fetchQuery, action.filterId);
			if (!parent || parent.type !== "filter") {
				return state;
			}

			// Create new condition node
			const newCondition: ConditionNode = {
				id: generateId(),
				type: "condition",
				attribute: "new_attribute",
				operator: "eq",
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.filterId, (node) => {
				const n = node as FilterNode;
				return {
					...n,
					conditions: [...n.conditions, newCondition],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newCondition.id,
				selectedNode: newCondition,
				loadedView: null, // Clear view info - query modified
			};
		}

		case "ADD_LINK_ENTITY": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity, link-entity, or filter)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (
				!parent ||
				(parent.type !== "entity" && parent.type !== "link-entity" && parent.type !== "filter")
			) {
				return state;
			}

			// Create new link-entity node
			// For filter parents, use "any" link type by default (common use case)
			const newLinkEntity: LinkEntityNode = {
				id: generateId(),
				type: "link-entity",
				name: "new_entity",
				from: "id_field",
				to: "parent_id_field",
				linkType: parent.type === "filter" ? "any" : "inner",
				attributes: [],
				orders: [],
				filters: [],
				links: [],
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const typedNode = node as EntityNode | LinkEntityNode | FilterNode;
				if (typedNode.type === "filter") {
					const n = typedNode as FilterNode;
					return {
						...n,
						links: [...(n.links || []), newLinkEntity],
					};
				} else {
					const n = typedNode as EntityNode | LinkEntityNode;
					return {
						...n,
						links: [...n.links, newLinkEntity],
					};
				}
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newLinkEntity.id,
				selectedNode: newLinkEntity,
				loadedView: null, // Clear view info - query modified
				layoutNeedsSync: true, // Link-entity may add columns
			};
		}

		case "REMOVE_NODE": {
			if (!state.fetchQuery) return state;

			// Remove the node from the tree
			const updatedFetch = removeNodeFromTree(state.fetchQuery, action.nodeId);
			if (!updatedFetch) return state;

			// If we're removing the selected node, clear selection
			const newSelectedId = state.selectedNodeId === action.nodeId ? null : state.selectedNodeId;
			const newSelectedNode = newSelectedId ? findNodeById(updatedFetch, newSelectedId) : null;

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newSelectedId,
				selectedNode: newSelectedNode,
				loadedView: null, // Clear view info - query modified
				layoutNeedsSync: true, // Removed node may affect columns
			};
		}

		case "UPDATE_NODE": {
			if (!state.fetchQuery) return state;

			// Find the node to update
			const targetNode = findNodeById(state.fetchQuery, action.nodeId);
			if (!targetNode) return state;

			// Update the tree immutably - apply updates to the target node
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.nodeId, (node) => {
				const n = node as Record<string, unknown>;

				// For FetchNode, merge updates into options property
				if (n.type === "fetch") {
					const fetchNode = node as unknown as FetchNode;
					return {
						...fetchNode,
						options: {
							...fetchNode.options,
							...action.updates,
						},
					} as FetchNode;
				}

				// For other nodes, apply updates directly
				return {
					...n,
					...action.updates,
				};
			});

			// Update selected node if it's the one we just modified
			const newSelectedNode =
				state.selectedNodeId === action.nodeId
					? findNodeById(updatedFetch, action.nodeId)
					: state.selectedNode;

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNode: newSelectedNode,
				loadedView: null, // Clear view info - query modified
				layoutNeedsSync: true, // Layout may need to sync with FetchXML changes
			};
		}

		case "SET_COLUMN_CONFIG": {
			return {
				...state,
				columnConfig: action.config,
				layoutNeedsSync: false,
			};
		}

		case "UPDATE_COLUMN_WIDTH": {
			if (!state.columnConfig) return state;
			return {
				...state,
				columnConfig: updateLayoutColumnWidth(state.columnConfig, action.columnName, action.width),
			};
		}

		case "REORDER_COLUMNS": {
			if (!state.columnConfig) return state;
			return {
				...state,
				columnConfig: reorderLayoutColumns(state.columnConfig, action.fromIndex, action.toIndex),
			};
		}

		case "SYNC_LAYOUT_WITH_FETCHXML": {
			if (!state.fetchQuery) return state;

			let newConfig: LayoutXmlConfig;
			if (state.columnConfig) {
				// Merge existing config with FetchXML changes (preserves widths/order where possible)
				newConfig = mergeLayoutWithFetchXml(
					state.columnConfig,
					state.fetchQuery,
					action.attributeTypeMap
				);
			} else {
				// Generate new config from FetchXML
				newConfig = generateLayoutFromFetchXml(state.fetchQuery, action.attributeTypeMap);
			}

			return {
				...state,
				columnConfig: newConfig,
				layoutNeedsSync: false,
			};
		}

		case "SET_SORT": {
			if (!state.fetchQuery?.entity) return state;

			const { attribute, descending, isMultiSort, entityName } = action;

			// Create new order node
			const newOrder: OrderNode = {
				id: generateId(),
				type: "order",
				attribute,
				descending: descending || undefined,
				entityname: entityName,
			};

			// Get current orders
			const currentOrders = state.fetchQuery.entity.orders || [];

			let newOrders: OrderNode[];

			if (isMultiSort) {
				// Multi-sort: check if this attribute already has an order
				const existingIndex = currentOrders.findIndex(
					(o) =>
						o.attribute === attribute && (o.entityname || undefined) === (entityName || undefined)
				);

				if (existingIndex >= 0) {
					// Update existing order's direction
					newOrders = currentOrders.map((o, i) =>
						i === existingIndex ? { ...o, descending: descending || undefined } : o
					);
				} else {
					// Add new order to the end
					newOrders = [...currentOrders, newOrder];
				}
			} else {
				// Single sort: replace all orders with just this one
				newOrders = [newOrder];
			}

			// Update the entity with new orders
			const newFetchQuery: FetchNode = {
				...state.fetchQuery,
				entity: {
					...state.fetchQuery.entity,
					orders: newOrders,
				},
			};

			return {
				...state,
				fetchQuery: newFetchQuery,
			};
		}

		default:
			return state;
	}
}

/** View info for load operations */
interface ViewLoadInfo {
	id: string;
	type: "system" | "personal";
	entitySetName: string;
	name: string;
}

interface BuilderContextValue extends BuilderState {
	setEntity: (entityName: string) => void;
	selectNode: (nodeId: NodeId) => void;
	addAttribute: (parentId: NodeId) => void;
	/** Add a specific attribute by name to an entity or link-entity */
	addAttributeByName: (parentId: NodeId, attributeName: string) => void;
	addAllAttributes: (parentId: NodeId) => void;
	addOrder: (parentId: NodeId) => void;
	addFilter: (parentId: NodeId) => void;
	addSubfilter: (filterId: NodeId) => void;
	addCondition: (filterId: NodeId) => void;
	addLinkEntity: (parentId: NodeId) => void;
	removeNode: (nodeId: NodeId) => void;
	updateNode: (nodeId: NodeId, updates: Record<string, unknown>) => void;
	newQuery: () => void;
	loadFetchXml: (xmlString: string) => ParseResult;
	/** Load a view's FetchXML while preserving view info for execution optimization */
	loadView: (xmlString: string, viewInfo: ViewLoadInfo, layoutXml?: string) => ParseResult;
	/** Clear the loaded view info (called when FetchXML is manually edited) */
	clearLoadedView: () => void;
	/** Set the column configuration (from parsed layoutxml or generated) */
	setColumnConfig: (config: LayoutXmlConfig) => void;
	/** Update a single column's width */
	updateColumnWidth: (columnName: string, width: number) => void;
	/** Reorder columns by moving one from fromIndex to toIndex */
	reorderColumns: (fromIndex: number, toIndex: number) => void;
	/** Sync layout with current FetchXML (regenerate/merge as needed) */
	syncLayoutWithFetchXml: (attributeTypeMap?: Map<string, string>) => void;
	/**
	 * Set sort order on an attribute. If isMultiSort is true, adds to existing sorts.
	 * If false, replaces all existing sorts.
	 */
	setSort: (
		attribute: string,
		descending: boolean,
		isMultiSort: boolean,
		entityName?: string
	) => void;
}

const BuilderContext = createContext<BuilderContextValue | null>(null);

export function BuilderProvider({ children }: { children: ReactNode }) {
	const [state, dispatch] = useReducer(builderReducer, initialState);

	const contextValue: BuilderContextValue = {
		...state,
		setEntity: (entityName: string) => dispatch({ type: "SET_ENTITY", entityName }),
		selectNode: (nodeId: NodeId) => dispatch({ type: "SELECT_NODE", nodeId }),
		addAttribute: (parentId: NodeId) => dispatch({ type: "ADD_ATTRIBUTE", parentId }),
		addAttributeByName: (parentId: NodeId, attributeName: string) =>
			dispatch({ type: "ADD_ATTRIBUTE_BY_NAME", parentId, attributeName }),
		addAllAttributes: (parentId: NodeId) => dispatch({ type: "ADD_ALL_ATTRIBUTES", parentId }),
		addOrder: (parentId: NodeId) => dispatch({ type: "ADD_ORDER", parentId }),
		addFilter: (parentId: NodeId) => dispatch({ type: "ADD_FILTER", parentId }),
		addSubfilter: (filterId: NodeId) => dispatch({ type: "ADD_SUBFILTER", filterId }),
		addCondition: (filterId: NodeId) => dispatch({ type: "ADD_CONDITION", filterId }),
		addLinkEntity: (parentId: NodeId) => dispatch({ type: "ADD_LINK_ENTITY", parentId }),
		removeNode: (nodeId: NodeId) => dispatch({ type: "REMOVE_NODE", nodeId }),
		updateNode: (nodeId: NodeId, updates: Record<string, unknown>) =>
			dispatch({ type: "UPDATE_NODE", nodeId, updates }),
		newQuery: () => dispatch({ type: "NEW_QUERY" }),
		loadFetchXml: (xmlString: string): ParseResult => {
			const result = parseFetchXml(xmlString);
			if (result.success && result.fetchNode) {
				dispatch({ type: "LOAD_FETCHXML", fetchNode: result.fetchNode });
			}
			return result;
		},
		loadView: (xmlString: string, viewInfo: ViewLoadInfo, layoutXml?: string): ParseResult => {
			const result = parseFetchXml(xmlString);
			if (result.success && result.fetchNode) {
				dispatch({
					type: "LOAD_VIEW",
					fetchNode: result.fetchNode,
					viewInfo: {
						id: viewInfo.id,
						type: viewInfo.type,
						originalFetchXml: xmlString,
						entitySetName: viewInfo.entitySetName,
						name: viewInfo.name,
					},
					layoutXml,
				});
			}
			return result;
		},
		clearLoadedView: () => dispatch({ type: "CLEAR_LOADED_VIEW" }),
		setColumnConfig: (config: LayoutXmlConfig) => dispatch({ type: "SET_COLUMN_CONFIG", config }),
		updateColumnWidth: (columnName: string, width: number) =>
			dispatch({ type: "UPDATE_COLUMN_WIDTH", columnName, width }),
		reorderColumns: (fromIndex: number, toIndex: number) =>
			dispatch({ type: "REORDER_COLUMNS", fromIndex, toIndex }),
		syncLayoutWithFetchXml: (attributeTypeMap?: Map<string, string>) =>
			dispatch({ type: "SYNC_LAYOUT_WITH_FETCHXML", attributeTypeMap }),
		setSort: (attribute: string, descending: boolean, isMultiSort: boolean, entityName?: string) =>
			dispatch({ type: "SET_SORT", attribute, descending, isMultiSort, entityName }),
	};

	return <BuilderContext.Provider value={contextValue}>{children}</BuilderContext.Provider>;
}

export function useBuilder() {
	const context = useContext(BuilderContext);
	if (!context) {
		throw new Error("useBuilder must be used within BuilderProvider");
	}
	return context;
}
