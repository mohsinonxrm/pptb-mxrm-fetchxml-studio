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

interface BuilderState {
	fetchQuery: FetchNode | null;
	selectedNodeId: NodeId | null;
	selectedNode: SelectedNode;
}

type BuilderAction =
	| { type: "SET_ENTITY"; entityName: string }
	| { type: "SELECT_NODE"; nodeId: NodeId }
	| { type: "ADD_ATTRIBUTE"; parentId: NodeId }
	| { type: "ADD_ALL_ATTRIBUTES"; parentId: NodeId }
	| { type: "ADD_ORDER"; parentId: NodeId }
	| { type: "ADD_FILTER"; parentId: NodeId }
	| { type: "ADD_SUBFILTER"; filterId: NodeId }
	| { type: "ADD_CONDITION"; filterId: NodeId }
	| { type: "ADD_LINK_ENTITY"; parentId: NodeId }
	| { type: "REMOVE_NODE"; nodeId: NodeId }
	| { type: "UPDATE_NODE"; nodeId: NodeId; updates: Record<string, unknown> }
	| { type: "NEW_QUERY" }
	| { type: "LOAD_FETCHXML"; fetchNode: FetchNode };

const initialState: BuilderState = {
	fetchQuery: null,
	selectedNodeId: null,
	selectedNode: null,
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
			// Load a pre-parsed FetchNode tree (from parser or view)
			const fetchNode = action.fetchNode;
			return {
				fetchQuery: fetchNode,
				selectedNodeId: fetchNode.entity.id,
				selectedNode: fetchNode.entity,
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
			};
		}

		case "ADD_LINK_ENTITY": {
			if (!state.fetchQuery) return state;

			// Find parent node (entity or link-entity)
			const parent = findNodeById(state.fetchQuery, action.parentId);
			if (!parent || (parent.type !== "entity" && parent.type !== "link-entity")) {
				return state;
			}

			// Create new link-entity node
			const newLinkEntity: LinkEntityNode = {
				id: generateId(),
				type: "link-entity",
				name: "new_entity",
				from: "id_field",
				to: "parent_id_field",
				linkType: "inner",
				attributes: [],
				orders: [],
				filters: [],
				links: [],
			};

			// Update the tree immutably
			const updatedFetch = updateNodeInTree(state.fetchQuery, action.parentId, (node) => {
				const n = node as EntityNode | LinkEntityNode;
				return {
					...n,
					links: [...n.links, newLinkEntity],
				};
			});

			return {
				...state,
				fetchQuery: updatedFetch,
				selectedNodeId: newLinkEntity.id,
				selectedNode: newLinkEntity,
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
			};
		}

		default:
			return state;
	}
}

interface BuilderContextValue extends BuilderState {
	setEntity: (entityName: string) => void;
	selectNode: (nodeId: NodeId) => void;
	addAttribute: (parentId: NodeId) => void;
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
}

const BuilderContext = createContext<BuilderContextValue | null>(null);

export function BuilderProvider({ children }: { children: ReactNode }) {
	const [state, dispatch] = useReducer(builderReducer, initialState);

	const contextValue: BuilderContextValue = {
		...state,
		setEntity: (entityName: string) => dispatch({ type: "SET_ENTITY", entityName }),
		selectNode: (nodeId: NodeId) => dispatch({ type: "SELECT_NODE", nodeId }),
		addAttribute: (parentId: NodeId) => dispatch({ type: "ADD_ATTRIBUTE", parentId }),
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
