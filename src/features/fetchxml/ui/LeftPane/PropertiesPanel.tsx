/**
 * Properties panel that displays editors for the selected tree node
 */

import { makeStyles, tokens } from "@fluentui/react-components";
import { debugLog } from "../../../../shared/utils/debug";
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
} from "../../model/nodes";
import { FetchOptionsEditor } from "./PropertiesPanel/editors/FetchOptionsEditor";
import { AttributeEditor } from "./PropertiesPanel/editors/AttributeEditor";
import { OrderEditor } from "./PropertiesPanel/editors/OrderEditor";
import { FilterEditor } from "./PropertiesPanel/editors/FilterEditor";
import { ConditionEditor } from "./PropertiesPanel/editors/ConditionEditor";
import { LinkEntityEditor } from "./PropertiesPanel/editors/LinkEntityEditor";

const useStyles = makeStyles({
	container: {
		overflow: "auto",
		height: "100%",
	},
	emptyState: {
		padding: "16px",
		color: tokens.colorNeutralForeground3,
		fontSize: "13px",
		fontStyle: "italic",
	},
	section: {
		padding: "16px",
	},
	label: {
		display: "block",
		fontSize: "13px",
		fontWeight: 600,
		marginBottom: "8px",
		color: tokens.colorNeutralForeground2,
	},
});

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

interface PropertiesPanelProps {
	selectedNode: SelectedNode;
	fetchQuery: FetchNode | null;
	onNodeUpdate: (nodeId: NodeId, updates: Record<string, unknown>) => void;
}

export function PropertiesPanel({ selectedNode, fetchQuery, onNodeUpdate }: PropertiesPanelProps) {
	const styles = useStyles();

	// Helper to find the parent entity/link-entity for a node
	const findParentEntity = (nodeId: NodeId): string | null => {
		if (!fetchQuery) return null;

		const findInNode = (
			node: unknown,
			targetId: NodeId,
			parentEntityName: string | null
		): string | null => {
			if (!node || typeof node !== "object") return null;

			const n = node as Record<string, unknown>;

			// If this is the target node, return the current parent entity name
			if (n.id === targetId) {
				debugLog("propertiesPanel", "Found target node:", {
					targetId,
					nodeType: n.type,
					nodeName: n.name,
					resolvedParentEntityName: parentEntityName,
				});
				return parentEntityName;
			}

			// If this is an entity or link-entity, update parent name
			// BUT only if it has a valid entity name (not a placeholder like "new_entity")
			let currentEntityName = parentEntityName;
			if (n.type === "entity" || n.type === "link-entity") {
				const entityName = n.name as string;
				// Only use the entity name if it's not a placeholder
				if (entityName && entityName !== "new_entity") {
					currentEntityName = entityName;
					debugLog("propertiesPanel", "Traversing entity/link-entity:", {
						nodeId: n.id,
						nodeType: n.type,
						entityName,
						updatedCurrentEntityName: currentEntityName,
					});
				}
			}

			// Search in arrays
			const arrayProps = ["attributes", "orders", "filters", "conditions", "subfilters", "links"];
			for (const prop of arrayProps) {
				if (Array.isArray(n[prop])) {
					for (const item of n[prop]) {
						const found = findInNode(item, targetId, currentEntityName);
						if (found) return found;
					}
				}
			}

			// Search in nested objects
			if (n.entity) {
				const found = findInNode(n.entity, targetId, currentEntityName);
				if (found) return found;
			}

			return null;
		};

		return findInNode(fetchQuery, nodeId, fetchQuery.entity.name);
	};

	if (!selectedNode) {
		return (
			<div className={styles.container}>
				<div className={styles.emptyState}>Select a node in the tree to view its properties</div>
			</div>
		);
	}

	const renderProperties = () => {
		switch (selectedNode.type) {
			case "fetch":
				return (
					<FetchOptionsEditor
						node={selectedNode}
						onUpdate={(updates) => onNodeUpdate(selectedNode.id, updates)}
					/>
				);

			case "entity":
				return (
					<div className={styles.section}>
						<div className={styles.label}>Entity: {selectedNode.name}</div>
						<div className={styles.emptyState}>
							Entity properties editor (Phase 5)
							<br />- enableprefiltering, prefilterparametername
						</div>
					</div>
				);

			case "all-attributes":
				return (
					<div className={styles.section}>
						<div className={styles.label}>All Attributes</div>
						<div className={styles.emptyState}>All-attributes toggle (Phase 5)</div>
					</div>
				);

			case "attribute":
				return (
					<AttributeEditor
						node={selectedNode}
						entityName={findParentEntity(selectedNode.id) ?? ""}
						onUpdate={(updates) => onNodeUpdate(selectedNode.id, updates)}
						isAggregateQuery={fetchQuery?.options.aggregate ?? false}
					/>
				);

			case "order":
				return (
					<OrderEditor
						node={selectedNode}
						entityName={findParentEntity(selectedNode.id) ?? ""}
						onUpdate={(updates) => onNodeUpdate(selectedNode.id, updates)}
					/>
				);

			case "filter":
				return (
					<FilterEditor
						node={selectedNode}
						onUpdate={(updates) => onNodeUpdate(selectedNode.id, updates)}
					/>
				);

			case "condition":
				return (
					<ConditionEditor
						node={selectedNode}
						entityName={findParentEntity(selectedNode.id) ?? ""}
						onUpdate={(updates) => onNodeUpdate(selectedNode.id, updates)}
						isAggregateQuery={fetchQuery?.options.aggregate ?? false}
					/>
				);

			case "link-entity":
				return (
					<LinkEntityEditor
						node={selectedNode}
						parentEntityName={findParentEntity(selectedNode.id) ?? ""}
						onUpdate={(updates) => onNodeUpdate(selectedNode.id, updates)}
					/>
				);

			default:
				return <div className={styles.emptyState}>Unknown node type</div>;
		}
	};

	return <div className={styles.container}>{renderProperties()}</div>;
}
