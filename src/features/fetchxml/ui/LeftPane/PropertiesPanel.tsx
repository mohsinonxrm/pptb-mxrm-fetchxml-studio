/**
 * Properties panel: the editor for the selected tree node. Query-rule findings (fetchxmlRules.ts)
 * are computed once per tree change and handed to each editor, so the panel shows the same problems
 * Validate and Execute report (#90b).
 */

import { useMemo } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
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
import { validateQueryRules } from "../../model/fetchxmlRules";
import {
	collectAllFilters,
	findNodeContainer,
	findParentEntityForNode,
} from "../../model/treeUtils";
import { useBuilder } from "../../state/builderContext";
import { FetchOptionsEditor } from "./PropertiesPanel/editors/FetchOptionsEditor";
import { EntityEditor } from "./PropertiesPanel/editors/EntityEditor";
import { AllAttributesEditor } from "./PropertiesPanel/editors/AllAttributesEditor";
import { AttributeEditor } from "./PropertiesPanel/editors/AttributeEditor";
import { OrderEditor } from "./PropertiesPanel/editors/OrderEditor";
import { FilterEditor } from "./PropertiesPanel/editors/FilterEditor";
import { ConditionEditor } from "./PropertiesPanel/editors/ConditionEditor";
import { LinkEntityEditor } from "./PropertiesPanel/editors/LinkEntityEditor";

const useStyles = makeStyles({
	container: {
		overflow: "auto",
		height: "100%",
		minHeight: 0, // Important for flex child with overflow
		scrollbarGutter: "stable",
	},
	emptyState: {
		padding: tokens.spacingHorizontalL,
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase300,
		fontStyle: "italic",
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
	const { removeNode } = useBuilder();
	const findings = useMemo(() => (fetchQuery ? validateQueryRules(fetchQuery) : []), [fetchQuery]);

	if (!selectedNode || !fetchQuery) {
		return (
			<div className={styles.container}>
				<div className={styles.emptyState}>Select a node in the tree to view its properties.</div>
			</div>
		);
	}

	const nodeFindings = findings.filter((f) => f.nodeId === selectedNode.id);
	const isAggregateQuery = fetchQuery.options.aggregate === true;
	const update = (updates: Record<string, unknown>) => onNodeUpdate(selectedNode.id, updates);
	const location = findNodeContainer(fetchQuery, selectedNode.id);
	const missing = <div className={styles.emptyState}>This node is no longer in the query.</div>;

	const renderEditor = () => {
		switch (selectedNode.type) {
			case "fetch":
				return <FetchOptionsEditor node={selectedNode} findings={nodeFindings} onUpdate={update} />;

			case "entity":
				return <EntityEditor node={selectedNode} onUpdate={update} />;

			case "all-attributes":
				return (
					<AllAttributesEditor isAggregateQuery={isAggregateQuery} onRemove={() => removeNode(selectedNode.id)} />
				);

			case "attribute":
				if (!location) return missing;
				return (
					<AttributeEditor
						node={selectedNode}
						container={location.container}
						isRootEntity={location.isRoot}
						isAggregateQuery={isAggregateQuery}
						fetchQuery={fetchQuery}
						findings={nodeFindings}
						onUpdate={update}
					/>
				);

			case "order":
				if (!location) return missing;
				return (
					<OrderEditor
						node={selectedNode}
						container={location.container}
						isRootEntity={location.isRoot}
						fetchQuery={fetchQuery}
						findings={nodeFindings}
						fetchFindings={findings.filter((f) => f.nodeId === fetchQuery.id)}
						onUpdate={update}
					/>
				);

			case "filter":
				return (
					<FilterEditor
						node={selectedNode}
						unionElsewhere={collectAllFilters(fetchQuery).some((f) => f.id !== selectedNode.id && f.hint === "union")}
						findings={nodeFindings}
						onUpdate={update}
					/>
				);

			case "condition":
				return (
					<ConditionEditor
						node={selectedNode}
						entityName={location?.container.name ?? findParentEntityForNode(fetchQuery, selectedNode.id) ?? ""}
						fetchQuery={fetchQuery}
						findings={nodeFindings}
						onUpdate={update}
					/>
				);

			case "link-entity":
				return (
					<LinkEntityEditor
						node={selectedNode}
						parentEntityName={findParentEntityForNode(fetchQuery, selectedNode.id) ?? fetchQuery.entity.name}
						inFilter={collectAllFilters(fetchQuery).some((f) => f.links?.some((l) => l.id === selectedNode.id))}
						findings={nodeFindings}
						onUpdate={update}
					/>
				);

			default:
				return <div className={styles.emptyState}>Unknown node type.</div>;
		}
	};

	// Keyed by node id, so per-node UI state resets when the selection changes
	return (
		<div className={styles.container} key={selectedNode.id}>
			{renderEditor()}
		</div>
	);
}
