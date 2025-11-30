/**
 * Tree view for FetchXML query builder
 * Displays hierarchical structure: Entity -> Attributes, Orders, Filters, Link-Entities
 */

import { useState, useEffect } from "react";
import {
	Tree,
	TreeItem,
	TreeItemLayout,
	Button,
	makeStyles,
	Menu,
	MenuTrigger,
	MenuPopover,
	MenuList,
	MenuItem,
	tokens,
	type TreeOpenChangeData,
} from "@fluentui/react-components";
import {
	Table20Regular,
	Column20Regular,
	Filter20Regular,
	ArrowSort20Regular,
	Link20Regular,
	Add20Regular,
	Delete20Regular,
} from "@fluentui/react-icons";
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
import { operatorRequiresValue } from "../../model/operators";
import { debugLog, debugWarn } from "../../../../shared/utils/debug";

/**
 * Format a value for display, avoiding scientific notation for numbers
 */
function formatDisplayValue(val: unknown): string {
	if (typeof val === "number") {
		if (Math.abs(val) < 1e-10 && val !== 0) {
			// Very small numbers - use high precision
			return val.toFixed(20).replace(/\.?0+$/, "");
		} else if (Number.isInteger(val)) {
			// Integer - no decimal point needed
			return val.toString();
		} else {
			// Decimal - preserve precision but remove trailing zeros
			return val.toFixed(10).replace(/\.?0+$/, "");
		}
	}
	if (Array.isArray(val)) {
		// Format array values (for between, in, etc.)
		return val.map((v) => formatDisplayValue(v)).join(", ");
	}
	return String(val);
}

const useStyles = makeStyles({
	container: {
		flex: 1,
		overflow: "auto",
		padding: "8px",
	},
	treeItem: {
		"&:hover": {
			backgroundColor: tokens.colorNeutralBackground1Hover,
		},
	},
	actions: {
		display: "flex",
		gap: "4px",
		opacity: 0,
		transition: "opacity 0.2s",
		"&:focus-within": {
			opacity: 1,
		},
	},
	treeItemWithActions: {
		"&:hover .actions": {
			opacity: 1,
		},
	},
});

interface TreeViewProps {
	fetchQuery: FetchNode;
	selectedNodeId: NodeId | null;
	onNodeSelect: (nodeId: NodeId) => void;
	onAddAttribute: (parentId: NodeId) => void;
	onAddAllAttributes: (parentId: NodeId) => void;
	onAddOrder: (parentId: NodeId) => void;
	onAddFilter: (parentId: NodeId) => void;
	onAddSubfilter: (filterId: NodeId) => void;
	onAddCondition: (filterId: NodeId) => void;
	onAddLinkEntity: (parentId: NodeId) => void;
	onRemoveNode: (nodeId: NodeId) => void;
}

export function TreeView({
	fetchQuery,
	selectedNodeId,
	onNodeSelect,
	onAddAttribute,
	onAddAllAttributes,
	onAddOrder,
	onAddFilter,
	onAddSubfilter,
	onAddCondition,
	onAddLinkEntity,
	onRemoveNode,
}: TreeViewProps) {
	const styles = useStyles();

	// Manage which tree items are open/expanded
	const [openItems, setOpenItems] = useState<Set<NodeId>>(
		new Set([fetchQuery.id, fetchQuery.entity.id])
	);

	// Helper to find path from root to a node (all ancestor IDs)
	const findPathToNode = (targetId: NodeId): NodeId[] => {
		const path: NodeId[] = [];

		const findInNode = (node: unknown, target: NodeId): boolean => {
			if (!node || typeof node !== "object") return false;

			const n = node as Record<string, unknown>;

			if (n.id === target) {
				debugLog("treeExpansion", "Found target node:", n.id);
				// Don't add the target to path, only ancestors
				return true;
			}

			// Add current node to path before searching children
			const currentId = n.id as NodeId;
			debugLog("treeExpansion", "Checking node:", currentId, "type:", n.type);

			// Search in all possible child arrays
			const arrayProps = ["attributes", "orders", "filters", "conditions", "subfilters", "links"];
			for (const prop of arrayProps) {
				if (Array.isArray(n[prop])) {
					for (const item of n[prop]) {
						if (findInNode(item, target)) {
							// Found in this branch, add current to path
							path.unshift(currentId);
							return true;
						}
					}
				}
			}

			// Search in nested entity
			if (n.entity && findInNode(n.entity, target)) {
				// Found in entity, add current to path
				path.unshift(currentId);
				return true;
			}

			return false;
		};

		findInNode(fetchQuery, targetId);
		debugLog("treeExpansion", "Final path:", path);
		return path;
	};

	// Auto-expand ancestors when a new node is selected
	useEffect(() => {
		if (selectedNodeId) {
			debugLog("treeExpansion", "Node selected, expanding path:", selectedNodeId);
			const pathToNode = findPathToNode(selectedNodeId);
			debugLog("treeExpansion", "Path to node:", pathToNode);

			if (pathToNode.length === 0) {
				debugWarn("treeExpansion", "Path is empty! Node might not be in tree yet.");
				return;
			}

			setOpenItems((prev) => {
				const newSet = new Set(prev);
				const sizeBefore = newSet.size;
				// Add all ancestors to open items (except the leaf node itself)
				pathToNode.forEach((id) => newSet.add(id));
				const itemsAdded = Array.from(newSet).filter((id) => !prev.has(id));
				debugLog("treeExpansion", "Updated openItems:", {
					sizeBefore,
					sizeAfter: newSet.size,
					added: itemsAdded,
					addedCount: itemsAdded.length,
					allOpen: Array.from(newSet),
				});

				// Always return the new set to ensure re-render even if no new items
				return newSet;
			});

			// Scroll the selected node into view after a short delay to allow DOM updates
			// Note: Fluent UI Tree doesn't set aria-selected="true" reliably, so we use
			// the last tree item as the target (newly added nodes appear at the end)
			setTimeout(() => {
				const allTreeItems = document.querySelectorAll('[role="treeitem"]');

				if (allTreeItems.length > 0) {
					const targetElement = allTreeItems[allTreeItems.length - 1];
					targetElement.scrollIntoView({ behavior: "smooth", block: "nearest" });
					debugLog(
						"treeExpansion",
						`Scrolled to item ${allTreeItems.length} of ${allTreeItems.length}`
					);
				} else {
					debugWarn("treeExpansion", "No tree items found in DOM for scrolling");
				}
			}, 250);
		}
	}, [selectedNodeId]); // eslint-disable-line react-hooks/exhaustive-deps

	const handleOpenChange = (_event: unknown, data: TreeOpenChangeData) => {
		debugLog("treeExpansion", "handleOpenChange called:", {
			open: data.open,
			value: data.value,
			type: data.type,
			openItemsCount: data.openItems.size,
			openItems: Array.from(data.openItems),
		});
		// data.openItems is already the updated Set from Fluent UI
		// Use it directly to respect user's expand/collapse actions
		setOpenItems(data.openItems as Set<NodeId>);
	};

	const renderFetchNode = (fetch: FetchNode) => {
		return (
			<TreeItem
				key={fetch.id}
				itemType="branch"
				value={fetch.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout iconBefore={<Table20Regular />} onClick={() => onNodeSelect(fetch.id)}>
					Fetch Query
				</TreeItemLayout>

				<Tree>{renderEntityNode(fetch.entity, true)}</Tree>
			</TreeItem>
		);
	};

	const renderEntityNode = (entity: EntityNode, isRootEntity: boolean = true) => {
		// Only allow one top-level filter per entity
		const hasFilter = entity.filters.length > 0;

		return (
			<TreeItem
				key={entity.id}
				itemType="branch"
				value={entity.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout
					iconBefore={<Table20Regular />}
					actions={
						<div className="actions">
							<Menu>
								<MenuTrigger disableButtonEnhancement>
									<Button appearance="subtle" size="small" icon={<Add20Regular />} title="Add..." />
								</MenuTrigger>
								<MenuPopover>
									<MenuList>
										<MenuItem onClick={() => onAddAllAttributes(entity.id)}>
											All Attributes
										</MenuItem>
										<MenuItem onClick={() => onAddAttribute(entity.id)}>Attribute</MenuItem>
										<MenuItem onClick={() => onAddOrder(entity.id)}>Order By</MenuItem>
										<MenuItem onClick={() => onAddFilter(entity.id)} disabled={hasFilter}>
											Filter
										</MenuItem>
										<MenuItem onClick={() => onAddLinkEntity(entity.id)}>Link-Entity</MenuItem>
									</MenuList>
								</MenuPopover>
							</Menu>
							{!isRootEntity && (
								<Button
									appearance="subtle"
									size="small"
									icon={<Delete20Regular />}
									onClick={(e) => {
										e.stopPropagation();
										onRemoveNode(entity.id);
									}}
									title="Remove"
								/>
							)}
						</div>
					}
					onClick={() => onNodeSelect(entity.id)}
				>
					{entity.name}
				</TreeItemLayout>

				<Tree>
					{/* All Attributes */}
					{entity.allAttributes?.enabled && renderAllAttributesNode(entity.allAttributes)}

					{/* Attributes */}
					{entity.attributes.map((attr) => renderAttributeNode(attr))}

					{/* Orders */}
					{entity.orders.map((order) => renderOrderNode(order))}

					{/* Filters */}
					{entity.filters.map((filter) => renderFilterNode(filter))}

					{/* Link-Entities */}
					{entity.links.map((link) => renderLinkEntityNode(link))}
				</Tree>
			</TreeItem>
		);
	};

	const renderAllAttributesNode = (node: AllAttributesNode) => {
		return (
			<TreeItem key={node.id} itemType="leaf" value={node.id} className={styles.treeItem}>
				<TreeItemLayout
					iconBefore={<Column20Regular />}
					actions={
						<Button
							appearance="subtle"
							size="small"
							icon={<Delete20Regular />}
							onClick={(e) => {
								e.stopPropagation();
								onRemoveNode(node.id);
							}}
							title="Remove"
						/>
					}
					onClick={() => onNodeSelect(node.id)}
				>
					All Attributes
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderAttributeNode = (attr: AttributeNode) => {
		const label = attr.alias ? `${attr.name} (as ${attr.alias})` : attr.name;
		const suffix = attr.aggregate ? ` [${attr.aggregate}]` : "";

		return (
			<TreeItem
				key={attr.id}
				itemType="leaf"
				value={attr.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout
					iconBefore={<Column20Regular />}
					actions={
						<div className="actions">
							<Button
								appearance="subtle"
								size="small"
								icon={<Delete20Regular />}
								onClick={(e) => {
									e.stopPropagation();
									onRemoveNode(attr.id);
								}}
								title="Remove Attribute"
							/>
						</div>
					}
					onClick={() => onNodeSelect(attr.id)}
				>
					{label}
					{suffix}
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderOrderNode = (order: OrderNode) => {
		const direction = order.descending ? "↓" : "↑";
		return (
			<TreeItem
				key={order.id}
				itemType="leaf"
				value={order.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout
					iconBefore={<ArrowSort20Regular />}
					actions={
						<div className="actions">
							<Button
								appearance="subtle"
								size="small"
								icon={<Delete20Regular />}
								onClick={(e) => {
									e.stopPropagation();
									onRemoveNode(order.id);
								}}
								title="Remove Order"
							/>
						</div>
					}
					onClick={() => onNodeSelect(order.id)}
				>
					{order.attribute} {direction}
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderFilterNode = (filter: FilterNode) => {
		return (
			<TreeItem
				key={filter.id}
				itemType="branch"
				value={filter.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout
					iconBefore={<Filter20Regular />}
					actions={
						<div className="actions">
							<Menu>
								<MenuTrigger disableButtonEnhancement>
									<Button appearance="subtle" size="small" icon={<Add20Regular />} title="Add..." />
								</MenuTrigger>
								<MenuPopover>
									<MenuList>
										<MenuItem onClick={() => onAddCondition(filter.id)}>Condition</MenuItem>
										<MenuItem onClick={() => onAddSubfilter(filter.id)}>Sub-Filter</MenuItem>
										<MenuItem onClick={() => onAddLinkEntity(filter.id)}>
											Link-Entity (any/not any/all)
										</MenuItem>
									</MenuList>
								</MenuPopover>
							</Menu>
							<Button
								appearance="subtle"
								size="small"
								icon={<Delete20Regular />}
								onClick={(e) => {
									e.stopPropagation();
									onRemoveNode(filter.id);
								}}
								title="Remove Filter"
							/>
						</div>
					}
					onClick={() => onNodeSelect(filter.id)}
				>
					Filter ({filter.conjunction})
				</TreeItemLayout>

				<Tree>
					{filter.conditions.map((cond) => renderConditionNode(cond))}
					{filter.subfilters.map((subfilter) => renderFilterNode(subfilter))}
					{filter.links?.map((link) => renderLinkEntityNode(link))}
				</Tree>
			</TreeItem>
		);
	};

	const renderConditionNode = (condition: ConditionNode) => {
		// Format the display text based on operator and value
		const requiresValue = operatorRequiresValue(condition.operator);
		let displayText = `${condition.attribute} ${condition.operator}`;

		if (
			requiresValue &&
			condition.value !== undefined &&
			condition.value !== null &&
			condition.value !== ""
		) {
			displayText += ` = ${formatDisplayValue(condition.value)}`;
		}

		return (
			<TreeItem
				key={condition.id}
				itemType="leaf"
				value={condition.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout
					actions={
						<div className="actions">
							<Button
								appearance="subtle"
								size="small"
								icon={<Delete20Regular />}
								onClick={(e) => {
									e.stopPropagation();
									onRemoveNode(condition.id);
								}}
								title="Remove Condition"
							/>
						</div>
					}
					onClick={() => onNodeSelect(condition.id)}
				>
					{displayText}
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderLinkEntityNode = (link: LinkEntityNode) => {
		// Attributes are not allowed in "not any" and "not all" link types
		const canHaveAttributes = link.linkType !== "not any" && link.linkType !== "not all";
		// Only allow one top-level filter per link-entity
		const hasFilter = link.filters.length > 0;

		return (
			<TreeItem
				key={link.id}
				itemType="branch"
				value={link.id}
				className={styles.treeItemWithActions}
			>
				<TreeItemLayout
					iconBefore={<Link20Regular />}
					actions={
						<div className="actions">
							<Menu>
								<MenuTrigger disableButtonEnhancement>
									<Button appearance="subtle" size="small" icon={<Add20Regular />} title="Add..." />
								</MenuTrigger>
								<MenuPopover>
									<MenuList>
										<MenuItem
											onClick={() => onAddAllAttributes(link.id)}
											disabled={!canHaveAttributes}
										>
											All Attributes
										</MenuItem>
										<MenuItem onClick={() => onAddAttribute(link.id)} disabled={!canHaveAttributes}>
											Attribute
										</MenuItem>
										<MenuItem onClick={() => onAddOrder(link.id)}>Order By</MenuItem>
										<MenuItem onClick={() => onAddFilter(link.id)} disabled={hasFilter}>
											Filter
										</MenuItem>
										<MenuItem onClick={() => onAddLinkEntity(link.id)}>Link-Entity</MenuItem>
									</MenuList>
								</MenuPopover>
							</Menu>
							<Button
								appearance="subtle"
								size="small"
								icon={<Delete20Regular />}
								onClick={(e) => {
									e.stopPropagation();
									onRemoveNode(link.id);
								}}
								title="Remove Link"
							/>
						</div>
					}
					onClick={() => onNodeSelect(link.id)}
				>
					{link.alias || link.name} ({link.linkType})
				</TreeItemLayout>

				<Tree>
					{link.allAttributes?.enabled && renderAllAttributesNode(link.allAttributes)}
					{link.attributes.map((attr) => renderAttributeNode(attr))}
					{link.orders.map((order) => renderOrderNode(order))}
					{link.filters.map((filter) => renderFilterNode(filter))}
					{link.links.map((nestedLink) => renderLinkEntityNode(nestedLink))}
				</Tree>
			</TreeItem>
		);
	};

	return (
		<div className={styles.container}>
			<Tree
				aria-label="FetchXML Query Structure"
				openItems={openItems}
				onOpenChange={handleOpenChange}
			>
				{renderFetchNode(fetchQuery)}
			</Tree>
		</div>
	);
}
