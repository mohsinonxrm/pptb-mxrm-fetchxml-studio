/**
 * Tree view for FetchXML query builder
 * Displays hierarchical structure: Entity -> Attributes, Orders, Filters, Link-Entities
 */

import { useState, useEffect, useCallback, useMemo } from "react";
import {
	Badge,
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
import { NON_RETURNING_LINK_TYPES, validateQueryRules } from "../../model/fetchxmlRules";
import { debugLog, debugWarn } from "../../../../shared/utils/debug";
import { SelectAttributesDialog } from "../Dialogs/SelectAttributesDialog";
import { metadataCache } from "../../state/cache";
import type { AttributeMetadata } from "../../api/pptbClient";
import { loadEntityAttributes } from "../../api/dataverseMetadata";

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
		minHeight: 0, // Important for flex child with overflow
		// Ensure scrollbars are always visible when content overflows
		scrollbarGutter: "stable",
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
	badges: {
		display: "flex",
		gap: tokens.spacingHorizontalXXS,
		alignItems: "center",
	},
});

type NodeBadge = { text: string; color?: "danger" | "brand" | "informative"; title?: string };

/** The badges whose condition holds */
const badgeList = (...items: (NodeBadge | false | undefined)[]): NodeBadge[] => items.filter((b): b is NodeBadge => !!b);

const FUNCTION_TEXT: Record<string, string> = {
	count: "Count",
	countcolumn: "Count values",
	sum: "Sum",
	avg: "Average",
	min: "Min",
	max: "Max",
};

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
	onUpdateAttributes: (parentId: NodeId, attributeNames: string[]) => void;
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
	onUpdateAttributes,
}: TreeViewProps) {
	const styles = useStyles();

	// Query-rule errors per node, the same ones the properties panel and Validate show (#90b)
	const errorsByNode = useMemo(() => {
		const map = new Map<NodeId, string[]>();
		for (const f of validateQueryRules(fetchQuery)) {
			if (f.severity !== "error") continue;
			map.set(f.nodeId, [...(map.get(f.nodeId) ?? []), f.message]);
		}
		return map;
	}, [fetchQuery]);

	const renderBadges = (nodeId: NodeId, badges: NodeBadge[] = []) => {
		const errors = errorsByNode.get(nodeId);
		const all = errors ? [...badges, { text: "Error", color: "danger" as const, title: errors.join(" ") }] : badges;
		if (all.length === 0) return undefined;
		return (
			<span className={styles.badges}>
				{all.map((b) => (
					<Badge key={b.text} size="small" appearance="tint" color={b.color ?? "informative"} title={b.title}>
						{b.text}
					</Badge>
				))}
			</span>
		);
	};

	// Dialog state
	const [selectAttrsDialogOpen, setSelectAttrsDialogOpen] = useState(false);
	const [selectAttrsEntityId, setSelectAttrsEntityId] = useState<NodeId | null>(null);
	const [selectAttrsDisplayName, setSelectAttrsDisplayName] = useState<string>("");
	const [selectAttrsAttributes, setSelectAttrsAttributes] = useState<AttributeMetadata[]>([]);
	const [selectAttrsCurrentNames, setSelectAttrsCurrentNames] = useState<string[]>([]);

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
		// Only a new selection expands and scrolls: findPathToNode reads the current tree, and re-running
		// on every tree edit would keep re-expanding and scrolling while the user works (#65)
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [selectedNodeId]);

	// Handle opening the Select Attributes dialog
	const handleOpenSelectAttributes = useCallback(
		async (entityNode: EntityNode | LinkEntityNode) => {
			const entityName = entityNode.name;
			const entityDisplayName =
				metadataCache.getEntityMetadata(entityName)?.DisplayName?.UserLocalizedLabel?.Label ||
				entityName;

			// Get currently selected attribute names
			const currentAttributeNames = entityNode.attributes.map((attr) => attr.name);

			// The attributes to offer, per the "Advanced Find only" setting (#57; cached by the loader)
			let attributes: AttributeMetadata[] = [];
			try {
				attributes = await loadEntityAttributes(entityName);
			} catch (error) {
				console.error("Failed to load attributes:", error);
			}

			setSelectAttrsEntityId(entityNode.id);
			setSelectAttrsDisplayName(entityDisplayName);
			setSelectAttrsAttributes(attributes);
			setSelectAttrsCurrentNames(currentAttributeNames);
			setSelectAttrsDialogOpen(true);
		},
		[]
	);

	// Handle applying selected attributes from dialog
	const handleApplySelectedAttributes = useCallback(
		(nodeId: NodeId, selectedAttributes: string[]) => {
			onUpdateAttributes(nodeId, selectedAttributes);
		},
		[onUpdateAttributes]
	);

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
				<TreeItemLayout
					iconBefore={<Table20Regular />}
					aside={renderBadges(
						fetch.id,
						badgeList(
							fetch.options.aggregate === true && { text: "Aggregate", color: "brand" },
							fetch.options.distinct && { text: "Distinct" },
							fetch.options.dataSource === "retained" && { text: "Retained" },
							(fetch.options.queryHints?.length ?? 0) > 0 && { text: "Hints" },
						),
					)}
					onClick={() => onNodeSelect(fetch.id)}
				>
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
										<MenuItem onClick={() => handleOpenSelectAttributes(entity)}>
											Select Attributes
										</MenuItem>
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
					aside={renderBadges(entity.id)}
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
					aside={renderBadges(node.id)}
					onClick={() => onNodeSelect(node.id)}
				>
					All Attributes
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderAttributeNode = (attr: AttributeNode) => {
		const label = attr.alias ? `${attr.name} (as ${attr.alias})` : attr.name;
		const badges: NodeBadge[] = [];
		if (attr.groupby) badges.push({ text: attr.dategrouping ? `Group by · ${attr.dategrouping}` : "Group by", color: "brand" });
		if (attr.aggregate) badges.push({ text: FUNCTION_TEXT[attr.aggregate] ?? attr.aggregate, color: "brand" });
		if (attr.aggregate === "countcolumn" && attr.distinct) badges.push({ text: "Unique" });
		if (attr.rowAggregate) badges.push({ text: "Child count" });

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
					aside={renderBadges(attr.id, badges)}
					onClick={() => onNodeSelect(attr.id)}
				>
					{label}
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderOrderNode = (order: OrderNode) => {
		const direction = order.descending ? "↓" : "↑";
		// Aggregate queries sort by alias (#90); show entityname prefix for link-entity orders
		const target = order.alias ? `${order.alias} (alias)` : (order.attribute ?? "");
		const displayName = order.entityname ? `${order.entityname}.${target}` : target;
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
					aside={renderBadges(order.id)}
					onClick={() => onNodeSelect(order.id)}
				>
					{displayName} {direction}
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
					aside={renderBadges(
						filter.id,
						badgeList(
							filter.hint === "union" && { text: "Union" },
							filter.isQuickFindFields && { text: "Quick find" },
						),
					)}
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

		if (requiresValue && condition.valueof) {
			// Compared to another attribute
			displayText += ` → ${condition.valueof}`;
		} else if (
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
					aside={renderBadges(condition.id)}
					onClick={() => onNodeSelect(condition.id)}
				>
					{displayText}
				</TreeItemLayout>
			</TreeItem>
		);
	};

	const renderLinkEntityNode = (link: LinkEntityNode) => {
		// exists / in / any / not any / all / not all return nothing: attributes and orders fail (0x80041121)
		const canHaveAttributes = !NON_RETURNING_LINK_TYPES.has(link.linkType);
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
										<MenuItem
											onClick={() => handleOpenSelectAttributes(link)}
											disabled={!canHaveAttributes}
										>
											Select Attributes
										</MenuItem>
										<MenuItem onClick={() => onAddOrder(link.id)} disabled={!canHaveAttributes}>
											Order By
										</MenuItem>
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
					aside={renderBadges(link.id)}
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

			<SelectAttributesDialog
				open={selectAttrsDialogOpen}
				onOpenChange={setSelectAttrsDialogOpen}
				entityDisplayName={selectAttrsDisplayName}
				nodeId={selectAttrsEntityId || ""}
				attributes={selectAttrsAttributes}
				selectedAttributeNames={selectAttrsCurrentNames}
				onApply={handleApplySelectedAttributes}
			/>
		</div>
	);
}
