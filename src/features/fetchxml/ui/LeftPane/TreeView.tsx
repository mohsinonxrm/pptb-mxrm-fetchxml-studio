/**
 * Tree view for FetchXML query builder
 * Displays hierarchical structure: Entity -> Attributes, Orders, Filters, Link-Entities
 */

import { useState, useEffect, useCallback, useMemo, useRef, type ReactElement, type ReactNode } from "react";
import {
	Badge,
	Tree,
	TreeItem,
	TreeItemLayout,
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	makeStyles,
	Menu,
	MenuDivider,
	MenuTrigger,
	MenuPopover,
	MenuList,
	MenuItem,
	mergeClasses,
	Tooltip,
	tokens,
	type PositioningVirtualElement,
	type TreeOpenChangeData,
} from "@fluentui/react-components";
import {
	Table20Regular,
	TableSimple20Regular,
	CheckboxChecked20Regular,
	TextAdd20Regular,
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
import { countDescendants, deleteConfirmationText, linkEntityLabel, showsNoAttributesHint } from "../../model/treeLabels";
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
		overflowX: "hidden",
		overflowY: "auto",
		padding: "8px",
		minHeight: 0, // Important for flex child with overflow
		minWidth: 0,
		// Ensure scrollbars are always visible when content overflows
		scrollbarGutter: "stable",
	},
	tree: {
		width: "100%",
		maxWidth: "100%",
		minWidth: 0,
		overflowX: "hidden",
		"&:hover": {
			backgroundColor: tokens.colorTransparentBackground,
		},
	},
	treeItemLayout: {
		width: "100%",
		maxWidth: "100%",
		minWidth: 0,
		boxSizing: "border-box",
		whiteSpace: "nowrap",
		"& .fui-TreeItemLayout__main": {
			display: "block",
			flex: "1 1 0%",
			minWidth: 0,
			maxWidth: "100%",
			overflow: "hidden",
			textOverflow: "ellipsis",
			whiteSpace: "nowrap",
		},
		"& .fui-TreeItemLayout__aside": {
			flexShrink: 0,
		},
		"& .fui-TreeItemLayout__actions": {
			flexShrink: 0,
			marginInlineStart: "auto",
		},
	},
	label: {
		display: "block",
		flex: "1 1 0%",
		minWidth: 0,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	selectedTreeItem: {
		backgroundColor: tokens.colorNeutralBackground1Selected,
		color: tokens.colorNeutralForeground1Selected,
		"&:hover": {
			backgroundColor: tokens.colorNeutralBackground1Selected,
		},
	},
	actions: {
		display: "flex",
		flexShrink: 0,
		gap: "4px",
		opacity: 1,
	},
	treeItemWithActions: {
		width: "100%",
		maxWidth: "100%",
		minWidth: 0,
		"&:hover": {
			backgroundColor: tokens.colorTransparentBackground,
		},
	},
	badges: {
		display: "flex",
		gap: tokens.spacingHorizontalXXS,
		alignItems: "center",
	},
	// "(no attributes)" after a link-entity (#24)
	hint: {
		color: tokens.colorNeutralForeground4,
		marginLeft: "6px",
		fontStyle: "italic",
	},
	// Delete in the context menu (#24)
	danger: {
		color: tokens.colorPaletteRedForeground1,
		"& svg": {
			color: tokens.colorPaletteRedForeground1,
		},
	},
});

type NodeBadge = { text: string; color?: "danger" | "brand" | "informative"; title?: string };

/** The badges whose condition holds */
const badgeList = (...items: (NodeBadge | false | undefined)[]): NodeBadge[] => items.filter((b): b is NodeBadge => !!b);

/** An item of a row's Add menu (#24) */
interface AddItem {
	key: string;
	label: string;
	icon: ReactElement;
	disabled?: boolean;
	onClick: () => void;
}

/** What a delete removes, for the confirmation (#24) */
interface DeleteTarget {
	id: NodeId;
	/** "Delete link-entity contact?" */
	title: string;
	/** Items inside the node: 0 deletes without asking */
	descendants: number;
}

/** One tree row */
interface RowSpec {
	id: NodeId;
	icon: ReactElement;
	label: ReactNode;
	/** Grey italic text after the label */
	hint?: string;
	badges?: NodeBadge[];
	addItems?: AddItem[];
	/** Omitted for rows that can't be deleted (the root entity) */
	deleteTarget?: DeleteTarget;
	children?: (ReactElement | false | undefined)[];
}

interface ContextMenuState {
	open: boolean;
	target?: PositioningVirtualElement;
	addItems?: AddItem[];
	deleteTarget?: DeleteTarget;
}

interface NodeTooltipState {
	nodeId: NodeId;
	text: string;
	x: number;
	y: number;
}

/** The primary key of an entity, from the cached metadata */
const primaryIdOf = (logicalName: string): string | undefined =>
	metadataCache.getEntityMetadata(logicalName)?.PrimaryIdAttribute ??
	metadataCache.getAllEntityMetadata()?.find((e) => e.LogicalName === logicalName)?.PrimaryIdAttribute;

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
	const handleNodeSelect = (nodeId: NodeId) => (event: React.MouseEvent) => {
		// The expand chevron toggles a branch; it should not change the active properties node.
		if (
			event.target instanceof Element &&
			event.target.closest(".fui-TreeItemLayout__expandIcon")
		) {
			return;
		}

		event.preventDefault();
		onNodeSelect(nodeId);
	};

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
	const treeRef = useRef<HTMLDivElement>(null);
	const [openItems, setOpenItems] = useState<Set<NodeId>>(
		new Set([fetchQuery.id, fetchQuery.entity.id])
	);

	// The Add menus render in a portal inside the row, and React events bubble through portals: a
	// menu click reached the row's onClick and re-selected the row, so the node the menu had just
	// added (and selected) lost the selection (#24, #25)
	const stopRowSelect = (e: React.MouseEvent) => e.stopPropagation();

	// Right-click / Shift+F10 menu (#24): one menu for the tree, opened at the row
	const [contextMenu, setContextMenu] = useState<ContextMenuState>({ open: false });
	const contextReturnFocus = useRef<HTMLElement | null>(null);
	const [nodeTooltip, setNodeTooltip] = useState<NodeTooltipState | null>(null);
	// A node with items inside waits for the confirmation (#24)
	const [pendingDelete, setPendingDelete] = useState<DeleteTarget | null>(null);

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

			// Bring the selected node into view once its ancestors have expanded (#25). Only this tree,
			// and only when the node is off-screen ("nearest"): selecting a visible node doesn't scroll.
			// (Before, it always scrolled to the last tree item, so selecting any node jumped to the bottom.)
			const timer = setTimeout(() => {
				const item = treeRef.current?.querySelector<HTMLElement>(
					`[data-fui-tree-item-value="${CSS.escape(selectedNodeId)}"]`,
				);
				item?.scrollIntoView({ behavior: "smooth", block: "nearest" });
			}, 250);
			return () => clearTimeout(timer);
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

	// ============ ADD / DELETE (#24) ============

	/** Delete a node; one with items inside asks first (#24) */
	const requestDelete = (target: DeleteTarget) => {
		if (target.descendants > 0) setPendingDelete(target);
		else onRemoveNode(target.id);
	};

	/** The Add items of an entity or link-entity, in the order of #24 */
	const entityAddItems = (entity: EntityNode | LinkEntityNode): AddItem[] => {
		// exists / in / any / not any / all / not all return nothing: attributes and orders fail (0x80041121)
		const canHaveAttributes = entity.type === "entity" || !NON_RETURNING_LINK_TYPES.has(entity.linkType);
		// Only one top-level filter per entity / link-entity
		const hasFilter = entity.filters.length > 0;
		return [
			{ key: "all", label: "All attributes", icon: <TableSimple20Regular />, disabled: !canHaveAttributes, onClick: () => onAddAllAttributes(entity.id) },
			{
				key: "select",
				label: "Select attributes…",
				icon: <CheckboxChecked20Regular />,
				disabled: !canHaveAttributes,
				onClick: () => void handleOpenSelectAttributes(entity),
			},
			{ key: "blank", label: "Blank attribute", icon: <TextAdd20Regular />, disabled: !canHaveAttributes, onClick: () => onAddAttribute(entity.id) },
			{ key: "order", label: "Order by", icon: <ArrowSort20Regular />, disabled: !canHaveAttributes, onClick: () => onAddOrder(entity.id) },
			{ key: "filter", label: "Filter", icon: <Filter20Regular />, disabled: hasFilter, onClick: () => onAddFilter(entity.id) },
			{ key: "link", label: "Link-entity", icon: <Link20Regular />, onClick: () => onAddLinkEntity(entity.id) },
		];
	};

	const filterAddItems = (filter: FilterNode): AddItem[] => [
		{ key: "condition", label: "Condition", icon: <Filter20Regular />, onClick: () => onAddCondition(filter.id) },
		{ key: "subfilter", label: "Sub-filter", icon: <Filter20Regular />, onClick: () => onAddSubfilter(filter.id) },
		{ key: "link", label: "Link-entity (any/not any/all)", icon: <Link20Regular />, onClick: () => onAddLinkEntity(filter.id) },
	];

	const renderMenuItems = (items: AddItem[]) =>
		items.map((item) => (
			<MenuItem key={item.key} icon={item.icon} disabled={item.disabled} onClick={item.onClick}>
				{item.label}
			</MenuItem>
		));

	/** Opens the row's context menu at a point (right-click, Shift+F10 or the context-menu key) */
	const openContextMenu = (x: number, y: number, row: RowSpec, returnFocusTo: HTMLElement | null) => {
		onNodeSelect(row.id);
		contextReturnFocus.current = returnFocusTo;
		setContextMenu({
			open: true,
			addItems: row.addItems,
			deleteTarget: row.deleteTarget,
			target: {
				getBoundingClientRect: () => ({ x, y, top: y, left: x, bottom: y, right: x, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect,
			},
		});
	};

	/**
	 * One tree row: its Add menu and Delete button, the context menu, and a chevron only when it has
	 * items inside (#24)
	 */
	const renderRow = (row: RowSpec) => {
		const children = row.children?.filter(Boolean) ?? [];
		const hasChildren = children.length > 0;
		const hasContextMenu = !!row.addItems?.length || !!row.deleteTarget;
		const fullLabel = typeof row.label === "string" ? [row.label, row.hint].filter(Boolean).join(" ") : undefined;
		return (
			<TreeItem
				key={row.id}
				itemType={hasChildren ? "branch" : "leaf"}
				value={row.id}
				className={styles.treeItemWithActions}
				onKeyDown={(e: React.KeyboardEvent<HTMLDivElement>) => {
					// Only the focused row, not the rows it contains (keyboard events bubble up the tree)
					if (!hasContextMenu || e.target !== e.currentTarget) return;
					if (e.key === "ContextMenu" || (e.shiftKey && e.key === "F10")) {
						e.preventDefault();
						const layout = e.currentTarget.querySelector<HTMLElement>(":scope > .fui-TreeItemLayout") ?? e.currentTarget;
						const rect = layout.getBoundingClientRect();
						openContextMenu(rect.left + 24, rect.bottom, row, e.currentTarget);
					}
				}}
			>
				<TreeItemLayout
					className={mergeClasses(
						styles.treeItemLayout,
						selectedNodeId === row.id && styles.selectedTreeItem,
					)}
					iconBefore={row.icon}
					actions={
						row.addItems || row.deleteTarget ? (
							<div className="actions">
								{row.addItems && (
									<Menu>
										<MenuTrigger disableButtonEnhancement>
											<Button appearance="subtle" size="small" icon={<Add20Regular />} title="Add…" aria-label="Add…" />
										</MenuTrigger>
										<MenuPopover onClick={stopRowSelect}>
											<MenuList>{renderMenuItems(row.addItems)}</MenuList>
										</MenuPopover>
									</Menu>
								)}
								{row.deleteTarget && (
									<Button
										appearance="subtle"
										size="small"
										icon={<Delete20Regular />}
										onClick={(e) => {
											e.stopPropagation();
											requestDelete(row.deleteTarget!);
										}}
										title="Delete"
										aria-label="Delete"
									/>
								)}
							</div>
						) : undefined
					}
					aside={renderBadges(row.id, row.badges)}
					onClick={handleNodeSelect(row.id)}
					onContextMenu={
						hasContextMenu
							? (e: React.MouseEvent<HTMLDivElement>) => {
									e.preventDefault();
									e.stopPropagation();
									openContextMenu(e.clientX, e.clientY, row, e.currentTarget.closest<HTMLElement>('[role="treeitem"]'));
								}
							: undefined
					}
				>
					<span
						className={styles.label}
						onPointerEnter={(event) => {
							const label = event.currentTarget;
							if (!fullLabel || label.scrollWidth <= label.clientWidth) return;
							setNodeTooltip({
								nodeId: row.id,
								text: fullLabel,
								x: event.clientX,
								y: event.clientY,
							});
						}}
						onPointerMove={(event) => {
							setNodeTooltip((current) =>
								current?.nodeId === row.id
									? { ...current, x: event.clientX, y: event.clientY }
									: current,
							);
						}}
						onPointerLeave={() => setNodeTooltip((current) => current?.nodeId === row.id ? null : current)}
					>
						{row.label}
						{row.hint && <span className={styles.hint}>{row.hint}</span>}
					</span>
				</TreeItemLayout>

				{hasChildren && <Tree className={styles.tree}>{children}</Tree>}
			</TreeItem>
		);
	};

	// ============ NODES ============

	const renderFetchNode = (fetch: FetchNode) =>
		renderRow({
			id: fetch.id,
			icon: <Table20Regular />,
			label: "Fetch query",
			badges: badgeList(
				fetch.options.aggregate === true && { text: "Aggregate", color: "brand" },
				fetch.options.distinct && { text: "Distinct" },
				fetch.options.dataSource === "retained" && { text: "Retained" },
				(fetch.options.queryHints?.length ?? 0) > 0 && { text: "Hints" },
			),
			children: [renderEntityNode(fetch.entity)],
		});

	/** Children shared by <entity> and <link-entity> */
	const containerChildren = (container: EntityNode | LinkEntityNode) => [
		container.allAttributes?.enabled && renderAllAttributesNode(container.allAttributes),
		...container.attributes.map((attr) => renderAttributeNode(attr)),
		...container.orders.map((order) => renderOrderNode(order)),
		...container.filters.map((filter) => renderFilterNode(filter, container.name)),
		...container.links.map((link) => renderLinkEntityNode(link, container.name)),
	];

	// The root entity: no Delete (New query replaces it)
	const renderEntityNode = (entity: EntityNode) =>
		renderRow({
			id: entity.id,
			icon: <Table20Regular />,
			label: entity.name,
			addItems: entityAddItems(entity),
			children: containerChildren(entity),
		});

	const renderAllAttributesNode = (node: AllAttributesNode) =>
		renderRow({
			id: node.id,
			icon: <TableSimple20Regular />,
			label: "All attributes",
			deleteTarget: { id: node.id, title: "Delete all attributes?", descendants: 0 },
		});

	const renderAttributeNode = (attr: AttributeNode) => {
		const badges: NodeBadge[] = [];
		if (attr.groupby) badges.push({ text: attr.dategrouping ? `Group by · ${attr.dategrouping}` : "Group by", color: "brand" });
		if (attr.aggregate) badges.push({ text: FUNCTION_TEXT[attr.aggregate] ?? attr.aggregate, color: "brand" });
		if (attr.aggregate === "countcolumn" && attr.distinct) badges.push({ text: "Unique" });
		if (attr.rowAggregate) badges.push({ text: "Child count" });
		return renderRow({
			id: attr.id,
			icon: <Column20Regular />,
			label: attr.alias ? `${attr.name} (as ${attr.alias})` : attr.name,
			badges,
			deleteTarget: { id: attr.id, title: `Delete attribute ${attr.name}?`, descendants: 0 },
		});
	};

	const renderOrderNode = (order: OrderNode) => {
		const direction = order.descending ? "↓" : "↑";
		// Aggregate queries sort by alias (#90); show entityname prefix for link-entity orders
		const target = order.alias ? `${order.alias} (alias)` : (order.attribute ?? "");
		const displayName = order.entityname ? `${order.entityname}.${target}` : target;
		return renderRow({
			id: order.id,
			icon: <ArrowSort20Regular />,
			label: `${displayName} ${direction}`,
			deleteTarget: { id: order.id, title: `Delete order ${displayName}?`, descendants: 0 },
		});
	};

	const renderFilterNode = (filter: FilterNode, containerName: string): ReactElement =>
		renderRow({
			id: filter.id,
			icon: <Filter20Regular />,
			label: `Filter (${filter.conjunction})`,
			badges: badgeList(filter.hint === "union" && { text: "Union" }, filter.isQuickFindFields && { text: "Quick find" }),
			addItems: filterAddItems(filter),
			deleteTarget: { id: filter.id, title: `Delete filter (${filter.conjunction})?`, descendants: countDescendants(filter) },
			children: [
				...filter.conditions.map((cond) => renderConditionNode(cond)),
				...filter.subfilters.map((subfilter) => renderFilterNode(subfilter, containerName)),
				...(filter.links ?? []).map((link) => renderLinkEntityNode(link, containerName)),
			],
		});

	const renderConditionNode = (condition: ConditionNode) => {
		// Format the display text based on operator and value
		const requiresValue = operatorRequiresValue(condition.operator);
		let displayText = `${condition.attribute} ${condition.operator}`;

		if (requiresValue && condition.valueof) {
			// Compared to another attribute
			displayText += ` → ${condition.valueof}`;
		} else if (requiresValue && condition.value !== undefined && condition.value !== null && condition.value !== "") {
			displayText += ` = ${formatDisplayValue(condition.value)}`;
		}

		return renderRow({
			id: condition.id,
			icon: <Filter20Regular />,
			label: displayText,
			deleteTarget: { id: condition.id, title: `Delete condition ${condition.attribute}?`, descendants: 0 },
		});
	};

	/** @param parentName - The entity / link-entity this link-entity joins from */
	const renderLinkEntityNode = (link: LinkEntityNode, parentName: string): ReactElement =>
		renderRow({
			id: link.id,
			icon: <Link20Regular />,
			// Names the lookup when there's no alias (#25)
			label: linkEntityLabel(link, primaryIdOf(parentName)),
			hint: showsNoAttributesHint(link) ? "(no attributes)" : undefined,
			addItems: entityAddItems(link),
			deleteTarget: { id: link.id, title: `Delete link-entity ${link.alias || link.name}?`, descendants: countDescendants(link) },
			children: containerChildren(link),
		});

	return (
		<div ref={treeRef} className={styles.container}>
			<Tree
				className={styles.tree}
				aria-label="FetchXML Query Structure"
				openItems={openItems}
				onOpenChange={handleOpenChange}
			>
				{renderFetchNode(fetchQuery)}
			</Tree>

			<Tooltip
				content={nodeTooltip?.text ?? ""}
				relationship="inaccessible"
				visible={nodeTooltip !== null}
				positioning={{
					target: nodeTooltip
						? {
								getBoundingClientRect: () => {
									const { x, y } = nodeTooltip;
									return { x, y, top: y, left: x, bottom: y, right: x, width: 0, height: 0, toJSON: () => ({}) } as DOMRect;
								},
							}
						: undefined,
					position: "below",
					align: "start",
					offset: 8,
				}}
			/>

			{/* Context menu: the row's Add items, then Delete (#24) */}
			<Menu
				open={contextMenu.open}
				onOpenChange={(_, data) => {
					setContextMenu((current) => ({ ...current, open: data.open }));
					if (!data.open) {
						// Back to the row, unless the chosen item moved the focus (a dialog)
						const row = contextReturnFocus.current;
						requestAnimationFrame(() => {
							if (row?.isConnected && (!document.activeElement || document.activeElement === document.body)) row.focus();
						});
					}
				}}
				positioning={{ target: contextMenu.target, position: "below", align: "start" }}
			>
				<MenuPopover onClick={stopRowSelect}>
					<MenuList>
						{contextMenu.addItems && renderMenuItems(contextMenu.addItems)}
						{contextMenu.addItems && contextMenu.deleteTarget && <MenuDivider />}
						{contextMenu.deleteTarget && (
							<MenuItem
								className={styles.danger}
								icon={<Delete20Regular />}
								onClick={() => contextMenu.deleteTarget && requestDelete(contextMenu.deleteTarget)}
							>
								Delete
							</MenuItem>
						)}
					</MenuList>
				</MenuPopover>
			</Menu>

			{/* Deleting a node with items inside asks first (#24) */}
			<Dialog open={pendingDelete !== null} onOpenChange={(_, data) => !data.open && setPendingDelete(null)}>
				<DialogSurface>
					<DialogBody>
						<DialogTitle>{pendingDelete?.title}</DialogTitle>
						<DialogContent>{pendingDelete && deleteConfirmationText(pendingDelete.descendants)}</DialogContent>
						<DialogActions>
							<Button appearance="secondary" onClick={() => setPendingDelete(null)}>
								Cancel
							</Button>
							<Button
								appearance="primary"
								onClick={() => {
									if (pendingDelete) onRemoveNode(pendingDelete.id);
									setPendingDelete(null);
								}}
							>
								Delete
							</Button>
						</DialogActions>
					</DialogBody>
				</DialogSurface>
			</Dialog>

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
