/**
 * DataGrid for displaying FetchXML query results with virtualization
 * Uses rich cell renderers based on attribute metadata for Power Apps-like experience
 * Displays formatted values from OData annotations when available
 * Uses @fluentui-contrib/react-data-grid-react-window for 1D virtualization with multi-select and resizable columns
 * Sort state is derived from FetchXML orders - clicking headers triggers onSortChange callback
 */

import { useMemo, useCallback, useEffect, useState, useRef, forwardRef, useImperativeHandle } from "react";
import {
	createTableColumn,
	makeStyles,
	tokens,
	TableCellLayout,
	Skeleton,
	SkeletonItem,
	Menu,
	MenuPopover,
	MenuList,
	MenuItem,
	MenuDivider,
	Popover,
	PopoverSurface,
	Button,
	Text,
	Tooltip,
	TagGroup,
	Tag,
	SearchBox,
	Toaster,
	Toast,
	ToastTitle,
	useToastController,
	useId,
	type PositioningVirtualElement,
} from "@fluentui/react-components";
import {
	DataGrid,
	DataGridHeader,
	DataGridHeaderCell,
	DataGridBody,
	DataGridRow,
	DataGridCell,
} from "@fluentui-contrib/react-data-grid-react-window";
import type { RowRenderer } from "@fluentui-contrib/react-data-grid-react-window";
import { useScrollbarWidth, useFluent } from "@fluentui/react-components";
import type { SortDirection } from "@fluentui/react-components";
import {
	ArrowSortUp16Regular,
	ArrowSortDown16Regular,
	ArrowSortUp20Regular,
	ArrowSortDown20Regular,
	Copy20Regular,
	Filter16Regular,
	Filter20Regular,
	FilterDismiss20Regular,
} from "@fluentui/react-icons";
import type { AttributeMetadata } from "../../api/pptbClient";
import { getCellRenderer } from "./getCellRenderer";
import { buildRowIdentity } from "./rowIdentity";
import { getFormattedValue, filterDisplayableColumns } from "./FormattedValueUtils";
import type { ConditionNode, FetchNode } from "../../model/nodes";
import { columnFilterTarget, type ColumnFilterTarget, type NewCondition } from "../../model/columnFilters";
import { ALL_OPERATORS, operatorRequiresValue } from "../../model/operators";
import { collectAllLinkEntities } from "../../model/treeUtils";
import { ConditionEditor } from "../LeftPane/PropertiesPanel/editors/ConditionEditor";
import { HighlightText } from "./HighlightText";
import { SearchHighlightContext } from "./searchHighlightContext";
import type { LayoutXmlConfig } from "../../model/layoutxml";
import type { DisplaySettings } from "../../model/displaySettings";
import { metadataCache } from "../../state/cache";
import { collectQueryColumns, joinedLinkAliases, resolveColumn, sortedColumns } from "../../model/queryColumns";
import {
	cellText,
	columnToText,
	formattedText,
	rawText,
	rowCountLabel,
	rowsToTsv,
	writeClipboard,
	type CopyColumn,
	type CopyValueMode,
} from "./copyResults";

/** Sort change event data passed to parent */
export interface SortChangeData {
	/** Result column key (queryColumns.ts); the caller maps it to an order */
	columnId: string;
	/** New sort direction */
	direction: SortDirection;
	/** Whether Shift key was held (multi-column sort) */
	isMultiSort: boolean;
}

/** Lets the results command bar copy rows (#50) */
export interface ResultsGridHandle {
	copyRows: (scope: "selected" | "all", mode: CopyValueMode) => void;
}

/** What a raw value is, for the cell menu's "Copy raw value" */
function rawValueKind(attributeType: string | undefined): string | undefined {
	switch (attributeType) {
		case "Lookup":
		case "Customer":
		case "Owner":
		case "Uniqueidentifier":
			return "GUID";
		case "DateTime":
			return "ISO date";
		case "Picklist":
		case "State":
		case "Status":
			return "option value";
		case "Boolean":
			return "true/false";
		case "Integer":
		case "BigInt":
		case "Decimal":
		case "Double":
		case "Money":
			return "number";
		default:
			return undefined;
	}
}

/** A right-clicked cell */
interface CellMenuState {
	open: boolean;
	/** The mouse position (or the cell, for the keyboard's context menu key) */
	target?: PositioningVirtualElement | HTMLElement;
	item?: Record<string, unknown>;
	columnId?: string;
}

const ROW_HEIGHT = 42;

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		backgroundColor: tokens.colorNeutralBackground1,
		position: "relative",
	},
	gridWrapper: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		backgroundColor: tokens.colorNeutralBackground1,
		position: "relative",
		minHeight: 0, // Important for flex child with overflow
	},
	gridContent: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden", // Hidden to prevent container scroll - DataGrid handles its own scrolling
		width: "100%",
		minHeight: 0, // Important for flex child with overflow
	},
	infoBar: {
		display: "flex",
		alignItems: "center",
		justifyContent: "space-between",
		padding: "8px 16px",
		borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
		backgroundColor: tokens.colorNeutralBackground1,
		fontSize: tokens.fontSizeBase200,
		color: tokens.colorNeutralForeground2,
		minHeight: "32px",
		zIndex: 1,
	},
	selectionWarning: {
		color: tokens.colorPaletteYellowForeground1,
		marginLeft: tokens.spacingHorizontalM,
	},
	emptyState: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase300,
	},
	cellContent: {
		padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	row: {
		boxSizing: "border-box",
		height: "44px",
	},
	body: {
		scrollbarGutter: "stable",
	},
	headerCellContent: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
	},
	sortIndicator: {
		display: "flex",
		alignItems: "center",
		marginLeft: "auto",
	},
	headerCell: {},
	filterIndicator: {
		flexShrink: 0,
		color: tokens.colorBrandForeground1,
	},
	filterBar: {
		display: "flex",
		alignItems: "center",
		flexWrap: "wrap",
		gap: tokens.spacingHorizontalS,
		padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
		borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
	},
	search: {
		marginLeft: "auto",
		minWidth: "140px",
		width: "240px",
		flexShrink: 1,
	},
	filterSurface: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
		width: "320px",
		maxWidth: "calc(100vw - 32px)",
	},
	filterNote: {
		color: tokens.colorNeutralForeground3,
	},
	filterActions: {
		display: "flex",
		justifyContent: "flex-end",
		gap: tokens.spacingHorizontalS,
	},
});

/** Header menu sort labels, by the attribute's type (as model-driven app grids word them) */
function sortLabels(attributeType: string | undefined): [string, string] {
	switch (attributeType) {
		case "DateTime":
			return ["Oldest to newest", "Newest to oldest"];
		case "Integer":
		case "BigInt":
		case "Decimal":
		case "Double":
		case "Money":
			return ["Smallest to largest", "Largest to smallest"];
		default:
			return ["A to Z", "Z to A"];
	}
}

/** Whether a condition has the value(s) its operator needs */
function hasRequiredValue(condition: ConditionNode): boolean {
	if (!condition.operator) return false;
	if (!operatorRequiresValue(condition.operator)) return true;
	const { value } = condition;
	if (Array.isArray(value)) return value.length > 0 && value.every((v) => v !== "" && v !== undefined && v !== null);
	return value !== undefined && value !== null && value !== "";
}

const sameTarget = (c: ConditionNode, t: ColumnFilterTarget) =>
	c.attribute === t.attribute && (c.entityname || undefined) === t.entityname;

export interface QueryResult {
	columns: string[];
	rows: Record<string, unknown>[];
	totalRecordCount?: number;
	moreRecords?: boolean;
	pagingCookie?: string;
	entityLogicalName?: string; // NEW: for loading attribute metadata
	executionTimeMs?: number; // NEW: for execution timing display
}

interface ResultsGridProps {
	result: QueryResult | null;
	isLoading?: boolean;
	/** Whether more pages are currently being loaded (for infinite scroll / Retrieve All) */
	isLoadingMore?: boolean;
	/** Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>> */
	attributeMetadata?: Map<string, Map<string, AttributeMetadata>>;
	fetchQuery?: FetchNode | null; // For extracting aliases and order state
	/**
	 * The entity's primary key attribute name from Dataverse metadata (e.g. "activityid" for activity
	 * entities, "accountid" for account). Used for stable row identity and selection.
	 */
	primaryIdAttribute?: string;
	onSelectedCountChange?: (count: number) => void;
	/** Callback when selection changes, provides the selected record IDs (GUIDs) */
	onSelectionChange?: (recordIds: string[]) => void;
	/** Column layout configuration for order and widths */
	columnConfig?: LayoutXmlConfig | null;
	/** Callback when a column is resized */
	onColumnResize?: (columnName: string, newWidth: number) => void;
	/** Callback when user clicks a column header to sort */
	onSortChange?: (data: SortChangeData) => void;
	/** Callback when user scrolls near bottom and more records are available (infinite scroll) */
	onLoadMore?: () => void;
	/** Display settings (logical names, value format) */
	displaySettings?: DisplaySettings;
	/** Quick search over the loaded rows (#23) */
	searchText?: string;
	onSearchTextChange?: (text: string) => void;
	/** The column filters' conditions (#23, columnFilters.ts) */
	columnFilters?: ConditionNode[];
	/** Add a column filter's condition and run the query again */
	onAddColumnFilter?: (condition: NewCondition) => void;
	/** Remove column filters' conditions and run the query again */
	onRemoveColumnFilters?: (conditionIds: string[]) => void;
	/** Rows shown after the quick search (for the Copy menu) */
	onVisibleCountChange?: (count: number) => void;
}

export const ResultsGrid = forwardRef<ResultsGridHandle, ResultsGridProps>(function ResultsGrid({
	result,
	isLoading,
	isLoadingMore,
	attributeMetadata,
	fetchQuery,
	primaryIdAttribute,
	onSelectedCountChange,
	onSelectionChange,
	columnConfig,
	onColumnResize,
	onSortChange,
	onLoadMore,
	displaySettings,
	searchText = "",
	onSearchTextChange,
	columnFilters,
	onAddColumnFilter,
	onRemoveColumnFilters,
	onVisibleCountChange,
}: ResultsGridProps, ref) {
	const styles = useStyles();
	const { targetDocument } = useFluent();
	const scrollbarWidth = useScrollbarWidth({ targetDocument });
	const [selectedItems, setSelectedItems] = useState<Set<string | number>>(new Set());
	const containerRef = useRef<HTMLDivElement>(null);
	const headerRef = useRef<HTMLDivElement>(null);
	const [gridDimensions, setGridDimensions] = useState({ width: 0, height: 0 });
	const [headerHeight, setHeaderHeight] = useState(0);

	// Measure container dimensions for virtualization
	useEffect(() => {
		const container = containerRef.current;
		if (!container) {
			console.log("[ResultsGrid] ResizeObserver: container ref is null");
			return;
		}

		console.log("[ResultsGrid] ResizeObserver: setting up observer on container", container);

		const observer = new ResizeObserver((entries) => {
			for (const entry of entries) {
				const { width, height } = entry.contentRect;
				console.log("[ResultsGrid] ResizeObserver: dimensions updated", { width, height });
				setGridDimensions({ width, height });
			}
		});

		observer.observe(container);
		return () => observer.disconnect();
	}, [result]); // Re-run when result changes to ensure container is mounted

	// Measure header height for proper body height calculation
	useEffect(() => {
		const header = headerRef.current;
		if (!header) return;

		const observer = new ResizeObserver(() => {
			const height = header.offsetHeight;
			console.log("[ResultsGrid] Header height:", height);
			setHeaderHeight(height);
		});

		observer.observe(header);
		return () => observer.disconnect();
	}, [result]);

	// Clear selection when results change
	useEffect(() => {
		setSelectedItems(new Set());
	}, [result]);

	// Sort arrows: the result columns the query's orders sort by (queryColumns.ts, #114)
	const sortStateMap = useMemo(() => {
		const map = new Map<string, { direction: SortDirection }>();
		if (!fetchQuery?.entity) return map;
		for (const [key, direction] of sortedColumns(fetchQuery)) {
			map.set(key, { direction });
			// Root lookups also come back as _name_value
			const column = resolveColumn(fetchQuery, key);
			if (column?.isRoot && !column.attribute.alias) map.set(`_${column.attribute.name}_value`, { direction });
		}
		return map;
	}, [fetchQuery]);

	// Unique, stable row IDs; selection is translated back to real record IDs (#59).
	const rowIdentity = useMemo(
		() => buildRowIdentity(result?.rows ?? [], primaryIdAttribute),
		[result?.rows, primaryIdAttribute],
	);
	const getRowId = rowIdentity.getRowId;

	// Result columns the query asks for, keyed as Dataverse returns them (queryColumns.ts, #114).
	// Link-entities without an alias come back as "{name}{N}.attribute".
	const requestedAttributes = useMemo(() => {
		const keys = new Set<string>();
		if (!fetchQuery) return keys;
		for (const column of collectQueryColumns(fetchQuery)) keys.add(column.key);
		return keys;
	}, [fetchQuery]);

	// Build comprehensive column metadata map from FetchXML query structure
	// Maps column name (as it appears in results) to display info
	interface ColumnDisplayInfo {
		/** Original attribute logical name */
		attributeName: string;
		/** Entity logical name this attribute belongs to */
		entityName: string;
		/** Explicit alias provided in FetchXML (if any) */
		alias?: string;
		/** Link-entity alias (if from link-entity) */
		linkEntityAlias?: string;
		/** Whether this is a lookup column (N:1 navigation) */
		isLookupColumn?: boolean;
		/** The lookup attribute on parent entity that joins to this link-entity (for N:1 relationships) */
		lookupAttribute?: string;
	}

	const columnDisplayMap = useMemo(() => {
		const map = new Map<string, ColumnDisplayInfo>();
		if (!fetchQuery?.entity?.name) return map;

		const rootEntityName = fetchQuery.entity.name;

		// Collect from root entity attributes
		if (fetchQuery.entity.attributes) {
			fetchQuery.entity.attributes.forEach((attr) => {
				const columnKey = attr.alias || attr.name;
				map.set(columnKey, {
					attributeName: attr.name,
					entityName: rootEntityName,
					alias: attr.alias,
				});
				// Also map lookup column format (_attributename_value) for lookup fields
				map.set(`_${attr.name}_value`, {
					attributeName: attr.name,
					entityName: rootEntityName,
					isLookupColumn: true,
				});
			});
		}

		// Collect from link-entities recursively; Dataverse names an unaliased link-entity "{name}{N}" (#114)
		const linkAliases = joinedLinkAliases(fetchQuery);
		const collectFromLinks = (
			links: import("../../model/nodes").LinkEntityNode[],
			parentLookupAttr?: string,
		) => {
			links.forEach((link) => {
				const linkAlias = linkAliases.get(link.id) ?? link.alias ?? link.name;
				// For N:1 relationships, 'to' is the lookup attribute on the parent entity (e.g., primarycontactid)
				// For 1:N relationships, 'to' is the PK on the parent entity
				// We use 'to' as it represents the attribute on the parent that links to this entity
				const lookupAttr = link.to || parentLookupAttr;

				if (link.attributes) {
					link.attributes.forEach((attr) => {
						// For link-entity attributes, the column in results may be:
						// 1. attr.alias (if explicitly aliased)
						// 2. linkAlias.attributeName (standard format)
						const columnKey = attr.alias || `${linkAlias}.${attr.name}`;
						map.set(columnKey, {
							attributeName: attr.name,
							entityName: link.name,
							alias: attr.alias,
							linkEntityAlias: linkAlias,
							lookupAttribute: lookupAttr,
						});

					});
				}

				if (link.links) {
					collectFromLinks(link.links, link.to);
				}
			});
		};

		if (fetchQuery.entity.links) {
			collectFromLinks(fetchQuery.entity.links);
		}

		return map;
	}, [fetchQuery]);

	// Memoize column definitions with display names and formatted values
	// Uses columnConfig for ordering and widths when available
	const { columns, columnDisplayNames, columnAttributeTypes, copyColumns } = useMemo(() => {
		if (!result || result.columns.length === 0)
			return {
				columns: [],
				columnDisplayNames: new Map<string, string>(),
				columnAttributeTypes: new Map<string, string | undefined>(),
				copyColumns: [] as CopyColumn[],
			};

		// Filter out OData and CRM annotation columns
		let displayableColumns = filterDisplayableColumns(result.columns);

		// Further filter to ONLY columns requested in FetchXML (if fetchQuery is available)
		if (fetchQuery && requestedAttributes.size > 0) {
			displayableColumns = displayableColumns.filter((col) => {
				// Check if this column or its base attribute (for lookups) is requested
				// For lookup columns (_attributename_value), check the base attribute name
				if (col.startsWith("_") && col.endsWith("_value")) {
					const baseAttr = col.slice(1, -6); // Remove _ prefix and _value suffix
					return requestedAttributes.has(baseAttr);
				}
				// For aliased or regular columns, check directly
				return requestedAttributes.has(col);
			});
		}

		// If columnConfig is provided, use it to order the columns
		if (columnConfig && columnConfig.columns.length > 0) {
			const configOrder = columnConfig.columns.map((c) => c.name);

			// Build a map that handles lookup field naming convention
			// FetchXML/config uses: primarycontactid, but Dataverse returns: _primarycontactid_value
			const getConfigIndex = (col: string): number => {
				// Direct match first
				const directIndex = configOrder.indexOf(col);
				if (directIndex >= 0) return directIndex;

				// If it's a lookup column (_xxx_value), try matching the base attribute name
				if (col.startsWith("_") && col.endsWith("_value")) {
					const baseAttr = col.slice(1, -6);
					const baseIndex = configOrder.indexOf(baseAttr);
					if (baseIndex >= 0) return baseIndex;
				}

				return -1; // Not in config
			};

			// Sort displayable columns by config order, append any not in config at end
			const inConfig = displayableColumns.filter((col) => getConfigIndex(col) >= 0);
			const notInConfig = displayableColumns.filter((col) => getConfigIndex(col) < 0);

			displayableColumns = [
				...inConfig.sort((a, b) => getConfigIndex(a) - getConfigIndex(b)),
				...notInConfig,
			];
		}

		const displayNames = new Map<string, string>();
		const attrTypes = new Map<string, string | undefined>();
		// The grid's columns as copied, in grid order (#50)
		const copyCols: CopyColumn[] = [];
		const columnDefs = displayableColumns.flatMap((col) => {
			// Get column display info from our comprehensive map
			const displayInfo = columnDisplayMap.get(col);

			// Helper to get attribute metadata from multi-entity map
			const getAttributeMetadata = (
				entityName: string,
				attrName: string,
			): AttributeMetadata | undefined => {
				return attributeMetadata?.get(entityName)?.get(attrName);
			};

			// Resolve attribute metadata based on display info
			let attribute: AttributeMetadata | undefined;

			if (displayInfo) {
				attribute = getAttributeMetadata(displayInfo.entityName, displayInfo.attributeName);
			} else if (col.startsWith("_") && col.endsWith("_value")) {
				// Fallback for lookup columns not in our map
				const baseAttr = col.slice(1, -6);
				const rootEntity = fetchQuery?.entity?.name;
				if (rootEntity) {
					attribute = getAttributeMetadata(rootEntity, baseAttr);
				}
			} else {
				// Fallback: try root entity
				const rootEntity = fetchQuery?.entity?.name;
				if (rootEntity) {
					attribute = getAttributeMetadata(rootEntity, col);
				}
			}

			// Determine display name with comprehensive logic:
			// 1. For aliased columns: Use the alias as display name
			// 2. For lookup columns (_attr_value): Use attribute display name (e.g., "Primary Contact")
			// 3. For link-entity columns: "{Attribute Display Name} ({Lookup Attribute Display Name})"
			// 4. For root entity columns: Use attribute display name
			// 5. Fallback: Clean up column name
			// If displaySettings.useLogicalNames is true, show logical names instead
			let displayName: string;

			// Helper to get logical name (column key or attribute LogicalName)
			const getLogicalName = (): string => {
				if (displayInfo?.isLookupColumn) {
					return displayInfo.attributeName || col.slice(1, -6);
				}
				if (displayInfo?.linkEntityAlias) {
					return `${displayInfo.linkEntityAlias}.${displayInfo.attributeName || col}`;
				}
				return displayInfo?.attributeName || col;
			};

			if (displaySettings?.useLogicalNames) {
				// Show logical name
				displayName = getLogicalName();
			} else if (displayInfo?.alias) {
				// User-specified alias takes precedence
				displayName = displayInfo.alias;
			} else if (displayInfo?.isLookupColumn) {
				// Lookup column: show attribute display name (represents the lookup field)
				displayName =
					attribute?.DisplayName?.UserLocalizedLabel?.Label ||
					displayInfo.attributeName ||
					col.slice(1, -6); // Remove _ prefix and _value suffix
			} else if (displayInfo?.linkEntityAlias) {
				// Link-entity column: "Attribute Display Name (Related Entity Display Name)"
				// e.g., "Email (Contact)" or "Email (Primary Contact)"
				const attrDisplayName =
					attribute?.DisplayName?.UserLocalizedLabel?.Label || displayInfo.attributeName;

				// Determine what to show in parentheses:
				// - For N:1 (many-to-one): Show the lookup attribute display name (e.g., "Primary Contact")
				// - For 1:N (one-to-many): Show the related entity display name (e.g., "Contact")
				let relatedDisplayName = displayInfo.linkEntityAlias;

				if (displayInfo.lookupAttribute) {
					// N:1 relationship - get the lookup attribute display name from root entity
					const rootEntityName = fetchQuery?.entity?.name;
					if (rootEntityName) {
						const lookupAttr = getAttributeMetadata(rootEntityName, displayInfo.lookupAttribute);
						if (lookupAttr?.DisplayName?.UserLocalizedLabel?.Label) {
							relatedDisplayName = lookupAttr.DisplayName.UserLocalizedLabel.Label;
						}
					}
				} else {
					// 1:N relationship - get the entity display name
					const entityMetadata = metadataCache.getEntityMetadata(displayInfo.entityName);
					if (entityMetadata?.DisplayName?.UserLocalizedLabel?.Label) {
						relatedDisplayName = entityMetadata.DisplayName.UserLocalizedLabel.Label;
					}
				}

				displayName = `${attrDisplayName} (${relatedDisplayName})`;
			} else if (attribute?.DisplayName?.UserLocalizedLabel?.Label) {
				// Root entity attribute with metadata display name
				displayName = attribute.DisplayName.UserLocalizedLabel.Label;
			} else if (col.startsWith("_") && col.endsWith("_value")) {
				// Fallback for lookup columns without metadata
				displayName = col.slice(1, -6);
			} else {
				// Final fallback: use column name as-is
				displayName = col;
			}

			// Check if this column is sorted (from FetchXML orders)
			const sortInfo = sortStateMap.get(col);
			// Whether a column filter targets this column (#23)
			const filterTarget = fetchQuery ? columnFilterTarget(fetchQuery, col) : undefined;
			const isFiltered =
				!!filterTarget && !("reason" in filterTarget) && !!columnFilters?.some((c) => sameTarget(c, filterTarget));
			const valueMode = displaySettings?.valueDisplayMode ?? "formatted";

			// For "both" mode, check if this column has formatted values in the result set
			// Only create a raw column if at least one record has a different formatted value
			const hasFormattedValues =
				valueMode === "both" &&
				result?.rows?.some((record: Record<string, unknown>) => {
					const rawValue = record[col];
					const formattedValue = getFormattedValue(record, col);
					return formattedValue !== undefined && formattedValue !== rawValue;
				});

			// Track display name and attribute type for column sizing options
			displayNames.set(col, displayName);
			attrTypes.set(col, attribute?.AttributeType);

			// For "both" mode with formatted values, create two columns
			if (valueMode === "both" && hasFormattedValues) {
				copyCols.push(
					{ id: col, header: displayName, sourceCol: col, showsRaw: false },
					{ id: `${col}__raw`, header: `${displayName} (Raw)`, sourceCol: col, showsRaw: true },
				);
				displayNames.set(`${col}__raw`, `${displayName} (Raw)`);
				attrTypes.set(`${col}__raw`, attribute?.AttributeType);
				// Create formatted value column
				const formattedColumn = createTableColumn<Record<string, unknown>>({
					columnId: col,
					compare: (a, b) => {
						const aVal = String(getFormattedValue(a, col) ?? a[col] ?? "");
						const bVal = String(getFormattedValue(b, col) ?? b[col] ?? "");
						return aVal.localeCompare(bVal);
					},
					renderHeaderCell: () => (
						<span className={styles.headerCellContent}>
							<span style={{ fontWeight: 600 }}>{displayName}</span>
							{isFiltered && <Filter16Regular className={styles.filterIndicator} aria-label="Filtered" />}
							{/* Show sort indicator based on FetchXML order state */}
							{sortInfo && (
								<span className={styles.sortIndicator}>
									{sortInfo.direction === "ascending" ? (
										<ArrowSortUp16Regular />
									) : (
										<ArrowSortDown16Regular />
									)}
								</span>
							)}
						</span>
					),
					renderCell: (item) => {
						const rawValue = item[col];
						const formattedValue = getFormattedValue(item, col);
						return (
							<TableCellLayout truncate>
								{getCellRenderer(attribute?.AttributeType, rawValue, formattedValue, attribute)}
							</TableCellLayout>
						);
					},
				});

				// Create raw value column
				const rawColumn = createTableColumn<Record<string, unknown>>({
					columnId: `${col}__raw`,
					compare: (a, b) => {
						const aVal = String(a[col] ?? "");
						const bVal = String(b[col] ?? "");
						return aVal.localeCompare(bVal);
					},
					renderHeaderCell: () => (
						<span className={styles.headerCellContent}>
							<span style={{ fontWeight: 600 }}>{displayName} (Raw)</span>
						</span>
					),
					renderCell: (item) => {
						const rawValue = item[col];
						return (
							<TableCellLayout truncate>
								{rawValue === null || rawValue === undefined ? (
									<span>—</span>
								) : (
									<span>
										<HighlightText text={String(rawValue)} />
									</span>
								)}
							</TableCellLayout>
						);
					},
				});

				return [formattedColumn, rawColumn];
			}

			// Standard single column (formatted or raw mode)
			copyCols.push({ id: col, header: displayName, sourceCol: col, showsRaw: valueMode === "raw" });
			return createTableColumn<Record<string, unknown>>({
				columnId: col,
				compare: (a, b) => {
					const aVal = String(getFormattedValue(a, col) ?? a[col] ?? "");
					const bVal = String(getFormattedValue(b, col) ?? b[col] ?? "");
					return aVal.localeCompare(bVal);
				},
				renderHeaderCell: () => (
					<span className={styles.headerCellContent}>
						<span style={{ fontWeight: 600 }}>{displayName}</span>
						{isFiltered && <Filter16Regular className={styles.filterIndicator} aria-label="Filtered" />}
						{/* Show sort indicator based on FetchXML order state */}
						{sortInfo && (
							<span className={styles.sortIndicator}>
								{sortInfo.direction === "ascending" ? (
									<ArrowSortUp16Regular />
								) : (
									<ArrowSortDown16Regular />
								)}
							</span>
						)}
					</span>
				),
				renderCell: (item) => {
					const rawValue = item[col];
					const formattedValue = getFormattedValue(item, col);

					// Determine what to display based on value display mode
					if (valueMode === "raw") {
						// Show raw value only (skip cell renderer formatting)
						return (
							<TableCellLayout truncate>
								{rawValue === null || rawValue === undefined ? (
									<span>—</span>
								) : (
									<span>
										<HighlightText text={String(rawValue)} />
									</span>
								)}
							</TableCellLayout>
						);
					} else {
						// Default: formatted (use cell renderer)
						return (
							<TableCellLayout truncate>
								{getCellRenderer(attribute?.AttributeType, rawValue, formattedValue, attribute)}
							</TableCellLayout>
						);
					}
				},
			});
		});
		return {
			columns: columnDefs,
			columnDisplayNames: displayNames,
			columnAttributeTypes: attrTypes,
			copyColumns: copyCols,
		};
	}, [
		result,
		attributeMetadata,
		columnDisplayMap,
		fetchQuery,
		requestedAttributes,
		columnConfig,
		styles,
		sortStateMap,
		displaySettings,
		columnFilters,
	]);

	// Selection handlers
	const handleSelectionChange = useCallback(
		(_e: unknown, data: { selectedItems: Set<string | number> }) => {
			setSelectedItems(data.selectedItems);
			onSelectedCountChange?.(data.selectedItems.size);
			// Record actions receive primary key values only — never synthetic row IDs.
			onSelectionChange?.(rowIdentity.toRecordIds(data.selectedItems));
		},
		[onSelectedCountChange, onSelectionChange, rowIdentity],
	);

	// ---- Header menu: sort and column filters (#23) ----
	const [headerMenu, setHeaderMenu] = useState<{ open: boolean; anchor?: HTMLElement; columnId?: string }>({
		open: false,
	});
	const [filterFor, setFilterFor] = useState<{
		anchor: HTMLElement;
		header: string;
		target: ColumnFilterTarget;
	} | null>(null);
	const [draftCondition, setDraftCondition] = useState<ConditionNode | null>(null);

	// Header clicks open the header menu; Shift+click still adds a sort directly.
	// We calculate direction ourselves based on FetchXML state, not Fluent's sortDirection
	const handleSortChange = useCallback(
		(
			e: React.MouseEvent,
			data: { sortColumn: string | number | undefined; sortDirection: SortDirection },
		) => {
			const columnKey = String(data.sortColumn);
			if (!e.shiftKey) {
				const anchor = (e.target as HTMLElement).closest<HTMLElement>('[role="columnheader"]');
				setHeaderMenu({ open: true, anchor: anchor ?? (e.currentTarget as HTMLElement), columnId: columnKey });
				return;
			}
			if (!onSortChange) return;

			// The "(Raw)" twin column sorts like its column
			const columnId = columnKey.replace(/__raw$/, "");
			const isMultiSort = true;

			// Get current sort info for this column from FetchXML state
			const sortInfo = sortStateMap.get(columnId);

			// Calculate direction: toggle if already sorted, otherwise ascending
			let direction: SortDirection;
			if (sortInfo) {
				// Column is already sorted - toggle direction
				direction = sortInfo.direction === "ascending" ? "descending" : "ascending";
			} else {
				// Column not sorted yet - start with ascending
				direction = "ascending";
			}

			onSortChange({ columnId, direction, isMultiSort });
		},
		[onSortChange, sortStateMap],
	);

	// ---- Quick search over the loaded rows (#23) ----
	const search = searchText.trim();
	const visibleRows = useMemo(() => {
		const rows = result?.rows ?? [];
		if (!search) return rows;
		const q = search.toLowerCase();
		// Matches the values as shown, in any column
		return rows.filter((r) => copyColumns.some((c) => cellText(r, c).toLowerCase().includes(q)));
	}, [result, search, copyColumns]);

	useEffect(() => {
		onVisibleCountChange?.(visibleRows.length);
	}, [visibleRows.length, onVisibleCountChange]);

	// Rows the search hides don't stay selected: record actions work on the selection
	useEffect(() => {
		if (!search || selectedItems.size === 0) return;
		const visible = new Set<string | number>(visibleRows.map((r) => getRowId(r)));
		const kept = new Set([...selectedItems].filter((id) => visible.has(id)));
		if (kept.size !== selectedItems.size) handleSelectionChange(undefined, { selectedItems: kept });
	}, [search, visibleRows, selectedItems, getRowId, handleSelectionChange]);

	// ---- Copy (#50) ----
	const toasterId = useId("results-toaster");
	const { dispatchToast } = useToastController(toasterId);
	const [cellMenu, setCellMenu] = useState<CellMenuState>({ open: false });
	/** The cell with keyboard focus, for Ctrl+C */
	const focusedCellRef = useRef<{ item: Record<string, unknown>; columnId: string } | null>(null);

	const copyColumnById = useCallback(
		(columnId: string) => copyColumns.find((c) => c.id === columnId),
		[copyColumns],
	);

	const selectedRecords = useCallback(
		() => (result?.rows ?? []).filter((r) => selectedItems.has(getRowId(r))),
		[result, selectedItems, getRowId],
	);

	const copy = useCallback(
		async (text: string, what: string) => {
			const ok = await writeClipboard(text);
			dispatchToast(
				<Toast>
					<ToastTitle>{ok ? `${what} copied` : "Couldn't copy to the clipboard"}</ToastTitle>
				</Toast>,
				{ intent: ok ? "success" : "error", timeout: ok ? 2000 : 4000 },
			);
		},
		[dispatchToast],
	);

	const copyRecords = useCallback(
		(records: Record<string, unknown>[], mode: CopyValueMode) => {
			if (records.length === 0) return;
			void copy(rowsToTsv(records, copyColumns, mode), rowCountLabel(records.length));
		},
		[copy, copyColumns],
	);

	useImperativeHandle(
		ref,
		() => ({
			copyRows: (scope, mode) => copyRecords(scope === "all" ? visibleRows : selectedRecords(), mode),
		}),
		[copyRecords, visibleRows, selectedRecords],
	);

	const openCellMenu = useCallback((e: React.MouseEvent<HTMLElement>, item: Record<string, unknown>, columnId: string) => {
		e.preventDefault();
		// Open at the mouse; the context menu key reports (0, 0), so open at the cell instead
		const { clientX: x, clientY: y } = e;
		const target: PositioningVirtualElement | HTMLElement =
			x === 0 && y === 0
				? e.currentTarget
				: { getBoundingClientRect: () => ({ x, y, top: y, left: x, bottom: y, right: x, width: 0, height: 0 }) };
		setCellMenu({ open: true, target, item, columnId });
	}, []);

	// Ctrl+C copies the selected rows, or else the focused cell, as shown. Text the user has
	// highlighted copies as usual.
	const handleGridKeyDown = useCallback(
		(e: React.KeyboardEvent) => {
			if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== "c" || e.altKey || e.shiftKey) return;
			if (window.getSelection()?.toString()) return;
			if (selectedItems.size > 0) {
				e.preventDefault();
				copyRecords(selectedRecords(), "shown");
				return;
			}
			const focused = focusedCellRef.current;
			const column = focused && copyColumnById(focused.columnId);
			if (focused && column) {
				e.preventDefault();
				void copy(cellText(focused.item, column), column.header);
			}
		},
		[selectedItems, copyRecords, selectedRecords, copyColumnById, copy],
	);

	// The right-clicked cell's values, for the cell menu
	const cellMenuContent = useMemo(() => {
		const { item, columnId } = cellMenu;
		const column = columnId ? copyColumnById(columnId) : undefined;
		if (!item || !column) return undefined;
		return {
			item,
			column,
			shown: cellText(item, column),
			raw: rawText(item, column.sourceCol),
			rawKind: rawValueKind(columnAttributeTypes.get(column.id)),
		};
	}, [cellMenu, copyColumnById, columnAttributeTypes]);

	// ---- Column filters (#23) ----
	const headerMenuColumnKey = headerMenu.columnId?.replace(/__raw$/, "");
	const headerMenuTarget = useMemo(
		() => (fetchQuery && headerMenuColumnKey ? columnFilterTarget(fetchQuery, headerMenuColumnKey) : undefined),
		[fetchQuery, headerMenuColumnKey],
	);
	const headerMenuFilters =
		headerMenuTarget && !("reason" in headerMenuTarget)
			? (columnFilters ?? []).filter((c) => sameTarget(c, headerMenuTarget))
			: [];
	const [sortAscLabel, sortDescLabel] = sortLabels(
		headerMenu.columnId ? columnAttributeTypes.get(headerMenu.columnId) : undefined,
	);

	const sortFromMenu = (direction: SortDirection) => {
		if (headerMenuColumnKey) onSortChange?.({ columnId: headerMenuColumnKey, direction, isMultiSort: false });
	};

	const openFilter = () => {
		if (!headerMenu.anchor || !headerMenuColumnKey || !headerMenuTarget || "reason" in headerMenuTarget) return;
		setFilterFor({
			anchor: headerMenu.anchor,
			header: copyColumnById(headerMenuColumnKey)?.header ?? headerMenuTarget.attribute,
			target: headerMenuTarget,
		});
		setDraftCondition({
			id: "column-filter-draft",
			type: "condition",
			attribute: headerMenuTarget.attribute,
			entityname: headerMenuTarget.entityname,
			operator: "" as ConditionNode["operator"],
		});
	};

	const applyFilter = () => {
		if (!draftCondition || !hasRequiredValue(draftCondition)) return;
		const { attribute, entityname, operator, value } = draftCondition;
		onAddColumnFilter?.({ attribute, entityname, operator, value });
		setFilterFor(null);
	};

	const linkEntities = useMemo(() => collectAllLinkEntities(fetchQuery ?? null), [fetchQuery]);

	// "Annual Revenue greater than 100000", "Status Reason equals Active"
	const filterLabel = useCallback(
		(c: ConditionNode): string => {
			const link = c.entityname ? linkEntities.find((l) => (l.alias || l.name) === c.entityname) : undefined;
			const entityName = link?.name ?? fetchQuery?.entity.name ?? "";
			const attribute = attributeMetadata?.get(entityName)?.get(c.attribute);
			const name = displaySettings?.useLogicalNames
				? `${c.entityname ? `${c.entityname}.` : ""}${c.attribute}`
				: `${attribute?.DisplayName?.UserLocalizedLabel?.Label ?? c.attribute}${c.entityname ? ` (${c.entityname})` : ""}`;
			const operator = (ALL_OPERATORS.find((o) => o.value === c.operator)?.label ?? c.operator).toLowerCase();
			const values = c.value === undefined || c.value === null ? [] : Array.isArray(c.value) ? c.value : [c.value];
			// Choice values show as their labels, taken from a loaded row with the same value
			const column = copyColumns.find((col) => {
				const target = fetchQuery ? columnFilterTarget(fetchQuery, col.sourceCol) : undefined;
				return !!target && !("reason" in target) && sameTarget(c, target);
			});
			const valueText = (v: unknown) => {
				const row = column && result?.rows.find((r) => rawText(r, column.sourceCol) === String(v));
				return row && column ? formattedText(row, column.sourceCol) : String(v);
			};
			return [name, operator, values.map(valueText).join(", ")].filter(Boolean).join(" ");
		},
		[linkEntities, fetchQuery, attributeMetadata, displaySettings, copyColumns, result],
	);

	// Above the grid: the column filters (left) and the quick search (right)
	const hasFilters = !!columnFilters && columnFilters.length > 0;
	const showSearch = !!onSearchTextChange && ((result?.rows.length ?? 0) > 0 || !!searchText);
	const filterBar =
		hasFilters || showSearch ? (
			<div className={styles.filterBar}>
				{hasFilters && (
					<>
						<Text size={200} weight="semibold">
							Filters
						</Text>
						<TagGroup
							size="small"
							aria-label="Column filters"
							onDismiss={(_, data) => onRemoveColumnFilters?.([String(data.value)])}
						>
							{columnFilters.map((c) => {
								const label = filterLabel(c);
								return (
									<Tag key={c.id} value={c.id} dismissible dismissIcon={{ "aria-label": `Remove filter: ${label}` }}>
										{label}
									</Tag>
								);
							})}
						</TagGroup>
						{columnFilters.length > 1 && (
							<Button
								size="small"
								appearance="transparent"
								onClick={() => onRemoveColumnFilters?.(columnFilters.map((c) => c.id))}
							>
								Clear all
							</Button>
						)}
					</>
				)}
				{showSearch && (
					<Tooltip
						content="Searches the loaded rows. To search every record, use Filter by… in a column header."
						relationship="description"
					>
						<SearchBox
							className={styles.search}
							size="small"
							placeholder="Search loaded rows"
							aria-label="Search loaded rows"
							value={searchText}
							onChange={(_, data) => onSearchTextChange?.(data.value)}
						/>
					</Tooltip>
				)}
			</div>
		) : null;

	// Row renderer function for virtualization with scrolling indicator support
	const renderRow: RowRenderer<Record<string, unknown>> = useCallback(
		({ item, rowId }, style, _index, isScrolling) => (
			<DataGridRow<Record<string, unknown>> key={rowId} style={style} className={styles.row}>
				{({ renderCell, columnId }) => (
					<DataGridCell
						focusMode="group"
						onContextMenu={(e: React.MouseEvent<HTMLElement>) => openCellMenu(e, item, String(columnId))}
						onFocus={() => {
							focusedCellRef.current = { item, columnId: String(columnId) };
						}}
						onBlur={() => {
							focusedCellRef.current = null;
						}}
					>
						{isScrolling ? (
							<Skeleton style={{ width: "100%" }}>
								<SkeletonItem shape="rectangle" animation="pulse" appearance="translucent" />
							</Skeleton>
						) : (
							renderCell(item)
						)}
					</DataGridCell>
				)}
			</DataGridRow>
		),
		[styles.row, openCellMenu],
	);

	// Build column sizing options from columnConfig or defaults
	const columnSizingOptions = useMemo(() => {
		const options: Record<string, { minWidth: number; defaultWidth: number; idealWidth?: number }> =
			{};

		// Minimum width implied by the attribute's data type (data content is often wider than the header)
		const typeBasedMinWidth = (attrType: string | undefined): number => {
			switch (attrType) {
				case "Uniqueidentifier":
					return 260; // GUIDs: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (36 chars)
				case "DateTime":
					return 160; // e.g. "Apr 3, 2026, 09:00 AM"
				case "Boolean":
					return 100; // switch + label
				case "Lookup":
				case "Customer":
				case "Owner":
					return 140; // formatted name, usually a person or account name
				case "Money":
					return 120;
				default:
					return 100;
			}
		};

		// Create a map of column widths from config
		// Also map lookup field variants (_xxx_value -> config width for xxx)
		const configWidths = new Map<string, number>();
		if (columnConfig) {
			for (const col of columnConfig.columns) {
				configWidths.set(col.name, col.width);
				// Also map the lookup column variant
				configWidths.set(`_${col.name}_value`, col.width);
			}
		}

		// Set sizing for each column
		// When no config width is available, estimate from the display name so headers don't overlap
		for (const col of columns) {
			const columnId = String(col.columnId);
			const configWidth = configWidths.get(columnId);

			if (configWidth !== undefined) {
				// Use the layout-XML / user-resized width as-is
				options[columnId] = {
					minWidth: 80,
					defaultWidth: configWidth,
					idealWidth: configWidth,
				};
			} else {
				// No config width – estimate from BOTH header display name and attribute data type.
				// Take the larger of the two so neither the header nor typical data values overflow.
				const displayName = columnDisplayNames.get(columnId) ?? columnId;
				const headerEstimate = Math.ceil(displayName.length * 8) + 52;
				const dataMin = typeBasedMinWidth(columnAttributeTypes.get(columnId));
				const estimated = Math.max(headerEstimate, dataMin);
				const dynamicWidth = Math.max(100, Math.min(estimated, 320));
				options[columnId] = {
					minWidth: Math.max(80, Math.min(dataMin, 200)),
					defaultWidth: dynamicWidth,
					idealWidth: dynamicWidth,
				};
			}
		}

		return options;
	}, [columns, columnConfig, columnDisplayNames, columnAttributeTypes]);

	// Handle column resize callback
	// Maps lookup column names back to config names for storage
	const handleColumnResize = useCallback(
		(
			_e: KeyboardEvent | TouchEvent | MouseEvent | undefined,
			data: { columnId: string | number; width: number },
		) => {
			if (onColumnResize) {
				let columnName = String(data.columnId);
				// If it's a lookup column (_xxx_value), store width under the base attribute name
				if (columnName.startsWith("_") && columnName.endsWith("_value")) {
					columnName = columnName.slice(1, -6);
				}
				onColumnResize(columnName, data.width);
			}
		},
		[onColumnResize],
	);

	// Infinite scroll: load more when user scrolls near bottom
	const handleItemsRendered = useCallback(
		({ visibleStopIndex }: { visibleStartIndex: number; visibleStopIndex: number }) => {
			// While searching, the visible rows aren't the loaded rows: don't load pages behind the user's back
			if (!result || !onLoadMore || isLoadingMore || search) return;

			// If user has scrolled to within 10 rows of the end, and there are more records
			const rowCount = result.rows.length;
			const threshold = Math.min(10, Math.max(1, Math.floor(rowCount * 0.1))); // 10 rows or 10% of data, whichever is smaller

			if (visibleStopIndex >= rowCount - threshold && result.moreRecords) {
				console.log("[ResultsGrid] Near bottom, triggering load more:", {
					visibleStopIndex,
					rowCount,
					threshold,
					moreRecords: result.moreRecords,
				});
				onLoadMore();
			}
		},
		[result, onLoadMore, isLoadingMore, search],
	);

	// Note: Data is already sorted by Dataverse based on FetchXML orders
	// No local sorting needed - we display rows as returned

	// The column filters stay visible (and removable) while loading and when they leave no rows
	const withFilterBar = (content: React.ReactNode) =>
		filterBar ? (
			<div className={styles.container}>
				{filterBar}
				{content}
			</div>
		) : (
			content
		);

	if (isLoading) {
		return withFilterBar(<div className={styles.emptyState}>Loading results...</div>);
	}

	if (!result) {
		return withFilterBar(<div className={styles.emptyState}>Click Execute to run the query</div>);
	}

	if (result.rows.length === 0) {
		return withFilterBar(<div className={styles.emptyState}>No records found</div>);
	}

	console.log("[ResultsGrid] Rendering grid:", {
		rowCount: result.rows.length,
		columnCount: columns.length,
		gridDimensions,
		primaryIdAttribute,
		firstRowId: result.rows.length > 0 ? getRowId(result.rows[0]) : "none",
		columnIds: columns.slice(0, 5).map((c) => c.columnId),
		sortStateMapSize: sortStateMap.size,
	});

	return (
		<div className={styles.container}>
			{filterBar}
			<div className={styles.gridWrapper} onKeyDown={handleGridKeyDown}>
				<div ref={containerRef} className={styles.gridContent}>
					<SearchHighlightContext.Provider value={search}>
						<DataGrid
							aria-label={result?.entityLogicalName ?? "Results"}
							items={visibleRows}
							columns={columns}
							sortable
							// Controlled sortState with a sentinel column ID that never matches any real column.
							// This suppresses Fluent's built-in sort indicator entirely so only our custom
							// renderHeaderCell arrows (driven by FetchXML order state) are ever visible.
							sortState={{ sortColumn: "__sort_managed_externally__", sortDirection: "ascending" }}
							resizableColumns
							resizableColumnsOptions={{ autoFitColumns: false }}
							columnSizingOptions={columnSizingOptions}
							onColumnResize={handleColumnResize}
							onSortChange={handleSortChange}
							selectionMode="multiselect"
							selectedItems={selectedItems}
							onSelectionChange={handleSelectionChange}
							getRowId={getRowId}
						>
							<DataGridHeader ref={headerRef} style={{ paddingRight: scrollbarWidth }}>
								<DataGridRow style={{ minHeight: "43px", maxHeight: "43px" }}>
									{({ renderHeaderCell }) => (
										<DataGridHeaderCell className={styles.headerCell}>
											{renderHeaderCell()}
										</DataGridHeaderCell>
									)}
								</DataGridRow>
							</DataGridHeader>
							<DataGridBody<Record<string, unknown>>
								itemSize={ROW_HEIGHT}
								height={Math.max(0, gridDimensions.height - headerHeight)}
								listProps={{
									useIsScrolling: true,
									className: styles.body,
									onItemsRendered: handleItemsRendered,
								}}
							>
								{renderRow}
							</DataGridBody>
						</DataGrid>
					</SearchHighlightContext.Provider>
				</div>
				<div className={styles.infoBar}>
					<span>
						Rows: {search ? `${visibleRows.length} of ${result.rows.length} loaded match the search` : result.rows.length}
						{result.moreRecords && !isLoadingMore && " (more available)"}
						{isLoadingMore && " (loading more...)"}
						{result.executionTimeMs !== undefined && ` | Executed in ${result.executionTimeMs}ms`}
					</span>
					<span>
						Selected: {selectedItems.size}
						{selectedItems.size > 0 &&
							selectedItems.size === result.rows.length &&
							result.moreRecords && (
								<span className={styles.selectionWarning}>
									⚠ Selection limited to loaded records. Use "Retrieve all pages" to load more.
								</span>
							)}
					</span>
				</div>
			</div>
			{cellMenuContent && (
				<Menu
					open={cellMenu.open}
					onOpenChange={(_, data) => setCellMenu((m) => ({ ...m, open: data.open }))}
					positioning={{ target: cellMenu.target, position: "below", align: "start" }}
				>
					<MenuPopover>
						<MenuList>
							<MenuItem
								icon={<Copy20Regular />}
								onClick={() => void copy(cellMenuContent.shown, cellMenuContent.column.header)}
							>
								Copy value
							</MenuItem>
							{cellMenuContent.raw !== "" && cellMenuContent.raw !== cellMenuContent.shown && (
								<MenuItem onClick={() => void copy(cellMenuContent.raw, `${cellMenuContent.column.header} (raw)`)}>
									Copy raw value{cellMenuContent.rawKind ? ` (${cellMenuContent.rawKind})` : ""}
								</MenuItem>
							)}
							<MenuDivider />
							<MenuItem onClick={() => copyRecords([cellMenuContent.item], "shown")}>Copy row</MenuItem>
							<MenuItem disabled={selectedItems.size === 0} onClick={() => copyRecords(selectedRecords(), "shown")}>
								Copy selected rows ({selectedItems.size})
							</MenuItem>
							<MenuItem
								onClick={() =>
									void copy(columnToText(result.rows, cellMenuContent.column), `${cellMenuContent.column.header} column`)
								}
							>
								Copy column
							</MenuItem>
						</MenuList>
					</MenuPopover>
				</Menu>
			)}
			<Menu
				open={headerMenu.open}
				onOpenChange={(_, data) => setHeaderMenu((m) => ({ ...m, open: data.open }))}
				positioning={{ target: headerMenu.anchor, position: "below", align: "start" }}
			>
				<MenuPopover>
					<MenuList>
						<MenuItem icon={<ArrowSortUp20Regular />} disabled={!onSortChange} onClick={() => sortFromMenu("ascending")}>
							{sortAscLabel}
						</MenuItem>
						<MenuItem icon={<ArrowSortDown20Regular />} disabled={!onSortChange} onClick={() => sortFromMenu("descending")}>
							{sortDescLabel}
						</MenuItem>
						<MenuDivider />
						{headerMenuTarget && "reason" in headerMenuTarget ? (
							<Tooltip content={headerMenuTarget.reason} relationship="description" positioning="after">
								<MenuItem icon={<Filter20Regular />} disabled>
									Filter by…
								</MenuItem>
							</Tooltip>
						) : (
							<MenuItem icon={<Filter20Regular />} disabled={!onAddColumnFilter || !headerMenuTarget} onClick={openFilter}>
								Filter by…
							</MenuItem>
						)}
						{headerMenuFilters.length > 0 && (
							<MenuItem
								icon={<FilterDismiss20Regular />}
								onClick={() => onRemoveColumnFilters?.(headerMenuFilters.map((c) => c.id))}
							>
								Clear filter
							</MenuItem>
						)}
					</MenuList>
				</MenuPopover>
			</Menu>
			<Popover
				open={!!filterFor}
				onOpenChange={(_, data) => {
					if (!data.open) setFilterFor(null);
				}}
				positioning={{ target: filterFor?.anchor, position: "below", align: "start" }}
				withArrow
				trapFocus
			>
				<PopoverSurface className={styles.filterSurface} aria-label={filterFor ? `Filter by ${filterFor.header}` : undefined}>
					{filterFor && draftCondition && fetchQuery && (
						<>
							<Text weight="semibold">Filter by {filterFor.header}</Text>
							<ConditionEditor
								columnFilter
								node={draftCondition}
								entityName={filterFor.target.entityName}
								fetchQuery={fetchQuery}
								findings={[]}
								onUpdate={(updates) => setDraftCondition((d) => (d ? { ...d, ...updates } : d))}
							/>
							<Text size={200} className={styles.filterNote}>
								Adds a condition to the query and runs it again, so it covers every record.
							</Text>
							<div className={styles.filterActions}>
								<Button onClick={() => setFilterFor(null)}>Cancel</Button>
								<Button appearance="primary" disabled={!hasRequiredValue(draftCondition)} onClick={applyFilter}>
									Apply
								</Button>
							</div>
						</>
					)}
				</PopoverSurface>
			</Popover>
			<Toaster toasterId={toasterId} position="bottom-end" />
		</div>
	);
});
