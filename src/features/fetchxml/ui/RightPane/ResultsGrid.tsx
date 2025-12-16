/**
 * DataGrid for displaying FetchXML query results with virtualization
 * Uses rich cell renderers based on attribute metadata for Power Apps-like experience
 * Displays formatted values from OData annotations when available
 * Uses @fluentui-contrib/react-data-grid-react-window for 1D virtualization with multi-select and resizable columns
 * Sort state is derived from FetchXML orders - clicking headers triggers onSortChange callback
 */

import { useMemo, useCallback, useEffect, useState, useRef } from "react";
import {
	createTableColumn,
	makeStyles,
	tokens,
	TableCellLayout,
	Skeleton,
	SkeletonItem,
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
import { ArrowSortUp16Regular, ArrowSortDown16Regular } from "@fluentui/react-icons";
import type { AttributeMetadata } from "../../api/pptbClient";
import { getCellRenderer } from "./DataGridCellRenderers";
import { getFormattedValue, filterDisplayableColumns } from "./FormattedValueUtils";
import type { FetchNode, AttributeNode, OrderNode } from "../../model/nodes";
import type { LayoutXmlConfig } from "../../model/layoutxml";
import type { DisplaySettings } from "../../model/displaySettings";

/** Sort change event data passed to parent */
export interface SortChangeData {
	/** Column being sorted (may include entityname prefix for link-entity attributes) */
	columnId: string;
	/** New sort direction */
	direction: SortDirection;
	/** Whether Shift key was held (multi-column sort) */
	isMultiSort: boolean;
	/** The entityname for link-entity attributes, undefined for root entity */
	entityName?: string;
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
	// Hide Fluent's default sort indicator - we render our own in renderHeaderCell
	headerCell: {
		"& > button > span:last-child": {
			display: "none",
		},
	},
});

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
}

export function ResultsGrid({
	result,
	isLoading,
	isLoadingMore,
	attributeMetadata,
	fetchQuery,
	onSelectedCountChange,
	onSelectionChange,
	columnConfig,
	onColumnResize,
	onSortChange,
	onLoadMore,
	displaySettings,
}: ResultsGridProps) {
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

	// Derive sort state from FetchXML orders (entity + link-entity orders)
	// This builds a map from column name to sort direction
	const sortStateMap = useMemo(() => {
		const map = new Map<string, { direction: SortDirection; entityName?: string }>();
		if (!fetchQuery?.entity) return map;

		// Build a reverse map from attribute name to alias for root entity
		const attrToAliasMap = new Map<string, string>();
		fetchQuery.entity.attributes?.forEach((attr) => {
			if (attr.alias) {
				attrToAliasMap.set(attr.name, attr.alias);
			}
		});

		// Process root entity orders
		// Orders with entityname refer to link-entity attributes
		fetchQuery.entity.orders?.forEach((order: OrderNode) => {
			const direction: SortDirection = order.descending ? "descending" : "ascending";
			if (order.entityname) {
				// This order references a link-entity attribute
				// The column in results might be: alias.attributeName or _attributeName_value for lookups
				map.set(`${order.entityname}.${order.attribute}`, {
					direction,
					entityName: order.entityname,
				});
				// Also map just the attribute name in case it appears that way
				map.set(order.attribute, { direction, entityName: order.entityname });
			} else {
				// Root entity attribute
				map.set(order.attribute, { direction });
				// Also check for lookup field naming convention
				map.set(`_${order.attribute}_value`, { direction });
				// Also map by alias if this attribute has one
				const alias = attrToAliasMap.get(order.attribute);
				if (alias) {
					map.set(alias, { direction });
				}
			}
		});

		// Process link-entity orders (orders defined inside link-entity)
		const processLinkOrders = (links: typeof fetchQuery.entity.links) => {
			links?.forEach((link) => {
				const alias = link.alias || link.name;
				link.orders?.forEach((order: OrderNode) => {
					const direction: SortDirection = order.descending ? "descending" : "ascending";
					// Link-entity attribute columns appear as alias.attributeName or entityname.attributeName
					map.set(`${alias}.${order.attribute}`, { direction, entityName: alias });
					map.set(`${link.name}.${order.attribute}`, { direction, entityName: alias });
				});
				// Recurse into nested link-entities
				if (link.links) {
					processLinkOrders(link.links);
				}
			});
		};
		processLinkOrders(fetchQuery.entity.links);

		return map;
	}, [fetchQuery]);

	// Get primary ID column name (typically {entity}id) from the FetchXML root entity
	const primaryIdColumn = useMemo(() => {
		// Use the root entity from FetchXML if available, otherwise fall back to result entity
		const entityName = fetchQuery?.entity?.name || result?.entityLogicalName;
		if (!entityName) return null;
		return `${entityName}id`;
	}, [fetchQuery, result]);

	// Generate stable row IDs using primary key GUID
	const getRowId = useCallback(
		(item: Record<string, unknown>) => {
			// CRITICAL: Use the GUID from the primary ID column for stable row identity
			if (primaryIdColumn && item[primaryIdColumn]) {
				const id = item[primaryIdColumn];
				// Ensure it's a valid GUID string (format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)
				if (typeof id === "string" && id.length > 0) {
					return id;
				}
			}
			// Fallback: this should NOT happen if FetchXML includes the primary ID
			// Use index as last resort (will break selection if data changes)
			console.warn("Row without primary ID found - selection may not work correctly");
			return `row_${Math.random()}`;
		},
		[primaryIdColumn]
	);

	// Build set of requested attributes from FetchXML query
	const requestedAttributes = useMemo(() => {
		const attrs = new Set<string>();
		if (!fetchQuery) return attrs;

		// Collect from root entity attributes
		const collectAttributes = (attributes: Array<AttributeNode>) => {
			attributes.forEach((attr) => {
				// Add both the original attribute name and alias (if exists)
				attrs.add(attr.name);
				if (attr.alias) {
					attrs.add(attr.alias);
				}
			});
		};

		if (fetchQuery.entity.attributes) {
			collectAttributes(fetchQuery.entity.attributes);
		}

		// Collect from link-entities recursively
		const collectFromLinks = (links: Array<any>, linkAlias?: string) => {
			links.forEach((link) => {
				const currentAlias = link.alias || linkAlias;
				if (link.attributes) {
					link.attributes.forEach((attr: AttributeNode) => {
						// For link-entity attributes, the column name in results may be:
						// 1. attr.alias (if explicitly aliased)
						// 2. linkAlias.attributeName (if link has alias)
						// 3. entityname.attributeName (Dataverse default)
						attrs.add(attr.name); // Base attribute name
						if (attr.alias) {
							attrs.add(attr.alias); // Explicit alias
						}
						if (currentAlias) {
							attrs.add(`${currentAlias}.${attr.name}`); // link-entity prefixed
						}
						if (link.name) {
							attrs.add(`${link.name}.${attr.name}`); // entity name prefixed
						}
					});
				}
				if (link.links) {
					collectFromLinks(link.links, currentAlias);
				}
			});
		};

		if (fetchQuery.entity.links) {
			collectFromLinks(fetchQuery.entity.links);
		}

		return attrs;
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

		// Collect from link-entities recursively
		const collectFromLinks = (
			links: import("../../model/nodes").LinkEntityNode[],
			parentLookupAttr?: string
		) => {
			links.forEach((link) => {
				const linkAlias = link.alias || link.name;
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

						// Also map entityname.attributeName variant
						if (link.name !== linkAlias) {
							map.set(`${link.name}.${attr.name}`, {
								attributeName: attr.name,
								entityName: link.name,
								linkEntityAlias: linkAlias,
								lookupAttribute: lookupAttr,
							});
						}
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
	const columns = useMemo(() => {
		if (!result || result.columns.length === 0) return [];

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

		return displayableColumns.flatMap((col) => {
			// Get column display info from our comprehensive map
			const displayInfo = columnDisplayMap.get(col);

			// Helper to get attribute metadata from multi-entity map
			const getAttributeMetadata = (
				entityName: string,
				attrName: string
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
				// Link-entity column: "Attribute Display Name (Lookup Attribute Display Name)"
				// e.g., "Email (Primary Contact)" where Email is from contact entity, Primary Contact is the lookup
				const attrDisplayName =
					attribute?.DisplayName?.UserLocalizedLabel?.Label || displayInfo.attributeName;

				// Get the lookup attribute display name from root entity metadata
				let lookupDisplayName = displayInfo.lookupAttribute || displayInfo.linkEntityAlias;
				if (displayInfo.lookupAttribute) {
					const rootEntityName = fetchQuery?.entity?.name;
					if (rootEntityName) {
						const lookupAttr = getAttributeMetadata(rootEntityName, displayInfo.lookupAttribute);
						if (lookupAttr?.DisplayName?.UserLocalizedLabel?.Label) {
							lookupDisplayName = lookupAttr.DisplayName.UserLocalizedLabel.Label;
						}
					}
				}

				displayName = `${attrDisplayName} (${lookupDisplayName})`;
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

			// For "both" mode with formatted values, create two columns
			if (valueMode === "both" && hasFormattedValues) {
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
							<TableCellLayout>
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
							<TableCellLayout>
								{rawValue === null || rawValue === undefined ? (
									<span>—</span>
								) : (
									<span>{String(rawValue)}</span>
								)}
							</TableCellLayout>
						);
					},
				});

				return [formattedColumn, rawColumn];
			}

			// Standard single column (formatted or raw mode)
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
							<TableCellLayout>
								{rawValue === null || rawValue === undefined ? (
									<span>—</span>
								) : (
									<span>{String(rawValue)}</span>
								)}
							</TableCellLayout>
						);
					} else {
						// Default: formatted (use cell renderer)
						return (
							<TableCellLayout>
								{getCellRenderer(attribute?.AttributeType, rawValue, formattedValue, attribute)}
							</TableCellLayout>
						);
					}
				},
			});
		});
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
	]);

	// Selection handlers
	const handleSelectionChange = useCallback(
		(_e: unknown, data: { selectedItems: Set<string | number> }) => {
			setSelectedItems(data.selectedItems);
			onSelectedCountChange?.(data.selectedItems.size);
			// Convert Set to array of string IDs for the callback
			const recordIds = Array.from(data.selectedItems).map(String);
			onSelectionChange?.(recordIds);
		},
		[onSelectedCountChange, onSelectionChange]
	);

	// Sort change handler: notifies parent to update FetchXML orders
	// Shift+click adds/toggles column in multi-sort, regular click replaces all
	// We calculate direction ourselves based on FetchXML state, not Fluent's sortDirection
	const handleSortChange = useCallback(
		(
			e: React.MouseEvent,
			data: { sortColumn: string | number | undefined; sortDirection: SortDirection }
		) => {
			if (!onSortChange) return;

			const columnId = String(data.sortColumn);
			const isMultiSort = e.shiftKey;

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

			onSortChange({
				columnId,
				direction,
				isMultiSort,
				entityName: sortInfo?.entityName,
			});
		},
		[onSortChange, sortStateMap]
	);

	// Row renderer function for virtualization with scrolling indicator support
	const renderRow: RowRenderer<Record<string, unknown>> = useCallback(
		({ item, rowId }, style, _index, isScrolling) => (
			<DataGridRow<Record<string, unknown>> key={rowId} style={style} className={styles.row}>
				{({ renderCell }) => (
					<DataGridCell focusMode="group">
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
		[styles.row]
	);

	// Build column sizing options from columnConfig or defaults
	const columnSizingOptions = useMemo(() => {
		const options: Record<string, { minWidth: number; defaultWidth: number; idealWidth?: number }> =
			{};

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
		for (const col of columns) {
			const columnId = String(col.columnId);
			const configWidth = configWidths.get(columnId);
			const width = configWidth ?? 150; // Default to 150 if not in config
			options[columnId] = {
				minWidth: 80,
				defaultWidth: width,
				idealWidth: width,
			};
		}

		return options;
	}, [columns, columnConfig]);

	// Handle column resize callback
	// Maps lookup column names back to config names for storage
	const handleColumnResize = useCallback(
		(
			_e: KeyboardEvent | TouchEvent | MouseEvent | undefined,
			data: { columnId: string | number; width: number }
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
		[onColumnResize]
	);

	// Infinite scroll: load more when user scrolls near bottom
	const handleItemsRendered = useCallback(
		({ visibleStopIndex }: { visibleStartIndex: number; visibleStopIndex: number }) => {
			if (!result || !onLoadMore || isLoadingMore) return;

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
		[result, onLoadMore, isLoadingMore]
	);

	// Note: Data is already sorted by Dataverse based on FetchXML orders
	// No local sorting needed - we display rows as returned

	if (isLoading) {
		return <div className={styles.emptyState}>Loading results...</div>;
	}

	if (!result) {
		return <div className={styles.emptyState}>Click Execute to run the query</div>;
	}

	if (result.rows.length === 0) {
		return <div className={styles.emptyState}>No records found</div>;
	}

	console.log("[ResultsGrid] Rendering grid:", {
		rowCount: result.rows.length,
		columnCount: columns.length,
		gridDimensions,
		primaryIdColumn,
		firstRowId: result.rows.length > 0 ? getRowId(result.rows[0]) : "none",
		columnIds: columns.slice(0, 5).map((c) => c.columnId),
		sortStateMapSize: sortStateMap.size,
	});

	return (
		<div className={styles.container}>
			<div className={styles.gridWrapper}>
				<div ref={containerRef} className={styles.gridContent}>
					<DataGrid
						aria-label={result?.entityLogicalName ?? "Results"}
						items={result.rows}
						columns={columns}
						sortable
						// Don't set sortState - we manage sort indicators ourselves in renderHeaderCell
						// This avoids double arrows (Fluent's + ours) and gives us control over multi-sort display
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
						<DataGridHeader ref={headerRef as any} style={{ paddingRight: scrollbarWidth }}>
							<DataGridRow style={{ minHeight: "43px", maxHeight: "43px" }}>
								{({ renderHeaderCell }) => (
									<DataGridHeaderCell className={styles.headerCell}>
										{renderHeaderCell()}
									</DataGridHeaderCell>
								)}
							</DataGridRow>
						</DataGridHeader>
						<DataGridBody<Record<string, unknown>>
							className={styles.body}
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
				</div>
				<div className={styles.infoBar}>
					<span>
						Rows: {result.rows.length}
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
		</div>
	);
}
