/**
 * DataGrid for displaying FetchXML query results with virtualization
 * Uses rich cell renderers based on attribute metadata for Power Apps-like experience
 * Displays formatted values from OData annotations when available
 * Uses @fluentui-contrib/react-data-grid-react-window for 1D virtualization with multi-select and resizable columns
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
import type { AttributeMetadata } from "../../api/pptbClient";
import { getCellRenderer } from "./DataGridCellRenderers";
import { getFormattedValue, filterDisplayableColumns } from "./FormattedValueUtils";
import type { FetchNode, AttributeNode } from "../../model/nodes";

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
	attributeMetadata?: Map<string, AttributeMetadata>;
	fetchQuery?: FetchNode | null; // For extracting aliases
	onSelectedCountChange?: (count: number) => void;
}

export function ResultsGrid({
	result,
	isLoading,
	attributeMetadata,
	fetchQuery,
	onSelectedCountChange,
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

	// Build alias map from FetchXML query structure
	const aliasMap = useMemo(() => {
		const map = new Map<string, string>();
		if (!fetchQuery) return map;

		// Collect aliases from root entity attributes
		const collectAttributeAliases = (attributes: Array<AttributeNode>) => {
			attributes.forEach((attr) => {
				if (attr.alias) {
					map.set(attr.alias, attr.name); // Map alias -> original attribute name
				}
			});
		};

		// Collect from root entity
		if (fetchQuery.entity.attributes) {
			collectAttributeAliases(fetchQuery.entity.attributes);
		}

		// Collect from link-entities recursively
		const collectFromLinks = (links: Array<any>, linkAlias?: string) => {
			links.forEach((link) => {
				const currentAlias = link.alias || linkAlias;
				if (link.attributes) {
					link.attributes.forEach((attr: AttributeNode) => {
						if (attr.alias) {
							map.set(attr.alias, attr.name); // Explicit alias -> attribute name
						}
						// Also map link-entity prefixed names
						if (currentAlias) {
							map.set(`${currentAlias}.${attr.name}`, attr.name);
						}
						if (link.name) {
							map.set(`${link.name}.${attr.name}`, attr.name);
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

		return map;
	}, [fetchQuery]);

	// Memoize column definitions with display names and formatted values
	// Memoize column definitions with display names and formatted values
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

		return displayableColumns.map((col) => {
			// Check if this column is an alias
			const originalAttributeName = aliasMap.get(col) || col;

			// For lookup columns (_attributename_value), try to get metadata for base attribute
			let attribute: AttributeMetadata | undefined;
			if (col.startsWith("_") && col.endsWith("_value")) {
				const baseAttr = col.slice(1, -6);
				attribute = attributeMetadata?.get(baseAttr);
			} else {
				attribute = attributeMetadata?.get(originalAttributeName) || attributeMetadata?.get(col);
			}

			// Determine display name: prefer alias from FetchXML, then metadata display name, then column name
			let displayName: string;
			if (aliasMap.has(col)) {
				// This IS an alias - use it as the display name
				displayName = col;
			} else if (attribute?.DisplayName?.UserLocalizedLabel?.Label) {
				// Use metadata display name
				displayName = attribute.DisplayName.UserLocalizedLabel.Label;
			} else {
				// Fall back to column name (cleaned up for lookups)
				if (col.startsWith("_") && col.endsWith("_value")) {
					displayName = col.slice(1, -6); // Remove _ prefix and _value suffix
				} else {
					displayName = col;
				}
			}

			return createTableColumn<Record<string, unknown>>({
				columnId: col,
				compare: (a, b) => {
					const aVal = String(getFormattedValue(a, col) ?? a[col] ?? "");
					const bVal = String(getFormattedValue(b, col) ?? b[col] ?? "");
					return aVal.localeCompare(bVal);
				},
				renderHeaderCell: () => displayName,
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
		});
	}, [result, attributeMetadata, aliasMap, fetchQuery, requestedAttributes]);

	// Selection handlers
	const handleSelectionChange = useCallback(
		(_e: unknown, data: { selectedItems: Set<string | number> }) => {
			setSelectedItems(data.selectedItems);
			onSelectedCountChange?.(data.selectedItems.size);
		},
		[onSelectedCountChange]
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
						resizableColumns
						resizableColumnsOptions={{ autoFitColumns: false }}
						columnSizingOptions={{
							...columns.reduce((acc, col) => {
								acc[col.columnId] = { minWidth: 150, defaultWidth: 200 };
								return acc;
							}, {} as Record<string, { minWidth: number; defaultWidth: number }>),
						}}
						selectionMode="multiselect"
						selectedItems={selectedItems}
						onSelectionChange={handleSelectionChange}
						getRowId={getRowId}
					>
						<DataGridHeader ref={headerRef as any} style={{ paddingRight: scrollbarWidth }}>
							<DataGridRow style={{ minHeight: "43px", maxHeight: "43px" }}>
								{({ renderHeaderCell }) => (
									<DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
								)}
							</DataGridRow>
						</DataGridHeader>
						<DataGridBody<Record<string, unknown>>
							className={styles.body}
							itemSize={ROW_HEIGHT}
							height={Math.max(0, gridDimensions.height - headerHeight)}
							listProps={{ useIsScrolling: true, className: styles.body }}
						>
							{renderRow}
						</DataGridBody>
					</DataGrid>
				</div>
				<div className={styles.infoBar}>
					<span>
						Rows: {result.rows.length}
						{result.executionTimeMs !== undefined && ` | Executed in ${result.executionTimeMs}ms`}
					</span>
					<span>Selected: {selectedItems.size}</span>
				</div>
			</div>
		</div>
	);
}
