/**
 * DataGrid for displaying FetchXML query results
 * Uses rich cell renderers based on attribute metadata for Power Apps-like experience
 * 
 * NOTE: Currently using non-virtualized DataGrid from @fluentui/react-components
 * For large datasets (>1000 rows), consider switching back to virtualized version:
 * - Import from: @fluentui-contrib/react-data-grid-react-window
 * - Requires: height calculation and DataGridBody props (itemSize, height, width)
 * - Benefits: Better performance with thousands of rows
 */

import { useMemo, useCallback, useEffect, useState } from "react";
import {
	TableCellLayout,
	createTableColumn,
	makeStyles,
	useScrollbarWidth,
	useFluent,
	tokens,
	DataGrid,
	DataGridHeader,
	DataGridHeaderCell,
	DataGridBody,
	DataGridRow,
	DataGridCell,
} from "@fluentui/react-components";
import type { AttributeMetadata } from "../../api/pptbClient";
import { getCellRenderer } from "./DataGridCellRenderers";
import { ResultsCommandBar } from "./ResultsCommandBar";
import { usePptbContext } from "../../../../shared/hooks/usePptbContext";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		backgroundColor: tokens.colorNeutralBackground1,
		position: "relative",
	},
	commandBarWrapper: {
		backgroundColor: tokens.colorNeutralBackground1,
		borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
		boxShadow: tokens.shadow4,
		zIndex: 1,
		position: "relative",
	},
	gridWrapper: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		backgroundColor: tokens.colorNeutralBackground1,
	},
	gridContent: {
		flex: 1,
		overflow: "auto",
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
	},
	emptyState: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase300,
	},
});

export interface QueryResult {
	columns: string[];
	rows: Record<string, unknown>[];
	totalRecordCount?: number;
	moreRecords?: boolean;
	pagingCookie?: string;
	entityLogicalName?: string; // NEW: for loading attribute metadata
}

interface ResultsGridProps {
	result: QueryResult | null;
	isLoading?: boolean;
	attributeMetadata?: Map<string, AttributeMetadata>; // NEW: attribute metadata map
}

export function ResultsGrid({ result, isLoading, attributeMetadata }: ResultsGridProps) {
	const styles = useStyles();
	const { targetDocument } = useFluent();
	const pptbContext = usePptbContext();
	const scrollbarWidth = useScrollbarWidth({ targetDocument });
	const [selectedItems, setSelectedItems] = useState<Set<string | number>>(new Set());

	// Clear selection when results change
	useEffect(() => {
		setSelectedItems(new Set());
	}, [result]);

	// Handle selection change from DataGrid
	const onSelectionChange = useCallback((_e: unknown, data: { selectedItems: Set<string | number> }) => {
		setSelectedItems(data.selectedItems);
	}, []);

	// Get primary ID column name (typically {entity}id)
	const primaryIdColumn = useMemo(() => {
		if (!result?.entityLogicalName) return null;
		return `${result.entityLogicalName}id`;
	}, [result]);

	// Get selected record(s)
	const selectedRecords = useMemo(() => {
		if (!result) return [];
		return result.rows.filter((_, index) => selectedItems.has(index));
	}, [result, selectedItems]);

	// Command bar handlers
	const handleOpen = useCallback(() => {
		if (selectedRecords.length !== 1 || !pptbContext.environmentUrl || !result?.entityLogicalName || !primaryIdColumn) {
			return;
		}

		const record = selectedRecords[0];
		const recordId = record[primaryIdColumn] as string;
		if (!recordId) return;

		// Construct Dataverse record URL
		const url = `${pptbContext.environmentUrl}/main.aspx?etn=${result.entityLogicalName}&id=${recordId}&pagetype=entityrecord`;
		window.open(url, "_blank");
	}, [selectedRecords, pptbContext.environmentUrl, result, primaryIdColumn]);

	const handleCopyUrl = useCallback(async () => {
		if (selectedRecords.length !== 1 || !pptbContext.environmentUrl || !result?.entityLogicalName || !primaryIdColumn) {
			return;
		}

		const record = selectedRecords[0];
		const recordId = record[primaryIdColumn] as string;
		if (!recordId) return;

		const url = `${pptbContext.environmentUrl}/main.aspx?etn=${result.entityLogicalName}&id=${recordId}&pagetype=entityrecord`;
		
		try {
			await navigator.clipboard.writeText(url);
			console.log("✅ URL copied to clipboard:", url);
		} catch (error) {
			console.error("Failed to copy URL:", error);
		}
	}, [selectedRecords, pptbContext.environmentUrl, result, primaryIdColumn]);

	const handleActivate = useCallback(() => {
		console.log("Activate not yet implemented. Selected records:", selectedRecords.length);
		// TODO: Implement with PPTB Dataverse API update
	}, [selectedRecords]);

	const handleDeactivate = useCallback(() => {
		console.log("Deactivate not yet implemented. Selected records:", selectedRecords.length);
		// TODO: Implement with PPTB Dataverse API update
	}, [selectedRecords]);

	const handleDelete = useCallback(() => {
		console.log("Delete not yet implemented. Selected records:", selectedRecords.length);
		// TODO: Implement with PPTB Dataverse API delete
	}, [selectedRecords]);

	const handleExport = useCallback(() => {
		console.log("Export not yet implemented");
		// TODO: Implement CSV export
	}, []);

	// Memoize columns with display names and rich renderers
	const columns = useMemo(
		() =>
			result
				? result.columns.map((col) => {
						const attribute = attributeMetadata?.get(col);
						const displayName = attribute?.DisplayName?.UserLocalizedLabel?.Label || col;

						return createTableColumn<Record<string, unknown>>({
							columnId: col,
							compare: (a, b) => {
								const aVal = String(a[col] ?? "");
								const bVal = String(b[col] ?? "");
								return aVal.localeCompare(bVal);
							},
							renderHeaderCell: () => displayName,
							renderCell: (item) => (
								<TableCellLayout>
									{getCellRenderer(attribute?.AttributeType, item[col], attribute)}
								</TableCellLayout>
							),
						});
				  })
				: [],
		[result, attributeMetadata]
	);

	// Memoize the row render function for performance
	const renderRow = useCallback(
		({ item, rowId }: { item: Record<string, unknown>; rowId: number | string }) => (
			<DataGridRow<Record<string, unknown>>
				key={rowId}
				selectionCell={{
					checkboxIndicator: { "aria-label": "Select row" },
				}}
			>
				{({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
			</DataGridRow>
		),
		[]
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

	return (
		<div className={styles.container}>
			<div className={styles.commandBarWrapper}>
				<ResultsCommandBar
					selectedCount={selectedItems.size}
					onOpen={handleOpen}
					onCopyUrl={handleCopyUrl}
					onActivate={handleActivate}
					onDeactivate={handleDeactivate}
					onDelete={handleDelete}
					onExport={handleExport}
					entityName={result.entityLogicalName}
				/>
			</div>
			<div className={styles.gridWrapper}>
				<div className={styles.gridContent}>
					<DataGrid
						items={result.rows}
						columns={columns}
						sortable
						resizableColumns
						focusMode="composite"
						selectionMode="multiselect"
						selectedItems={selectedItems}
						onSelectionChange={onSelectionChange}
					>
						<DataGridHeader style={{ paddingRight: scrollbarWidth }}>
							<DataGridRow
								selectionCell={{
									checkboxIndicator: { "aria-label": "Select all rows" },
								}}
							>
								{({ renderHeaderCell }) => (
									<DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
								)}
							</DataGridRow>
						</DataGridHeader>
						<DataGridBody<Record<string, unknown>>>
							{renderRow}
						</DataGridBody>
					</DataGrid>
				</div>
				<div className={styles.infoBar}>
					<span>Rows: {result.rows.length}</span>
					<span>Selected: {selectedItems.size}</span>
				</div>
			</div>
		</div>
	);
}
