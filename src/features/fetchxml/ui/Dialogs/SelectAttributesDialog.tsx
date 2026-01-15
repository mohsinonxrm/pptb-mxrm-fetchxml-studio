/**
 * Select Attributes Dialog
 * Allows bulk selection/deselection of attributes for an entity or link-entity
 */

import { useState, useMemo, useCallback, useEffect } from "react";
import {
	Dialog,
	DialogSurface,
	DialogBody,
	DialogTitle,
	DialogContent,
	DialogActions,
	Button,
	SearchBox,
	makeStyles,
	tokens,
	DataGrid,
	DataGridHeader,
	DataGridRow,
	DataGridHeaderCell,
	DataGridBody,
	DataGridCell,
	TableCellLayout,
	createTableColumn,
	type DataGridProps,
	type InputOnChangeData,
	type SearchBoxChangeEvent,
	type TableColumnDefinition,
} from "@fluentui/react-components";
import type { AttributeMetadata } from "../../api/pptbClient";
import type { NodeId } from "../../model/nodes";

const useStyles = makeStyles({
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
		minHeight: "400px",
		maxHeight: "600px",
	},
	searchBox: {
		width: "100%",
	},
	dataGridContainer: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		minHeight: "300px",
		overflow: "hidden",
	},
	dataGrid: {
		minWidth: "100%",
		display: "flex",
		flexDirection: "column",
		height: "100%",
	},
	dataGridBody: {
		overflow: "auto",
		flex: 1,
	},
});

interface AttributeRow {
	logicalName: string;
	displayName: string;
	dataType: string;
	isSelected: boolean;
}

interface SelectAttributesDialogProps {
	open: boolean;
	onOpenChange: (open: boolean) => void;
	entityDisplayName: string;
	nodeId: NodeId;
	attributes: AttributeMetadata[];
	selectedAttributeNames: string[];
	onApply: (nodeId: NodeId, selectedAttributes: string[]) => void;
}

export function SelectAttributesDialog({
	open,
	onOpenChange,
	entityDisplayName,
	nodeId,
	attributes,
	selectedAttributeNames,
	onApply,
}: SelectAttributesDialogProps) {
	const styles = useStyles();
	const [searchQuery, setSearchQuery] = useState("");
	const [selectedRows, setSelectedRows] = useState<Set<string>>(new Set(selectedAttributeNames));

	// Reset selected rows when dialog opens with new data
	useEffect(() => {
		if (open) {
			setSelectedRows(new Set(selectedAttributeNames));
			setSearchQuery("");
		}
	}, [open, selectedAttributeNames]);

	// Transform attributes to rows
	const allRows = useMemo<AttributeRow[]>(() => {
		return attributes
			.filter((attr) => attr.IsValidForAdvancedFind?.Value !== false)
			.map((attr) => ({
				logicalName: attr.LogicalName,
				displayName: attr.DisplayName?.UserLocalizedLabel?.Label || attr.LogicalName,
				dataType: attr.AttributeType || "Unknown",
				isSelected: selectedAttributeNames.includes(attr.LogicalName),
			}))
			.sort((a, b) => a.displayName.localeCompare(b.displayName));
	}, [attributes, selectedAttributeNames]);

	// Filter rows based on search query
	const filteredRows = useMemo(() => {
		if (!searchQuery.trim()) {
			return allRows;
		}

		const query = searchQuery.toLowerCase();
		return allRows.filter(
			(row) =>
				row.logicalName.toLowerCase().includes(query) ||
				row.displayName.toLowerCase().includes(query) ||
				row.dataType.toLowerCase().includes(query)
		);
	}, [allRows, searchQuery]);

	// Define columns
	const columns = useMemo<TableColumnDefinition<AttributeRow>[]>(
		() => [
			createTableColumn<AttributeRow>({
				columnId: "logicalName",
				compare: (a, b) => a.logicalName.localeCompare(b.logicalName),
				renderHeaderCell: () => "Logical Name",
				renderCell: (item) => <TableCellLayout>{item.logicalName}</TableCellLayout>,
			}),
			createTableColumn<AttributeRow>({
				columnId: "displayName",
				compare: (a, b) => a.displayName.localeCompare(b.displayName),
				renderHeaderCell: () => "Display Name",
				renderCell: (item) => <TableCellLayout>{item.displayName}</TableCellLayout>,
			}),
			createTableColumn<AttributeRow>({
				columnId: "dataType",
				compare: (a, b) => a.dataType.localeCompare(b.dataType),
				renderHeaderCell: () => "Data Type",
				renderCell: (item) => <TableCellLayout>{item.dataType}</TableCellLayout>,
			}),
		],
		[]
	);

	const handleSearchChange = useCallback((_ev: SearchBoxChangeEvent, data: InputOnChangeData) => {
		setSearchQuery(data.value);
	}, []);

	const handleSelectionChange: DataGridProps["onSelectionChange"] = useCallback(
		(_e: unknown, data: { selectedItems: Iterable<unknown> }) => {
			setSelectedRows(data.selectedItems as Set<string>);
		},
		[]
	);

	const handleApply = useCallback(() => {
		const selectedAttributes = Array.from(selectedRows);
		onApply(nodeId, selectedAttributes);
		onOpenChange(false);
	}, [selectedRows, nodeId, onApply, onOpenChange]);

	const handleCancel = useCallback(() => {
		onOpenChange(false);
	}, [onOpenChange]);

	return (
		<Dialog open={open} onOpenChange={(_event, data) => onOpenChange(data.open)}>
			<DialogSurface aria-describedby={undefined}>
				<DialogBody>
					<DialogTitle>Select Attributes - {entityDisplayName}</DialogTitle>
					<DialogContent className={styles.content}>
						<SearchBox
							className={styles.searchBox}
							placeholder="Search attributes..."
							value={searchQuery}
							onChange={handleSearchChange}
						/>
						<div className={styles.dataGridContainer}>
							<DataGrid
								className={styles.dataGrid}
								items={filteredRows}
								columns={columns}
								sortable
								selectionMode="multiselect"
								selectedItems={selectedRows}
								onSelectionChange={handleSelectionChange}
								getRowId={(item) => item.logicalName}
							>
								<DataGridHeader>
									<DataGridRow
										selectionCell={{
											checkboxIndicator: { "aria-label": "Select all attributes" },
										}}
									>
										{({ renderHeaderCell }) => (
											<DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
										)}
									</DataGridRow>
								</DataGridHeader>
								<DataGridBody<AttributeRow> className={styles.dataGridBody}>
									{({ item, rowId }) => (
										<DataGridRow<AttributeRow>
											key={rowId}
											selectionCell={{
												checkboxIndicator: { "aria-label": "Select attribute" },
											}}
										>
											{({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
										</DataGridRow>
									)}
								</DataGridBody>
							</DataGrid>
						</div>
					</DialogContent>
					<DialogActions>
						<Button appearance="primary" onClick={handleApply}>
							Apply
						</Button>
						<Button appearance="secondary" onClick={handleCancel}>
							Cancel
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
