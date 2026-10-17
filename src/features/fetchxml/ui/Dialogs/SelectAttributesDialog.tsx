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
	Dropdown,
	Option,
	Switch,
	Text,
	Tooltip,
	makeStyles,
	mergeClasses,
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
	type TableColumnSizingOptions,
} from "@fluentui/react-components";
import type { AttributeMetadata } from "../../api/pptbClient";
import type { NodeId } from "../../model/nodes";

const ALL_TYPES = "All types";

const useStyles = makeStyles({
	// Wide enough for long display and logical names side by side (#52); never wider than the window
	surface: {
		width: "min(1000px, 94vw)",
		maxWidth: "min(1000px, 94vw)",
	},
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
		minHeight: "420px",
		maxHeight: "62vh",
	},
	filters: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalM,
		flexWrap: "wrap",
	},
	searchBox: {
		flex: 1,
		minWidth: "220px",
	},
	typeFilter: {
		minWidth: "150px",
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
	// One line per cell: long names end with an ellipsis and show in full in a tooltip (#52)
	truncate: {
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	mono: {
		fontFamily: tokens.fontFamilyMonospace,
		fontSize: tokens.fontSizeBase200,
	},
	count: {
		marginRight: "auto",
		alignSelf: "center",
		color: tokens.colorNeutralForeground3,
	},
});

interface AttributeRow {
	logicalName: string;
	displayName: string;
	dataType: string;
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

/** The attribute's type as people know it: AttributeTypeName tells Virtual types apart (Image, File, MultiSelectPicklist) */
const typeLabel = (attr: AttributeMetadata) =>
	(attr.AttributeTypeName?.Value ?? attr.AttributeType ?? "").replace(/Type$/, "") || "Unknown";

// Resizable columns (#52); the dialog's width fits these defaults
const COLUMN_SIZING: TableColumnSizingOptions = {
	displayName: { defaultWidth: 300, idealWidth: 300, minWidth: 120 },
	logicalName: { defaultWidth: 300, idealWidth: 300, minWidth: 120 },
	dataType: { defaultWidth: 140, idealWidth: 140, minWidth: 90 },
};

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
	const [typeFilter, setTypeFilter] = useState(ALL_TYPES);
	const [selectedOnly, setSelectedOnly] = useState(false);
	const [selectedRows, setSelectedRows] = useState<Set<string>>(new Set(selectedAttributeNames));

	// Reset selected rows and filters when dialog opens with new data
	useEffect(() => {
		if (open) {
			setSelectedRows(new Set(selectedAttributeNames));
			setSearchQuery("");
			setTypeFilter(ALL_TYPES);
			setSelectedOnly(false);
		}
	}, [open, selectedAttributeNames]);

	// Transform attributes to rows (already the attributes to offer: the "Advanced Find only" setting, #57)
	const allRows = useMemo<AttributeRow[]>(
		() =>
			attributes
				.map((attr) => ({
					logicalName: attr.LogicalName,
					displayName: attr.DisplayName?.UserLocalizedLabel?.Label || attr.LogicalName,
					dataType: typeLabel(attr),
				}))
				.sort((a, b) => a.displayName.localeCompare(b.displayName)),
		[attributes],
	);
	const types = useMemo(() => [...new Set(allRows.map((r) => r.dataType))].sort(), [allRows]);

	// Search (display or logical name), type and "Selected only"
	const filteredRows = useMemo(() => {
		const query = searchQuery.trim().toLowerCase();
		return allRows.filter(
			(row) =>
				(!query || row.logicalName.toLowerCase().includes(query) || row.displayName.toLowerCase().includes(query)) &&
				(typeFilter === ALL_TYPES || row.dataType === typeFilter) &&
				(!selectedOnly || selectedRows.has(row.logicalName)),
		);
	}, [allRows, searchQuery, typeFilter, selectedOnly, selectedRows]);
	// Attributes already in the query but not listed (e.g. hidden by "Advanced Find only") stay selected
	const listedSelectedCount = useMemo(() => allRows.filter((r) => selectedRows.has(r.logicalName)).length, [allRows, selectedRows]);

	// Define columns
	const columns = useMemo<TableColumnDefinition<AttributeRow>[]>(() => {
		const cell = (text: string, mono?: boolean) => (
			<TableCellLayout truncate>
				<Tooltip content={text} relationship="description">
					<span className={mono ? mergeClasses(styles.truncate, styles.mono) : styles.truncate}>{text}</span>
				</Tooltip>
			</TableCellLayout>
		);
		return [
			createTableColumn<AttributeRow>({
				columnId: "displayName",
				compare: (a, b) => a.displayName.localeCompare(b.displayName),
				renderHeaderCell: () => "Display name",
				renderCell: (item) => cell(item.displayName),
			}),
			createTableColumn<AttributeRow>({
				columnId: "logicalName",
				compare: (a, b) => a.logicalName.localeCompare(b.logicalName),
				renderHeaderCell: () => "Logical name",
				renderCell: (item) => cell(item.logicalName, true),
			}),
			createTableColumn<AttributeRow>({
				columnId: "dataType",
				compare: (a, b) => a.dataType.localeCompare(b.dataType),
				renderHeaderCell: () => "Type",
				renderCell: (item) => cell(item.dataType),
			}),
		];
	}, [styles.truncate, styles.mono]);

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
			<DialogSurface className={styles.surface} aria-describedby={undefined}>
				<DialogBody>
					<DialogTitle>Select attributes · {entityDisplayName}</DialogTitle>
					<DialogContent className={styles.content}>
						<div className={styles.filters}>
							<SearchBox
								className={styles.searchBox}
								placeholder="Search display or logical name"
								value={searchQuery}
								onChange={handleSearchChange}
							/>
							<Dropdown
								className={styles.typeFilter}
								aria-label="Filter by type"
								value={typeFilter}
								selectedOptions={[typeFilter]}
								onOptionSelect={(_, data) => setTypeFilter(data.optionValue ?? ALL_TYPES)}
							>
								{[ALL_TYPES, ...types].map((type) => (
									<Option key={type} value={type}>
										{type}
									</Option>
								))}
							</Dropdown>
							<Switch label="Selected only" checked={selectedOnly} onChange={(_, data) => setSelectedOnly(data.checked)} />
						</div>
						<div className={styles.dataGridContainer}>
							<DataGrid
								className={styles.dataGrid}
								items={filteredRows}
								columns={columns}
								sortable
								resizableColumns
								columnSizingOptions={COLUMN_SIZING}
								selectionMode="multiselect"
								selectedItems={selectedRows}
								onSelectionChange={handleSelectionChange}
								getRowId={(item) => item.logicalName}
								size="small"
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
												checkboxIndicator: { "aria-label": `Select ${item.displayName}` },
											}}
										>
											{({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
										</DataGridRow>
									)}
								</DataGridBody>
							</DataGrid>
						</div>
					</DialogContent>
					<DialogActions fluid>
						<Text className={styles.count} size={200}>
							{listedSelectedCount} of {allRows.length} selected
						</Text>
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
