/**
 * Command bar for DataGrid with row actions
 * Similar to Power Apps Model Driven Apps view command bar
 */

import { useMemo, useState, useCallback } from "react";
import {
	Toolbar,
	ToolbarButton,
	ToolbarDivider,
	makeStyles,
	Menu,
	MenuTrigger,
	MenuPopover,
	MenuList,
	MenuItem,
	Input,
	tokens,
} from "@fluentui/react-components";
import {
	Open20Regular,
	Link20Regular,
	CheckmarkCircle20Regular,
	DismissCircle20Regular,
	Delete20Regular,
	ArrowExport20Regular,
	ColumnTriple20Regular,
	Add20Regular,
	Search20Regular,
} from "@fluentui/react-icons";
import { ColumnConfigDialog } from "./ColumnConfigDialog";
import type { LayoutColumn } from "../../model/layoutxml";
import type { AttributeMetadata } from "../../api/pptbClient";

const useStyles = makeStyles({
	toolbar: {
		padding: "0",
	},
	menuSearch: {
		padding: "8px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
	},
	menuList: {
		maxHeight: "300px",
		overflowY: "auto",
	},
});

export interface CommandBarProps {
	selectedCount: number;
	onOpen?: () => void;
	onCopyUrl?: () => void;
	onActivate?: () => void;
	onDeactivate?: () => void;
	onDelete?: () => void;
	onExport?: () => void;
	entityName?: string;
	columns?: LayoutColumn[];
	onReorderColumns?: (columns: LayoutColumn[]) => void;
	/** Available attributes from metadata for adding columns */
	availableAttributes?: AttributeMetadata[];
	/** Currently selected attributes in the query */
	selectedAttributes?: string[];
	/** Callback when user wants to add a column */
	onAddColumn?: (attributeName: string) => void;
}

export function ResultsCommandBar({
	selectedCount,
	onOpen,
	onCopyUrl,
	onActivate,
	onDeactivate,
	onDelete,
	onExport,
	entityName,
	columns,
	onReorderColumns,
	availableAttributes,
	selectedAttributes,
	onAddColumn,
}: CommandBarProps) {
	const styles = useStyles();
	const hasSelection = selectedCount > 0;
	const singleSelection = selectedCount === 1;
	const [columnDialogOpen, setColumnDialogOpen] = useState(false);
	const [addColumnSearch, setAddColumnSearch] = useState("");

	const handleColumnsClick = useCallback(() => {
		setColumnDialogOpen(true);
	}, []);

	const handleColumnDialogClose = useCallback(() => {
		setColumnDialogOpen(false);
	}, []);

	const handleReorderColumns = useCallback(
		(reorderedColumns: LayoutColumn[]) => {
			onReorderColumns?.(reorderedColumns);
			setColumnDialogOpen(false);
		},
		[onReorderColumns]
	);

	// Filter attributes that aren't already in the query
	const addableAttributes = useMemo(() => {
		if (!availableAttributes) return [];
		const selectedSet = new Set(selectedAttributes || []);
		return availableAttributes
			.filter((attr) => !selectedSet.has(attr.LogicalName))
			.filter((attr) => {
				// Filter by search term
				if (!addColumnSearch) return true;
				const searchLower = addColumnSearch.toLowerCase();
				const displayName = attr.DisplayName?.UserLocalizedLabel?.Label || "";
				return (
					attr.LogicalName.toLowerCase().includes(searchLower) ||
					displayName.toLowerCase().includes(searchLower)
				);
			})
			.sort((a, b) => {
				const aName = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
				const bName = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
				return aName.localeCompare(bName);
			});
	}, [availableAttributes, selectedAttributes, addColumnSearch]);

	const handleAddColumnSelect = useCallback(
		(attributeName: string) => {
			onAddColumn?.(attributeName);
			setAddColumnSearch("");
		},
		[onAddColumn]
	);

	// Determine if entity supports statecode/statuscode (most common pattern)
	const supportsActivation = useMemo(() => {
		if (!entityName) return false;
		// Most entities support activation, but some don't (e.g., activitypointer, annotation)
		const noActivationEntities = [
			"activitypointer",
			"annotation",
			"note",
			"connection",
			"connectionrole",
			"savedquery",
			"userquery",
		];
		return !noActivationEntities.includes(entityName.toLowerCase());
	}, [entityName]);

	return (
		<Toolbar className={styles.toolbar} size="medium" aria-label="Record actions">
			<ToolbarButton
				appearance="subtle"
				icon={<Open20Regular />}
				disabled={!singleSelection}
				onClick={onOpen}
				aria-label="Open record"
			>
				Open
			</ToolbarButton>

			<ToolbarButton
				appearance="subtle"
				icon={<Link20Regular />}
				disabled={!singleSelection}
				onClick={onCopyUrl}
				aria-label="Copy record URL"
			>
				Copy URL
			</ToolbarButton>

			<ToolbarDivider />

			{supportsActivation && (
				<>
					<ToolbarButton
						appearance="subtle"
						icon={<CheckmarkCircle20Regular />}
						disabled={!hasSelection}
						onClick={onActivate}
						aria-label="Activate selected records"
					>
						Activate
					</ToolbarButton>

					<ToolbarButton
						appearance="subtle"
						icon={<DismissCircle20Regular />}
						disabled={!hasSelection}
						onClick={onDeactivate}
						aria-label="Deactivate selected records"
					>
						Deactivate
					</ToolbarButton>

					<ToolbarDivider />
				</>
			)}

			<ToolbarButton
				appearance="subtle"
				icon={<Delete20Regular />}
				disabled={!hasSelection}
				onClick={onDelete}
				aria-label="Delete selected records"
			>
				Delete
			</ToolbarButton>

			<ToolbarDivider />

			<ToolbarButton
				appearance="subtle"
				icon={<ArrowExport20Regular />}
				onClick={onExport}
				aria-label="Export to CSV"
			>
				Export
			</ToolbarButton>

			{columns && columns.length > 0 && (
				<>
					<ToolbarDivider />
					<ToolbarButton
						appearance="subtle"
						icon={<ColumnTriple20Regular />}
						onClick={handleColumnsClick}
						aria-label="Configure columns"
					>
						Columns
					</ToolbarButton>
				</>
			)}

			{availableAttributes && availableAttributes.length > 0 && onAddColumn && (
				<Menu>
					<MenuTrigger disableButtonEnhancement>
						<ToolbarButton appearance="subtle" icon={<Add20Regular />} aria-label="Add column">
							Add Column
						</ToolbarButton>
					</MenuTrigger>
					<MenuPopover>
						<div className={styles.menuSearch}>
							<Input
								contentBefore={<Search20Regular />}
								placeholder="Search attributes..."
								value={addColumnSearch}
								onChange={(_e, data) => setAddColumnSearch(data.value)}
								size="small"
							/>
						</div>
						<MenuList className={styles.menuList}>
							{addableAttributes.length === 0 ? (
								<MenuItem disabled>
									{addColumnSearch ? "No matching attributes" : "All attributes already added"}
								</MenuItem>
							) : (
								addableAttributes.slice(0, 50).map((attr) => (
									<MenuItem
										key={attr.LogicalName}
										onClick={() => handleAddColumnSelect(attr.LogicalName)}
									>
										{attr.DisplayName?.UserLocalizedLabel?.Label || attr.LogicalName}
									</MenuItem>
								))
							)}
							{addableAttributes.length > 50 && (
								<MenuItem disabled>
									...and {addableAttributes.length - 50} more (use search)
								</MenuItem>
							)}
						</MenuList>
					</MenuPopover>
				</Menu>
			)}

			{columns && columns.length > 0 && (
				<ColumnConfigDialog
					open={columnDialogOpen}
					columns={columns}
					onClose={handleColumnDialogClose}
					onReorder={handleReorderColumns}
				/>
			)}
		</Toolbar>
	);
}
