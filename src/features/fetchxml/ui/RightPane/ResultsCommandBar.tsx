/**
 * Command bar for DataGrid with row actions
 * Similar to Power Apps Model Driven Apps view command bar
 */

import { useMemo } from "react";
import { Toolbar, ToolbarButton, ToolbarDivider, makeStyles } from "@fluentui/react-components";
import {
	Open20Regular,
	Link20Regular,
	CheckmarkCircle20Regular,
	DismissCircle20Regular,
	Delete20Regular,
	ArrowExport20Regular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
	toolbar: {
		padding: "0",
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
}: CommandBarProps) {
	const styles = useStyles();
	const hasSelection = selectedCount > 0;
	const singleSelection = selectedCount === 1;

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
		</Toolbar>
	);
}
