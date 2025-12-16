/**
 * Column Configuration Dialog
 * Allows users to reorder columns by moving them up/down
 */

import { useState, useCallback, useEffect } from "react";
import {
	Dialog,
	DialogSurface,
	DialogTitle,
	DialogBody,
	DialogContent,
	DialogActions,
	Button,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { ArrowUp24Regular, ArrowDown24Regular } from "@fluentui/react-icons";
import type { LayoutColumn } from "../../model/layoutxml";

const useStyles = makeStyles({
	columnList: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
		maxHeight: "400px",
		overflowY: "auto",
		padding: tokens.spacingVerticalS,
		border: `1px solid ${tokens.colorNeutralStroke1}`,
		borderRadius: tokens.borderRadiusMedium,
	},
	columnItem: {
		display: "flex",
		alignItems: "center",
		justifyContent: "space-between",
		padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		borderRadius: tokens.borderRadiusSmall,
		"&:hover": {
			backgroundColor: tokens.colorNeutralBackground1Hover,
		},
	},
	columnName: {
		flex: 1,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	buttonGroup: {
		display: "flex",
		gap: tokens.spacingHorizontalXS,
	},
	indexBadge: {
		minWidth: "24px",
		textAlign: "center",
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
		marginRight: tokens.spacingHorizontalS,
	},
});

interface ColumnConfigDialogProps {
	/** Whether the dialog is open */
	open: boolean;
	/** Columns to display for reordering */
	columns: LayoutColumn[];
	/** Called when dialog is closed without applying */
	onClose: () => void;
	/** Called when Apply is clicked with the new column order */
	onReorder: (columns: LayoutColumn[]) => void;
}

export function ColumnConfigDialog({ open, columns, onClose, onReorder }: ColumnConfigDialogProps) {
	const styles = useStyles();
	const [localColumns, setLocalColumns] = useState<LayoutColumn[]>([]);

	// Initialize local state when dialog opens or columns change
	useEffect(() => {
		if (open && columns.length > 0) {
			setLocalColumns([...columns]);
		}
	}, [open, columns]);

	const moveColumn = useCallback((index: number, direction: "up" | "down") => {
		setLocalColumns((prev) => {
			const newColumns = [...prev];
			const targetIndex = direction === "up" ? index - 1 : index + 1;

			if (targetIndex < 0 || targetIndex >= newColumns.length) {
				return prev;
			}

			// Swap columns
			[newColumns[index], newColumns[targetIndex]] = [newColumns[targetIndex], newColumns[index]];
			return newColumns;
		});
	}, []);

	const handleApply = useCallback(() => {
		onReorder(localColumns);
	}, [localColumns, onReorder]);

	if (columns.length === 0) {
		return null;
	}

	return (
		<Dialog open={open} onOpenChange={(_e, data) => !data.open && onClose()}>
			<DialogSurface>
				<DialogBody>
					<DialogTitle>Configure Columns</DialogTitle>
					<DialogContent>
						<p
							style={{
								marginBottom: tokens.spacingVerticalM,
								color: tokens.colorNeutralForeground2,
							}}
						>
							Use the arrow buttons to reorder columns. Changes apply when you click Apply.
						</p>
						<div className={styles.columnList}>
							{localColumns.map((column, index) => (
								<div key={column.name} className={styles.columnItem}>
									<span className={styles.indexBadge}>{index + 1}</span>
									<span className={styles.columnName}>{column.displayName || column.name}</span>
									<span
										style={{
											color: tokens.colorNeutralForeground3,
											fontSize: tokens.fontSizeBase200,
											marginRight: tokens.spacingHorizontalS,
										}}
									>
										{column.width}px
									</span>
									<div className={styles.buttonGroup}>
										<Button
											appearance="subtle"
											size="small"
											icon={<ArrowUp24Regular />}
											disabled={index === 0}
											onClick={() => moveColumn(index, "up")}
											title="Move up"
										/>
										<Button
											appearance="subtle"
											size="small"
											icon={<ArrowDown24Regular />}
											disabled={index === localColumns.length - 1}
											onClick={() => moveColumn(index, "down")}
											title="Move down"
										/>
									</div>
								</div>
							))}
						</div>
					</DialogContent>
					<DialogActions>
						<Button appearance="secondary" onClick={onClose}>
							Cancel
						</Button>
						<Button appearance="primary" onClick={handleApply}>
							Apply
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
