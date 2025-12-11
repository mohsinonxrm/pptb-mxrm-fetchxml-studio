/**
 * Edit Columns Panel (Drawer)
 * Shows selected columns with reorder/remove options
 * Matches Power Apps Model-Driven Apps UX
 */

import { useState, useCallback, useEffect, useRef } from "react";
import {
	DrawerBody,
	DrawerHeader,
	DrawerHeaderTitle,
	OverlayDrawer,
	Button,
	makeStyles,
	tokens,
	Text,
	Toolbar,
	ToolbarButton,
} from "@fluentui/react-components";
import {
	Dismiss24Regular,
	ArrowUp20Regular,
	ArrowDown20Regular,
	Delete20Regular,
	Add20Regular,
	ArrowReset20Regular,
	ReOrder16Regular,
} from "@fluentui/react-icons";
import type { LayoutColumn } from "../../model/layoutxml";

const useStyles = makeStyles({
	drawer: {
		width: "400px",
	},
	headerActions: {
		display: "flex",
		gap: tokens.spacingHorizontalS,
	},
	columnList: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
		marginTop: tokens.spacingVerticalM,
	},
	columnItem: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		borderRadius: tokens.borderRadiusMedium,
		cursor: "grab",
		userSelect: "none",
		"&:hover": {
			backgroundColor: tokens.colorNeutralBackground1Hover,
			border: `1px solid ${tokens.colorNeutralStroke1Hover}`,
		},
	},
	columnItemSelected: {
		backgroundColor: tokens.colorNeutralBackground1Selected,
		border: `1px solid ${tokens.colorBrandStroke1}`,
	},
	columnItemDragging: {
		opacity: 0.5,
		cursor: "grabbing",
	},
	columnItemDragOver: {
		borderTop: `2px solid ${tokens.colorBrandStroke1}`,
		marginTop: "-2px",
	},
	dragHandle: {
		cursor: "grab",
		color: tokens.colorNeutralForeground3,
		display: "flex",
		alignItems: "center",
	},
	columnName: {
		flex: 1,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	columnActions: {
		display: "flex",
		gap: tokens.spacingHorizontalXXS,
	},
	footer: {
		display: "flex",
		justifyContent: "flex-end",
		gap: tokens.spacingHorizontalS,
		padding: tokens.spacingVerticalM,
		borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
		marginTop: "auto",
	},
	emptyState: {
		display: "flex",
		flexDirection: "column",
		alignItems: "center",
		justifyContent: "center",
		padding: tokens.spacingVerticalXXL,
		color: tokens.colorNeutralForeground3,
		textAlign: "center",
	},
});

export interface EditColumnsPanelProps {
	/** Whether the panel is open */
	open: boolean;
	/** Columns to display for editing */
	columns: LayoutColumn[];
	/** Entity display name for the title */
	entityDisplayName?: string;
	/** Called when panel is closed */
	onClose: () => void;
	/** Called when Apply is clicked with the updated columns */
	onApply: (columns: LayoutColumn[]) => void;
	/** Called when user wants to add columns */
	onAddColumns: () => void;
	/** Called to reset columns to default */
	onResetToDefault?: () => void;
}

export function EditColumnsPanel({
	open,
	columns,
	entityDisplayName,
	onClose,
	onApply,
	onAddColumns,
	onResetToDefault,
}: EditColumnsPanelProps) {
	const styles = useStyles();
	const [localColumns, setLocalColumns] = useState<LayoutColumn[]>([]);
	const [selectedIndex, setSelectedIndex] = useState<number | null>(null);
	const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
	const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);
	const dragNodeRef = useRef<HTMLDivElement | null>(null);

	// Initialize local state when panel opens or columns change
	useEffect(() => {
		if (open && columns.length > 0) {
			setLocalColumns([...columns]);
			setSelectedIndex(null);
		}
	}, [open, columns]);

	// Drag and drop handlers
	const handleDragStart = useCallback((e: React.DragEvent<HTMLDivElement>, index: number) => {
		setDraggedIndex(index);
		dragNodeRef.current = e.currentTarget;
		e.dataTransfer.effectAllowed = "move";
		e.dataTransfer.setData("text/plain", index.toString());
		// Add a slight delay to allow the drag image to be captured
		setTimeout(() => {
			if (dragNodeRef.current) {
				dragNodeRef.current.style.opacity = "0.5";
			}
		}, 0);
	}, []);

	const handleDragEnd = useCallback(() => {
		setDraggedIndex(null);
		setDragOverIndex(null);
		if (dragNodeRef.current) {
			dragNodeRef.current.style.opacity = "1";
			dragNodeRef.current = null;
		}
	}, []);

	const handleDragOver = useCallback(
		(e: React.DragEvent<HTMLDivElement>, index: number) => {
			e.preventDefault();
			e.dataTransfer.dropEffect = "move";
			if (draggedIndex !== null && draggedIndex !== index) {
				setDragOverIndex(index);
			}
		},
		[draggedIndex]
	);

	const handleDragLeave = useCallback(() => {
		setDragOverIndex(null);
	}, []);

	const handleDrop = useCallback(
		(e: React.DragEvent<HTMLDivElement>, targetIndex: number) => {
			e.preventDefault();
			if (draggedIndex === null || draggedIndex === targetIndex) {
				setDragOverIndex(null);
				return;
			}

			setLocalColumns((prev) => {
				const newColumns = [...prev];
				const [removed] = newColumns.splice(draggedIndex, 1);
				// Adjust target index if we removed from before it
				const adjustedTarget = draggedIndex < targetIndex ? targetIndex - 1 : targetIndex;
				newColumns.splice(adjustedTarget, 0, removed);
				return newColumns;
			});

			// Update selection to follow the moved item
			setSelectedIndex(draggedIndex < targetIndex ? targetIndex - 1 : targetIndex);
			setDraggedIndex(null);
			setDragOverIndex(null);
		},
		[draggedIndex]
	);

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
		// Update selection to follow the moved item
		setSelectedIndex((prev) =>
			prev === index ? (direction === "up" ? index - 1 : index + 1) : prev
		);
	}, []);

	const removeColumn = useCallback((index: number) => {
		setLocalColumns((prev) => prev.filter((_, i) => i !== index));
		setSelectedIndex(null);
	}, []);

	const handleApply = useCallback(() => {
		onApply(localColumns);
	}, [localColumns, onApply]);

	const handleColumnClick = useCallback((index: number) => {
		setSelectedIndex((prev) => (prev === index ? null : index));
	}, []);

	return (
		<OverlayDrawer
			open={open}
			onOpenChange={(_e, data) => !data.open && onClose()}
			position="end"
			size="medium"
			className={styles.drawer}
		>
			<DrawerHeader>
				<DrawerHeaderTitle
					action={
						<Button
							appearance="subtle"
							aria-label="Close"
							icon={<Dismiss24Regular />}
							onClick={onClose}
						/>
					}
				>
					Edit columns{entityDisplayName ? `: ${entityDisplayName}` : ""}
				</DrawerHeaderTitle>
			</DrawerHeader>
			<DrawerBody>
				<Toolbar size="small">
					<ToolbarButton icon={<Add20Regular />} onClick={onAddColumns} aria-label="Add columns">
						Add columns
					</ToolbarButton>
					{onResetToDefault && (
						<ToolbarButton
							icon={<ArrowReset20Regular />}
							onClick={onResetToDefault}
							aria-label="Reset to default"
						>
							Reset to default
						</ToolbarButton>
					)}
				</Toolbar>

				{localColumns.length === 0 ? (
					<div className={styles.emptyState}>
						<Text>No columns selected</Text>
						<Text size={200}>Click "Add columns" to add columns to the view</Text>
					</div>
				) : (
					<div className={styles.columnList}>
						{localColumns.map((column, index) => {
							const isDragging = draggedIndex === index;
							const isDragOver = dragOverIndex === index;
							return (
								<div
									key={column.name}
									className={`${styles.columnItem} ${
										selectedIndex === index ? styles.columnItemSelected : ""
									} ${isDragging ? styles.columnItemDragging : ""} ${
										isDragOver ? styles.columnItemDragOver : ""
									}`}
									onClick={() => handleColumnClick(index)}
									role="button"
									tabIndex={0}
									draggable
									onDragStart={(e) => handleDragStart(e, index)}
									onDragEnd={handleDragEnd}
									onDragOver={(e) => handleDragOver(e, index)}
									onDragLeave={handleDragLeave}
									onDrop={(e) => handleDrop(e, index)}
									onKeyDown={(e) => {
										if (e.key === "Enter" || e.key === " ") {
											handleColumnClick(index);
										}
									}}
								>
									<span className={styles.dragHandle}>
										<ReOrder16Regular />
									</span>
									<span className={styles.columnName}>{column.displayName || column.name}</span>
									<div className={styles.columnActions}>
										<Button
											appearance="subtle"
											size="small"
											icon={<ArrowUp20Regular />}
											disabled={index === 0}
											onClick={(e) => {
												e.stopPropagation();
												moveColumn(index, "up");
											}}
											title="Move up"
											aria-label="Move up"
										/>
										<Button
											appearance="subtle"
											size="small"
											icon={<ArrowDown20Regular />}
											disabled={index === localColumns.length - 1}
											onClick={(e) => {
												e.stopPropagation();
												moveColumn(index, "down");
											}}
											title="Move down"
											aria-label="Move down"
										/>
										<Button
											appearance="subtle"
											size="small"
											icon={<Delete20Regular />}
											onClick={(e) => {
												e.stopPropagation();
												removeColumn(index);
											}}
											title="Remove"
											aria-label="Remove column"
										/>
									</div>
								</div>
							);
						})}
					</div>
				)}

				<div className={styles.footer}>
					<Button appearance="secondary" onClick={onClose}>
						Cancel
					</Button>
					<Button appearance="primary" onClick={handleApply}>
						Apply
					</Button>
				</div>
			</DrawerBody>
		</OverlayDrawer>
	);
}
