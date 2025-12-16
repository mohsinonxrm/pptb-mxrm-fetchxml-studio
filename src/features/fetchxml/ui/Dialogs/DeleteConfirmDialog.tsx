/**
 * Delete Confirmation Dialog
 * Shows confirmation before deleting records
 * Supports batch delete with progress tracking and ETA
 */

import { useCallback, useState, useEffect } from "react";
import {
	Dialog,
	DialogSurface,
	DialogTitle,
	DialogBody,
	DialogContent,
	DialogActions,
	Button,
	Text,
	Spinner,
	makeStyles,
	tokens,
	ProgressBar,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
} from "@fluentui/react-components";
import {
	Warning20Regular,
	CheckmarkCircle20Regular,
	DismissCircle20Regular,
} from "@fluentui/react-icons";
import type { BatchDeleteProgress, BatchDeleteResult } from "../../api/pptbClient";

const useStyles = makeStyles({
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
	},
	warningRow: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorPaletteRedForeground1,
	},
	successRow: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorPaletteGreenForeground1,
	},
	recordInfo: {
		backgroundColor: tokens.colorNeutralBackground3,
		padding: tokens.spacingVerticalS,
		borderRadius: tokens.borderRadiusMedium,
	},
	errorMessage: {
		color: tokens.colorPaletteRedForeground1,
		marginTop: tokens.spacingVerticalS,
	},
	progressSection: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
	},
	progressStats: {
		display: "flex",
		justifyContent: "space-between",
		alignItems: "center",
		marginTop: tokens.spacingVerticalXS,
	},
	eta: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
	},
	resultSummary: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
	},
});

export interface DeleteConfirmDialogProps {
	/** Whether the dialog is open */
	open: boolean;
	/** Record display name or description */
	recordName?: string;
	/** Entity display name */
	entityDisplayName?: string;
	/** Number of records to delete (for multi-select) */
	recordCount: number;
	/** Called when dialog is closed */
	onClose: () => void;
	/** Called when delete is confirmed - returns batch result for batch deletes */
	onConfirm: () => Promise<BatchDeleteResult | void>;
	/** Progress callback for batch deletes */
	onProgress?: BatchDeleteProgress;
}

type DialogState = "confirm" | "deleting" | "success" | "partial" | "error";

export function DeleteConfirmDialog({
	open,
	recordName,
	entityDisplayName,
	recordCount,
	onClose,
	onConfirm,
	onProgress,
}: DeleteConfirmDialogProps) {
	const styles = useStyles();
	const [state, setState] = useState<DialogState>("confirm");
	const [error, setError] = useState<string | null>(null);
	const [result, setResult] = useState<BatchDeleteResult | null>(null);
	const [progress, setProgress] = useState<BatchDeleteProgress | null>(null);

	// Update progress from prop
	useEffect(() => {
		if (onProgress) {
			setProgress(onProgress);
		}
	}, [onProgress]);

	const handleConfirm = useCallback(async () => {
		setState("deleting");
		setError(null);
		setProgress(null);
		setResult(null);
		try {
			const res = await onConfirm();
			if (res) {
				// Batch delete result
				setResult(res);
				if (res.failed === 0) {
					setState("success");
				} else if (res.succeeded === 0) {
					setState("error");
					setError(`All ${res.failed} delete(s) failed`);
				} else {
					setState("partial");
				}
			} else {
				// Simple delete (void return)
				setState("success");
			}
		} catch (err) {
			setError(err instanceof Error ? err.message : "Failed to delete record(s)");
			setState("error");
		}
	}, [onConfirm]);

	const handleClose = useCallback(() => {
		if (state !== "deleting") {
			setState("confirm");
			setError(null);
			setResult(null);
			setProgress(null);
			onClose();
		}
	}, [state, onClose]);

	/**
	 * Intelligently format ETA based on duration:
	 * - < 1 minute: show seconds
	 * - < 1 hour: show minutes and seconds
	 * - < 12 hours: show hours and minutes
	 * - >= 12 hours: show estimated completion date/time
	 */
	const formatEta = (seconds: number): string => {
		if (seconds < 60) {
			return `~${seconds}s remaining`;
		}

		const minutes = Math.floor(seconds / 60);
		const hours = Math.floor(minutes / 60);

		if (minutes < 60) {
			const secs = seconds % 60;
			return `~${minutes}m ${secs}s remaining`;
		}

		if (hours < 12) {
			const mins = minutes % 60;
			return `~${hours}h ${mins}m remaining`;
		}

		// >= 12 hours: show estimated completion date/time
		const completionDate = new Date(Date.now() + seconds * 1000);
		const isToday = completionDate.toDateString() === new Date().toDateString();
		const isTomorrow =
			completionDate.toDateString() === new Date(Date.now() + 86400000).toDateString();

		const timeStr = completionDate.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

		if (isToday) {
			return `Est. completion: today at ${timeStr}`;
		} else if (isTomorrow) {
			return `Est. completion: tomorrow at ${timeStr}`;
		} else {
			const dateStr = completionDate.toLocaleDateString([], { month: "short", day: "numeric" });
			return `Est. completion: ${dateStr} at ${timeStr}`;
		}
	};

	const isSingle = recordCount === 1;
	const title = isSingle ? "Delete Record" : `Delete ${recordCount} Records`;
	const message = isSingle
		? `Are you sure you want to delete this ${entityDisplayName || "record"}?`
		: `Are you sure you want to delete ${recordCount} ${entityDisplayName || "records"}?`;

	const renderContent = () => {
		switch (state) {
			case "confirm":
				return (
					<>
						<div className={styles.warningRow}>
							<Warning20Regular />
							<Text weight="semibold">This action cannot be undone.</Text>
						</div>

						<Text>{message}</Text>

						{isSingle && recordName && (
							<div className={styles.recordInfo}>
								<Text weight="semibold">{recordName}</Text>
							</div>
						)}
					</>
				);

			case "deleting":
				return (
					<div className={styles.progressSection}>
						<MessageBar intent="info">
							<MessageBarBody>
								<MessageBarTitle>Deleting records...</MessageBarTitle>
								{progress && (
									<>
										Processing batch {progress.batchesCompleted} of {progress.totalBatches}
									</>
								)}
							</MessageBarBody>
						</MessageBar>

						{progress && (
							<>
								<ProgressBar value={progress.completed / progress.total} max={1} />
								<div className={styles.progressStats}>
									<Text size={200}>
										{progress.completed} of {progress.total} ({progress.succeeded} succeeded,{" "}
										{progress.failed} failed)
									</Text>
									{progress.estimatedSecondsRemaining !== undefined &&
										progress.estimatedSecondsRemaining > 0 && (
											<Text className={styles.eta}>
												{formatEta(progress.estimatedSecondsRemaining)}
											</Text>
										)}
								</div>
							</>
						)}

						{!progress && (
							<div
								style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS }}
							>
								<Spinner size="small" />
								<Text>Deleting...</Text>
							</div>
						)}
					</div>
				);

			case "success":
				return (
					<>
						<div className={styles.successRow}>
							<CheckmarkCircle20Regular />
							<Text weight="semibold">
								{result
									? `Successfully deleted ${result.succeeded} record(s)!`
									: "Record(s) deleted successfully!"}
							</Text>
						</div>
					</>
				);

			case "partial":
				return (
					<>
						<MessageBar intent="warning">
							<MessageBarBody>
								<MessageBarTitle>Partial Success</MessageBarTitle>
								{result && (
									<>
										{result.succeeded} of {result.succeeded + result.failed} records deleted.
										{result.failed} failed.
									</>
								)}
							</MessageBarBody>
						</MessageBar>

						{result && result.errors.length > 0 && (
							<div className={styles.resultSummary}>
								<Text weight="semibold" size={200}>
									Errors:
								</Text>
								{result.errors.slice(0, 5).map((err, i) => (
									<Text key={i} size={200} className={styles.errorMessage}>
										{err}
									</Text>
								))}
								{result.errors.length > 5 && (
									<Text size={200} className={styles.errorMessage}>
										...and {result.errors.length - 5} more errors
									</Text>
								)}
							</div>
						)}
					</>
				);

			case "error":
				return (
					<>
						<div className={styles.warningRow}>
							<DismissCircle20Regular />
							<Text weight="semibold">Delete Failed</Text>
						</div>

						<Text className={styles.errorMessage}>{error}</Text>

						{result && result.errors.length > 0 && (
							<div className={styles.resultSummary}>
								<Text weight="semibold" size={200}>
									Errors:
								</Text>
								{result.errors.slice(0, 5).map((err, i) => (
									<Text key={i} size={200} className={styles.errorMessage}>
										{err}
									</Text>
								))}
								{result.errors.length > 5 && (
									<Text size={200} className={styles.errorMessage}>
										...and {result.errors.length - 5} more errors
									</Text>
								)}
							</div>
						)}
					</>
				);
		}
	};

	const renderActions = () => {
		switch (state) {
			case "confirm":
				return (
					<>
						<Button appearance="secondary" onClick={handleClose}>
							Cancel
						</Button>
						<Button appearance="primary" onClick={handleConfirm}>
							Delete
						</Button>
					</>
				);

			case "deleting":
				return (
					<Button appearance="secondary" disabled>
						Cancel
					</Button>
				);

			case "success":
			case "partial":
			case "error":
				return (
					<Button appearance="primary" onClick={handleClose}>
						Close
					</Button>
				);
		}
	};

	return (
		<Dialog open={open} onOpenChange={(_, data) => !data.open && handleClose()}>
			<DialogSurface>
				<DialogTitle>{title}</DialogTitle>
				<DialogBody>
					<DialogContent className={styles.content}>{renderContent()}</DialogContent>
				</DialogBody>
				<DialogActions>{renderActions()}</DialogActions>
			</DialogSurface>
		</Dialog>
	);
}
