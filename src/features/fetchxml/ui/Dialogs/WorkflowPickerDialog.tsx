/**
 * Workflow Picker Dialog
 * Allows user to select and run an on-demand workflow on selected records
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
	Radio,
	RadioGroup,
	ProgressBar,
	makeStyles,
	tokens,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
} from "@fluentui/react-components";
import { Play20Regular, CheckmarkCircle20Regular, Warning20Regular } from "@fluentui/react-icons";
import type { WorkflowInfo, WorkflowBatchProgress } from "../../api/pptbClient";

const useStyles = makeStyles({
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
	},
	workflowList: {
		maxHeight: "300px",
		overflowY: "auto",
		border: `1px solid ${tokens.colorNeutralStroke1}`,
		borderRadius: tokens.borderRadiusMedium,
		padding: tokens.spacingVerticalS,
	},
	workflowOption: {
		display: "flex",
		flexDirection: "column",
		padding: tokens.spacingVerticalXS,
	},
	workflowDescription: {
		marginLeft: "28px", // Align with radio label
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
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
	},
	eta: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
	},
	resultSection: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
	},
	successRow: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorPaletteGreenForeground1,
	},
	errorRow: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorPaletteRedForeground1,
	},
	emptyState: {
		display: "flex",
		flexDirection: "column",
		alignItems: "center",
		justifyContent: "center",
		padding: tokens.spacingVerticalXXL,
		color: tokens.colorNeutralForeground3,
	},
	loadingState: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		padding: tokens.spacingVerticalXL,
		gap: tokens.spacingHorizontalS,
	},
});

export interface WorkflowPickerDialogProps {
	/** Whether the dialog is open */
	open: boolean;
	/** Number of records selected */
	recordCount: number;
	/** Entity display name */
	entityDisplayName?: string;
	/** Pre-selected workflow (from menu selection) */
	preSelectedWorkflow?: WorkflowInfo;
	/** Called when dialog is closed */
	onClose: () => void;
	/** Called to fetch available workflows */
	onFetchWorkflows: () => Promise<WorkflowInfo[]>;
	/** Called when workflow execution is confirmed */
	onExecute: (
		workflow: WorkflowInfo,
		onProgress: (progress: WorkflowBatchProgress) => void
	) => Promise<{ succeeded: number; failed: number; errors: string[] }>;
}

type DialogState = "loading" | "select" | "executing" | "complete" | "error";

export function WorkflowPickerDialog({
	open,
	recordCount,
	entityDisplayName,
	preSelectedWorkflow,
	onClose,
	onFetchWorkflows,
	onExecute,
}: WorkflowPickerDialogProps) {
	const styles = useStyles();
	const [state, setState] = useState<DialogState>("loading");
	const [workflows, setWorkflows] = useState<WorkflowInfo[]>([]);
	const [selectedWorkflowId, setSelectedWorkflowId] = useState<string | null>(null);
	const [error, setError] = useState<string | null>(null);
	const [progress, setProgress] = useState<WorkflowBatchProgress | null>(null);
	const [result, setResult] = useState<{
		succeeded: number;
		failed: number;
		errors: string[];
	} | null>(null);

	// Load workflows when dialog opens
	useEffect(() => {
		if (open) {
			setState("loading");
			setSelectedWorkflowId(null);
			setError(null);
			setResult(null);
			setProgress(null);

			onFetchWorkflows()
				.then((wfs) => {
					setWorkflows(wfs);
					setState("select");
					// Auto-select: preSelectedWorkflow takes priority, then single workflow
					if (preSelectedWorkflow) {
						// Verify the pre-selected workflow exists in the fetched list
						const found = wfs.find((w) => w.workflowid === preSelectedWorkflow.workflowid);
						if (found) {
							setSelectedWorkflowId(found.workflowid);
						}
					} else if (wfs.length === 1) {
						setSelectedWorkflowId(wfs[0].workflowid);
					}
				})
				.catch((err) => {
					setError(err instanceof Error ? err.message : "Failed to load workflows");
					setState("error");
				});
		}
	}, [open, onFetchWorkflows, preSelectedWorkflow]);

	const handleExecute = useCallback(async () => {
		const workflow = workflows.find((w) => w.workflowid === selectedWorkflowId);
		if (!workflow) return;

		setState("executing");
		setProgress(null);
		setError(null);

		try {
			const res = await onExecute(workflow, (progressUpdate) => {
				// Progress now comes from executeWorkflowBatch with batch info and ETA
				setProgress(progressUpdate);
			});
			setResult(res);
			setState("complete");
		} catch (err) {
			setError(err instanceof Error ? err.message : "Failed to execute workflow");
			setState("error");
		}
	}, [workflows, selectedWorkflowId, onExecute]);

	const handleClose = useCallback(() => {
		if (state !== "executing") {
			onClose();
		}
	}, [state, onClose]);

	const selectedWorkflow = workflows.find((w) => w.workflowid === selectedWorkflowId);

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

	const renderContent = () => {
		switch (state) {
			case "loading":
				return (
					<div className={styles.loadingState}>
						<Spinner size="small" />
						<Text>Loading available workflows...</Text>
					</div>
				);

			case "select":
				if (workflows.length === 0) {
					return (
						<div className={styles.emptyState}>
							<Text weight="semibold">No workflows available</Text>
							<Text size={200}>
								There are no on-demand workflows configured for {entityDisplayName || "this entity"}
								.
							</Text>
						</div>
					);
				}

				return (
					<>
						<Text>
							Select a workflow to run on {recordCount} {recordCount === 1 ? "record" : "records"}:
						</Text>

						<div className={styles.workflowList}>
							<RadioGroup
								value={selectedWorkflowId || ""}
								onChange={(_, data) => setSelectedWorkflowId(data.value)}
							>
								{workflows.map((wf) => (
									<div key={wf.workflowid} className={styles.workflowOption}>
										<Radio value={wf.workflowid} label={wf.name} />
										{wf.description && (
											<Text className={styles.workflowDescription}>{wf.description}</Text>
										)}
									</div>
								))}
							</RadioGroup>
						</div>
					</>
				);

			case "executing":
				return (
					<div className={styles.progressSection}>
						<MessageBar intent="info">
							<MessageBarBody>
								<MessageBarTitle>Running workflow...</MessageBarTitle>
								Executing <strong>{selectedWorkflow?.name}</strong> on {recordCount} records
								{progress && ` (batch ${progress.batchesCompleted} of ${progress.totalBatches})`}
							</MessageBarBody>
						</MessageBar>
						{progress ? (
							<>
								<ProgressBar value={progress.completed} max={progress.total} thickness="large" />
								<div className={styles.progressStats}>
									<Text size={200}>
										{progress.completed} of {progress.total} completed ({progress.succeeded}{" "}
										succeeded, {progress.failed} failed)
									</Text>
									{progress.estimatedSecondsRemaining !== undefined &&
										progress.estimatedSecondsRemaining > 0 && (
											<Text className={styles.eta}>
												{formatEta(progress.estimatedSecondsRemaining)}
											</Text>
										)}
								</div>
							</>
						) : (
							<div
								style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS }}
							>
								<Spinner size="small" />
								<Text>Starting workflow execution...</Text>
							</div>
						)}
					</div>
				);

			case "complete":
				return (
					<div className={styles.resultSection}>
						{result && result.succeeded > 0 && (
							<div className={styles.successRow}>
								<CheckmarkCircle20Regular />
								<Text weight="semibold">
									{result.succeeded} record{result.succeeded !== 1 ? "s" : ""} processed
									successfully
								</Text>
							</div>
						)}

						{result && result.failed > 0 && (
							<>
								<div className={styles.errorRow}>
									<Warning20Regular />
									<Text weight="semibold">
										{result.failed} record{result.failed !== 1 ? "s" : ""} failed
									</Text>
								</div>
								{result.errors.length > 0 && (
									<MessageBar intent="error">
										<MessageBarBody>
											<MessageBarTitle>Errors:</MessageBarTitle>
											{result.errors.slice(0, 5).map((err, i) => (
												<Text key={i} size={200} block>
													{err}
												</Text>
											))}
											{result.errors.length > 5 && (
												<Text size={200}>...and {result.errors.length - 5} more</Text>
											)}
										</MessageBarBody>
									</MessageBar>
								)}
							</>
						)}
					</div>
				);

			case "error":
				return (
					<MessageBar intent="error">
						<MessageBarBody>
							<MessageBarTitle>Error</MessageBarTitle>
							{error}
						</MessageBarBody>
					</MessageBar>
				);
		}
	};

	const renderActions = () => {
		switch (state) {
			case "loading":
				return (
					<Button appearance="secondary" onClick={handleClose}>
						Cancel
					</Button>
				);

			case "select":
				return (
					<>
						<Button appearance="secondary" onClick={handleClose}>
							Cancel
						</Button>
						<Button
							appearance="primary"
							icon={<Play20Regular />}
							onClick={handleExecute}
							disabled={!selectedWorkflowId || workflows.length === 0}
						>
							Run Workflow
						</Button>
					</>
				);

			case "executing":
				return (
					<Button appearance="secondary" disabled>
						Running...
					</Button>
				);

			case "complete":
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
				<DialogTitle>Run Workflow</DialogTitle>
				<DialogBody>
					<DialogContent className={styles.content}>{renderContent()}</DialogContent>
				</DialogBody>
				<DialogActions>{renderActions()}</DialogActions>
			</DialogSurface>
		</Dialog>
	);
}
