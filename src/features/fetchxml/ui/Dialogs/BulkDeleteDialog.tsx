/**
 * Bulk Delete Dialog
 * Shows confirmation and progress for bulk delete operations
 * Supports both selected records and "all records from view" scenarios
 */

import { useCallback, useState } from "react";
import {
	Dialog,
	DialogSurface,
	DialogTitle,
	DialogBody,
	DialogContent,
	DialogActions,
	Button,
	Text,
	Input,
	Spinner,
	Link,
	makeStyles,
	tokens,
	Field,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Checkbox,
} from "@fluentui/react-components";
import {
	Warning20Regular,
	CheckmarkCircle20Regular,
	Copy20Regular,
	ErrorCircle20Regular,
} from "@fluentui/react-icons";

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
	dangerWarning: {
		backgroundColor: tokens.colorPaletteRedBackground1,
		padding: tokens.spacingVerticalM,
		borderRadius: tokens.borderRadiusMedium,
		border: `1px solid ${tokens.colorPaletteRedBorder1}`,
	},
	successRow: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorPaletteGreenForeground1,
	},
	infoBox: {
		backgroundColor: tokens.colorNeutralBackground3,
		padding: tokens.spacingVerticalM,
		borderRadius: tokens.borderRadiusMedium,
	},
	jobUrlRow: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		marginTop: tokens.spacingVerticalS,
	},
	errorMessage: {
		color: tokens.colorPaletteRedForeground1,
		marginTop: tokens.spacingVerticalS,
	},
	confirmCheckbox: {
		marginTop: tokens.spacingVerticalS,
	},
});

export interface BulkDeleteDialogProps {
	/** Whether the dialog is open */
	open: boolean;
	/** Number of records to delete (0 means ALL from view) */
	recordCount: number;
	/** Entity display name */
	entityDisplayName?: string;
	/** Whether this is deleting ALL records from view (no selection) */
	isAllRecords?: boolean;
	/** Total records in view (when isAllRecords=true) */
	totalViewRecords?: number;
	/** Called when dialog is closed */
	onClose: () => void;
	/** Called when bulk delete is confirmed */
	onConfirm: (jobName: string) => Promise<{ asyncOperationId: string; jobUrl: string }>;
}

type DialogState = "confirm" | "submitting" | "success" | "error";

export function BulkDeleteDialog({
	open,
	recordCount,
	entityDisplayName,
	isAllRecords = false,
	totalViewRecords: _totalViewRecords, // Intentionally unused - we don't show count for "all records"
	onClose,
	onConfirm,
}: BulkDeleteDialogProps) {
	const styles = useStyles();
	const [state, setState] = useState<DialogState>("confirm");
	const [jobName, setJobName] = useState(
		`Bulk Delete ${entityDisplayName || "Records"} - ${new Date().toLocaleString()}`
	);
	const [error, setError] = useState<string | null>(null);
	const [result, setResult] = useState<{ asyncOperationId: string; jobUrl: string } | null>(null);
	const [copied, setCopied] = useState(false);
	const [confirmAllRecords, setConfirmAllRecords] = useState(false);

	const handleConfirm = useCallback(async () => {
		setState("submitting");
		setError(null);
		try {
			const res = await onConfirm(jobName);
			setResult(res);
			setState("success");
		} catch (err) {
			setError(err instanceof Error ? err.message : "Failed to submit bulk delete job");
			setState("error");
		}
	}, [onConfirm, jobName]);

	const handleClose = useCallback(() => {
		if (state !== "submitting") {
			setState("confirm");
			setError(null);
			setResult(null);
			setCopied(false);
			setConfirmAllRecords(false);
			onClose();
		}
	}, [state, onClose]);

	const handleCopyUrl = useCallback(async () => {
		if (result?.jobUrl) {
			try {
				await navigator.clipboard.writeText(result.jobUrl);
				setCopied(true);
				setTimeout(() => setCopied(false), 2000);
			} catch {
				// Fallback for older browsers
				const textArea = document.createElement("textarea");
				textArea.value = result.jobUrl;
				document.body.appendChild(textArea);
				textArea.select();
				document.execCommand("copy");
				document.body.removeChild(textArea);
				setCopied(true);
				setTimeout(() => setCopied(false), 2000);
			}
		}
	}, [result]);

	const handleOpenJob = useCallback(() => {
		if (result?.jobUrl) {
			window.open(result.jobUrl, "_blank");
		}
	}, [result]);

	const renderContent = () => {
		switch (state) {
			case "confirm":
				return (
					<>
						{isAllRecords ? (
							<>
								<MessageBar intent="error">
									<MessageBarBody>
										<MessageBarTitle>Danger Zone</MessageBarTitle>
										This will delete ALL records matching your current FetchXML query!
									</MessageBarBody>
								</MessageBar>

								<div className={styles.dangerWarning}>
									<div className={styles.warningRow}>
										<ErrorCircle20Regular />
										<Text weight="semibold">
											You have not selected any records. This will delete ALL records returned from
											the view.
										</Text>
									</div>
									<Text style={{ marginTop: tokens.spacingVerticalS, display: "block" }}>
										This is a destructive operation that cannot be undone. The bulk delete job will
										delete every record that matches your FetchXML query.
									</Text>
								</div>

								<Checkbox
									className={styles.confirmCheckbox}
									checked={confirmAllRecords}
									onChange={(_, data) => setConfirmAllRecords(data.checked === true)}
									label="I understand that this will delete ALL records from the view"
								/>
							</>
						) : (
							<>
								<div className={styles.warningRow}>
									<Warning20Regular />
									<Text weight="semibold">This action cannot be undone.</Text>
								</div>

								<Text>
									You are about to delete {recordCount} {entityDisplayName || "records"}. This will
									create a background job to process the deletion.
								</Text>
							</>
						)}

						<Field label="Job Name" required>
							<Input
								value={jobName}
								onChange={(_, data) => setJobName(data.value)}
								placeholder="Enter a name for the bulk delete job"
							/>
						</Field>

						<Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
							You can track the progress of the job in the System Jobs area.
						</Text>
					</>
				);

			case "submitting":
				return (
					<div className={styles.infoBox}>
						<div style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS }}>
							<Spinner size="small" />
							<Text>Submitting bulk delete job...</Text>
						</div>
					</div>
				);

			case "success":
				return (
					<>
						<div className={styles.successRow}>
							<CheckmarkCircle20Regular />
							<Text weight="semibold">Bulk delete job submitted successfully!</Text>
						</div>

						<div className={styles.infoBox}>
							<Text>
								The bulk delete job has been submitted and will run in the background. You can track
								its progress using the link below.
							</Text>

							{result?.jobUrl && (
								<div className={styles.jobUrlRow}>
									<Link onClick={handleOpenJob}>Open Job Status</Link>
									<Button
										appearance="subtle"
										size="small"
										icon={<Copy20Regular />}
										onClick={handleCopyUrl}
									>
										{copied ? "Copied!" : "Copy URL"}
									</Button>
								</div>
							)}

							{result?.asyncOperationId && (
								<Text
									size={200}
									style={{
										marginTop: tokens.spacingVerticalS,
										color: tokens.colorNeutralForeground3,
									}}
								>
									Job ID: {result.asyncOperationId}
								</Text>
							)}
						</div>
					</>
				);

			case "error":
				return (
					<>
						<div className={styles.warningRow}>
							<Warning20Regular />
							<Text weight="semibold">Failed to submit bulk delete job</Text>
						</div>

						<Text className={styles.errorMessage}>{error}</Text>

						<Text size={200}>
							Please check your permissions and try again. You need the prvBulkDelete privilege to
							submit bulk delete jobs.
						</Text>
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
						<Button
							appearance="primary"
							onClick={handleConfirm}
							disabled={!jobName.trim() || (isAllRecords && !confirmAllRecords)}
						>
							Submit Bulk Delete
						</Button>
					</>
				);

			case "submitting":
				return (
					<Button appearance="secondary" disabled>
						Cancel
					</Button>
				);

			case "success":
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
				<DialogTitle>
					{isAllRecords
						? `Bulk Delete ALL ${entityDisplayName || "Records"} from View`
						: `Bulk Delete ${recordCount} Records`}
				</DialogTitle>
				<DialogBody>
					<DialogContent className={styles.content}>{renderContent()}</DialogContent>
				</DialogBody>
				<DialogActions>{renderActions()}</DialogActions>
			</DialogSurface>
		</Dialog>
	);
}
