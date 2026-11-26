/**
 * Bulk Delete Dialog
 * Shows confirmation and progress for bulk delete operations
 * Supports both selected records and "all records from view" scenarios
 */

import { useCallback, useEffect, useState } from "react";
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
	Checkbox,
} from "@fluentui/react-components";
import { BarTitle } from "../../../../shared/components/BarTitle";
import {
	Warning20Regular,
	CheckmarkCircle20Regular,
	Copy20Regular,
	ErrorCircle20Regular,
} from "@fluentui/react-icons";
import { formatDataverseError } from "../../api/dataverseErrors";
import { openUrl } from "../../api/host";
import { useCopy } from "../../../../shared/feedback/useCopy";
import { TargetScope } from "../../../../shared/components/TargetScope";
import { defaultBulkDeleteJobName } from "./bulkDeleteJobName";
import { withReason } from "../../../../shared/components/WithReason";
import { useTargetConfirmation } from "../../../../shared/hooks/useTargetConfirmation";
import { useFocusWhen } from "../../../../shared/hooks/useFocusWhen";

const useStyles = makeStyles({
	// A grid item's min-width is auto, so the long warning text pushed the MessageBar body 1px past the
	// dialog and showed a horizontal scrollbar (#152)
	messageBarBody: {
		minWidth: 0,
	},
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
	/** Entity logical name (the confirmation names it with the environment, #69) */
	entityLogicalName?: string;
	/** Whether this is deleting ALL records from view (no selection) */
	isAllRecords?: boolean;
	/** Rows loaded in the grid (when isAllRecords=true); not the total, which isn't known */
	totalViewRecords?: number;
	/** Whether Dataverse said more records match than were loaded (when isAllRecords=true) */
	moreRecords?: boolean;
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
	entityLogicalName,
	isAllRecords = false,
	totalViewRecords,
	moreRecords,
	onClose,
	onConfirm,
}: BulkDeleteDialogProps) {
	const styles = useStyles();
	const [state, setState] = useState<DialogState>("confirm");
	const [jobName, setJobName] = useState(() => defaultBulkDeleteJobName(entityDisplayName));
	const [error, setError] = useState<string | null>(null);
	const [result, setResult] = useState<{ asyncOperationId: string; jobUrl: string } | null>(null);
	const copy = useCopy();
	const [confirmAllRecords, setConfirmAllRecords] = useState(false);
	// Type-to-confirm for Production, more than 100 records or an unknown total (#69)
	const scope = useTargetConfirmation();
	// Cancel first (owner, #279): after Fluent's own first focus, which a parent's effect follows
	const cancelRef = useFocusWhen(open && state === "confirm");
	// Focus follows the buttons: the running Cancel, then Close (#279)
	const runningRef = useFocusWhen(state === "submitting");
	const closeRef = useFocusWhen(state === "success" || state === "error");

	// The dialog stays mounted, so the default name is made each time it opens, for the entity and
	// time of this job (#182). A name typed while it's open is kept.
	useEffect(() => {
		if (open) setJobName(defaultBulkDeleteJobName(entityDisplayName));
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [open]);

	const handleConfirm = useCallback(async () => {
		setState("submitting");
		setError(null);
		try {
			const res = await onConfirm(jobName);
			setResult(res);
			setState("success");
		} catch (err) {
			setError(formatDataverseError(err, "Failed to submit bulk delete job")); // (#96)
			setState("error");
		}
	}, [onConfirm, jobName]);

	const handleClose = useCallback(() => {
		if (state !== "submitting") {
			setState("confirm");
			setError(null);
			setResult(null);
			setConfirmAllRecords(false);
			onClose();
		}
	}, [state, onClose]);

	const handleCopyUrl = useCallback(async () => {
		// The ToolBox clipboard first, then the browser's (#70); says so in a toast (#277)
		if (result?.jobUrl) await copy(result.jobUrl, "the job link");
	}, [result, copy]);

	const handleOpenJob = useCallback(() => {
		// In the connection's browser profile (#70)
		if (result?.jobUrl) {
			void openUrl(result.jobUrl);
		}
	}, [result]);

	const renderContent = () => {
		switch (state) {
			case "confirm":
				return (
					<>
						{/* Environment and scope (#69): with no selection, only the loaded count is known */}
						<TargetScope
							entityLogicalName={entityLogicalName}
							entityDisplayName={entityDisplayName}
							scope={
								isAllRecords
									? {
											kind: "allMatching",
											loadedCount: totalViewRecords,
											moreRecords,
										}
									: { kind: "records", count: recordCount }
							}
							onConfirmationChange={scope.onConfirmationChange}
						/>

						{isAllRecords ? (
							<>
								<MessageBar intent="error">
									<MessageBarBody className={styles.messageBarBody}>
										<BarTitle>Danger zone</BarTitle>
										This will delete ALL records matching your current FetchXML
										query!
									</MessageBarBody>
								</MessageBar>

								<div className={styles.dangerWarning}>
									<div className={styles.warningRow}>
										<ErrorCircle20Regular />
										<Text weight="semibold">
											You have not selected any records. This will delete ALL
											records returned from the view.
										</Text>
									</div>
									<Text
										style={{
											marginTop: tokens.spacingVerticalS,
											display: "block",
										}}
									>
										This is a destructive operation that cannot be undone. The
										bulk delete job will delete every record that matches your
										FetchXML query.
									</Text>
								</div>

								<Checkbox
									className={styles.confirmCheckbox}
									checked={confirmAllRecords}
									onChange={(_, data) =>
										setConfirmAllRecords(data.checked === true)
									}
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
									You are about to delete {recordCount}{" "}
									{entityDisplayName || "records"}. This will create a background
									job to process the deletion.
								</Text>
							</>
						)}

						<Field label="Job name" required>
							<Input
								value={jobName}
								onChange={(_, data) => setJobName(data.value)}
								placeholder="Enter a name for the bulk delete job"
							/>
						</Field>

						<Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
							You can follow the job's progress in Dataverse's system jobs.
						</Text>
					</>
				);

			case "submitting":
				return (
					<div className={styles.infoBox}>
						<div
							style={{
								display: "flex",
								alignItems: "center",
								gap: tokens.spacingHorizontalS,
							}}
						>
							<Spinner size="small" />
							<Text>Submitting the bulk delete job…</Text>
						</div>
					</div>
				);

			case "success":
				return (
					<>
						<div className={styles.successRow}>
							<CheckmarkCircle20Regular />
							<Text weight="semibold">Bulk delete job submitted.</Text>
						</div>

						<div className={styles.infoBox}>
							<Text>
								The bulk delete job has been submitted and will run in the
								background. You can track its progress using the link below.
							</Text>

							{result?.jobUrl && (
								<div className={styles.jobUrlRow}>
									<Link onClick={handleOpenJob}>Open job status</Link>
									<Button
										appearance="subtle"
										size="small"
										icon={<Copy20Regular />}
										onClick={handleCopyUrl}
									>
										Copy link
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
							Check your permissions and try again. You need the prvBulkDelete
							privilege to submit bulk delete jobs.
						</Text>
					</>
				);
		}
	};

	/** Why Submit can't be pressed yet (#279): what the dialog asks for, in its order */
	const submitReason = !jobName.trim()
		? "Enter a job name"
		: isAllRecords && !confirmAllRecords
			? "Confirm that this deletes all records"
			: scope.reason;

	const renderActions = () => {
		switch (state) {
			case "confirm":
				return (
					<>
						{/* The safe action is focused first (owner, #279) */}
						<Button appearance="secondary" ref={cancelRef} onClick={handleClose}>
							Cancel
						</Button>
						{withReason(
							<Button
								appearance="primary"
								onClick={handleConfirm}
								disabledFocusable={!!submitReason || !scope.confirmed}
							>
								Submit bulk delete
							</Button>,
							submitReason,
						)}
					</>
				);

			case "submitting":
				return withReason(
					<Button ref={runningRef} appearance="secondary" disabledFocusable>
						Cancel
					</Button>,
					"Wait for the job to be submitted",
				);

			case "success":
			case "error":
				return (
					<Button ref={closeRef} appearance="primary" onClick={handleClose}>
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
						? `Bulk delete all matching ${entityDisplayName || "records"}`
						: `Bulk delete ${recordCount} records`}
				</DialogTitle>
				<DialogBody>
					<DialogContent className={styles.content}>{renderContent()}</DialogContent>
				</DialogBody>
				<DialogActions>{renderActions()}</DialogActions>
			</DialogSurface>
		</Dialog>
	);
}
