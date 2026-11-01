/**
 * Activate / Deactivate confirmation (#143). Reads the entity's state model from metadata and plans
 * the change (model/recordState.ts): simple entities get their one target; entities with more
 * states or status reasons get a status reason picker grouped by state, limited by status reason
 * transitions when the entity enforces them. Shows the environment and scope (#69), then progress
 * and per-record errors like delete.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	Dropdown,
	Field,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Option,
	OptionGroup,
	ProgressBar,
	Spinner,
	Text,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { CheckmarkCircle20Regular } from "@fluentui/react-icons";
import type { BatchDeleteProgress, BatchDeleteResult } from "../../api/pptbClient";
import { getEntityStateModel, getRecordStatuses } from "../../api/pptbClient";
import { formatDataverseError } from "../../api/dataverseErrors";
import { planStateChange, stateNote, type StateAction, type StatePlan, type StateTarget } from "../../model/recordState";
import { TargetScope } from "../../../../shared/components/TargetScope";
import { UndoNote } from "../../../../shared/components/UndoNote";

const useStyles = makeStyles({
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
	},
	row: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
	success: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorPaletteGreenForeground1,
	},
	errors: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
	},
	errorText: {
		color: tokens.colorPaletteRedForeground1,
	},
});

export interface SetStateDialogProps {
	open: boolean;
	action: StateAction;
	recordIds: string[];
	entityLogicalName?: string;
	entityDisplayName?: string;
	/** Primary key, to read the records' current status reasons when transitions are enforced */
	primaryIdAttribute?: string;
	onClose: () => void;
	/** Runs the change; resolves with per-record results */
	onExecute: (target: StateTarget, onProgress: (progress: BatchDeleteProgress) => void) => Promise<BatchDeleteResult>;
}

type Phase = "loading" | "confirm" | "running" | "done" | "error";

const targetKey = (target: StateTarget) => `${target.statecode}:${target.statuscode}`;

export function SetStateDialog({
	open,
	action,
	recordIds,
	entityLogicalName,
	entityDisplayName,
	primaryIdAttribute,
	onClose,
	onExecute,
}: SetStateDialogProps) {
	const styles = useStyles();
	const [phase, setPhase] = useState<Phase>("loading");
	const [plan, setPlan] = useState<StatePlan | null>(null);
	const [target, setTarget] = useState<StateTarget | null>(null);
	const [error, setError] = useState<string | null>(null);
	const [progress, setProgress] = useState<BatchDeleteProgress | null>(null);
	const [result, setResult] = useState<BatchDeleteResult | null>(null);
	// Type-to-confirm for Production or more than 100 records (#69)
	const [scopeConfirmed, setScopeConfirmed] = useState(false);

	// Read when planning, so a metadata reload while the result shows doesn't plan again
	const primaryIdRef = useRef(primaryIdAttribute);
	primaryIdRef.current = primaryIdAttribute;

	const verb = action === "activate" ? "Activate" : "Deactivate";
	const count = recordIds.length;

	// Plan when the dialog opens: the state model, then (if transitions are enforced) the records'
	// current status reasons
	useEffect(() => {
		if (!open || !entityLogicalName) return;
		let current = true;
		setPhase("loading");
		setPlan(null);
		setTarget(null);
		setError(null);
		setProgress(null);
		setResult(null);
		(async () => {
			const model = await getEntityStateModel(entityLogicalName);
			if (!model) return { kind: "none", reason: "This entity has no Status (statecode) column." } as StatePlan;
			const primaryId = primaryIdRef.current;
			const statuses =
				model.enforceTransitions && primaryId ? await getRecordStatuses(entityLogicalName, primaryId, recordIds) : undefined;
			return planStateChange(model, action, statuses, entityLogicalName);
		})()
			.then((next) => {
				if (!current) return;
				setPlan(next);
				if (next.kind === "simple") setTarget(next.target);
				if (next.kind === "choose") setTarget(next.defaultTarget);
				setPhase("confirm");
			})
			.catch((err: unknown) => {
				if (!current) return;
				setError(formatDataverseError(err, "Couldn't read the entity's status options"));
				setPhase("error");
			});
		return () => {
			current = false;
		};
	}, [open, entityLogicalName, recordIds, action]);

	const selectedChoice = useMemo(() => {
		if (plan?.kind !== "choose" || !target) return undefined;
		const group = plan.groups.find((g) => g.statecode === target.statecode);
		const choice = group?.choices.find((c) => c.statuscode === target.statuscode);
		return group && choice ? { group, choice } : undefined;
	}, [plan, target]);

	const handleRun = useCallback(async () => {
		if (!target) return;
		setPhase("running");
		setProgress(null);
		try {
			const res = await onExecute(target, setProgress);
			setResult(res);
			if (res.succeeded === 0 && res.failed > 0) {
				setError(`All ${res.failed} records failed.`);
				setPhase("error");
			} else {
				setPhase("done");
			}
		} catch (err) {
			setError(formatDataverseError(err, `${verb} failed`)); // (#96)
			setPhase("error");
		}
	}, [target, onExecute, verb]);

	const handleClose = useCallback(() => {
		if (phase !== "running") onClose();
	}, [phase, onClose]);

	const note = target ? stateNote(entityLogicalName, target.statecode) : undefined;

	const renderErrors = () =>
		result && result.errors.length > 0 ? (
			<div className={styles.errors}>
				{result.errors.slice(0, 5).map((err, i) => (
					<Text key={i} size={200} className={styles.errorText}>
						{err}
					</Text>
				))}
				{result.errors.length > 5 && (
					<Text size={200} className={styles.errorText}>
						...and {result.errors.length - 5} more
					</Text>
				)}
			</div>
		) : null;

	const renderContent = () => {
		switch (phase) {
			case "loading":
				return (
					<div className={styles.row}>
						<Spinner size="small" />
						<Text>Reading the status options...</Text>
					</div>
				);
			case "confirm":
				return (
					<>
						{/* Environment and scope (#69) */}
						<TargetScope
							entityLogicalName={entityLogicalName}
							entityDisplayName={entityDisplayName}
							scope={{ kind: "records", count }}
							onConfirmationChange={setScopeConfirmed}
						/>
						{plan && plan.kind !== "none" && (
							<UndoNote undo="possible">
								To undo, select the records and run {action === "activate" ? "Deactivate" : "Activate"}
								{plan.transitionsEnforced ? ", if this entity's status reason transitions allow it." : "."}
								{plan.kind === "choose" && " The previous status reason isn't restored, so set it again if needed."}
							</UndoNote>
						)}
						{plan?.kind === "none" && (
							<MessageBar intent="warning" layout="multiline">
								<MessageBarBody>
									<MessageBarTitle>Nothing to change</MessageBarTitle>
									{plan.reason}
								</MessageBarBody>
							</MessageBar>
						)}
						{plan?.kind === "simple" && (
							<Text>
								Sets Status to <strong>{plan.stateLabel}</strong> and Status Reason to{" "}
								<strong>{plan.statusLabel}</strong>.
							</Text>
						)}
						{plan?.kind === "choose" && (
							<>
								{plan.note && (
									<MessageBar intent="info" layout="multiline">
										<MessageBarBody>{plan.note}</MessageBarBody>
									</MessageBar>
								)}
								<Field label="Status reason" hint="Grouped by Status. The default for each Status is marked.">
									<Dropdown
										value={selectedChoice ? `${selectedChoice.group.label}: ${selectedChoice.choice.label}` : ""}
										selectedOptions={target ? [targetKey(target)] : []}
										onOptionSelect={(_, data) => {
											const [statecode, statuscode] = (data.optionValue ?? "").split(":").map(Number);
											if (Number.isFinite(statecode) && Number.isFinite(statuscode)) setTarget({ statecode, statuscode });
										}}
									>
										{plan.groups.map((group) => (
											<OptionGroup key={group.statecode} label={group.label}>
												{group.choices.map((choice) => (
													<Option
														key={choice.statuscode}
														value={targetKey({ statecode: group.statecode, statuscode: choice.statuscode })}
														text={`${group.label}: ${choice.label}`}
													>
														{choice.isDefault ? `${choice.label} (default)` : choice.label}
													</Option>
												))}
											</OptionGroup>
										))}
									</Dropdown>
								</Field>
							</>
						)}
						{note && (
							<MessageBar intent="warning" layout="multiline">
								<MessageBarBody>{note}</MessageBarBody>
							</MessageBar>
						)}
					</>
				);
			case "running":
				return (
					<>
						<MessageBar intent="info">
							<MessageBarBody>
								<MessageBarTitle>Updating records...</MessageBarTitle>
								{progress && `Batch ${progress.batchesCompleted} of ${progress.totalBatches}`}
							</MessageBarBody>
						</MessageBar>
						{progress ? (
							<>
								<ProgressBar value={progress.completed / progress.total} max={1} />
								<Text size={200}>
									{progress.completed} of {progress.total} ({progress.succeeded} succeeded, {progress.failed} failed)
								</Text>
							</>
						) : (
							<div className={styles.row}>
								<Spinner size="small" />
								<Text>Updating...</Text>
							</div>
						)}
					</>
				);
			case "done":
				return result && result.failed > 0 ? (
					<>
						<MessageBar intent="warning" layout="multiline">
							<MessageBarBody>
								<MessageBarTitle>Partly done</MessageBarTitle>
								{result.succeeded} of {result.succeeded + result.failed} records updated. {result.failed} failed.
							</MessageBarBody>
						</MessageBar>
						{renderErrors()}
					</>
				) : (
					<div className={styles.success}>
						<CheckmarkCircle20Regular />
						<Text weight="semibold">
							{result?.succeeded === 1 ? "1 record updated." : `${result?.succeeded ?? 0} records updated.`}
						</Text>
					</div>
				);
			case "error":
				return (
					<>
						<MessageBar intent="error" layout="multiline">
							<MessageBarBody>
								<MessageBarTitle>{verb} failed</MessageBarTitle>
								{error}
							</MessageBarBody>
						</MessageBar>
						{renderErrors()}
					</>
				);
		}
	};

	const renderActions = () => {
		switch (phase) {
			case "loading":
			case "confirm":
				return (
					<>
						<Button appearance="secondary" onClick={handleClose}>
							Cancel
						</Button>
						<Button
							appearance="primary"
							onClick={handleRun}
							disabled={phase !== "confirm" || !target || plan?.kind === "none" || !scopeConfirmed}
						>
							{verb}
						</Button>
					</>
				);
			case "running":
				return (
					<Button appearance="secondary" disabled>
						Cancel
					</Button>
				);
			case "done":
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
				<DialogBody>
					<DialogTitle>{`${verb} ${count === 1 ? "1 record" : `${count} records`}`}</DialogTitle>
					<DialogContent className={styles.content}>{renderContent()}</DialogContent>
					<DialogActions>{renderActions()}</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
