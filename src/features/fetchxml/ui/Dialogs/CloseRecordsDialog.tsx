/**
 * Close dialog for a close message (#149): WinOpportunity / LoseOpportunity, CloseIncident,
 * WinQuote / CloseQuote, FulfillSalesOrder / CancelSalesOrder. Starts with the environment and
 * scope (#69); the status reason choices come from the target state in metadata; the fields are
 * the close activity's (model/closeActions.ts). The same values go to every selected record, one
 * request and one close activity each. Then progress and per-record errors like delete.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
	Button,
	Combobox,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	Dropdown,
	Field,
	Input,
	Link,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Option,
	ProgressBar,
	SpinButton,
	Spinner,
	Text,
	Textarea,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { CheckmarkCircle20Regular } from "@fluentui/react-icons";
import type { BatchDeleteProgress, BatchDeleteResult, CompetitorOption } from "../../api/pptbClient";
import { getEntityStateModel, searchCompetitors } from "../../api/pptbClient";
import { formatDataverseError } from "../../api/dataverseErrors";
import { openUrl } from "../../api/host";
import {
	RESOLUTION_MAX_LENGTH,
	closeStatusChoices,
	closedSummary,
	suggestedCloseStatus,
	validateCloseValues,
	type CloseField,
	type CloseMessageAction,
	type CloseStatusChoices,
	type CloseValues,
} from "../../model/closeActions";
import { TargetScope } from "../../../../shared/components/TargetScope";
import { UndoNote } from "../../../../shared/components/UndoNote";

const useStyles = makeStyles({
	surface: {
		maxWidth: "560px",
	},
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
	note: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
		lineHeight: tokens.lineHeightBase200,
	},
	code: {
		fontFamily: tokens.fontFamilyMonospace,
		fontSize: tokens.fontSizeBase200,
	},
	option: {
		display: "flex",
		flexDirection: "column",
		minWidth: 0,
	},
	optionSecondary: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
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

export interface CloseRecordsDialogProps {
	open: boolean;
	action: CloseMessageAction | null;
	recordIds: string[];
	entityLogicalName?: string;
	entityDisplayName?: string;
	onClose: () => void;
	/** Runs the message on every record with these values; resolves with per-record results */
	onExecute: (values: CloseValues, onProgress: (progress: BatchDeleteProgress) => void) => Promise<BatchDeleteResult>;
}

type Phase = "loading" | "confirm" | "running" | "done" | "error";

const today = () => {
	const now = new Date();
	return new Date(now.getFullYear(), now.getMonth(), now.getDate());
};

const DEFAULT_BILLABLE_MINUTES = 30;

export function CloseRecordsDialog({
	open,
	action,
	recordIds,
	entityLogicalName,
	entityDisplayName,
	onClose,
	onExecute,
}: CloseRecordsDialogProps) {
	const styles = useStyles();
	const [phase, setPhase] = useState<Phase>("loading");
	const [choices, setChoices] = useState<CloseStatusChoices | null>(null);
	const [error, setError] = useState<string | null>(null);
	const [progress, setProgress] = useState<BatchDeleteProgress | null>(null);
	const [result, setResult] = useState<BatchDeleteResult | null>(null);
	// Type-to-confirm for Production or more than 100 records (#69)
	const [scopeConfirmed, setScopeConfirmed] = useState(false);

	// Values
	const [status, setStatus] = useState<number | undefined>();
	const [statusTouched, setStatusTouched] = useState(false);
	const [revenue, setRevenue] = useState("");
	const [date, setDate] = useState<Date | null>(today);
	const [description, setDescription] = useState("");
	const [competitor, setCompetitor] = useState<CompetitorOption | undefined>();
	const [competitorText, setCompetitorText] = useState("");
	const [competitors, setCompetitors] = useState<CompetitorOption[]>([]);
	const [competitorsLoading, setCompetitorsLoading] = useState(false);
	const [resolution, setResolution] = useState("");
	const [resolutionTouched, setResolutionTouched] = useState(false);
	const [billable, setBillable] = useState<number | null>(DEFAULT_BILLABLE_MINUTES);

	const count = recordIds.length;
	const fields = useMemo(() => new Set<CloseField>(action?.fields ?? []), [action]);

	// Reset and read the status reasons when the dialog opens
	useEffect(() => {
		if (!open || !entityLogicalName || !action) return;
		let current = true;
		setPhase("loading");
		setChoices(null);
		setError(null);
		setProgress(null);
		setResult(null);
		setStatus(undefined);
		setStatusTouched(false);
		setRevenue("");
		setDate(today());
		setDescription("");
		setCompetitor(undefined);
		setCompetitorText("");
		setResolution("");
		setResolutionTouched(false);
		setBillable(DEFAULT_BILLABLE_MINUTES);
		getEntityStateModel(entityLogicalName)
			.then((model) => {
				if (!current) return;
				if (!model) throw new Error("This entity has no Status (statecode) column.");
				const next = closeStatusChoices(model, action);
				setChoices(next);
				setStatus(next.preselected);
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
	}, [open, entityLogicalName, action]);

	// Naming a competitor suggests Out-Sold, until the user picks a status reason
	useEffect(() => {
		if (!action || !choices || statusTouched) return;
		setStatus(suggestedCloseStatus(action, choices, competitor?.id));
	}, [action, choices, competitor, statusTouched]);

	// Competitor search (Lose opportunity), shortly after typing stops
	const searchSeq = useRef(0);
	useEffect(() => {
		if (!open || !fields.has("competitor") || phase !== "confirm") return;
		const seq = ++searchSeq.current;
		const timer = window.setTimeout(() => {
			setCompetitorsLoading(true);
			searchCompetitors(competitor && competitorText === competitor.name ? "" : competitorText)
				.then((list) => {
					if (seq === searchSeq.current) setCompetitors(list);
				})
				.catch((err: unknown) => {
					console.warn("Competitor search failed:", err);
					if (seq === searchSeq.current) setCompetitors([]);
				})
				.finally(() => {
					if (seq === searchSeq.current) setCompetitorsLoading(false);
				});
		}, 250);
		return () => window.clearTimeout(timer);
	}, [open, fields, phase, competitorText, competitor]);

	const values = useMemo<CloseValues | undefined>(
		() =>
			status === undefined
				? undefined
				: {
						statuscode: status,
						revenue,
						date,
						description,
						competitorId: competitor?.id,
						resolution,
						billableMinutes: billable,
					},
		[status, revenue, date, description, competitor, resolution, billable],
	);
	const errors = action && values ? validateCloseValues(action, values) : {};
	const valid = !!values && Object.keys(errors).length === 0;

	const handleRun = useCallback(async () => {
		if (!values || !action) return;
		setPhase("running");
		setProgress(null);
		try {
			const res = await onExecute(values, setProgress);
			setResult(res);
			if (res.succeeded === 0 && res.failed > 0) {
				setError(`All ${res.failed} records failed.`);
				setPhase("error");
			} else {
				setPhase("done");
			}
		} catch (err) {
			setError(formatDataverseError(err, `${action.label} failed`)); // (#96)
			setPhase("error");
		}
	}, [values, action, onExecute]);

	const handleClose = useCallback(() => {
		if (phase !== "running") onClose();
	}, [phase, onClose]);

	if (!action) return null;

	const statusLabel = action.statusLabel ?? "Status reason";
	const stateLabel = choices?.state?.label ?? "";
	const chosen = choices?.choices.find((c) => c.statuscode === status);

	const renderStatus = () => {
		if (!choices) return null;
		if (choices.choices.length === 0) {
			return (
				<MessageBar intent="warning" layout="multiline">
					<MessageBarBody>
						<MessageBarTitle>Nothing to choose</MessageBarTitle>
						The target Status has no status reasons in metadata.
					</MessageBarBody>
				</MessageBar>
			);
		}
		if (choices.choices.length === 1) {
			return (
				<Text>
					Sets Status to <strong>{stateLabel}</strong> and {statusLabel} to <strong>{choices.choices[0].label}</strong>.
				</Text>
			);
		}
		return (
			<Field label={statusLabel} required hint={`Status reasons of ${stateLabel}, from metadata. The default is marked.`}>
				<Dropdown
					value={chosen?.label ?? ""}
					selectedOptions={status !== undefined ? [String(status)] : []}
					onOptionSelect={(_, data) => {
						const value = Number(data.optionValue);
						if (Number.isFinite(value)) {
							setStatus(value);
							setStatusTouched(true);
						}
					}}
				>
					{choices.choices.map((c) => (
						<Option key={c.statuscode} value={String(c.statuscode)} text={c.label}>
							{c.isDefault ? `${c.label} (default)` : c.label}
						</Option>
					))}
				</Dropdown>
			</Field>
		);
	};

	const renderField = (kind: CloseField) => {
		switch (kind) {
			case "revenue":
				return (
					<Field
						key={kind}
						label="Actual revenue"
						validationState={errors.revenue ? "error" : "none"}
						validationMessage={errors.revenue}
						hint={
							count > 1
								? "Applied to every selected opportunity. Leave empty to keep each one's own value."
								: "Leave empty to keep the opportunity's own value."
						}
					>
						<Input value={revenue} onChange={(_, d) => setRevenue(d.value)} inputMode="decimal" placeholder="Optional" />
					</Field>
				);
			case "date":
				return (
					<Field key={kind} label="Close date" required validationState={errors.date ? "error" : "none"} validationMessage={errors.date}>
						<DatePicker value={date} onSelectDate={(d) => setDate(d ?? null)} placeholder="Select a date..." allowTextInput />
					</Field>
				);
			case "description":
				return (
					<Field key={kind} label={fields.has("resolution") ? "Remarks" : "Description"}>
						<Textarea
							value={description}
							onChange={(_, d) => setDescription(d.value)}
							maxLength={action.descriptionMaxLength}
							rows={2}
							resize="vertical"
							placeholder="Optional"
						/>
					</Field>
				);
			case "competitor":
				return (
					<Field key={kind} label="Competitor" hint="Optional. Naming a competitor selects Out-Sold unless you chose a status reason.">
						<Combobox
							placeholder="Search competitors..."
							value={competitorText}
							selectedOptions={competitor ? [competitor.id] : []}
							freeform
							clearable
							onChange={(event) => {
								setCompetitorText(event.target.value);
								if (competitor && event.target.value !== competitor.name) setCompetitor(undefined);
							}}
							onOptionSelect={(_, data) => {
								const next = competitors.find((c) => c.id === data.optionValue);
								setCompetitor(next);
								setCompetitorText(next?.name ?? "");
							}}
						>
							{competitorsLoading && competitors.length === 0 ? (
								<Option disabled value="__loading__" text="">
									Searching...
								</Option>
							) : competitors.length === 0 ? (
								<Option disabled value="__none__" text="">
									No competitors found
								</Option>
							) : (
								competitors.map((c) => (
									<Option key={c.id} value={c.id} text={c.name}>
										<span className={styles.option}>
											<span>{c.name}</span>
											{c.website && <span className={styles.optionSecondary}>{c.website}</span>}
										</span>
									</Option>
								))
							)}
						</Combobox>
					</Field>
				);
			case "resolution":
				return (
					<Field
						key={kind}
						label="Resolution"
						required
						validationState={errors.resolution && resolutionTouched ? "error" : "none"}
						validationMessage={resolutionTouched ? errors.resolution : undefined}
						hint={`The case resolution's subject (up to ${RESOLUTION_MAX_LENGTH} characters).`}
					>
						<Input
							value={resolution}
							maxLength={RESOLUTION_MAX_LENGTH}
							onChange={(_, d) => setResolution(d.value)}
							onBlur={() => setResolutionTouched(true)}
						/>
					</Field>
				);
			case "billable":
				return (
					<Field
						key={kind}
						label="Billable time (minutes)"
						validationState={errors.billable ? "error" : "none"}
						validationMessage={errors.billable}
						hint={count > 1 ? "Recorded on each case's resolution." : undefined}
					>
						<SpinButton
							value={billable}
							min={0}
							step={15}
							onChange={(_, d) => {
								if (d.value !== undefined) setBillable(d.value);
								else if (d.displayValue !== undefined) setBillable(d.displayValue.trim() === "" ? null : Number(d.displayValue));
							}}
						/>
					</Field>
				);
		}
	};

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
						{/* The way back (#69): Activate reopens closed opportunities, cases and orders (verified,
						    see CLOSE_MESSAGE_ENTITIES); it doesn't apply to quotes */}
						{entityLogicalName?.toLowerCase() === "quote" ? (
							<UndoNote undo="none">There's no undo: won or closed quotes can't be reactivated here.</UndoNote>
						) : (
							<UndoNote undo="possible">To undo, run Activate on the records this changes. The close activities stay.</UndoNote>
						)}
						{action.note && <Text className={styles.note}>{action.note}</Text>}
						{renderStatus()}
						{action.fields.map(renderField)}
						<Text className={styles.note}>
							Runs{" "}
							<Link inline onClick={() => void openUrl(action.docUrl)}>
								{action.message}
							</Link>{" "}
							once per record, in batches like Delete. Each record gets its own{" "}
							<span className={styles.code}>{action.closeEntity}</span> activity with these values.
						</Text>
					</>
				);
			case "running":
				return (
					<>
						<MessageBar intent="info">
							<MessageBarBody>
								<MessageBarTitle>Running {action.message}...</MessageBarTitle>
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
								<Text>Closing...</Text>
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
								{result.succeeded} of {result.succeeded + result.failed} {action.nouns[1]} closed. {result.failed} failed.
							</MessageBarBody>
						</MessageBar>
						{renderErrors()}
						<Text className={styles.note}>Successful records stay closed; nothing is rolled back. Run again on the failed ones after fixing them.</Text>
					</>
				) : (
					<div className={styles.success}>
						<CheckmarkCircle20Regular />
						<Text weight="semibold">{closedSummary(action, result?.succeeded ?? 0)}</Text>
					</div>
				);
			case "error":
				return (
					<>
						<MessageBar intent="error" layout="multiline">
							<MessageBarBody>
								<MessageBarTitle>{action.label} failed</MessageBarTitle>
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
						<Button appearance="primary" onClick={handleRun} disabled={phase !== "confirm" || !valid || !scopeConfirmed}>
							{action.label}
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
			<DialogSurface className={styles.surface}>
				<DialogBody>
					<DialogTitle>{action.title(count)}</DialogTitle>
					<DialogContent className={styles.content}>{renderContent()}</DialogContent>
					<DialogActions>{renderActions()}</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
