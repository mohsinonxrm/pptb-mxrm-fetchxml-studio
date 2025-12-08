/**
 * SaveViewDialog - Dialog for saving views to Dataverse
 * Handles validation, name/description input, and view type selection
 */

import { useState, useEffect, useCallback, useMemo } from "react";
import {
	Dialog,
	DialogSurface,
	DialogBody,
	DialogTitle,
	DialogContent,
	DialogActions,
	Button,
	Input,
	Textarea,
	Field,
	Radio,
	RadioGroup,
	MessageBar,
	MessageBarTitle,
	MessageBarBody,
	Spinner,
	makeStyles,
	tokens,
	Divider,
	Text,
} from "@fluentui/react-components";
import {
	DismissRegular,
	InfoRegular,
	WarningRegular,
	ErrorCircleRegular,
	CheckmarkCircleRegular,
} from "@fluentui/react-icons";
import {
	validateFetchXmlExpression,
	checkSavedQueryPrivileges,
	checkUserQueryPrivileges,
	createSavedQuery,
	updateSavedQuery,
	createUserQuery,
	updateUserQuery,
	publishSavedQuery,
	addSolutionComponent,
	getUnmanagedSolutions,
	COMPONENT_TYPE_SAVEDQUERY,
	type ValidatorIssue,
	type Solution,
} from "../../api/pptbClient";
import { SolutionPicker } from "./SolutionPicker";

const useStyles = makeStyles({
	surface: {
		maxWidth: "500px",
		width: "100%",
	},
	content: {
		display: "flex",
		flexDirection: "column",
		gap: "16px",
	},
	fieldGroup: {
		display: "flex",
		flexDirection: "column",
		gap: "12px",
	},
	validationSection: {
		display: "flex",
		flexDirection: "column",
		gap: "8px",
	},
	validationHeader: {
		display: "flex",
		alignItems: "center",
		gap: "8px",
	},
	validationList: {
		display: "flex",
		flexDirection: "column",
		gap: "6px",
		maxHeight: "150px",
		overflowY: "auto",
		paddingLeft: "4px",
	},
	validationItem: {
		display: "flex",
		alignItems: "flex-start",
		gap: "8px",
		fontSize: "13px",
	},
	severityIcon: {
		flexShrink: 0,
		marginTop: "2px",
	},
	severityLow: {
		color: tokens.colorPaletteBlueBackground2,
	},
	severityMedium: {
		color: tokens.colorPaletteYellowForeground1,
	},
	severityHigh: {
		color: tokens.colorPaletteRedForeground1,
	},
	severityCritical: {
		color: tokens.colorPaletteRedForeground1,
	},
	solutionInfo: {
		display: "flex",
		flexDirection: "column",
		gap: "4px",
		padding: "8px 12px",
		backgroundColor: tokens.colorNeutralBackground3,
		borderRadius: "4px",
		marginTop: "8px",
	},
	solutionLabel: {
		fontSize: "12px",
		color: tokens.colorNeutralForeground3,
	},
	solutionValue: {
		fontSize: "13px",
		fontWeight: 500,
	},
	overwriteWarning: {
		marginTop: "4px",
	},
	radioDescription: {
		fontSize: "12px",
		color: tokens.colorNeutralForeground3,
		marginLeft: "28px",
		marginTop: "-4px",
	},
	publishNote: {
		fontSize: "12px",
		color: tokens.colorNeutralForeground3,
		fontStyle: "italic",
	},
});

export type SaveViewDialogMode = "save" | "saveAs";

interface SaveViewDialogProps {
	/** Whether the dialog is open */
	open: boolean;
	/** Close callback */
	onClose: () => void;
	/** Save complete callback */
	onSaveComplete: (viewId: string, viewType: "system" | "personal", viewName: string) => void;
	/** Dialog mode: save (overwrite) or saveAs (new) */
	mode: SaveViewDialogMode;
	/** Whether to publish after save (system views only) */
	shouldPublish: boolean;
	/** FetchXML to save */
	fetchXml: string;
	/** LayoutXML configuration */
	layoutXml: string;
	/** Entity logical name */
	entityLogicalName: string;
	/** Entity object type code */
	objectTypeCode: number;
	/** Primary ID attribute */
	primaryIdAttribute: string;
	/** Currently loaded view (for Save mode) */
	loadedView: {
		id: string;
		type: "system" | "personal";
		name: string;
	} | null;
}

interface ValidationState {
	status: "idle" | "validating" | "valid" | "warning" | "error";
	messages: ValidatorIssue[];
	errorMessage?: string;
}

export function SaveViewDialog({
	open,
	onClose,
	onSaveComplete,
	mode,
	shouldPublish,
	fetchXml,
	layoutXml,
	entityLogicalName,
	objectTypeCode: _objectTypeCode, // Reserved for future use
	primaryIdAttribute: _primaryIdAttribute, // Reserved for future use
	loadedView,
}: SaveViewDialogProps) {
	const styles = useStyles();

	// Form state
	const [name, setName] = useState("");
	const [description, setDescription] = useState("");
	const [viewType, setViewType] = useState<"system" | "personal">("personal");

	// Solution state (for system views)
	const [selectedSolution, setSelectedSolution] = useState<Solution | null>(null);
	const [unmanagedSolutions, setUnmanagedSolutions] = useState<Solution[]>([]);
	const [solutionsLoading, setSolutionsLoading] = useState(false);

	// Privileges state - systemView privileges from checkSavedQueryPrivileges
	const [systemPrivileges, setSystemPrivileges] = useState<{
		canWrite: boolean;
		canPublish: boolean;
	} | null>(null);
	// Personal view privileges from checkUserQueryPrivileges
	const [personalPrivileges, setPersonalPrivileges] = useState<{
		canWrite: boolean;
	} | null>(null);

	// Validation state
	const [validation, setValidation] = useState<ValidationState>({
		status: "idle",
		messages: [],
	});

	// Save state
	const [isSaving, setIsSaving] = useState(false);
	const [saveError, setSaveError] = useState<string | null>(null);

	// Check if this is a "Save" to existing view
	const isOverwriteMode = mode === "save" && loadedView !== null;

	// Initialize form when dialog opens
	useEffect(() => {
		if (open) {
			if (isOverwriteMode && loadedView) {
				setName(loadedView.name);
				setViewType(loadedView.type);
			} else {
				setName("");
				setViewType("personal");
			}
			setDescription("");
			setSaveError(null);
			setValidation({ status: "idle", messages: [] });
		}
	}, [open, isOverwriteMode, loadedView]);

	// Check privileges when dialog opens
	useEffect(() => {
		if (open) {
			// Check both system and personal view privileges
			checkSavedQueryPrivileges().then(setSystemPrivileges);
			checkUserQueryPrivileges().then(setPersonalPrivileges);
		}
	}, [open]);

	// Load unmanaged solutions when system view is selected
	useEffect(() => {
		if (!open || viewType !== "system") {
			return;
		}

		async function loadSolutionInfo() {
			setSolutionsLoading(true);
			try {
				// Always load unmanaged solutions for the picker
				const solutions = await getUnmanagedSolutions();
				setUnmanagedSolutions(solutions);
			} catch (error) {
				console.error("Failed to load solution info:", error);
				setUnmanagedSolutions([]);
			} finally {
				setSolutionsLoading(false);
			}
		}

		loadSolutionInfo();
	}, [open, viewType]);

	// Auto-validate when dialog opens
	useEffect(() => {
		if (!open || !fetchXml) {
			return;
		}

		async function runValidation() {
			setValidation({ status: "validating", messages: [] });

			try {
				const result = await validateFetchXmlExpression(fetchXml);
				const messages = result.ValidationResults?.Messages || [];

				if (messages.length === 0) {
					setValidation({ status: "valid", messages: [] });
				} else {
					// Check for critical/high severity
					const hasCritical = messages.some((m) => m.Severity >= 2);
					setValidation({
						status: hasCritical ? "warning" : "warning",
						messages,
					});
				}
			} catch (error) {
				// Validation API call failed - this doesn't mean the FetchXML is invalid
				// Just skip validation and allow the user to proceed
				console.warn("FetchXML validation unavailable:", error);
				setValidation({
					status: "warning",
					messages: [],
					errorMessage: "Validation unavailable - FetchXML will be validated when saved",
				});
			}
		}

		runValidation();
	}, [open, fetchXml]);

	// Get severity icon
	const getSeverityIcon = useCallback(
		(severity: number) => {
			switch (severity) {
				case 0:
					return <InfoRegular className={`${styles.severityIcon} ${styles.severityLow}`} />;
				case 1:
					return <WarningRegular className={`${styles.severityIcon} ${styles.severityMedium}`} />;
				case 2:
				case 3:
					return <ErrorCircleRegular className={`${styles.severityIcon} ${styles.severityHigh}`} />;
				default:
					return <InfoRegular className={`${styles.severityIcon} ${styles.severityLow}`} />;
			}
		},
		[styles]
	);

	// Can save check
	const canSave = useMemo(() => {
		if (!name.trim()) return false;
		if (isSaving) return false;
		if (viewType === "system") {
			// System views require canWrite privilege (covers both create and update)
			if (!systemPrivileges?.canWrite) return false;
		} else {
			// Personal views require canWrite privilege
			if (!personalPrivileges?.canWrite) return false;
		}
		return true;
	}, [name, isSaving, viewType, systemPrivileges, personalPrivileges]);

	// Handle save
	const handleSave = useCallback(async () => {
		if (!canSave) return;

		setIsSaving(true);
		setSaveError(null);

		try {
			let viewId: string;

			if (viewType === "personal") {
				// Personal view
				if (isOverwriteMode && loadedView) {
					await updateUserQuery(loadedView.id, {
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						description: description.trim() || undefined,
					});
					viewId = loadedView.id;
				} else {
					viewId = await createUserQuery({
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						returnedtypecode: entityLogicalName,
						description: description.trim() || undefined,
						querytype: 0,
					});
				}
			} else {
				// System view
				if (isOverwriteMode && loadedView) {
					await updateSavedQuery(loadedView.id, {
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						description: description.trim() || undefined,
					});
					viewId = loadedView.id;
				} else {
					viewId = await createSavedQuery({
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						returnedtypecode: entityLogicalName,
						description: description.trim() || undefined,
						querytype: 0,
					});

					// Add to solution if selected
					if (selectedSolution) {
						await addSolutionComponent(
							viewId,
							COMPONENT_TYPE_SAVEDQUERY,
							selectedSolution.uniquename,
							false
						);
					}
				}

				// Publish if requested
				if (shouldPublish) {
					await publishSavedQuery(viewId);
				}
			}

			onSaveComplete(viewId, viewType, name.trim());
		} catch (error) {
			console.error("Failed to save view:", error);
			setSaveError(error instanceof Error ? error.message : "Failed to save view");
		} finally {
			setIsSaving(false);
		}
	}, [
		canSave,
		viewType,
		isOverwriteMode,
		loadedView,
		name,
		fetchXml,
		layoutXml,
		description,
		entityLogicalName,
		selectedSolution,
		shouldPublish,
		onSaveComplete,
	]);

	// Dialog title
	const dialogTitle = useMemo(() => {
		if (isOverwriteMode) {
			return shouldPublish ? "Save and Publish View" : "Save View";
		}
		return "Save As New View";
	}, [isOverwriteMode, shouldPublish]);

	return (
		<Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
			<DialogSurface className={styles.surface}>
				<DialogBody>
					<DialogTitle
						action={
							<Button
								appearance="subtle"
								aria-label="Close"
								icon={<DismissRegular />}
								onClick={onClose}
							/>
						}
					>
						{dialogTitle}
					</DialogTitle>
					<DialogContent className={styles.content}>
						{/* Validation Section */}
						<div className={styles.validationSection}>
							<div className={styles.validationHeader}>
								{validation.status === "validating" && (
									<>
										<Spinner size="tiny" />
										<Text>Validating FetchXML...</Text>
									</>
								)}
								{validation.status === "valid" && (
									<>
										<CheckmarkCircleRegular
											style={{ color: tokens.colorPaletteGreenForeground1 }}
										/>
										<Text>FetchXML is valid</Text>
									</>
								)}
								{validation.status === "warning" && (
									<>
										<WarningRegular style={{ color: tokens.colorPaletteYellowForeground1 }} />
										<Text>
											{validation.messages.length} validation{" "}
											{validation.messages.length === 1 ? "message" : "messages"}
										</Text>
									</>
								)}
								{validation.status === "error" && (
									<>
										<ErrorCircleRegular style={{ color: tokens.colorPaletteRedForeground1 }} />
										<Text>{validation.errorMessage}</Text>
									</>
								)}
							</div>

							{validation.messages.length > 0 && (
								<div className={styles.validationList}>
									{validation.messages.map((msg, idx) => (
										<div key={idx} className={styles.validationItem}>
											{getSeverityIcon(msg.Severity)}
											<span>{msg.LocalizedMessageText}</span>
										</div>
									))}
								</div>
							)}
						</div>

						<Divider />

						{/* Form Fields */}
						<div className={styles.fieldGroup}>
							<Field label="Name" required>
								<Input
									value={name}
									onChange={(_, data) => setName(data.value)}
									placeholder="Enter view name"
									disabled={isSaving}
								/>
							</Field>
							<Field label="Description">
								<Textarea
									value={description}
									onChange={(_, data) => setDescription(data.value)}
									placeholder="Optional description"
									disabled={isSaving}
									rows={2}
								/>
							</Field>
							{/* View Type Selection - only for Save As */}
							{!isOverwriteMode && (
								<Field label="View Type">
									<RadioGroup
										value={viewType}
										onChange={(_, data) => setViewType(data.value as "system" | "personal")}
										disabled={isSaving}
									>
										<Radio
											value="personal"
											label="Personal View"
											disabled={!personalPrivileges?.canWrite}
										/>
										<div className={styles.radioDescription}>
											{personalPrivileges?.canWrite
												? "Only visible to you"
												: "You don't have permission to create personal views"}
										</div>
										<Radio
											value="system"
											label="System View"
											disabled={!systemPrivileges?.canWrite}
										/>
										<div className={styles.radioDescription}>
											{systemPrivileges?.canWrite
												? "Visible to all users with access to this entity"
												: "You don't have permission to create system views"}
										</div>
									</RadioGroup>
								</Field>
							)}
							{/* Solution Picker for System Views */}
							{viewType === "system" && !isOverwriteMode && (
								<>
									{solutionsLoading ? (
										<div className={styles.solutionInfo}>
											<Spinner size="tiny" label="Loading solutions..." />
										</div>
									) : (
										<SolutionPicker
											solutions={unmanagedSolutions}
											selectedSolution={selectedSolution}
											onSolutionChange={setSelectedSolution}
											disabled={isSaving}
										/>
									)}
								</>
							)}{" "}
							{/* Overwrite Warning */}
							{isOverwriteMode && loadedView && (
								<MessageBar intent="warning" className={styles.overwriteWarning}>
									<MessageBarBody>
										<MessageBarTitle>Overwrite Warning</MessageBarTitle>
										This will overwrite the existing view "{loadedView.name}".
									</MessageBarBody>
								</MessageBar>
							)}
							{/* Publish Note */}
							{shouldPublish && viewType === "system" && (
								<Text className={styles.publishNote}>
									The view will be published after saving, making it immediately available to users.
								</Text>
							)}
						</div>

						{/* Save Error */}
						{saveError && (
							<MessageBar intent="error">
								<MessageBarBody>
									<MessageBarTitle>Save Failed</MessageBarTitle>
									{saveError}
								</MessageBarBody>
							</MessageBar>
						)}
					</DialogContent>
					<DialogActions>
						<Button appearance="secondary" onClick={onClose} disabled={isSaving}>
							Cancel
						</Button>
						<Button appearance="primary" onClick={handleSave} disabled={!canSave}>
							{isSaving ? <Spinner size="tiny" /> : shouldPublish ? "Save and Publish" : "Save"}
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
