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
	Checkbox,
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
	validateSavedQuery,
	checkSavedQueryPrivileges,
	checkUserQueryPrivileges,
	createSavedQuery,
	updateSavedQuery,
	createUserQuery,
	updateUserQuery,
	publishEntityViews,
	addSolutionComponent,
	getUnmanagedSolutions,
	checkUserQueryWriteAccess,
	getSolutionLabel,
	getDefaultPublicView,
	whoAmI,
	COMPONENT_TYPE_SAVEDQUERY,
	type ValidatorIssue,
	type Solution,
} from "../../api/pptbClient";
import { formatDataverseError } from "../../api/dataverseErrors";
import { holdClose } from "../../api/host";
import { SolutionPicker } from "./SolutionPicker";
import {
	canOverwriteView,
	canSetDefaultView,
	countQuickFindPlaceholders,
	getValidationStatus,
	isSharedView,
} from "../../model/savedViews";
import { TargetScope } from "../../../../shared/components/TargetScope";
import { UndoNote } from "../../../../shared/components/UndoNote";

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
	accessCheck: {
		display: "flex",
		alignItems: "center",
		gap: "8px",
	},
	placeholderNote: {
		fontSize: "12px",
		color: tokens.colorNeutralForeground3,
	},
	code: {
		fontFamily: tokens.fontFamilyMonospace,
	},
});

export type SaveViewDialogMode = "save" | "saveAs";

/** The loaded view a Save overwrites */
export interface SaveViewDialogLoadedView {
	id: string;
	type: "system" | "personal";
	name: string;
	/** querytype of the view (0 = public view); ValidateSavedQuery checks the query for it (#93) */
	queryType?: number;
	/** System views: false when the view isn't customizable, so it can't be overwritten (#93) */
	isCustomizable?: boolean;
	/** The view's description, pre-filled on overwrite (#142) */
	description?: string;
	/** System views: the default view (isdefault, #142) */
	isDefault?: boolean;
	/** System views: managed (ismanaged) and its solution, for the managed view warning (#142) */
	isManaged?: boolean;
	solutionId?: string;
	/** Personal views: the owner, for views shared with the user (#142) */
	ownerId?: string;
	ownerName?: string;
}

/** What a save wrote besides the name, so the loaded view stays in step (#142) */
export interface SavedViewDetails {
	description?: string;
	isDefault?: boolean;
}

interface SaveViewDialogProps {
	/** Whether the dialog is open */
	open: boolean;
	/** Close callback */
	onClose: () => void;
	/** Save complete callback */
	onSaveComplete: (
		viewId: string,
		viewType: "system" | "personal",
		viewName: string,
		details: SavedViewDetails,
	) => void;
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
	/** Entity display name, for the default view hint (#142) */
	entityDisplayName?: string;
	/** Entity object type code */
	objectTypeCode: number;
	/** Primary ID attribute */
	primaryIdAttribute: string;
	/** Currently loaded view (for Save mode) */
	loadedView: SaveViewDialogLoadedView | null;
}

interface ValidationState {
	/** "unavailable": the ValidateFetchXmlExpression call itself failed (errorMessage says so) */
	status: "idle" | "validating" | "valid" | "warning" | "error" | "unavailable";
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
	entityDisplayName,
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

	// "Save a copy…" from a view that can't be overwritten switches to Save as (#142)
	const [savingCopy, setSavingCopy] = useState(false);
	// The connected user, to tell a personal view shared with them from their own (#142)
	const [currentUserId, setCurrentUserId] = useState<string | undefined>(undefined);
	// RetrievePrincipalAccess on a personal view the user may not own (#142): "unknown" when it
	// couldn't be determined, which allows the save (Dataverse decides)
	const [writeAccess, setWriteAccess] = useState<"checking" | "allowed" | "denied" | "unknown">("unknown");
	// Managed system view: its solution, named in the warning (#142)
	const [managedSolution, setManagedSolution] = useState<{ friendlyName: string; uniqueName: string } | null>(null);
	// Set as default view (#142)
	const [makeDefault, setMakeDefault] = useState(false);
	const [currentDefaultView, setCurrentDefaultView] = useState<{ id: string; name: string } | null>(null);

	// Check if this is a "Save" to existing view
	const isOverwriteRequest = mode === "save" && loadedView !== null;
	const isOverwriteMode = isOverwriteRequest && !savingCopy;
	// A system view whose iscustomizable managed property is false can't be updated (#93)
	const isNotCustomizable = isOverwriteMode && loadedView !== null && !canOverwriteView(loadedView);
	// A personal view shared with the user without write access (#142)
	const isReadOnlyShare = isOverwriteMode && writeAccess === "denied";
	const isOverwriteBlocked = isNotCustomizable || isReadOnlyShare;
	const isCheckingAccess = isOverwriteMode && writeAccess === "checking";
	const isShared = isOverwriteMode && loadedView !== null && isSharedView(loadedView, currentUserId);
	const isManagedOverwrite = isOverwriteMode && loadedView?.type === "system" && loadedView.isManaged === true;
	// Public system views only: the loaded one on overwrite, a new system view on Save as (#142)
	const showDefaultOption =
		!isOverwriteBlocked && (isOverwriteMode && loadedView ? canSetDefaultView(loadedView) : viewType === "system");
	const isAlreadyDefault = isOverwriteMode && loadedView?.isDefault === true;
	const hasQuickFindPlaceholders = useMemo(() => countQuickFindPlaceholders(fetchXml) > 0, [fetchXml]);

	// Initialize form when dialog opens
	useEffect(() => {
		if (open) {
			if (isOverwriteRequest && loadedView) {
				setName(loadedView.name);
				setViewType(loadedView.type);
				// Overwrite keeps the view's description unless it's edited (#142)
				setDescription(loadedView.description ?? "");
				setMakeDefault(loadedView.isDefault === true);
			} else {
				setName("");
				setViewType("personal");
				setDescription("");
				setMakeDefault(false);
			}
			setSavingCopy(false);
			setSaveError(null);
			setValidation({ status: "idle", messages: [] });
		}
	}, [open, isOverwriteRequest, loadedView]);

	// Who owns the view, and whether the user can write it (#142). Their own views skip the check.
	useEffect(() => {
		if (!open || !isOverwriteRequest || loadedView?.type !== "personal") {
			setWriteAccess("unknown");
			return;
		}
		let active = true;
		setWriteAccess("checking");
		void (async () => {
			const user = await whoAmI();
			if (!active) return;
			setCurrentUserId(user?.UserId);
			const ownView = !!user && !!loadedView.ownerId && !isSharedView(loadedView, user.UserId);
			const canWrite = ownView ? true : await checkUserQueryWriteAccess(loadedView.id);
			if (!active) return;
			setWriteAccess(canWrite === undefined ? "unknown" : canWrite ? "allowed" : "denied");
		})();
		return () => {
			active = false;
		};
	}, [open, isOverwriteRequest, loadedView]);

	// The managed view's solution (#142)
	useEffect(() => {
		setManagedSolution(null);
		if (!open || !isManagedOverwrite || !loadedView?.solutionId) return;
		let active = true;
		void getSolutionLabel(loadedView.solutionId).then((label) => {
			if (active) setManagedSolution(label);
		});
		return () => {
			active = false;
		};
	}, [open, isManagedOverwrite, loadedView]);

	// The current default public view, named in the hint (#142)
	useEffect(() => {
		if (!open || !showDefaultOption || !entityLogicalName) return;
		let active = true;
		void getDefaultPublicView(entityLogicalName).then((view) => {
			if (active) setCurrentDefaultView(view);
		});
		return () => {
			active = false;
		};
	}, [open, showDefaultOption, entityLogicalName]);

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
				// Performance advice: high / critical show as errors but never block saving (#93)
				setValidation({ status: getValidationStatus(messages), messages });
			} catch (error) {
				// Validation API call failed - this doesn't mean the FetchXML is invalid
				// Just skip validation and allow the user to proceed
				console.warn("FetchXML validation unavailable:", error);
				setValidation({
					status: "unavailable",
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
		if (isOverwriteBlocked || isCheckingAccess) return false;
		if (viewType === "system") {
			// System views require canWrite privilege (covers both create and update)
			if (!systemPrivileges?.canWrite) return false;
		} else {
			// Personal views require canWrite privilege
			if (!personalPrivileges?.canWrite) return false;
		}
		return true;
	}, [name, isSaving, isOverwriteBlocked, isCheckingAccess, viewType, systemPrivileges, personalPrivileges]);

	// Handle save
	const handleSave = useCallback(async () => {
		if (!canSave) return;

		setIsSaving(true);
		setSaveError(null);
		// The host warns before the tab closes while the view is saved and published (#70)
		const releaseClose = holdClose();

		try {
			let viewId: string;
			// Overwrite writes the description as shown, so clearing it clears the view's (#142)
			const savedDescription = description.trim();
			// Set as default view: sent only to make a view the default; Dataverse keeps one default
			// public view per entity ("You can set only one active public view as the default view").
			// https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#set-a-public-view-as-the-default-view
			const setAsDefault = showDefaultOption && makeDefault && !isAlreadyDefault;

			if (viewType === "personal") {
				// Personal view
				if (isOverwriteMode && loadedView) {
					await updateUserQuery(loadedView.id, {
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						description: savedDescription || null,
					});
					viewId = loadedView.id;
				} else {
					viewId = await createUserQuery({
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						returnedtypecode: entityLogicalName,
						description: savedDescription || undefined,
						querytype: 0,
					});
				}
			} else {
				// System view: Dataverse validates the query for the view's query type first, and its
				// error is shown as the save error (#93). Overwrites keep the view's querytype.
				// https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/validatesavedquery
				const queryType = isOverwriteMode && loadedView ? (loadedView.queryType ?? 0) : 0;
				await validateSavedQuery(fetchXml, queryType);

				if (isOverwriteMode && loadedView) {
					await updateSavedQuery(loadedView.id, {
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						description: savedDescription || null,
						...(setAsDefault ? { isdefault: true } : {}),
					});
					viewId = loadedView.id;
				} else {
					viewId = await createSavedQuery({
						name: name.trim(),
						fetchxml: fetchXml,
						layoutxml: layoutXml,
						returnedtypecode: entityLogicalName,
						description: savedDescription || undefined,
						querytype: 0,
						isdefault: setAsDefault,
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
					// Publishes the view's entity: the documented PublishXml form (#93 h)
					await publishEntityViews(entityLogicalName);
				}
			}

			onSaveComplete(viewId, viewType, name.trim(), {
				description: savedDescription || undefined,
				isDefault: viewType === "system" ? setAsDefault || isAlreadyDefault : undefined,
			});
		} catch (error) {
			console.error("Failed to save view:", error);
			setSaveError(formatDataverseError(error, "Failed to save view")); // (#96)
		} finally {
			releaseClose();
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
		showDefaultOption,
		makeDefault,
		isAlreadyDefault,
	]);

	// Save a copy of a view the user can't overwrite: Save as, named after it (#142)
	const handleSaveCopy = useCallback(() => {
		if (!loadedView) return;
		setSavingCopy(true);
		setName(`${loadedView.name} (copy)`);
		setViewType("personal");
		setMakeDefault(false);
		setSaveError(null);
	}, [loadedView]);

	// Dialog title
	const dialogTitle = useMemo(() => {
		if (isOverwriteMode) {
			return shouldPublish ? "Save and publish view" : "Save view";
		}
		return "Save as new view";
	}, [isOverwriteMode, shouldPublish]);

	// What the save writes, for the environment and scope block (#69)
	const scopeText = useMemo(() => {
		const publish = shouldPublish && viewType === "system" ? ", then published" : "";
		if (isOverwriteMode && loadedView) return `The ${loadedView.type} view "${loadedView.name}" (overwritten${publish})`;
		return `A new ${viewType} view${publish}`;
	}, [isOverwriteMode, loadedView, shouldPublish, viewType]);

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
						{/* Environment and scope (#69) */}
						<TargetScope entityLogicalName={entityLogicalName} scope={{ kind: "text", text: scopeText }} />
						{isOverwriteMode &&
							loadedView &&
							!isOverwriteBlocked &&
							!isCheckingAccess &&
							(isManagedOverwrite ? (
								// Removing the unmanaged layer brings the managed definition back:
								// https://learn.microsoft.com/power-apps/maker/data-platform/solution-layers#remove-an-unmanaged-layer
								<UndoNote undo="possible">
									To undo, remove the view's unmanaged layer in Power Apps (Solution layers): the managed definition
									comes back.
								</UndoNote>
							) : (
								<UndoNote undo="none">
									There's no undo: the view's current definition is replaced. To keep it, cancel and use Save as
									instead.
								</UndoNote>
							))}

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
										{/* ValidateFetchXmlExpression only checks performance (#93) */}
										<Text>No performance issues found</Text>
									</>
								)}
								{(validation.status === "warning" || validation.status === "error") && (
									<>
										{validation.status === "error" ? (
											<ErrorCircleRegular style={{ color: tokens.colorPaletteRedForeground1 }} />
										) : (
											<WarningRegular style={{ color: tokens.colorPaletteYellowForeground1 }} />
										)}
										<Text>
											{validation.messages.length} validation{" "}
											{validation.messages.length === 1 ? "message" : "messages"}
										</Text>
									</>
								)}
								{validation.status === "unavailable" && (
									<>
										<InfoRegular style={{ color: tokens.colorNeutralForeground3 }} />
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

						{/* RetrievePrincipalAccess on a personal view the user may not own (#142) */}
						{isCheckingAccess && loadedView && (
							<div className={styles.accessCheck}>
								<Spinner size="tiny" />
								<Text>Checking your access to "{loadedView.name}"…</Text>
							</div>
						)}

						{/* Form Fields */}
						{!isCheckingAccess && (
						<div className={styles.fieldGroup}>
							{/* Shared without write access: only a copy can be saved (#142) */}
							{isReadOnlyShare && loadedView && (
								<MessageBar intent="error" layout="multiline">
									<MessageBarBody>
										<MessageBarTitle>Can't overwrite this view</MessageBarTitle>
										{loadedView.ownerName
											? `${loadedView.ownerName} shared "${loadedView.name}" with you without write access. Save a copy instead.`
											: `"${loadedView.name}" was shared with you without write access. Save a copy instead.`}
									</MessageBarBody>
								</MessageBar>
							)}
							<Field label="Name" required>
								<Input
									value={name}
									onChange={(_, data) => setName(data.value)}
									placeholder="Enter view name"
									disabled={isSaving || isOverwriteBlocked}
								/>
							</Field>
							<Field
								label="Description"
								hint={isOverwriteMode && loadedView?.description ? "Loaded from the view." : undefined}
							>
								<Textarea
									value={description}
									onChange={(_, data) => setDescription(data.value)}
									placeholder="Optional description"
									disabled={isSaving || isOverwriteBlocked}
									rows={2}
								/>
							</Field>
							{/* View Type Selection - only for Save As */}
							{!isOverwriteMode && (
								<Field label="View type">
									<RadioGroup
										value={viewType}
										onChange={(_, data) => setViewType(data.value as "system" | "personal")}
										disabled={isSaving}
									>
										<Radio
											value="personal"
											label="Personal view"
											disabled={!personalPrivileges?.canWrite}
										/>
										<div className={styles.radioDescription}>
											{personalPrivileges?.canWrite
												? "Only visible to you"
												: "You don't have permission to create personal views"}
										</div>
										<Radio
											value="system"
											label="System view"
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
							{/* Set as default view: public system views only (#142) */}
							{showDefaultOption && (
								<Field
									hint={
										isAlreadyDefault
											? "This is the default view. To change the default, set another public view as default."
											: currentDefaultView && currentDefaultView.id !== loadedView?.id
												? `Replaces "${currentDefaultView.name}" as the default public view for ${entityDisplayName || entityLogicalName}.`
												: undefined
									}
								>
									<Checkbox
										label="Set as default view"
										checked={makeDefault}
										disabled={isSaving || isAlreadyDefault}
										onChange={(_, data) => setMakeDefault(data.checked === true)}
									/>
								</Field>
							)}
							{/* Not customizable: the view can only be saved as a copy (#93) */}
							{isNotCustomizable && loadedView && (
								<MessageBar intent="error" className={styles.overwriteWarning}>
									<MessageBarBody>
										<MessageBarTitle>Can't overwrite this view</MessageBarTitle>
										"{loadedView.name}" isn't customizable. Save a copy instead.
									</MessageBarBody>
								</MessageBar>
							)}
							{/* Managed system view: a warning, not a block (#142). Saving adds an unmanaged
							    layer, which sits on top of the managed layers and defines the view's behavior.
							    https://learn.microsoft.com/power-apps/maker/data-platform/solution-layers#remove-an-unmanaged-layer
							    https://learn.microsoft.com/power-platform/alm/solution-layers-alm */}
							{isManagedOverwrite && !isOverwriteBlocked && loadedView && (
								<MessageBar intent="warning" layout="multiline" className={styles.overwriteWarning}>
									<MessageBarBody>
										<MessageBarTitle>Managed view</MessageBarTitle>
										{managedSolution
											? `"${loadedView.name}" comes from the managed solution ${managedSolution.friendlyName} (${managedSolution.uniqueName}).`
											: `"${loadedView.name}" comes from a managed solution.`}{" "}
										Saving adds an unmanaged customization on top of it, which stays in effect until it's
										removed, even after the solution is updated.
									</MessageBarBody>
								</MessageBar>
							)}
							{/* A view someone else owns and shared with write access (#142) */}
							{isShared && !isOverwriteBlocked && loadedView && (
								<MessageBar intent="warning" layout="multiline" className={styles.overwriteWarning}>
									<MessageBarBody>
										<MessageBarTitle>Overwrite warning</MessageBarTitle>
										"{loadedView.name}" belongs to {loadedView.ownerName || "another user"}. Saving changes it for
										everyone it's shared with.
									</MessageBarBody>
								</MessageBar>
							)}
							{/* Overwrite Warning */}
							{isOverwriteMode && !isOverwriteBlocked && !isManagedOverwrite && !isShared && loadedView && (
								<MessageBar intent="warning" className={styles.overwriteWarning}>
									<MessageBarBody>
										<MessageBarTitle>Overwrite warning</MessageBarTitle>
										This will overwrite the existing view "{loadedView.name}".
									</MessageBarBody>
								</MessageBar>
							)}
							{/* Quick find placeholders are saved as written; the test value only runs (#142) */}
							{hasQuickFindPlaceholders && !isOverwriteBlocked && (
								<Text className={styles.placeholderNote}>
									<span className={styles.code}>{"{0}"}</span> placeholders are saved as written. The quick find test
									value is only used to run the query.
								</Text>
							)}
							{/* Publish Note */}
							{shouldPublish && viewType === "system" && (
								<Text className={styles.publishNote}>
									The view will be published after saving, making it immediately available to users.
								</Text>
							)}
						</div>
						)}

						{/* Save Error */}
						{saveError && (
							<MessageBar intent="error">
								<MessageBarBody>
									<MessageBarTitle>Save failed</MessageBarTitle>
									{saveError}
								</MessageBarBody>
							</MessageBar>
						)}
					</DialogContent>
					<DialogActions>
						{/* A view that can't be overwritten can be saved as a copy (#142) */}
						{isOverwriteBlocked && (
							<Button appearance="secondary" onClick={handleSaveCopy} disabled={isSaving}>
								Save a copy…
							</Button>
						)}
						<Button appearance="secondary" onClick={onClose} disabled={isSaving}>
							Cancel
						</Button>
						<Button appearance="primary" onClick={handleSave} disabled={!canSave}>
							{isSaving ? (
								<Spinner size="tiny" />
							) : shouldPublish && viewType === "system" ? (
								"Save and publish"
							) : (
								"Save"
							)}
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
