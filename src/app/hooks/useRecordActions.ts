/**
 * The record actions on the results (#301): Open, Copy URL, Delete, Delete all and Bulk delete, Run
 * workflow, Activate / Deactivate and Close, with their dialogs, the selection, and the reasons each
 * action is off.
 *
 * Two hooks, so AppShell's effect order stays as it was:
 * - useRecordActionAvailability checks the privileges (#179) and what Activate / Deactivate / Close
 *   allow (#143, #149) when the entity changes. AppShell calls it where these effects were, before
 *   the metadata loads.
 * - useRecordActions holds the dialogs and the handlers. It has no effects; AppShell calls it once
 *   the root entity's metadata and the query runner are there.
 *
 * Multi-request writes stop when the tool's connection changes in place (#175, #66): the sequential
 * delete passes connectionGuard() to each deleteRecord; the batch delete, Activate / Deactivate, Close
 * and Run workflow guard themselves in pptbClient.
 */
import { useCallback, useEffect, useMemo, useState } from "react";
import {
	isDataverseAvailable,
	getEnvironmentUrl,
	buildRecordUrl,
	deleteRecord,
	deleteRecordsBatch,
	checkBulkDeletePrivilege,
	submitBulkDelete,
	submitBulkDeleteFromFetchXml,
	checkWorkflowPrivileges,
	getOnDemandWorkflows,
	executeWorkflowBatch,
	checkDeletePrivilege,
	checkWritePrivilege,
	checkEntityPrivilege,
	getEntityStateModel,
	setRecordsState,
	closeRecords,
} from "../../features/fetchxml/api/pptbClient";
import type {
	EntityMetadata,
	WorkflowInfo,
	BatchDeleteProgress,
	BatchDeleteResult,
	WorkflowBatchProgress,
} from "../../features/fetchxml/api/pptbClient";
import {
	closeMessageReason,
	stateActionUnavailableReason,
	type StateAction,
	type StateTarget,
} from "../../features/fetchxml/model/recordState";
import {
	closeActionsFor,
	findCloseAction,
	type CloseAction,
	type CloseMessageAction,
	type CloseStateAction,
	type CloseValues,
} from "../../features/fetchxml/model/closeActions";
import type { CloseMenuItem } from "../../features/fetchxml/ui/RightPane/ResultsCommandBar";
import { generateFetchXml } from "../../features/fetchxml/model/fetchxml";
import type { PagingState } from "../../features/fetchxml/api/paging";
import { formatDataverseError } from "../../features/fetchxml/api/dataverseErrors";
import type { FetchNode, LinkEntityNode } from "../../features/fetchxml/model/nodes";
import type { QueryResult } from "../../features/fetchxml/ui/RightPane/ResultsGrid";
import {
	deleteAllUnavailableReason,
	deleteSelectionUnavailableReason,
	MAX_DIRECT_DELETE,
} from "../../features/fetchxml/model/deleteScope";
import {
	connectionGuard,
	copyText,
	holdClose,
	openUrl,
	whileHoldingClose,
} from "../../features/fetchxml/api/host";
import { debugLog } from "../../shared/utils/debug";

export interface RecordActionPrivileges {
	canDelete: boolean;
	canBulkDelete: boolean;
	canRunWorkflow: boolean;
	privilegesChecked: boolean;
}

export interface RecordActionAvailability {
	/** Delete (per entity), Bulk delete and Run workflow (per user); all off until checked (#179) */
	recordActionPrivileges: RecordActionPrivileges;
	/** Why Activate / Deactivate are off for the entity, per action (#143) */
	stateActionReasons: { activate?: string; deactivate?: string };
	/** The Close menu's items for close-message entities, or undefined (#149) */
	closeMenuItems: CloseMenuItem[] | undefined;
}

/**
 * The record action privileges and the Activate / Deactivate / Close availability for the entity
 * (#143, #149, #179). Checked again when the entity changes; a check for an earlier entity that
 * finishes last is dropped.
 */
export function useRecordActionAvailability(
	entityLogicalName: string | undefined,
): RecordActionAvailability {
	// Record action privileges state
	const [recordActionPrivileges, setRecordActionPrivileges] = useState<RecordActionPrivileges>({
		canDelete: false,
		canBulkDelete: false,
		canRunWorkflow: false,
		privilegesChecked: false,
	});

	// Why Activate / Deactivate are off for the root entity, per action (#143)
	const [stateActionReasons, setStateActionReasons] = useState<{
		activate?: string;
		deactivate?: string;
	}>({});

	// Close menu for close-message entities (#149): its items
	const [closeMenuItems, setCloseMenuItems] = useState<CloseMenuItem[] | undefined>();

	// Check record action privileges when entity changes
	useEffect(() => {
		// Until this entity's checks finish, nothing is allowed: keeping the previous entity's
		// privileges would enable Delete where it isn't allowed, or the other way round (#179)
		setRecordActionPrivileges({
			canDelete: false,
			canBulkDelete: false,
			canRunWorkflow: false,
			privilegesChecked: false,
		});
		if (!entityLogicalName || !isDataverseAvailable()) return;

		// A check for an entity that's no longer selected can finish last; its result is dropped (#179)
		let cancelled = false;
		const checkRecordPrivileges = async () => {
			try {
				// Delete is per entity; Bulk Delete and workflow privileges are per user
				const [canDelete, canBulkDelete, canRunWorkflow] = await Promise.all([
					checkDeletePrivilege(entityLogicalName),
					checkBulkDeletePrivilege(),
					checkWorkflowPrivileges(),
				]);
				if (cancelled) return;

				setRecordActionPrivileges({
					canDelete,
					canBulkDelete,
					canRunWorkflow,
					privilegesChecked: true,
				});

				debugLog("app", `🔐 Record action privileges for ${entityLogicalName}:`, {
					canDelete,
					canBulkDelete,
					canRunWorkflow,
				});
			} catch (error) {
				if (cancelled) return;
				console.error("Failed to check record action privileges:", error);
				setRecordActionPrivileges({
					canDelete: false,
					canBulkDelete: false,
					canRunWorkflow: false,
					privilegesChecked: true,
				});
			}
		};

		void checkRecordPrivileges();
		return () => {
			cancelled = true;
		};
	}, [entityLogicalName]);

	// Activate / Deactivate / Close availability (#143, #149): entities whose states need dedicated
	// messages (Close replaces Deactivate where it has actions), no statecode, or missing privileges:
	// Write on the entity, and Create on the close activity a message creates (the documented
	// privileges, e.g. CloseIncidentRequest: prvWriteIncident, prvCreateActivity). Undeterminable
	// answers leave the actions on; Dataverse decides.
	useEffect(() => {
		const activateClose = closeMessageReason(entityLogicalName, "activate");
		const deactivateClose = closeMessageReason(entityLogicalName, "deactivate");
		const actions = closeActionsFor(entityLogicalName);
		const toItems = (
			reasonFor: (action: CloseAction) => string | undefined,
		): CloseMenuItem[] | undefined =>
			actions?.map((a) => ({
				key: a.key,
				label: a.label,
				destructive: a.destructive,
				disabledReason: reasonFor(a),
			}));
		setCloseMenuItems(toItems(() => undefined));
		if (
			!entityLogicalName ||
			!isDataverseAvailable() ||
			(activateClose && deactivateClose && !actions)
		) {
			setStateActionReasons({ activate: activateClose, deactivate: deactivateClose });
			return;
		}
		let current = true;
		setStateActionReasons({ activate: activateClose, deactivate: deactivateClose });
		(async () => {
			const closeEntities = [
				...new Set(
					(actions ?? []).flatMap((a) => (a.kind === "message" ? [a.closeEntity] : [])),
				),
			];
			const [model, canWrite, canCreate] = await Promise.all([
				getEntityStateModel(entityLogicalName).catch((error: unknown) => {
					console.warn("Activate/Deactivate: couldn't read the state model:", error);
					return undefined;
				}),
				checkWritePrivilege(entityLogicalName),
				Promise.all(
					closeEntities.map(
						async (e) => [e, await checkEntityPrivilege(e, "Create")] as const,
					),
				).then((pairs) => new Map(pairs)),
			]);
			if (!current) return;
			const writeReason =
				canWrite === false
					? "You don't have the Write privilege on this entity, so you can't change its records' status."
					: undefined;
			setStateActionReasons({
				activate:
					activateClose ??
					(model !== undefined
						? stateActionUnavailableReason(model, "activate")
						: undefined) ??
					writeReason,
				deactivate:
					deactivateClose ??
					(model !== undefined
						? stateActionUnavailableReason(model, "deactivate")
						: undefined) ??
					writeReason,
			});
			setCloseMenuItems(
				toItems(
					(a) =>
						writeReason ??
						(a.kind === "message" && canCreate.get(a.closeEntity) === false
							? `You don't have the Create privilege on activities, and ${a.message} creates a ${a.closeEntity} activity.`
							: undefined),
				),
			);
		})();
		return () => {
			current = false;
		};
	}, [entityLogicalName]);

	return { recordActionPrivileges, stateActionReasons, closeMenuItems };
}

export interface DeleteDialogState {
	open: boolean;
	recordIds: string[];
	recordName?: string;
	isBatchDelete?: boolean; // Using batch DELETE for 4-100 records
}

export interface BulkDeleteDialogState {
	open: boolean;
	recordIds: string[];
	isAllRecords?: boolean;
	totalViewRecords?: number;
}

export interface WorkflowDialogState {
	open: boolean;
	recordIds: string[];
	preSelectedWorkflow?: WorkflowInfo;
}

export interface SetStateDialogState {
	open: boolean;
	action: StateAction;
	recordIds: string[];
	closeAction?: CloseStateAction;
}

export interface CloseDialogState {
	open: boolean;
	action: CloseMessageAction | null;
	recordIds: string[];
}

export interface RecordActionsInputs {
	/** The selected entity (the tree's root <entity>) */
	entityLogicalName: string | undefined;
	/** The builder's query: aggregate, distinct, 1:N / N:N link-entities, and Delete all's check (#174) */
	fetchQuery: FetchNode | null;
	/** The root entity's metadata: primary key and name, for record names, Bulk delete and Close */
	entityMetadata: EntityMetadata | null;
	/** The results shown in the grid, and the error of the run that produced them */
	queryResult: QueryResult | null;
	executeError: string | undefined;
	/** The current paging state; its xml is the FetchXML that ran, for Delete all (#142) */
	pagingRef: { readonly current: PagingState | null };
	/** The latest handleExecute (AppShell): the query runs again after records change */
	handleExecuteRef: { readonly current: () => Promise<void> };
	/** From useRecordActionAvailability */
	availability: RecordActionAvailability;
}

export interface RecordActions {
	canDelete: boolean;
	canBulkDelete: boolean;
	canRunWorkflow: boolean;
	isAggregateQuery: boolean;
	/** 1:N or N:N link-entities repeat rows, so record commands are off */
	hasOneToManyRelationship: boolean;
	/** Why every record action is off (no record IDs in the results), or undefined */
	recordActionsDisabledReason: string | undefined;
	/** Why Delete all is off (#174), or undefined */
	deleteAllDisabledReason: string | undefined;
	activateDisabledReason: string | undefined;
	deactivateDisabledReason: string | undefined;
	closeActions: CloseMenuItem[] | undefined;
	closeDisabledReason: string | undefined;

	/** The grid's selection */
	onSelectionChange: (recordIds: string[]) => void;
	getSelectedRecordIds: () => string[];

	/** The command bar's actions, on the given records */
	openRecord: (recordIds: string[]) => Promise<void>;
	copyRecordUrl: (recordIds: string[]) => Promise<void>;
	/** No records: Delete all; 1-100: the Delete dialog; more: the Bulk delete dialog */
	deleteRecords: (recordIds: string[]) => void;
	bulkDeleteRecords: (recordIds: string[]) => void;
	activateRecords: (recordIds: string[]) => void;
	deactivateRecords: (recordIds: string[]) => void;
	closeRecords: (actionKey: string, recordIds: string[]) => void;
	runSpecificWorkflow: (workflow: WorkflowInfo, recordIds: string[]) => void;
	fetchWorkflows: (forEntity?: string) => Promise<WorkflowInfo[]>;

	/** The Delete dialog */
	deleteDialogState: DeleteDialogState;
	deleteProgress: BatchDeleteProgress | null;
	confirmDelete: () => Promise<BatchDeleteResult | void>;
	dismissDeleteDialog: () => void;
	/** The Bulk delete dialog */
	bulkDeleteDialogState: BulkDeleteDialogState;
	confirmBulkDelete: (
		jobName: string,
	) => Promise<{ asyncOperationId: string; bulkDeleteOperationId: string; jobUrl: string }>;
	dismissBulkDeleteDialog: () => void;
	/** The Run workflow dialog */
	workflowDialogState: WorkflowDialogState;
	executeWorkflow: (
		workflow: WorkflowInfo,
		recordIds: string[],
		onProgress: (progress: WorkflowBatchProgress) => void,
	) => Promise<{ succeeded: number; failed: number; errors: string[] }>;
	dismissWorkflowDialog: () => void;
	/** The Activate / Deactivate dialog (#143), also for the Close menu's state changes (#149) */
	setStateDialogState: SetStateDialogState;
	executeStateChange: (
		target: StateTarget,
		onProgress: (progress: BatchDeleteProgress) => void,
	) => Promise<BatchDeleteResult>;
	dismissSetStateDialog: () => void;
	/** The close message dialog (#149) */
	closeDialogState: CloseDialogState;
	executeClose: (
		values: CloseValues,
		onProgress: (progress: BatchDeleteProgress) => void,
	) => Promise<BatchDeleteResult>;
	dismissCloseDialog: () => void;
}

export function useRecordActions({
	entityLogicalName,
	fetchQuery,
	entityMetadata,
	queryResult,
	executeError,
	pagingRef,
	handleExecuteRef,
	availability,
}: RecordActionsInputs): RecordActions {
	const { recordActionPrivileges, stateActionReasons, closeMenuItems } = availability;

	// Dialog states for record actions
	const [deleteDialogState, setDeleteDialogState] = useState<DeleteDialogState>({
		open: false,
		recordIds: [],
	});

	// Progress tracking for batch delete
	const [deleteProgress, setDeleteProgress] = useState<BatchDeleteProgress | null>(null);

	const [bulkDeleteDialogState, setBulkDeleteDialogState] = useState<BulkDeleteDialogState>({
		open: false,
		recordIds: [],
	});

	const [workflowDialogState, setWorkflowDialogState] = useState<WorkflowDialogState>({
		open: false,
		recordIds: [],
	});

	// Activate / Deactivate confirmation (#143); also the Close menu's plain state changes (#149)
	const [setStateDialogState, setSetStateDialogState] = useState<SetStateDialogState>({
		open: false,
		action: "activate",
		recordIds: [],
	});

	// The close message dialog for close-message entities (#149)
	const [closeDialogState, setCloseDialogState] = useState<CloseDialogState>({
		open: false,
		action: null,
		recordIds: [],
	});

	// Currently selected record IDs (from ResultsGrid)
	const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

	// Check if current query is an aggregate query (disables delete/workflow buttons)
	const isAggregateQuery = fetchQuery?.options?.aggregate === true;

	// Check if query has 1-N or N-N relationships that cause row duplication
	// When these relationships exist, record commands should be disabled
	const hasOneToManyRelationship = useMemo(() => {
		if (!fetchQuery?.entity?.links) return false;

		const checkLinks = (links: LinkEntityNode[]): boolean => {
			for (const link of links) {
				// Check if this link-entity is a 1-N or N-N relationship
				if (link.relationshipType === "1N" || link.relationshipType === "NN") {
					return true;
				}
				// Check nested link-entities
				if (link.links && checkLinks(link.links)) {
					return true;
				}
			}
			return false;
		};

		return checkLinks(fetchQuery.entity.links);
	}, [fetchQuery]);

	// Record actions need the root table's primary key in the results. Dataverse returns it for
	// normal queries even when not selected; aggregate and distinct results don't include it (#59).
	// Decide from the rows actually returned rather than from the query shape.
	const primaryIdAttribute = entityMetadata?.PrimaryIdAttribute;
	const isDistinctQuery = fetchQuery?.options?.distinct === true;
	const recordActionsDisabledReason = useMemo(() => {
		const rows = queryResult?.rows;
		if (!rows || rows.length === 0 || !primaryIdAttribute) return undefined;
		if (rows.some((row) => typeof row[primaryIdAttribute] === "string")) return undefined;
		if (isAggregateQuery) {
			return "Aggregate query results don't contain record IDs, so record actions aren't available.";
		}
		if (isDistinctQuery) {
			return "Distinct query results don't contain record IDs (Dataverse doesn't return the primary key for distinct queries), so record actions aren't available.";
		}
		return "These results don't include the table's primary key, so record actions aren't available.";
	}, [queryResult, primaryIdAttribute, isAggregateQuery, isDistinctQuery]);

	// "Delete all" (no selection) deletes what the query on screen matches (#174): the results must
	// come from a successful run of a query that matches the same records as the tree. Editing a
	// filter in the left pane doesn't re-run the query, so the old rows would still be on screen.
	const deleteAllDisabledReason = useMemo(
		() => deleteAllUnavailableReason(executeError ? null : queryResult?.query, fetchQuery),
		[executeError, queryResult?.query, fetchQuery],
	);

	/**
	 * Handle selection change from ResultsGrid
	 */
	const handleSelectionChange = useCallback((recordIds: string[]) => {
		setSelectedRecordIds(recordIds);
	}, []);

	/**
	 * Get selected record IDs - called by command bar actions
	 */
	const getSelectedRecordIds = useCallback(() => {
		return selectedRecordIds;
	}, [selectedRecordIds]);

	/**
	 * Open record in a new browser tab
	 */
	const handleOpenRecord = useCallback(
		async (recordIds: string[]) => {
			if (recordIds.length === 0 || !entityLogicalName) return;

			const envUrl = await getEnvironmentUrl();
			if (!envUrl) {
				console.error("Failed to get environment URL");
				return;
			}

			// Open each selected record in the connection's browser profile, already signed in to
			// the right tenant (#70)
			for (const recordId of recordIds) {
				const url = buildRecordUrl(entityLogicalName, recordId, envUrl);
				await openUrl(url);
			}
		},
		[entityLogicalName],
	);

	/**
	 * Copy record URL(s) to clipboard
	 */
	const handleCopyRecordUrl = useCallback(
		async (recordIds: string[]) => {
			if (recordIds.length === 0 || !entityLogicalName) return;

			const envUrl = await getEnvironmentUrl();
			if (!envUrl) {
				console.error("Failed to get environment URL");
				return;
			}

			const urls = recordIds.map((id) => buildRecordUrl(entityLogicalName, id, envUrl));
			const textToCopy = urls.join("\n");

			if (await copyText(textToCopy)) {
				debugLog("app", `📋 Copied ${urls.length} record URL(s) to clipboard`);
			} else {
				console.error("Failed to copy to clipboard");
			}
		},
		[entityLogicalName],
	);

	/**
	 * Open delete confirmation dialog
	 * Handles 3 scenarios:
	 * 1. No selection → Bulk delete ALL records from view (with warning)
	 * 2. 1-100 records → Batch DELETE requests (with progress)
	 * 3. >100 records → Bulk delete with FetchXML IN operator
	 */
	const handleDeleteRecords = useCallback(
		(recordIds: string[]) => {
			if (!entityLogicalName) return;

			// Scenario 1: No selection - delete ALL records from view
			if (recordIds.length === 0) {
				if (deleteAllDisabledReason) return; // the button explains why (#174)
				if (!recordActionPrivileges.canBulkDelete) {
					debugLog(
						"app",
						"Bulk delete privilege required to delete all records from view",
					);
					return;
				}
				const totalRecords = queryResult?.rows.length || 0;
				setBulkDeleteDialogState({
					open: true,
					recordIds: [],
					isAllRecords: true,
					totalViewRecords: totalRecords,
				});
				return;
			}

			// Scenario 3: More than 100 records - use bulk delete with IN operator. Without prvBulkDelete
			// it's blocked, not deleted record by record (#177); the command bar says why.
			if (recordIds.length > MAX_DIRECT_DELETE) {
				if (
					deleteSelectionUnavailableReason(
						recordIds.length,
						recordActionPrivileges.canBulkDelete,
					)
				)
					return;
				setBulkDeleteDialogState({
					open: true,
					recordIds,
					isAllRecords: false,
				});
				return;
			}

			// Scenario 2: 1-100 records - use batch DELETE or direct delete
			const recordName =
				recordIds.length === 1
					? (queryResult?.rows.find(
							(row) => row[entityMetadata?.PrimaryIdAttribute || ""] === recordIds[0],
						)?.[entityMetadata?.PrimaryNameAttribute || ""] as string | undefined)
					: undefined;

			// Use batch delete for 4+ records (more efficient than sequential)
			const isBatchDelete = recordIds.length >= 4;
			setDeleteDialogState({ open: true, recordIds, recordName, isBatchDelete });
		},
		[
			entityLogicalName,
			recordActionPrivileges.canBulkDelete,
			queryResult,
			entityMetadata,
			deleteAllDisabledReason,
		],
	);

	/**
	 * Explicitly trigger bulk delete dialog (from command bar menu)
	 * Bypasses the normal auto-routing logic and always uses bulk delete
	 */
	const handleBulkDeleteRecords = useCallback(
		(recordIds: string[]) => {
			if (!entityLogicalName || !recordActionPrivileges.canBulkDelete) return;

			const isAllRecords = recordIds.length === 0;
			if (isAllRecords && deleteAllDisabledReason) return; // the button explains why (#174)
			const totalRecords = queryResult?.rows.length || 0;

			setBulkDeleteDialogState({
				open: true,
				recordIds,
				isAllRecords,
				totalViewRecords: isAllRecords ? totalRecords : undefined,
			});
		},
		[
			entityLogicalName,
			recordActionPrivileges.canBulkDelete,
			queryResult,
			deleteAllDisabledReason,
		],
	);

	/**
	 * Execute delete for records (called from DeleteConfirmDialog)
	 * Supports both sequential delete (1-3 records) and batch delete (4-100 records)
	 */
	const handleConfirmDelete = useCallback(async (): Promise<BatchDeleteResult | void> => {
		if (!entityLogicalName || deleteDialogState.recordIds.length === 0) return;

		const recordIds = deleteDialogState.recordIds;
		setDeleteProgress(null);
		// The host warns before the tab closes while records are being deleted (#70)
		const releaseClose = holdClose();
		try {
			// For batch delete (4+ records): $batch requests with the metadata EntitySetName on
			// PPTB ≥ 1.2.7 (#61)
			if (deleteDialogState.isBatchDelete) {
				const result = await deleteRecordsBatch(entityLogicalName, recordIds, (progress) =>
					setDeleteProgress(progress),
				);

				if (result.succeeded > 0) {
					debugLog(
						"app",
						`✅ Batch deleted ${result.succeeded} of ${recordIds.length} records`,
					);
					// Re-execute query to refresh results
					void handleExecuteRef.current();
				}

				return result;
			}

			// Sequential delete for 1-3 records
			let successCount = 0;
			const errors: string[] = [];

			// Stop if the tool's connection changes in place (#175, #66): the next delete would run in
			// the new environment, where the same record IDs can exist
			const stillOnConnection = await connectionGuard();
			for (const recordId of recordIds) {
				try {
					await deleteRecord(entityLogicalName, recordId, stillOnConnection);
					successCount++;
				} catch (error) {
					errors.push(
						`Failed to delete ${recordId}: ${formatDataverseError(error)}`, // (#96)
					);
				}
			}

			if (successCount > 0) {
				debugLog("app", `✅ Deleted ${successCount} of ${recordIds.length} records`);
				// Re-execute query to refresh results
				void handleExecuteRef.current();
			}

			if (errors.length > 0) {
				console.error("Delete errors:", errors);
				return { succeeded: successCount, failed: errors.length, errors };
			}
		} finally {
			releaseClose();
		}
	}, [
		entityLogicalName,
		deleteDialogState.recordIds,
		deleteDialogState.isBatchDelete,
		handleExecuteRef,
	]);

	/**
	 * Submit bulk delete job (called from BulkDeleteDialog)
	 * Supports both selected records (IN operator) and all records from view
	 */
	const handleConfirmBulkDelete = useCallback(
		async (jobName: string) => {
			if (!entityLogicalName || !entityMetadata) {
				throw new Error("Missing entity information for bulk delete");
			}

			// If deleting ALL records from view (no selection)
			if (bulkDeleteDialogState.isAllRecords) {
				// The query that produced the results on screen, never the tree after later edits (#174)
				const executed = queryResult?.query;
				if (!executed || deleteAllDisabledReason) {
					throw new Error(
						deleteAllDisabledReason ??
							"Run the query first: Delete all deletes the records the query on screen matches.",
					);
				}
				// The FetchXML that ran: a quick find view's test value replaced {0} only there (#142), so
				// the tree's XML would still say {0}
				const currentFetchXml = pagingRef.current?.xml ?? generateFetchXml(executed);
				const result = await submitBulkDeleteFromFetchXml(currentFetchXml, jobName);
				debugLog(
					"app",
					`📤 Bulk delete job (all records) submitted: ${result.asyncOperationId}`,
				);
				return result;
			}

			// If deleting specific selected records
			if (bulkDeleteDialogState.recordIds.length === 0) {
				throw new Error("No records selected for bulk delete");
			}

			const recordIds = bulkDeleteDialogState.recordIds;
			const primaryIdAttribute = entityMetadata.PrimaryIdAttribute;

			const result = await submitBulkDelete(
				entityLogicalName,
				primaryIdAttribute,
				recordIds,
				jobName,
			);

			debugLog("app", `📤 Bulk delete job submitted: ${result.asyncOperationId}`);
			return result;
		},
		[
			entityLogicalName,
			entityMetadata,
			bulkDeleteDialogState.recordIds,
			bulkDeleteDialogState.isAllRecords,
			queryResult?.query,
			deleteAllDisabledReason,
			pagingRef,
		],
	);

	/**
	 * Run a specific workflow directly (from menu button)
	 * Opens the workflow dialog with the workflow pre-selected for confirmation
	 */
	const handleRunSpecificWorkflow = useCallback((workflow: WorkflowInfo, recordIds: string[]) => {
		if (recordIds.length === 0) return;
		// Open the workflow dialog with the selected workflow pre-selected
		setWorkflowDialogState({ open: true, recordIds, preSelectedWorkflow: workflow });
	}, []);

	/**
	 * Fetch available workflows for an entity: the one the Run workflow menu names (the results'
	 * entity, #182), or the selected entity
	 */
	const handleFetchWorkflows = useCallback(
		async (forEntity?: string): Promise<WorkflowInfo[]> => {
			const entity = forEntity ?? entityLogicalName;
			if (!entity) return [];
			return getOnDemandWorkflows(entity);
		},
		[entityLogicalName],
	);

	/**
	 * Execute workflow on selected records
	 */
	const handleExecuteWorkflow = useCallback(
		async (
			workflow: WorkflowInfo,
			recordIds: string[],
			onProgress: (progress: WorkflowBatchProgress) => void,
		): Promise<{ succeeded: number; failed: number; errors: string[] }> => {
			if (!entityLogicalName) {
				return { succeeded: 0, failed: recordIds.length, errors: ["Entity not selected"] };
			}
			// The host warns before the tab closes while workflows are being started (#70)
			return whileHoldingClose(() =>
				executeWorkflowBatch(workflow, recordIds, entityLogicalName, onProgress),
			);
		},
		[entityLogicalName],
	);

	/** Activate / Deactivate (#143): opens the confirmation, which plans the change from metadata */
	const openSetStateDialog = useCallback(
		(action: StateAction, recordIds: string[]) => {
			if (!entityLogicalName || recordIds.length === 0) return;
			setSetStateDialogState({ open: true, action, recordIds });
		},
		[entityLogicalName],
	);
	const handleActivateRecords = useCallback(
		(ids: string[]) => openSetStateDialog("activate", ids),
		[openSetStateDialog],
	);
	const handleDeactivateRecords = useCallback(
		(ids: string[]) => openSetStateDialog("deactivate", ids),
		[openSetStateDialog],
	);

	/** Runs the state change (#143): PATCH statecode / statuscode, holding the host's close warning */
	const handleExecuteStateChange = useCallback(
		async (
			target: StateTarget,
			onProgress: (progress: BatchDeleteProgress) => void,
		): Promise<BatchDeleteResult> => {
			if (!entityLogicalName) throw new Error("Entity not selected");
			const result = await whileHoldingClose(() =>
				setRecordsState(
					entityLogicalName,
					setStateDialogState.recordIds,
					target,
					onProgress,
				),
			);
			if (result.succeeded > 0) {
				debugLog(
					"app",
					`✅ State changed on ${result.succeeded} of ${setStateDialogState.recordIds.length} records`,
				);
				void handleExecuteRef.current();
			}
			return result;
		},
		[entityLogicalName, setStateDialogState.recordIds, handleExecuteRef],
	);

	/**
	 * Close menu (#149): plain state changes (Cancel case, Mark as paid, Cancel invoice) open the
	 * Activate / Deactivate confirmation limited to their state; close messages open their dialog
	 */
	const handleCloseRecords = useCallback(
		(actionKey: string, recordIds: string[]) => {
			const action = findCloseAction(entityLogicalName, actionKey);
			if (!action || recordIds.length === 0) return;
			if (action.kind === "state") {
				setSetStateDialogState({
					open: true,
					action: "deactivate",
					recordIds,
					closeAction: action,
				});
			} else {
				setCloseDialogState({ open: true, action, recordIds });
			}
		},
		[entityLogicalName],
	);

	/** Runs the close message on each record (#149), holding the host's close warning */
	const handleExecuteClose = useCallback(
		async (
			values: CloseValues,
			onProgress: (progress: BatchDeleteProgress) => void,
		): Promise<BatchDeleteResult> => {
			const action = closeDialogState.action;
			if (!entityLogicalName || !action) throw new Error("Entity not selected");
			const result = await whileHoldingClose(() =>
				closeRecords(
					entityLogicalName,
					entityMetadata?.PrimaryIdAttribute,
					closeDialogState.recordIds,
					action,
					values,
					onProgress,
				),
			);
			if (result.succeeded > 0) {
				debugLog(
					"app",
					`✅ ${action.message} on ${result.succeeded} of ${closeDialogState.recordIds.length} records`,
				);
				void handleExecuteRef.current();
			}
			return result;
		},
		[
			entityLogicalName,
			entityMetadata?.PrimaryIdAttribute,
			closeDialogState.action,
			closeDialogState.recordIds,
			handleExecuteRef,
		],
	);

	// Rows repeat for 1:N / N:N link-entities; Activate / Deactivate say so like their other reasons (#143)
	const oneToManyStateReason = hasOneToManyRelationship
		? "The query has a one-to-many or many-to-many link-entity, so rows repeat. Remove it to activate or deactivate records."
		: undefined;
	const oneToManyCloseReason = hasOneToManyRelationship
		? "The query has a one-to-many or many-to-many link-entity, so rows repeat. Remove it to close records."
		: undefined;

	return {
		canDelete: recordActionPrivileges.canDelete,
		canBulkDelete: recordActionPrivileges.canBulkDelete,
		canRunWorkflow: recordActionPrivileges.canRunWorkflow,
		isAggregateQuery,
		hasOneToManyRelationship,
		recordActionsDisabledReason,
		deleteAllDisabledReason,
		activateDisabledReason: stateActionReasons.activate ?? oneToManyStateReason,
		deactivateDisabledReason: stateActionReasons.deactivate ?? oneToManyStateReason,
		closeActions: closeMenuItems,
		closeDisabledReason: oneToManyCloseReason,

		onSelectionChange: handleSelectionChange,
		getSelectedRecordIds,

		openRecord: handleOpenRecord,
		copyRecordUrl: handleCopyRecordUrl,
		deleteRecords: handleDeleteRecords,
		bulkDeleteRecords: handleBulkDeleteRecords,
		activateRecords: handleActivateRecords,
		deactivateRecords: handleDeactivateRecords,
		closeRecords: handleCloseRecords,
		runSpecificWorkflow: handleRunSpecificWorkflow,
		fetchWorkflows: handleFetchWorkflows,

		deleteDialogState,
		deleteProgress,
		confirmDelete: handleConfirmDelete,
		// New functions every render, as the inline onClose handlers were
		dismissDeleteDialog: () => {
			setDeleteDialogState({ open: false, recordIds: [] });
			setDeleteProgress(null);
		},
		bulkDeleteDialogState,
		confirmBulkDelete: handleConfirmBulkDelete,
		dismissBulkDeleteDialog: () => setBulkDeleteDialogState({ open: false, recordIds: [] }),
		workflowDialogState,
		executeWorkflow: handleExecuteWorkflow,
		dismissWorkflowDialog: () => setWorkflowDialogState({ open: false, recordIds: [] }),
		setStateDialogState,
		executeStateChange: handleExecuteStateChange,
		dismissSetStateDialog: () => setSetStateDialogState((prev) => ({ ...prev, open: false })),
		closeDialogState,
		executeClose: handleExecuteClose,
		dismissCloseDialog: () => setCloseDialogState((prev) => ({ ...prev, open: false })),
	};
}
