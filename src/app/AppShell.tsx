/**
 * Main application shell with Fluent UI theming and layout
 */

import { useState, useEffect, useLayoutEffect, useMemo, useCallback, useRef } from "react";
import {
	FluentProvider,
	webLightTheme,
	webDarkTheme,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { usePptbContext } from "../shared/hooks/usePptbContext";
import { useLazyMetadata } from "../shared/hooks/useLazyMetadata";
import { loadAllEntityAttributes } from "../features/fetchxml/api/dataverseMetadata";
import { metadataScope, useAdvancedFindOnly } from "../features/fetchxml/state/metadataScope";
import { PreviewTabs } from "../features/fetchxml/ui/RightPane/PreviewTabs";
import type { RelatedEntityColumn } from "../features/fetchxml/ui/RightPane/AddColumnsPanel";
import { EntitySelector } from "../features/fetchxml/ui/Toolbar/EntitySelector";
import { SaveViewButton } from "../features/fetchxml/ui/Toolbar/SaveViewButton";
import { SendToToolButton } from "../features/fetchxml/ui/Toolbar/SendToToolButton";
import { isT2TSupported, resolveLaunchPrefill } from "../features/fetchxml/api/invocation";
import { useLaunchContext } from "../shared/hooks/useLaunchContext";
import { useSendToTools } from "../shared/hooks/useSendToTools";
import { TreeView } from "../features/fetchxml/ui/LeftPane/TreeView";
import { PropertiesPanel } from "../features/fetchxml/ui/LeftPane/PropertiesPanel";
import { clearMetadataLoadErrors } from "../features/fetchxml/ui/LeftPane/PropertiesPanel/metadataLoadErrors";
import { BuilderProvider } from "../features/fetchxml/state/builderStore";
import { useBuilder } from "../features/fetchxml/state/builderContext";
import { ThemeProvider } from "../shared/contexts/ThemeContext";
import { PaneErrorBoundary, ToolErrorBoundary } from "../shared/components/ErrorBoundary";
import { DeleteConfirmDialog } from "../features/fetchxml/ui/Dialogs/DeleteConfirmDialog";
import { BulkDeleteDialog } from "../features/fetchxml/ui/Dialogs/BulkDeleteDialog";
import { WorkflowPickerDialog } from "../features/fetchxml/ui/Dialogs/WorkflowPickerDialog";
import {
	executeFetchXml,
	executeNextLink,
	executeSystemView,
	executePersonalView,
	whoAmI,
	isDataverseAvailable,
	exportToExcel,
	downloadBase64File,
	checkPrivilegeByName,
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
	loadFormattingMetadata,
	loadUserTimeZone,
} from "../features/fetchxml/api/pptbClient";
import {
	closeMessageReason,
	stateActionUnavailableReason,
	type StateAction,
	type StateTarget,
} from "../features/fetchxml/model/recordState";
import {
	closeActionsFor,
	findCloseAction,
	type CloseAction,
	type CloseMessageAction,
	type CloseStateAction,
	type CloseValues,
} from "../features/fetchxml/model/closeActions";
import { SetStateDialog } from "../features/fetchxml/ui/Dialogs/SetStateDialog";
import { CloseRecordsDialog } from "../features/fetchxml/ui/Dialogs/CloseRecordsDialog";
import type { CloseMenuItem } from "../features/fetchxml/ui/RightPane/ResultsCommandBar";
import type {
	AttributeMetadata,
	LoadedViewInfo,
	EntityMetadata,
	RelationshipMetadata,
	WorkflowInfo,
	BatchDeleteProgress,
	BatchDeleteResult,
	WorkflowBatchProgress,
} from "../features/fetchxml/api/pptbClient";
import { generateFetchXml, addPagingToFetchXml } from "../features/fetchxml/model/fetchxml";
import {
	advancePaging,
	hasNextPage,
	readFetchPagingOptions,
	retrieveAllLimitReached,
	RETRIEVE_ALL_LIMIT_MESSAGE,
	type PagingNotice,
	type PagingState,
} from "../features/fetchxml/api/paging";
import type { FetchXmlResult } from "../features/fetchxml/api/pptbClient";
import { formatDataverseError } from "../features/fetchxml/api/dataverseErrors";
import { DATAVERSE_DEFAULT_PAGE_SIZE, fetchPageWithRecovery } from "../features/fetchxml/api/pageRecovery";
import { LONG_RUNNING_RETRY, withThrottleRetry, type ThrottleRetryOptions } from "../features/fetchxml/api/throttleRetry";
import { debugLog } from "../shared/utils/debug";
import { getQueryEntities } from "../features/fetchxml/model/treeUtils";
import { NON_RETURNING_LINK_TYPES } from "../features/fetchxml/model/fetchxmlRules";
import { planRemoveColumns, planSortByColumn, resolveColumn } from "../features/fetchxml/model/queryColumns";
import { orderByLayout, resultColumns } from "../features/fetchxml/model/resultColumns";
import { relationshipKey } from "../features/fetchxml/model/manyToMany";
import { columnFilterConditions, type NewCondition } from "../features/fetchxml/model/columnFilters";
import { toAttributeTypeInfo } from "../features/fetchxml/model/attributeCapabilities";
import type { FetchNode, LinkEntityNode } from "../features/fetchxml/model/nodes";
import { buildViewLayoutXml } from "../features/fetchxml/model/layoutxml";
import {
	applyQuickFindValue,
	countQuickFindPlaceholders,
	getViewLoadNotice,
	isUnmodifiedView,
	toQuickFindSearchValue,
} from "../features/fetchxml/model/savedViews";
import type { SavedViewDetails } from "../features/fetchxml/ui/Dialogs/SaveViewDialog";
import { QuickFindTestValue } from "../features/fetchxml/ui/RightPane/QuickFindTestValue";
import type { QueryResult } from "../features/fetchxml/ui/RightPane/ResultsGrid";
import { deleteAllUnavailableReason, deleteSelectionUnavailableReason, MAX_DIRECT_DELETE } from "../features/fetchxml/model/deleteScope";
import { SettingsDrawer } from "../features/fetchxml/ui/Settings/SettingsDrawer";
import { usePersistedSettings } from "../shared/hooks/usePersistedSettings";
import { treeSizeFor } from "../features/fetchxml/model/displaySettings";
import { connectionGuard, copyText, holdClose, openUrl, whileHoldingClose } from "../features/fetchxml/api/host";
import { metadataCache } from "../features/fetchxml/state/cache";
import { resetDataverseSessionCache } from "../features/fetchxml/api/sessionCache";
import { nextConnectionScope, type ConnectionScope } from "../shared/hooks/connectionIdentity";

const SIMPLE_PAGING_MESSAGE =
	"Dataverse didn't return a paging cookie for this query (for example, it's sorted by a link-entity attribute or uses distinct), so pages are fetched with page numbers only. Microsoft recommends paging cookies; this works, but is slower for large result sets.";

// ⚠️ IMPORTANT: makeStyles must be called OUTSIDE the component
// but tokens will automatically update when FluentProvider theme changes
// because Fluent UI v9 uses CSS variables under the hood
const useStyles = makeStyles({
	root: {
		// The theme background covers the whole tool; without it, surfaces with no background of
		// their own showed the page's white canvas in dark mode (#67). (Not on the FluentProvider:
		// Fluent copies its className onto portal containers.)
		backgroundColor: tokens.colorNeutralBackground1,
		color: tokens.colorNeutralForeground1,
		display: "flex",
		flexDirection: "row",
		height: "100%",
		width: "100%",
		overflow: "hidden",
		position: "absolute",
		top: 0,
		left: 0,
		right: 0,
		bottom: 0,
	},
	leftPane: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minWidth: "300px",
		maxWidth: "800px",
		height: "100%",
	},
	leftPaneTop: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minHeight: 0, // Important for flex child with overflow
	},
	horizontalResizeHandle: {
		height: "6px",
		cursor: "ns-resize",
		backgroundColor: tokens.colorNeutralBackground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		position: "relative",
		flexShrink: 0,
		":hover": {
			backgroundColor: tokens.colorNeutralBackground3Hover,
		},
		":active": {
			backgroundColor: tokens.colorBrandBackgroundPressed,
		},
	},
	horizontalGrip: {
		fontSize: "10px",
		color: tokens.colorNeutralForeground3,
		userSelect: "none",
		pointerEvents: "none",
		display: "flex",
		gap: "3px",
		alignItems: "center",
	},
	verticalResizeHandle: {
		width: "6px",
		cursor: "ew-resize",
		backgroundColor: tokens.colorNeutralBackground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		position: "relative",
		flexShrink: 0,
		":hover": {
			backgroundColor: tokens.colorNeutralBackground3Hover,
		},
		":active": {
			backgroundColor: tokens.colorBrandBackgroundPressed,
		},
	},
	verticalGrip: {
		fontSize: "10px",
		color: tokens.colorNeutralForeground3,
		userSelect: "none",
		pointerEvents: "none",
		display: "flex",
		flexDirection: "column",
		gap: "2px",
		lineHeight: "0.5",
	},
	leftPaneBottom: {
		display: "flex",
		flexDirection: "column",
		minHeight: "100px",
		overflow: "hidden",
		minWidth: 0, // Important for flex child with overflow
	},
	propertiesContent: {
		flex: 1,
		overflow: "hidden",
	},
	rightPane: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minWidth: 0,
	},
	toolbar: {
		padding: "12px 16px",
		borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		display: "flex",
		gap: "8px",
		alignItems: "center",
		flexShrink: 0,
	},
	placeholder: {
		padding: "16px",
		color: tokens.colorNeutralForeground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
	},
});

/**
 * What outlives a connection change (#66): AppContent is remounted for the new connection, so
 * this lives in AppShell.
 */
interface ConnectionCarryOver {
	/** The launch prefill (T2T) was applied: never apply it again */
	prefillConsumed: boolean;
	/** Reads the current query's FetchXML (set by AppContent) */
	readFetchXml?: () => string | undefined;
	/** The query to load into the remounted AppContent */
	fetchXml?: string;
}

export function AppShell() {
	const { theme, connectionKey, connectionName } = usePptbContext();
	// The PPTB host's theme (light / dark), resolved by the host (#67)
	const isDark = theme === "dark";

	// Native controls (scrollbars, date inputs, autofill) follow the theme too
	useEffect(() => {
		document.documentElement.style.colorScheme = isDark ? "dark" : "light";
	}, [isDark]);

	// Connection changed in place (PPTB ≥ 1.2.7, #66): everything loaded from the previous
	// environment is stale (metadata cache, entity lists, publishers and solutions, privileges,
	// results, the loaded view). Clear the cache and remount the builder and the app content as a
	// new session; only the query itself is kept. Derived while rendering, so the cache is clear
	// before the new tree's effects load metadata.
	const carryOver = useRef<ConnectionCarryOver>({ prefillConsumed: false });
	const [connectionScope, setConnectionScope] = useState<ConnectionScope>({ key: undefined, session: 0 });
	const nextScope = nextConnectionScope(connectionScope, connectionKey);
	if (nextScope.scope !== connectionScope) {
		if (nextScope.changed) {
			metadataCache.clear();
			// Failed metadata loads belonged to the previous connection (#201)
			clearMetadataLoadErrors();
			// WhoAmI and privilege answers belong to the previous user and environment (#71)
			resetDataverseSessionCache();
			carryOver.current.fetchXml = carryOver.current.readFetchXml?.();
		}
		setConnectionScope(nextScope.scope);
	}

	// Tell the user why the tool reset (a host notification; no new UI)
	const session = connectionScope.session;
	useEffect(() => {
		if (session === 0) return;
		void window.toolboxAPI?.utils?.showNotification?.({
			title: "Connection changed",
			body: connectionName
				? `FetchXML Studio reloaded its metadata for ${connectionName}. Results and the loaded view were cleared.`
				: "FetchXML Studio has no connection now. Results and the loaded view were cleared.",
			type: "info",
		});
		// Only when the session changes; the name is read for that session
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [session]);

	// After a render error, the tool is re-mounted like after a connection change (#185): with the
	// last good query, or with a new one
	const [restarts, setRestarts] = useState(0);
	const restart = (withQuery: boolean) => {
		carryOver.current.fetchXml = withQuery ? carryOver.current.readFetchXml?.() : undefined;
		setRestarts((n) => n + 1);
	};

	return (
		<FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
			<ThemeProvider isDark={isDark}>
				<ToolErrorBoundary readFetchXml={() => carryOver.current.readFetchXml?.()} onRestart={restart}>
					<BuilderProvider key={`${session}-${restarts}`}>
						<AppContent carryOver={carryOver.current} connectionKey={connectionKey} />
					</BuilderProvider>
				</ToolErrorBoundary>
			</ThemeProvider>
		</FluentProvider>
	);
}

/** connectionKey: "" with no connection, undefined until the host answers (usePptbContext) */
function AppContent({ carryOver, connectionKey }: { carryOver: ConnectionCarryOver; connectionKey?: string }) {
	const styles = useStyles();
	const builder = useBuilder();
	const { loadAttributes, loadEntityMetadata, loadRelationships } = useLazyMetadata();

	// Tool-to-Tool (T2T) launch context: detected once, shared across the app.
	// When present, another tool launched us with a prefill payload.
	const launch = useLaunchContext();

	// Tools we can send the current query to, discovered via the "fetchxml" capability.
	const { tools: sendToTools } = useSendToTools();

	// After a connection change (#66): load the query the previous session had, once
	useEffect(() => {
		const fetchXml = carryOver.fetchXml;
		if (!fetchXml) return;
		carryOver.fetchXml = undefined;
		const result = builder.loadFetchXml(fetchXml);
		if (!result.success) console.error("Connection change: couldn't reload the query:", result.errors);
		// builder is intentionally omitted — runs once per session (mount)
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	// Lets AppShell read the current query when the connection changes (#66)
	const currentFetchQuery = builder.fetchQuery;
	useEffect(() => {
		carryOver.readFetchXml = () => (currentFetchQuery ? generateFetchXml(currentFetchQuery) : undefined);
	}, [carryOver, currentFetchQuery]);

	// Consume the inbound prefill exactly once (per tool instance, not per connection, #66). When a
	// FetchXML query is supplied, load it (the root entity is derived from the FetchXML). When only
	// an entity is supplied, start a fresh query rooted at that entity (metadata-driven).
	useEffect(() => {
		if (launch.loading || carryOver.prefillConsumed) return;
		carryOver.prefillConsumed = true;

		if (!launch.context) return;

		void (async () => {
			try {
				const action = await resolveLaunchPrefill(launch.context);
				if (!action) return;

				if (action.kind === "entity") {
					builder.setEntity(action.entityLogicalName);
					return;
				}

				const result = builder.loadFetchXml(action.fetchXml);
				if (!result.success) {
					console.error("T2T prefill: failed to parse incoming FetchXML:", result.errors);
					await window.toolboxAPI?.utils?.showNotification?.({
						title: "Could not load incoming query",
						body: "The tool that launched FetchXML Studio sent a query that couldn't be parsed.",
						type: "error",
					});
				}
			} catch (error) {
				console.error("T2T prefill: failed to load launch context:", error);
			}
		})();
		// builder is intentionally omitted — consumption is guarded to run once.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [launch.loading, launch.context]);

	// State for query execution
	const [queryResult, setQueryResult] = useState<QueryResult | null>(null);
	const [isExecuting, setIsExecuting] = useState(false);
	const [isLoadingMore, setIsLoadingMore] = useState(false);
	// Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>>
	const [attributeMetadata, setAttributeMetadata] = useState<
		Map<string, Map<string, AttributeMetadata>>
	>(new Map());
	// State for entity metadata (for layoutxml generation)
	const [loadedEntityMetadata, setLoadedEntityMetadata] = useState<EntityMetadata | null>(null);
	// State for lookup relationships (many-to-one) and one-to-many relationships
	const [lookupRelationships, setLookupRelationships] = useState<RelationshipMetadata[]>([]);
	const [oneToManyRelationships, setOneToManyRelationships] = useState<RelationshipMetadata[]>([]);
	const [isLoadingRelationships, setIsLoadingRelationships] = useState(false);

	// Paging for infinite scroll / Retrieve All (#58, #60). The ref is what the async page loops read
	// (no stale closures); the state mirrors it for rendering.
	const pagingRef = useRef<PagingState | null>(null);
	const [pagingState, setPagingStateValue] = useState<PagingState | null>(null);
	const setPaging = (next: PagingState | null) => {
		pagingRef.current = next;
		setPagingStateValue(next);
	};
	// What happened to paging that the user should know about: a failed page, Stop, simple paging
	const [pagingNotice, setPagingNotice] = useState<PagingNotice | null>(null);
	// Run tokens (#60): results from an older execution, or pages from a stopped loop, are dropped.
	// resultRun changes with every execution and when the results are cleared (entity change, view,
	// New); pageRun also changes on Stop, which keeps the loaded rows.
	const resultRunRef = useRef(0);
	const pageRunRef = useRef(0);
	// Rows loaded for the current results, for Retrieve All's limit (#183)
	const loadedRowsRef = useRef(0);
	// While a Retrieve All page waits to retry because Dataverse is throttling: the wait (#183)
	const [pagingWaitSeconds, setPagingWaitSeconds] = useState<number | undefined>(undefined);
	// Unmounting (e.g. the connection changed, #66) stops page loops still running for these results
	useEffect(
		() => () => {
			resultRunRef.current++;
			pageRunRef.current++;
		},
		[],
	);

	// Export to Excel state
	const [exportStatus, setExportStatus] = useState<{
		isExporting: boolean;
		error?: string;
		hasPrivilege: boolean;
		privilegeChecked: boolean;
	}>({
		isExporting: false,
		hasPrivilege: false,
		privilegeChecked: false,
	});

	// Execution error (e.g. Dataverse API error surfaced to user)
	const [executeError, setExecuteError] = useState<string | undefined>();

	// Quick find test value (#142): replaces {0} in the executed FetchXML only
	const [quickFindValue, setQuickFindValue] = useState("");
	const [quickFindAddPercent, setQuickFindAddPercent] = useState(true);

	// Entity node of the tree the current results were produced for (see the clearing effect below)
	const resultEntityIdRef = useRef<string | undefined>(undefined);

	// Record action privileges state
	const [recordActionPrivileges, setRecordActionPrivileges] = useState<{
		canDelete: boolean;
		canBulkDelete: boolean;
		canRunWorkflow: boolean;
		privilegesChecked: boolean;
	}>({
		canDelete: false,
		canBulkDelete: false,
		canRunWorkflow: false,
		privilegesChecked: false,
	});

	// Dialog states for record actions
	const [deleteDialogState, setDeleteDialogState] = useState<{
		open: boolean;
		recordIds: string[];
		recordName?: string;
		isBatchDelete?: boolean; // Using batch DELETE for 4-100 records
	}>({ open: false, recordIds: [] });

	// Progress tracking for batch delete
	const [deleteProgress, setDeleteProgress] = useState<BatchDeleteProgress | null>(null);

	const [bulkDeleteDialogState, setBulkDeleteDialogState] = useState<{
		open: boolean;
		recordIds: string[];
		isAllRecords?: boolean;
		totalViewRecords?: number;
	}>({ open: false, recordIds: [] });

	const [workflowDialogState, setWorkflowDialogState] = useState<{
		open: boolean;
		recordIds: string[];
		preSelectedWorkflow?: WorkflowInfo;
	}>({ open: false, recordIds: [] });

	// Activate / Deactivate confirmation (#143); also the Close menu's plain state changes (#149)
	const [setStateDialogState, setSetStateDialogState] = useState<{
		open: boolean;
		action: StateAction;
		recordIds: string[];
		closeAction?: CloseStateAction;
	}>({ open: false, action: "activate", recordIds: [] });

	// Why Activate / Deactivate are off for the root entity, per action (#143)
	const [stateActionReasons, setStateActionReasons] = useState<{ activate?: string; deactivate?: string }>({});

	// Close menu for close-message entities (#149): its items, and the close message dialog
	const [closeMenuItems, setCloseMenuItems] = useState<CloseMenuItem[] | undefined>();
	const [closeDialogState, setCloseDialogState] = useState<{
		open: boolean;
		action: CloseMessageAction | null;
		recordIds: string[];
	}>({ open: false, action: null, recordIds: [] });

	// Currently selected record IDs (from ResultsGrid)
	const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

	// Settings drawer state
	const [settingsDrawerOpen, setSettingsDrawerOpen] = useState(false);
	const [displaySettings, setDisplaySettings] = usePersistedSettings();
	// The pickers follow "Advanced Find only" (#57). A layout effect, so the scope is set before the
	// children's effects load metadata.
	useLayoutEffect(() => {
		metadataScope.setAdvancedFindOnly(displaySettings.advancedFindOnly);
	}, [displaySettings.advancedFindOnly]);
	const advancedFindOnly = useAdvancedFindOnly();

	// State for resizable split
	const [topHeight, setTopHeight] = useState(58); // Percentage of left pane height for tree/properties
	const [leftPaneWidth, setLeftPaneWidth] = useState(480); // Width of left pane in pixels
	const [isDraggingHorizontal, setIsDraggingHorizontal] = useState(false); // Tree/Properties resize
	const [isDraggingVertical, setIsDraggingVertical] = useState(false); // Left/Right pane resize

	// Handle horizontal resize (tree/properties split)
	useEffect(() => {
		if (!isDraggingHorizontal) return;

		const handleMouseMove = (e: MouseEvent) => {
			const leftPane = document.querySelector("[data-left-pane]") as HTMLElement;
			if (!leftPane) return;

			const rect = leftPane.getBoundingClientRect();
			const newTopHeight = ((e.clientY - rect.top) / rect.height) * 100;

			// Constrain between 30% and 80%
			if (newTopHeight >= 30 && newTopHeight <= 80) {
				setTopHeight(newTopHeight);
			}
		};

		const handleMouseUp = () => {
			setIsDraggingHorizontal(false);
		};

		document.addEventListener("mousemove", handleMouseMove);
		document.addEventListener("mouseup", handleMouseUp);

		return () => {
			document.removeEventListener("mousemove", handleMouseMove);
			document.removeEventListener("mouseup", handleMouseUp);
		};
	}, [isDraggingHorizontal]);

	// Handle vertical resize (left/right pane split)
	useEffect(() => {
		if (!isDraggingVertical) return;

		const handleMouseMove = (e: MouseEvent) => {
			const appRoot = document.querySelector("[data-app-root]") as HTMLElement;
			if (!appRoot) return;

			const rect = appRoot.getBoundingClientRect();
			const newLeftPaneWidth = e.clientX - rect.left;

			// Constrain between 300px and 800px
			if (newLeftPaneWidth >= 300 && newLeftPaneWidth <= 800) {
				setLeftPaneWidth(newLeftPaneWidth);
			}
		};

		const handleMouseUp = () => {
			setIsDraggingVertical(false);
		};

		document.addEventListener("mousemove", handleMouseMove);
		document.addEventListener("mouseup", handleMouseUp);

		return () => {
			document.removeEventListener("mousemove", handleMouseMove);
			document.removeEventListener("mouseup", handleMouseUp);
		};
	}, [isDraggingVertical]);

	// Check Dataverse API on mount. WhoAmI is memoized per connection, so this shares the startup
	// request with the privilege checks (#71); the details are debug output only.
	useEffect(() => {
		if (!isDataverseAvailable()) {
			console.error("❌ Dataverse API not available on window object");
			return;
		}
		whoAmI()
			.then((result) => {
				if (result) debugLog("app", "WhoAmI succeeded", { UserId: result.UserId });
				else debugLog("app", "WhoAmI returned null - API may not be fully initialized");
			})
			.catch((error) => {
				console.error("❌ WhoAmI Error:", error);
			});
	}, []);

	// Check export privilege on mount
	useEffect(() => {
		const checkExportPrivilege = async () => {
			try {
				const user = await whoAmI();
				if (user) {
					const hasPrivilege = await checkPrivilegeByName(user.UserId, "prvExportToExcel");
					setExportStatus((prev) => ({
						...prev,
						hasPrivilege,
						privilegeChecked: true,
					}));
					debugLog("app", `📊 Export to Excel privilege: ${hasPrivilege ? "GRANTED" : "DENIED"}`);
				} else {
					setExportStatus((prev) => ({
						...prev,
						hasPrivilege: false,
						privilegeChecked: true,
					}));
				}
			} catch (error) {
				console.error("Failed to check export privilege:", error);
				setExportStatus((prev) => ({
					...prev,
					hasPrivilege: false,
					privilegeChecked: true,
				}));
			}
		};

		if (isDataverseAvailable()) {
			checkExportPrivilege();
		}
	}, []);

	// Check record action privileges when entity changes
	const entityLogicalName = builder.fetchQuery?.entity.name;
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
		const toItems = (reasonFor: (action: CloseAction) => string | undefined): CloseMenuItem[] | undefined =>
			actions?.map((a) => ({ key: a.key, label: a.label, destructive: a.destructive, disabledReason: reasonFor(a) }));
		setCloseMenuItems(toItems(() => undefined));
		if (!entityLogicalName || !isDataverseAvailable() || (activateClose && deactivateClose && !actions)) {
			setStateActionReasons({ activate: activateClose, deactivate: deactivateClose });
			return;
		}
		let current = true;
		setStateActionReasons({ activate: activateClose, deactivate: deactivateClose });
		(async () => {
			const closeEntities = [...new Set((actions ?? []).flatMap((a) => (a.kind === "message" ? [a.closeEntity] : [])))];
			const [model, canWrite, canCreate] = await Promise.all([
				getEntityStateModel(entityLogicalName).catch((error: unknown) => {
					console.warn("Activate/Deactivate: couldn't read the state model:", error);
					return undefined;
				}),
				checkWritePrivilege(entityLogicalName),
				Promise.all(closeEntities.map(async (e) => [e, await checkEntityPrivilege(e, "Create")] as const)).then((pairs) => new Map(pairs)),
			]);
			if (!current) return;
			const writeReason =
				canWrite === false ? "You don't have the Write privilege on this entity, so you can't change its records' status." : undefined;
			setStateActionReasons({
				activate: activateClose ?? (model !== undefined ? stateActionUnavailableReason(model, "activate") : undefined) ?? writeReason,
				deactivate: deactivateClose ?? (model !== undefined ? stateActionUnavailableReason(model, "deactivate") : undefined) ?? writeReason,
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

	// Root entity + all link-entity names. The root always comes from the query's <entity>
	// element, never from a position in a sorted list (#54).
	const queryEntities = useMemo(() => getQueryEntities(builder.fetchQuery), [builder.fetchQuery]);
	const rootEntityName = queryEntities.rootEntityName;

	// Only trust root metadata that belongs to the current root entity. While a new root's
	// metadata is loading (or after a stale response), this is null rather than another table's.
	const entityMetadata =
		loadedEntityMetadata && loadedEntityMetadata.LogicalName === rootEntityName
			? loadedEntityMetadata
			: null;

	// Load attribute metadata for all entities in the FetchXML query
	// This includes the root entity and all link-entities
	useEffect(() => {
		const { rootEntityName, entityNames } = queryEntities;
		if (!rootEntityName) {
			setAttributeMetadata(new Map());
			setLoadedEntityMetadata(null);
			return;
		}

		// Ignore responses that arrive after the query's entities have changed.
		let cancelled = false;

		// Metadata of every attribute of the query's entities, for looking up the attributes in the
		// query (grid formatting, columns, export); not filtered by "Advanced Find only" (#57)
		Promise.all(
			entityNames.map(async (entityName) => {
				try {
					const attributes = await loadAllEntityAttributes(entityName);
					const attrMap = new Map<string, AttributeMetadata>();
					attributes.forEach((attr) => {
						attrMap.set(attr.LogicalName, attr);
					});
					return { entityName, attrMap };
				} catch (error) {
					console.error(`Failed to load attributes for ${entityName}:`, error);
					return { entityName, attrMap: new Map<string, AttributeMetadata>() };
				}
			}),
		).then((results) => {
			if (cancelled) return;
			const multiEntityMap = new Map<string, Map<string, AttributeMetadata>>();
			results.forEach(({ entityName, attrMap }) => {
				multiEntityMap.set(entityName, attrMap);
			});
			setAttributeMetadata(multiEntityMap);
		});

		// Load root entity metadata (ObjectTypeCode, PrimaryIdAttribute, EntitySetName, ...)
		loadEntityMetadata(rootEntityName)
			.then((entity) => {
				if (!cancelled) setLoadedEntityMetadata(entity);
			})
			.catch((error) => {
				console.error("Failed to load entity metadata:", error);
				if (!cancelled) setLoadedEntityMetadata(null);
			});

		// Load relationships (many-to-one and one-to-many) for the Add Columns panel, per the
		// "Advanced Find only" setting (#57)
		setIsLoadingRelationships(true);
		loadRelationships(rootEntityName)
			.then((relationships) => {
				if (cancelled) return;
				// Store lookup relationships (many-to-one)
				setLookupRelationships(relationships.manyToOne);
				// Store one-to-many relationships for 1-N column support
				setOneToManyRelationships(relationships.oneToMany);
			})
			.catch((error) => {
				console.error("Failed to load relationships:", error);
				if (cancelled) return;
				setLookupRelationships([]);
				setOneToManyRelationships([]);
			})
			.finally(() => {
				if (!cancelled) setIsLoadingRelationships(false);
			});

		return () => {
			cancelled = true;
		};
		// queryEntities.key changes whenever the root or the set of entities changes.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [queryEntities.key, loadEntityMetadata, loadRelationships, advancedFindOnly]);

	// Clear query results when the fetch query structure changes (entity change, view clear, etc.)
	// We use the entity node id as a proxy - it changes when entity is re-selected or view is cleared.
	// Results produced for the new tree are kept: an editor-mode Execute loads the editor's FetchXML
	// into the tree and runs it straight away, and a fast response can arrive first (#102).
	useEffect(() => {
		const entityId = builder.fetchQuery?.entity?.id;
		if (entityId !== undefined && entityId === resultEntityIdRef.current) return;
		// Pages still loading for the old results must not bring them back (#60)
		resultRunRef.current++;
		pageRunRef.current++;
		setQueryResult(null);
		setPaging(null);
		setPagingNotice(null);
		setIsLoadingMore(false);
		setPagingWaitSeconds(undefined);
		// A first page still loading for the old results is dropped, so it can't end the spinner
		setIsExecuting(false);
	}, [builder.fetchQuery?.entity?.id]);

	// Sync layout with FetchXML when needed (e.g., after adding/removing attributes)
	// Also sync when columnConfig is null but there are attributes (ensures columns button works)
	const { syncLayoutWithFetchXml } = builder; // stable: builder actions are created once (#65)
	useEffect(() => {
		const hasAttributes = (builder.fetchQuery?.entity?.attributes?.length ?? 0) > 0;
		const needsSync = builder.layoutNeedsSync || (hasAttributes && !builder.columnConfig);

		if (needsSync && builder.fetchQuery) {
			// Build attribute type map from metadata for better column widths
			// Use root entity's attribute metadata
			const attributeTypeMap = new Map<string, string>();
			const rootEntityName = builder.fetchQuery.entity?.name;
			if (attributeMetadata && rootEntityName) {
				const rootEntityAttrs = attributeMetadata.get(rootEntityName);
				rootEntityAttrs?.forEach((attr, name) => {
					attributeTypeMap.set(name, attr.AttributeType || "");
				});
			}
			syncLayoutWithFetchXml(attributeTypeMap, entityMetadata?.PrimaryIdAttribute);
		}
	}, [
		builder.layoutNeedsSync,
		builder.fetchQuery,
		builder.columnConfig,
		attributeMetadata,
		entityMetadata,
		syncLayoutWithFetchXml,
	]);

	// Generate FetchXML from builder state — exactly what the user built, nothing injected (#59).
	const fetchXml = builder.fetchQuery ? generateFetchXml(builder.fetchQuery) : "";

	// Quick find views keep {0} in their conditions (#142): a test value is needed to execute them
	const quickFindPlaceholderCount = useMemo(() => countQuickFindPlaceholders(fetchXml), [fetchXml]);
	const quickFindSearchValue = toQuickFindSearchValue(quickFindValue, quickFindAddPercent);

	// Check if current query is an aggregate query (disables delete/workflow buttons)
	const isAggregateQuery = builder.fetchQuery?.options?.aggregate === true;

	// Check if query has 1-N or N-N relationships that cause row duplication
	// When these relationships exist, record commands should be disabled
	const hasOneToManyRelationship = useMemo(() => {
		if (!builder.fetchQuery?.entity?.links) return false;

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

		return checkLinks(builder.fetchQuery.entity.links);
	}, [builder.fetchQuery]);

	// Record actions need the root table's primary key in the results. Dataverse returns it for
	// normal queries even when not selected; aggregate and distinct results don't include it (#59).
	// Decide from the rows actually returned rather than from the query shape.
	const primaryIdAttribute = entityMetadata?.PrimaryIdAttribute;
	const isDistinctQuery = builder.fetchQuery?.options?.distinct === true;
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
		() => deleteAllUnavailableReason(executeError ? null : queryResult?.query, builder.fetchQuery),
		[executeError, queryResult?.query, builder.fetchQuery],
	);

	// Generate LayoutXML from column config for saving
	const layoutXml = useMemo(() => {
		if (!builder.fetchQuery || !entityMetadata) {
			return "";
		}
		// object, row id and jump from the root Entity's metadata (#93)
		return buildViewLayoutXml(builder.columnConfig, builder.fetchQuery, {
			objectTypeCode: entityMetadata.ObjectTypeCode,
			primaryIdAttribute: entityMetadata.PrimaryIdAttribute,
			primaryNameAttribute: entityMetadata.PrimaryNameAttribute,
		});
	}, [builder.columnConfig, builder.fetchQuery, entityMetadata]);

	// Handle save view completion - update loaded view state
	const handleSaveViewComplete = (
		viewId: string,
		viewType: "system" | "personal",
		viewName: string,
		details: SavedViewDetails,
	) => {
		// Update the builder's loaded view state so subsequent saves overwrite the same view.
		if (entityMetadata && builder.fetchQuery) {
			// Overwriting keeps the view's query type; Save As creates a public view (#93). The
			// solution and owner stay; description and default are what was saved (#142).
			const overwritten = builder.loadedView?.id === viewId ? builder.loadedView : null;
			builder.setLoadedView({
				id: viewId,
				type: viewType,
				entitySetName: entityMetadata.EntitySetName,
				name: viewName,
				originalFetchXml: generateFetchXml(builder.fetchQuery),
				queryType: overwritten ? overwritten.queryType : 0,
				isCustomizable: overwritten?.isCustomizable,
				description: details.description,
				isDefault: details.isDefault,
				isManaged: overwritten?.isManaged,
				solutionId: overwritten?.solutionId,
				ownerId: overwritten?.ownerId,
				ownerName: overwritten?.ownerName,
			});
		}
		debugLog("app", `✅ View saved: ${viewName} (${viewType}) - ${viewId}`);
	};

	/**
	 * Execute the tree's query and handle Retrieve All if enabled.
	 * @param synced - in editor mode the editor's FetchXML is loaded into the tree first (#102);
	 *   the new tree and its XML are passed here because builder state updates on the next render.
	 */
	const handleExecute = async (synced?: { xml: string; fetchQuery: FetchNode }) => {
		const query = synced?.fetchQuery ?? builder.fetchQuery;
		const builtXml = synced?.xml ?? fetchXml;
		if (!builtXml) return;
		// Quick find views (#142): the test value replaces {0} in the FetchXML that runs (and its later
		// pages); the tree, the editor and saved views keep {0}. Without a value nothing runs: the
		// test value field asks for one.
		// https://learn.microsoft.com/power-apps/developer/data-platform/quick-find
		const hasQuickFindPlaceholders = countQuickFindPlaceholders(builtXml) > 0;
		if (hasQuickFindPlaceholders && !quickFindSearchValue) return;
		const xmlToExecute = hasQuickFindPlaceholders ? applyQuickFindValue(builtXml, quickFindSearchValue) : builtXml;

		// A new run: anything still loading for the previous one is dropped (#60)
		const run = ++resultRunRef.current;
		pageRunRef.current++;

		setIsExecuting(true);
		setIsLoadingMore(false);
		setPagingWaitSeconds(undefined);
		setQueryResult(null);
		setPaging(null);
		setPagingNotice(null);
		setExecuteError(undefined);

		resultEntityIdRef.current = query?.entity.id;

		const startTime = performance.now();
		const entityLogicalName = query?.entity.name;
		const retrieveAll = builder.retrieveAllRecords;

		// Paging follows the FetchXML that runs, not the tree (#60). Microsoft: don't page a query that
		// uses top; Dataverse applies the limit.
		const options = readFetchPagingOptions(xmlToExecute);
		const hasTopLimit = options.top !== undefined;

		try {
			// Determine execution method based on loaded view state. A synced editor query
			// replaced the tree (and cleared the loaded view), so it always runs as FetchXML.
			let result: FetchXmlResult | undefined;
			// Set when the first page needed a smaller count (#183); later pages keep it
			let firstPageShrunkTo: PagingState | undefined;
			const loadedView = synced ? null : builder.loadedView;
			let useViewExecution = false;

			// A quick find view runs as FetchXML with the test value: by ID it would run with {0} as
			// stored, which matches nothing (verified live, #142)
			if (loadedView && !hasQuickFindPlaceholders) {
				if (isUnmodifiedView(xmlToExecute, loadedView.originalFetchXml)) {
					useViewExecution = true;
					debugLog(
						"app",
						`📋 Executing ${loadedView.type} view "${loadedView.name}" via ${
							loadedView.type === "system" ? "savedQuery" : "userQuery"
						}=${loadedView.id}`,
					);

					// Retried after service protection errors like every page (#183, #96); a view's page
					// size is Dataverse's, so it can't be made smaller
					if (loadedView.type === "system") {
						result = await withThrottleRetry(() => executeSystemView(loadedView.entitySetName, loadedView.id));
					} else {
						result = await withThrottleRetry(() => executePersonalView(loadedView.entitySetName, loadedView.id));
					}
				}
			}

			// If not using view execution (either no view, view was modified, or editor override), use FetchXML
			if (!result) {
				debugLog(
					"app",
					synced
						? "✏️ Executing editor FetchXML (loaded into the tree)"
						: hasQuickFindPlaceholders
							? "🔎 Executing quick find FetchXML with the test value"
							: loadedView
								? `📝 View "${loadedView.name}" was modified - executing its FetchXML`
								: "📡 Executing FetchXML query",
				);
				// The first page gets the same recovery as later pages (#183): retried after service
				// protection errors (#96), and requested again with half the count when it's over the
				// PPTB host's 16 MB response limit (#68). That is the common case: a 5,000-row first page
				// of a wide Entity. Without a count, Dataverse's default page size (5,000) is the starting
				// point; queries with top or aggregate can't take a count, so they only get the retry.
				const canShrink = !hasTopLimit && query?.options?.aggregate !== true;
				const firstPageState: PagingState = {
					mode: "fetchxml",
					xml: xmlToExecute,
					pageSize: canShrink ? (options.count ?? DATAVERSE_DEFAULT_PAGE_SIZE) : undefined,
					currentPage: (options.page ?? 1) - 1,
					nextPage: options.page ?? 1,
					moreRecords: true,
					isRetrieveAllInProgress: false,
					retrieveAll: false,
				};
				const { result: firstPageResult, state: fetchedFirst } = await fetchPageWithRecovery(
					firstPageState,
					// Unchanged, the query runs as written; after a halving, with the smaller count
					(s) => (s.pageSize === firstPageState.pageSize ? executeFetchXml(xmlToExecute) : executeFetchXml(addPagingToFetchXml(xmlToExecute, s.nextPage, undefined, s.pageSize))),
					{ onShrink: (s) => debugLog("app", `First page too large: retrying with count ${s.pageSize}`) },
				);
				result = firstPageResult;
				if (fetchedFirst.pageSize !== firstPageState.pageSize) firstPageShrunkTo = fetchedFirst;
			}
			if (run !== resultRunRef.current) return; // superseded while the first page was loading

			const executionTimeMs = Math.round(performance.now() - startTime);
			const moreRecords = !hasTopLimit && hasNextPage(useViewExecution ? "view" : "fetchxml", result);
			const firstPage = options.page ?? 1;

			loadedRowsRef.current = result.records.length;
			setQueryResult({
				query,
				rows: result.records.map((record) => ({ ...record })),
				totalRecordCount: result.totalRecordCount,
				totalRecordCountLimitExceeded: result.totalRecordCountLimitExceeded,
				moreRecords,
				pagingCookie: result.pagingCookie,
				entityLogicalName,
				executionTimeMs,
			});
			setPaging({
				mode: useViewExecution ? "view" : "fetchxml",
				xml: xmlToExecute,
				pageSize: firstPageShrunkTo?.pageSize ?? options.count,
				// Dataverse's default page size, for a smaller page after the 16 MB limit (#68)
				fullPageSize: options.count === undefined && moreRecords ? result.records.length : undefined,
				// After a smaller first page (#183), paging continues from the page it was fetched as
				currentPage: firstPageShrunkTo?.nextPage ?? firstPage,
				nextPage: (firstPageShrunkTo?.nextPage ?? firstPage) + 1,
				pagingCookie: result.pagingCookie,
				nextLink: result.nextLink,
				moreRecords,
				isRetrieveAllInProgress: retrieveAll && moreRecords,
				retrieveAll,
			});
			setIsExecuting(false);

			if (retrieveAll && moreRecords) await loadAllPages(run);
		} catch (error) {
			if (run !== resultRunRef.current) return;
			console.error("Failed to execute FetchXML:", error);
			// Documented Dataverse errors get actionable text; Dataverse's own message follows (#96)
			setExecuteError(formatDataverseError(error, "An unexpected error occurred while executing the query."));
			setQueryResult({ query, rows: [] });
			setIsExecuting(false);
		}
	};

	// The latest handleExecute, for callbacks that re-run the query after an await (record deletes):
	// they call the current version instead of listing a function that changes every render (#65)
	const handleExecuteRef = useRef(handleExecute);
	useLayoutEffect(() => {
		handleExecuteRef.current = handleExecute;
	});

	// Column filters from the results grid (#23): they change the query, which then runs again.
	// Running waits for the change to commit, so it runs the new query.
	const columnFilters = useMemo(
		() => columnFilterConditions(builder.fetchQuery, builder.columnFilterId),
		[builder.fetchQuery, builder.columnFilterId],
	);
	const [columnFilterRunRequest, setColumnFilterRunRequest] = useState(0);
	useEffect(() => {
		if (columnFilterRunRequest > 0) void handleExecuteRef.current();
	}, [columnFilterRunRequest]);
	const { addColumnFilter, removeColumnFilter, clearColumnFilters } = builder; // stable (#65)
	const handleAddColumnFilter = useCallback(
		(condition: NewCondition) => {
			addColumnFilter(condition);
			setColumnFilterRunRequest((n) => n + 1);
		},
		[addColumnFilter],
	);
	const handleRemoveColumnFilters = useCallback(
		(conditionIds: string[]) => {
			if (conditionIds.length === 0) return;
			if (conditionIds.length === columnFilters.length) clearColumnFilters();
			else conditionIds.forEach((id) => removeColumnFilter(id));
			setColumnFilterRunRequest((n) => n + 1);
		},
		[columnFilters.length, clearColumnFilters, removeColumnFilter],
	);

	/**
	 * Fetch the page after `state` (#58, #60): saved views follow @odata.nextLink; FetchXML sets page,
	 * the decoded paging cookie (when Dataverse returned one) and the executed query's own count.
	 */
	const fetchNextPage = async (
		state: PagingState,
		retry?: ThrottleRetryOptions,
	): Promise<{ result: FetchXmlResult; next: PagingState }> => {
		// Retried after service protection errors (#96); a page over the host's 16 MB limit is
		// requested again with half the count, which later pages keep (#68)
		const { result, state: fetched } = await fetchPageWithRecovery(
			state,
			(s) =>
				s.mode === "view"
					? executeNextLink(s.nextLink ?? "")
					: executeFetchXml(addPagingToFetchXml(s.xml, s.nextPage, s.pagingCookie, s.pageSize)),
			{ retry, onShrink: (s) => debugLog("app", `Page too large: retrying page ${s.nextPage} with count ${s.pageSize}`) },
		);
		return { result, next: advancePaging(fetched, result) };
	};

	/** Add a fetched page to the results of run `run` (ignored if the results changed meanwhile). */
	const appendPage = (run: number, state: PagingState, result: FetchXmlResult, next: PagingState) => {
		if (run !== resultRunRef.current) return false;
		loadedRowsRef.current += result.records.length;
		setQueryResult((prev) =>
			prev
				? { ...prev, rows: prev.rows.concat(result.records), moreRecords: next.moreRecords, pagingCookie: next.pagingCookie }
				: prev,
		);
		// A page loaded, so an earlier failure or Stop notice no longer applies. A FetchXML page
		// requested without a cookie means Dataverse returned none for this query: say so once.
		const simple: PagingNotice | null =
			state.mode === "fetchxml" && !state.pagingCookie ? { kind: "simple", message: SIMPLE_PAGING_MESSAGE } : null;
		setPagingNotice((prev) => (prev && prev.kind !== "error" && prev.kind !== "stopped" && prev.kind !== "limit" ? prev : simple));
		return true;
	};

	/**
	 * Retrieve All: fetch the remaining pages one after another (#60). Stops when there are no more
	 * pages, when Stop is pressed (pageRun changes), when the results change (resultRun changes) or
	 * at RETRIEVE_ALL_ROW_LIMIT rows (#183). Throttled pages get the long back-off, with the wait in
	 * the status line (#183).
	 */
	const loadAllPages = async (run: number) => {
		const pageRun = ++pageRunRef.current;
		const isCurrent = () => run === resultRunRef.current && pageRun === pageRunRef.current;
		let state = pagingRef.current;
		if (!state) return;
		state = { ...state, isRetrieveAllInProgress: true, retrieveAll: true };
		setIsLoadingMore(true);
		setPaging(state);
		// The host warns before the tab closes while Retrieve All runs (#70)
		const releaseClose = holdClose();
		const retry: ThrottleRetryOptions = {
			...LONG_RUNNING_RETRY,
			// Stop or new results during a wait: the page isn't fetched again
			shouldContinue: isCurrent,
			onRetry: ({ delayMs }) => {
				if (isCurrent()) setPagingWaitSeconds(Math.ceil(delayMs / 1000));
			},
		};
		try {
			while (state?.moreRecords) {
				if (retrieveAllLimitReached(loadedRowsRef.current)) {
					// Further pages load one at a time from here, as after Stop
					state = { ...state, isRetrieveAllInProgress: false, retrieveAll: false };
					setPaging(state);
					setPagingNotice({ kind: "limit", message: RETRIEVE_ALL_LIMIT_MESSAGE });
					debugLog("app", `Retrieve All stopped at ${loadedRowsRef.current} rows (limit)`);
					return;
				}
				const { result, next } = await fetchNextPage(state, retry);
				if (!isCurrent()) return;
				setPagingWaitSeconds(undefined);
				appendPage(run, state, result, next);
				state = { ...next, isRetrieveAllInProgress: next.moreRecords };
				setPaging(state);
			}
			debugLog("app", "✅ Retrieve All complete");
		} catch (error) {
			if (!isCurrent()) return;
			console.error("Failed during Retrieve All:", error);
			setPagingNotice({
				kind: "error",
				page: state?.nextPage,
				message: formatDataverseError(error), // (#96)
			});
		} finally {
			releaseClose();
			if (isCurrent()) {
				setPagingWaitSeconds(undefined);
				setIsLoadingMore(false);
				if (pagingRef.current) setPaging({ ...pagingRef.current, isRetrieveAllInProgress: false });
			}
		}
	};

	/** Stop Retrieve All (#60): the page in flight is discarded; the loaded rows stay. */
	const handleStopRetrieveAll = () => {
		const state = pagingRef.current;
		if (!state?.isRetrieveAllInProgress) return;
		pageRunRef.current++;
		setIsLoadingMore(false);
		setPagingWaitSeconds(undefined);
		// Further pages load one at a time from here (scrolling, or Retry after a failed page)
		setPaging({ ...state, isRetrieveAllInProgress: false, retrieveAll: false });
		setPagingNotice({ kind: "stopped", message: "" });
	};

	/**
	 * Load more records (for infinite scroll when Retrieve All is OFF)
	 */
	const handleLoadMore = async () => {
		const state = pagingRef.current;
		if (!state || !state.moreRecords || state.isRetrieveAllInProgress || isLoadingMore) return;

		const run = resultRunRef.current;
		const pageRun = ++pageRunRef.current;
		const isCurrent = () => run === resultRunRef.current && pageRun === pageRunRef.current;
		setIsLoadingMore(true);

		try {
			const { result, next } = await fetchNextPage(state);
			if (!isCurrent()) return;
			appendPage(run, state, result, next);
			setPaging(next);
			debugLog("app", `✅ Page ${next.currentPage} loaded: ${result.records.length} records`);
		} catch (error) {
			if (!isCurrent()) return;
			console.error("Failed to load more records:", error);
			setPagingNotice({
				kind: "error",
				page: state.nextPage,
				message: formatDataverseError(error), // (#96)
			});
		} finally {
			if (isCurrent()) setIsLoadingMore(false);
		}
	};

	/** Retry the page that failed: continue Retrieve All, or load the one page (#60). */
	const handleRetryPaging = () => {
		setPagingNotice(null);
		const state = pagingRef.current;
		if (!state) return;
		if (state.retrieveAll) void loadAllPages(resultRunRef.current);
		else void handleLoadMore();
	};

	/**
	 * Export to Excel using Dataverse ExportToExcel API
	 * Requires a saved view (system or personal) and prvExportToExcel privilege
	 */
	const handleExport = async () => {
		// Check privilege first
		if (!exportStatus.hasPrivilege) {
			setExportStatus((prev) => ({
				...prev,
				error:
					"You don't have permission to export to Excel. Contact your administrator to request the 'Export to Excel' privilege.",
			}));
			return;
		}

		// Check if we have a saved view
		if (!builder.loadedView) {
			setExportStatus((prev) => ({
				...prev,
				error: "Export to Excel requires a saved view. Please save your query as a view first.",
			}));
			return;
		}

		if (!fetchXml || !layoutXml) {
			setExportStatus((prev) => ({
				...prev,
				error: "Export to Excel requires FetchXML and LayoutXML",
			}));
			return;
		}

		// Clear any previous error and set exporting state
		setExportStatus((prev) => ({
			...prev,
			isExporting: true,
			error: undefined,
		}));

		try {
			debugLog(
				"app",
				`📤 Exporting to Excel via ${builder.loadedView.type} view "${builder.loadedView.name}"...`,
			);

			// Use view name for filename
			const viewName = builder.loadedView.name;

			const result = await exportToExcel(
				builder.loadedView.id,
				builder.loadedView.type,
				fetchXml,
				layoutXml,
				viewName,
			);

			// Trigger download
			await downloadBase64File(result.excelFile, result.filename);

			debugLog("app", `✅ Export complete: ${result.filename}`);

			// Clear exporting state on success
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
			}));
		} catch (error) {
			console.error("Export to Excel failed:", error);
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
				error: formatDataverseError(error, "Export to Excel failed. Please try again."), // (#96)
			}));
		}
	};

	/**
	 * Export to Excel locally using exceljs
	 * Exports the current result set with native Excel data types
	 * No view required - works directly with query results
	 */
	const handleExportLocal = async () => {
		if (!queryResult || !queryResult.rows.length) {
			setExportStatus((prev) => ({
				...prev,
				error: "No data to export. Execute a query first.",
			}));
			return;
		}

		setExportStatus((prev) => ({
			...prev,
			isExporting: true,
			error: undefined,
		}));

		try {
			debugLog("app", `📤 Exporting ${queryResult.rows.length} records locally...`);

			// Build column display names map
			const columnDisplayNames = new Map<string, string>();

			// The grid's columns, in the grid's order (resultColumns.ts, #91)
			const getAttribute = (entityName: string, attributeName: string) =>
				attributeMetadata?.get(entityName)?.get(attributeName);
			const columns = orderByLayout(
				resultColumns(builder.fetchQuery, queryResult.rows, getAttribute, queryResult.query),
				builder.columnConfig,
			).map((c) => c.key);

			for (const col of columns) {
				// Try to get display name from attribute metadata
				let displayName = col;

				if (attributeMetadata && entityLogicalName) {
					const entityAttrs = attributeMetadata.get(entityLogicalName);
					if (entityAttrs) {
						// Direct attribute lookup
						const attr = entityAttrs.get(col);
						if (attr?.DisplayName?.UserLocalizedLabel?.Label) {
							displayName = attr.DisplayName.UserLocalizedLabel.Label;
						}
						// Handle lookup columns (_primarycontactid_value -> primarycontactid)
						if (col.startsWith("_") && col.endsWith("_value")) {
							const baseAttr = col.slice(1, -6);
							const lookupAttr = entityAttrs.get(baseAttr);
							if (lookupAttr?.DisplayName?.UserLocalizedLabel?.Label) {
								displayName = lookupAttr.DisplayName.UserLocalizedLabel.Label;
							}
						}
					}
				}

				columnDisplayNames.set(col, displayName);
			}

			// Generate filename from entity or view name
			const fileName =
				builder.loadedView?.name ||
				entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
				entityLogicalName ||
				"export";

			// Date behaviours and formats, money precision and the user's time zone (#184); the export
			// still works without them (UTC, labelled; currency and decimals from formatted values)
			const entityNames = [...(attributeMetadata?.keys() ?? [])];
			// exceljs (~1 MB) loads on the first local export, not at startup (#187)
			const [formattingEntries, timeZone, { exportToExcelLocal, downloadExcelFile }] = await Promise.all([
				Promise.all(
					entityNames.map((name) =>
						loadFormattingMetadata(name)
							.then((formatting) => [name, formatting] as const)
							.catch((error: unknown) => {
								console.error(`Failed to load formatting metadata for ${name}:`, error);
								return [name, new Map()] as const;
							}),
					),
				),
				loadUserTimeZone(),
				import("../features/fetchxml/api/excelExport"),
			]);

			const { buffer, fileName: finalFileName } = await exportToExcelLocal({
				records: queryResult.rows,
				columns,
				columnDisplayNames,
				attributeMetadata,
				entityName: entityLogicalName,
				fileName,
				displaySettings,
				formattingMetadata: new Map(formattingEntries),
				timeZone,
			});

			// Trigger download
			await downloadExcelFile(buffer, finalFileName);

			debugLog("app", `✅ Local export complete: ${finalFileName}`);

			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
			}));
		} catch (error) {
			console.error("Local export failed:", error);
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
				error: error instanceof Error ? error.message : "Local export failed. Please try again.",
			}));
		}
	};

	// Clear export error after 10 seconds
	useEffect(() => {
		if (exportStatus.error) {
			const timer = setTimeout(() => {
				setExportStatus((prev) => ({ ...prev, error: undefined }));
			}, 10000);
			return () => clearTimeout(timer);
		}
	}, [exportStatus.error]);

	// ============ RECORD ACTION HANDLERS ============

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
					debugLog("app", "Bulk delete privilege required to delete all records from view");
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
				if (deleteSelectionUnavailableReason(recordIds.length, recordActionPrivileges.canBulkDelete)) return;
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
		[entityLogicalName, recordActionPrivileges.canBulkDelete, queryResult, entityMetadata, deleteAllDisabledReason],
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
		[entityLogicalName, recordActionPrivileges.canBulkDelete, queryResult, deleteAllDisabledReason],
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
				const result = await deleteRecordsBatch(
					entityLogicalName,
					recordIds,
					(progress) => setDeleteProgress(progress),
				);

				if (result.succeeded > 0) {
					debugLog("app", `✅ Batch deleted ${result.succeeded} of ${recordIds.length} records`);
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
					throw new Error(deleteAllDisabledReason ?? "Run the query first: Delete all deletes the records the query on screen matches.");
				}
				// The FetchXML that ran: a quick find view's test value replaced {0} only there (#142), so
				// the tree's XML would still say {0}
				const currentFetchXml = pagingRef.current?.xml ?? generateFetchXml(executed);
				const result = await submitBulkDeleteFromFetchXml(currentFetchXml, jobName);
				debugLog("app", `📤 Bulk delete job (all records) submitted: ${result.asyncOperationId}`);
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
	 * Fetch available workflows for the entity
	 */
	const handleFetchWorkflows = useCallback(async (): Promise<WorkflowInfo[]> => {
		if (!entityLogicalName) return [];
		return getOnDemandWorkflows(entityLogicalName);
	}, [entityLogicalName]);

	/**
	 * Execute workflow on selected records
	 */
	const handleExecuteWorkflow = useCallback(
		async (
			workflowId: string,
			recordIds: string[],
			onProgress: (progress: WorkflowBatchProgress) => void,
		): Promise<{ succeeded: number; failed: number; errors: string[] }> => {
			if (!entityLogicalName) {
				return { succeeded: 0, failed: recordIds.length, errors: ["Entity not selected"] };
			}
			// The host warns before the tab closes while workflows are being started (#70)
			return whileHoldingClose(() => executeWorkflowBatch(workflowId, recordIds, entityLogicalName, onProgress));
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
	const handleActivateRecords = useCallback((ids: string[]) => openSetStateDialog("activate", ids), [openSetStateDialog]);
	const handleDeactivateRecords = useCallback((ids: string[]) => openSetStateDialog("deactivate", ids), [openSetStateDialog]);

	/** Runs the state change (#143): PATCH statecode / statuscode, holding the host's close warning */
	const handleExecuteStateChange = useCallback(
		async (target: StateTarget, onProgress: (progress: BatchDeleteProgress) => void): Promise<BatchDeleteResult> => {
			if (!entityLogicalName) throw new Error("Entity not selected");
			const result = await whileHoldingClose(() =>
				setRecordsState(entityLogicalName, setStateDialogState.recordIds, target, onProgress),
			);
			if (result.succeeded > 0) {
				debugLog("app", `✅ State changed on ${result.succeeded} of ${setStateDialogState.recordIds.length} records`);
				void handleExecuteRef.current();
			}
			return result;
		},
		[entityLogicalName, setStateDialogState.recordIds],
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
				setSetStateDialogState({ open: true, action: "deactivate", recordIds, closeAction: action });
			} else {
				setCloseDialogState({ open: true, action, recordIds });
			}
		},
		[entityLogicalName],
	);

	/** Runs the close message on each record (#149), holding the host's close warning */
	const handleExecuteClose = useCallback(
		async (values: CloseValues, onProgress: (progress: BatchDeleteProgress) => void): Promise<BatchDeleteResult> => {
			const action = closeDialogState.action;
			if (!entityLogicalName || !action) throw new Error("Entity not selected");
			const result = await whileHoldingClose(() =>
				closeRecords(entityLogicalName, entityMetadata?.PrimaryIdAttribute, closeDialogState.recordIds, action, values, onProgress),
			);
			if (result.succeeded > 0) {
				debugLog("app", `✅ ${action.message} on ${result.succeeded} of ${closeDialogState.recordIds.length} records`);
				void handleExecuteRef.current();
			}
			return result;
		},
		[entityLogicalName, entityMetadata?.PrimaryIdAttribute, closeDialogState.action, closeDialogState.recordIds],
	);

	// Rows repeat for 1:N / N:N link-entities; Activate / Deactivate say so like their other reasons (#143)
	const oneToManyStateReason = hasOneToManyRelationship
		? "The query has a one-to-many or many-to-many link-entity, so rows repeat. Remove it to activate or deactivate records."
		: undefined;
	const oneToManyCloseReason = hasOneToManyRelationship
		? "The query has a one-to-many or many-to-many link-entity, so rows repeat. Remove it to close records."
		: undefined;

	// ============ END RECORD ACTION HANDLERS ============

	return (
		<div className={styles.root} data-app-root>
			{/* Left Pane */}
			<div className={styles.leftPane} data-left-pane style={{ width: `${leftPaneWidth}px` }}>
				{/* Top: Tree View */}
				<div className={styles.leftPaneTop} style={{ height: `${topHeight}%` }}>
					<PaneErrorBoundary where="Query scope" title="The query scope couldn't be shown" resetKey={connectionKey}>
					<EntitySelector
						selectedEntity={builder.fetchQuery?.entity.name || null}
						onEntityChange={builder.setEntity}
						onNewQuery={builder.newQuery}
						entityScopeMode={displaySettings.entityScopeMode}
						onEntityScopeModeChange={(entityScopeMode) => setDisplaySettings({ ...displaySettings, entityScopeMode })}
						advancedFindOnly={displaySettings.advancedFindOnly}
						onAdvancedFindOnlyChange={(advancedFindOnly) => setDisplaySettings({ ...displaySettings, advancedFindOnly })}
						onViewLoad={(viewInfo: LoadedViewInfo) => {
							// Load the view's FetchXML into the tree while preserving view info
							// Pass layoutxml for column configuration if available
							const result = builder.loadView(
								viewInfo.originalFetchXml,
								{
									id: viewInfo.id,
									type: viewInfo.type,
									entitySetName: viewInfo.entitySetName,
									name: viewInfo.name,
									queryType: viewInfo.queryType,
									isCustomizable: viewInfo.isCustomizable,
									description: viewInfo.description,
									isDefault: viewInfo.isDefault,
									isManaged: viewInfo.isManaged,
									solutionId: viewInfo.solutionId,
									ownerId: viewInfo.ownerId,
									ownerName: viewInfo.ownerName,
								},
								viewInfo.layoutxml,
							);
							// A test value belongs to the quick find view it was entered for (#142)
							setQuickFindValue("");
							// Say when the view couldn't be loaded, or the tree changed its query (#93)
							const notice = getViewLoadNotice(viewInfo.name, result);
							if (notice) {
								if (!result.success) console.error("Failed to parse view FetchXML:", result.errors);
								void window.toolboxAPI?.utils?.showNotification?.(notice);
							}
						}}
					/>
					</PaneErrorBoundary>
					{builder.fetchQuery ? (
						<PaneErrorBoundary where="Query tree" title="The query tree couldn't be shown" resetKey={builder.fetchQuery}>
						<TreeView
							fetchQuery={builder.fetchQuery}
							size={treeSizeFor(displaySettings.gridDensity)}
							iconStyle={displaySettings.treeIconStyle}
							attributeMetadata={attributeMetadata}
							selectedNodeId={builder.selectedNodeId}
							onNodeSelect={builder.selectNode}
							onAddAttribute={builder.addAttribute}
							onAddAllAttributes={builder.addAllAttributes}
							onAddOrder={builder.addOrder}
							onAddFilter={builder.addFilter}
							onAddSubfilter={builder.addSubfilter}
							onAddCondition={builder.addCondition}
							onAddLinkEntity={builder.addLinkEntity}
							onRemoveNode={builder.removeNode}
							onUpdateAttributes={builder.updateAttributes}
						/>
						</PaneErrorBoundary>
					) : (
						<div className={styles.placeholder}>Select an entity to begin</div>
					)}
				</div>
				{/* Resize Handle */}
				<div
					className={styles.horizontalResizeHandle}
					onMouseDown={() => setIsDraggingHorizontal(true)}
					title="Drag to resize"
				>
					<span className={styles.horizontalGrip}>
						<span>•</span>
						<span>•</span>
						<span>•</span>
					</span>
				</div>{" "}
				{/* Bottom: Properties Panel */}
				<div className={styles.leftPaneBottom} style={{ height: `${100 - topHeight}%` }}>
					<div className={styles.toolbar}>
						<span style={{ fontWeight: 600 }}>Properties</span>
					</div>
					<div className={styles.propertiesContent}>
						<PaneErrorBoundary where="Properties panel" title="The properties couldn't be shown" resetKey={builder.selectedNodeId}>
						<PropertiesPanel
							selectedNode={builder.selectedNode}
							fetchQuery={builder.fetchQuery}
							onNodeUpdate={builder.updateNode}
						/>
						</PaneErrorBoundary>
					</div>
				</div>
			</div>
			{/* Vertical Resize Handle */}
			<div
				className={styles.verticalResizeHandle}
				onMouseDown={() => setIsDraggingVertical(true)}
				title="Drag to resize"
			>
				<span className={styles.verticalGrip}>
					<span>•</span>
					<span>•</span>
					<span>•</span>
				</span>
			</div>{" "}
			{/* Right Pane */}
			<div className={styles.rightPane}>
				<PreviewTabs
					xml={fetchXml}
					layoutXml={layoutXml}
					result={queryResult}
					isExecuting={isExecuting}
					isLoadingMore={isLoadingMore}
					onExecute={handleExecute}
					onExport={handleExport}
					onExportLocal={handleExportLocal}
					canExport={!!builder.loadedView && exportStatus.hasPrivilege}
					isExporting={exportStatus.isExporting}
					exportError={exportStatus.error}
					onDismissExportError={() => setExportStatus((prev) => ({ ...prev, error: undefined }))}
					executeError={executeError}
					onDismissExecuteError={() => setExecuteError(undefined)}
					exportDisabledReason={
						!builder.loadedView
							? "Save as a view first to enable export"
							: !exportStatus.hasPrivilege
								? "You don't have the prvExportToExcel privilege"
								: undefined
					}
					onParseToTree={builder.loadFetchXml}
					isCallee={launch.isCallee}
					attributeMetadata={attributeMetadata}
					fetchQuery={builder.fetchQuery}
					primaryIdAttribute={entityMetadata?.PrimaryIdAttribute}
					connectionKey={connectionKey}
					selectedNodeId={builder.selectedNodeId}
					onSelectNode={builder.selectNode}
					onAddOrder={builder.addOrder}
					columnConfig={builder.columnConfig}
					onColumnResize={builder.updateColumnWidth}
					onLoadMore={handleLoadMore}
					pagingNotice={pagingNotice}
					pagingWaitSeconds={pagingWaitSeconds}
					isRetrieveAllInProgress={pagingState?.isRetrieveAllInProgress ?? false}
					onStopRetrieveAll={handleStopRetrieveAll}
					onRetryPaging={handleRetryPaging}
					onDismissPagingNotice={() => setPagingNotice(null)}
					entityDisplayName={
						entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityMetadata?.LogicalName
					}
					lookupRelationships={lookupRelationships}
					oneToManyRelationships={oneToManyRelationships}
					isLoadingRelationships={isLoadingRelationships}
					onLoadRelatedAttributes={loadAttributes}
					quickFindTestValue={
						quickFindPlaceholderCount > 0 ? (
							<QuickFindTestValue
								value={quickFindValue}
								onValueChange={setQuickFindValue}
								addPercent={quickFindAddPercent}
								onAddPercentChange={setQuickFindAddPercent}
								searchValue={quickFindSearchValue}
								disabled={isExecuting}
							/>
						) : undefined
					}
					executeDisabledReason={
						quickFindPlaceholderCount > 0 && !quickFindSearchValue
							? "Enter a quick find value to execute this query."
							: undefined
					}
					saveViewButton={
						<SaveViewButton
							fetchXml={fetchXml}
							layoutXml={layoutXml}
							entityLogicalName={builder.fetchQuery?.entity?.name || ""}
							entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
							objectTypeCode={entityMetadata?.ObjectTypeCode || 0}
							primaryIdAttribute={entityMetadata?.PrimaryIdAttribute || ""}
							loadedView={builder.loadedView}
							onSaveComplete={handleSaveViewComplete}
							disabled={!fetchXml || !entityMetadata}
						/>
					}
					sendToToolButton={
						isT2TSupported() && sendToTools.length > 0 ? (
							<SendToToolButton
								fetchXml={fetchXml}
								entityLogicalName={builder.fetchQuery?.entity?.name || ""}
								tools={sendToTools}
								disabled={!fetchXml || !builder.fetchQuery?.entity?.name}
							/>
						) : undefined
					}
					onReorderColumns={(columns) => {
						// Set the column config with new order
						if (builder.columnConfig) {
							builder.setColumnConfig({
								...builder.columnConfig,
								columns,
							});
						}
					}}
					onRemoveColumns={(columnKeys) => {
						// Each column maps to its attribute (root or link-entity, aliased or not); sorts on a
						// removed alias and outer link-entities left empty go too (queryColumns.ts, #114)
						if (!builder.fetchQuery) return;
						const plan = planRemoveColumns(builder.fetchQuery, columnKeys);
						plan.removeNodeIds.forEach((id) => builder.removeNode(id));
					}}
					onAddColumn={(attributeName) => {
						// Add the attribute to the root entity
						if (builder.fetchQuery?.entity.id) {
							builder.addAttributeByName(builder.fetchQuery.entity.id, attributeName);
						}
					}}
					onAddRelatedColumns={(relatedColumns: RelatedEntityColumn[]) => {
						// Add related entity columns - this requires creating link-entities
						if (!builder.fetchQuery?.entity.id) return;

						// Group by relationship to create link-entities
						const byRelationship = new Map<string, RelatedEntityColumn[]>();
						for (const col of relatedColumns) {
							// An intersect entity's two joins share a schema name (#57)
							const key = relationshipKey(col.relationship);
							if (!byRelationship.has(key)) {
								byRelationship.set(key, []);
							}
							byRelationship.get(key)!.push(col);
						}

						// For each relationship, find or create link-entity and add attributes
						for (const [, cols] of byRelationship) {
							const rel = cols[0].relationship;
							const relType = cols[0].relationshipType;

							let fromAttr: string;
							let toAttr: string;
							let relatedEntity: string;
							let linkType: "inner" | "outer";
							let relTypeForBuilder: "N1" | "1N";

							if (relType === "N1") {
								// For N:1 (Many-to-One) lookup relationships:
								// - ReferencingAttribute = FK on current entity (e.g., primarycontactid on account)
								// - ReferencedAttribute = PK on related entity (e.g., contactid on contact)
								// In FetchXML link-entity:
								// - from = attribute on the LINKED entity (the PK)
								// - to = attribute on the PARENT entity (the FK)
								fromAttr = rel.ReferencedAttribute; // PK on linked (contact.contactid)
								toAttr = rel.ReferencingAttribute; // FK on parent (account.primarycontactid)
								relatedEntity = rel.ReferencedEntity; // contact
								linkType = "outer"; // N:1 typically use outer join to not filter parent records
								relTypeForBuilder = "N1";
							} else {
								// For 1:N (One-to-Many) relationships:
								// - ReferencingEntity = the related entity (the "many" side that has the FK)
								// - ReferencedEntity = the current entity (the "one" side)
								// - ReferencingAttribute = FK on the related entity
								// - ReferencedAttribute = PK on the current entity
								// In FetchXML link-entity:
								// - from = FK on the LINKED entity (the FK to parent)
								// - to = PK on the PARENT entity
								fromAttr = rel.ReferencingAttribute; // FK on linked entity
								toAttr = rel.ReferencedAttribute; // PK on parent (e.g., accountid)
								relatedEntity = rel.ReferencingEntity; // The related entity (e.g., contact)
								linkType = "outer"; // 1:N use outer join to include parent records without children
								relTypeForBuilder = "1N";
							}

							// Reuse a link-entity for this relationship, unless its type returns no attributes
							// (exists / in / any / not any / all / not all reject them: 0x80041121)
							const existingLinkEntity = builder.fetchQuery.entity.links.find(
								(le) =>
									le.from === fromAttr &&
									le.to === toAttr &&
									le.name === relatedEntity &&
									!NON_RETURNING_LINK_TYPES.has(le.linkType),
							);

							let linkEntityId: string;

							if (existingLinkEntity) {
								linkEntityId = existingLinkEntity.id;
							} else {
								// Create the link-entity with config - returns the ID immediately
								linkEntityId = builder.addLinkEntityWithConfig(
									builder.fetchQuery.entity.id,
									relatedEntity,
									fromAttr,
									toAttr,
									linkType,
									relTypeForBuilder,
								);
							}

							// Add attributes to the link-entity using the ID
							for (const col of cols) {
								builder.addAttributeByName(linkEntityId, col.attribute.LogicalName);
							}
						}
					}}
					columnFilters={columnFilters}
					onAddColumnFilter={handleAddColumnFilter}
					onRemoveColumnFilters={handleRemoveColumnFilters}
					onSortChange={(data) => {
						// The column maps to the right <order>: alias in aggregate queries, entityname for
						// link-entity columns, nothing for types Dataverse can't sort (queryColumns.ts, #114)
						if (!builder.fetchQuery) return;
						const column = resolveColumn(builder.fetchQuery, data.columnId);
						const metadata = column
							? attributeMetadata?.get(column.container.name)?.get(column.attribute.name)
							: undefined;
						const plan = planSortByColumn(
							builder.fetchQuery,
							data.columnId,
							metadata ? toAttributeTypeInfo(metadata) : undefined,
						);
						if ("reason" in plan) {
							void window.toolboxAPI?.utils?.showNotification?.({
								title: "Can't sort by this column",
								body: plan.reason,
								type: "warning",
							});
							return;
						}
						builder.setSort(plan, data.direction === "descending", data.isMultiSort);
					}}
					// Record action handlers
					onOpenRecord={handleOpenRecord}
					onCopyRecordUrl={handleCopyRecordUrl}
					onDeleteRecords={handleDeleteRecords}
					onBulkDeleteRecords={handleBulkDeleteRecords}
					onActivateRecords={handleActivateRecords}
					onDeactivateRecords={handleDeactivateRecords}
					activateDisabledReason={stateActionReasons.activate ?? oneToManyStateReason}
					deactivateDisabledReason={stateActionReasons.deactivate ?? oneToManyStateReason}
					closeActions={closeMenuItems}
					closeDisabledReason={oneToManyCloseReason}
					onCloseRecords={handleCloseRecords}
					onRunSpecificWorkflow={handleRunSpecificWorkflow}
					canDelete={recordActionPrivileges.canDelete}
					canBulkDelete={recordActionPrivileges.canBulkDelete}
					canRunWorkflow={recordActionPrivileges.canRunWorkflow}
					onFetchWorkflows={handleFetchWorkflows}
					isAggregateQuery={isAggregateQuery}
					hasOneToManyRelationship={hasOneToManyRelationship}
					recordActionsDisabledReason={recordActionsDisabledReason}
					deleteAllDisabledReason={deleteAllDisabledReason}
					onSelectionChange={handleSelectionChange}
					getSelectedRecordIds={getSelectedRecordIds}
					onOpenSettings={() => setSettingsDrawerOpen(true)}
					displaySettings={displaySettings}
					onGridDensityChange={(gridDensity) => setDisplaySettings({ ...displaySettings, gridDensity })}
				/>
			</div>
			{/* Delete Confirmation Dialog */}
			<DeleteConfirmDialog
				open={deleteDialogState.open}
				recordName={deleteDialogState.recordName}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityLogicalName || "record"
				}
				entityLogicalName={entityLogicalName}
				recordCount={deleteDialogState.recordIds.length}
				onClose={() => {
					setDeleteDialogState({ open: false, recordIds: [] });
					setDeleteProgress(null);
				}}
				onConfirm={handleConfirmDelete}
				onProgress={deleteProgress ?? undefined}
			/>
			{/* Bulk Delete Dialog */}
			<BulkDeleteDialog
				open={bulkDeleteDialogState.open}
				recordCount={bulkDeleteDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityLogicalName || "records"
				}
				entityLogicalName={entityLogicalName}
				isAllRecords={bulkDeleteDialogState.isAllRecords}
				totalViewRecords={bulkDeleteDialogState.totalViewRecords}
				moreRecords={queryResult?.moreRecords}
				onClose={() => setBulkDeleteDialogState({ open: false, recordIds: [] })}
				onConfirm={handleConfirmBulkDelete}
			/>
			{/* Workflow Picker Dialog */}
			<WorkflowPickerDialog
				open={workflowDialogState.open}
				recordCount={workflowDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityLogicalName || "records"
				}
				entityLogicalName={entityLogicalName}
				preSelectedWorkflow={workflowDialogState.preSelectedWorkflow}
				onClose={() => setWorkflowDialogState({ open: false, recordIds: [] })}
				onFetchWorkflows={handleFetchWorkflows}
				onExecute={(workflow, onProgress) =>
					handleExecuteWorkflow(workflow.workflowid, workflowDialogState.recordIds, onProgress)
				}
			/>
			{/* Activate / Deactivate Dialog (#143) */}
			<SetStateDialog
				open={setStateDialogState.open}
				action={setStateDialogState.action}
				recordIds={setStateDialogState.recordIds}
				entityLogicalName={entityLogicalName}
				entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
				primaryIdAttribute={entityMetadata?.PrimaryIdAttribute}
				closeAction={setStateDialogState.closeAction}
				onClose={() => setSetStateDialogState((prev) => ({ ...prev, open: false }))}
				onExecute={handleExecuteStateChange}
			/>
			{/* Close message Dialog (#149) */}
			<CloseRecordsDialog
				open={closeDialogState.open}
				action={closeDialogState.action}
				recordIds={closeDialogState.recordIds}
				entityLogicalName={entityLogicalName}
				entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
				onClose={() => setCloseDialogState((prev) => ({ ...prev, open: false }))}
				onExecute={handleExecuteClose}
			/>
			{/* Settings Drawer */}
			<SettingsDrawer
				open={settingsDrawerOpen}
				settings={displaySettings}
				onClose={() => setSettingsDrawerOpen(false)}
				onSettingsChange={setDisplaySettings}
			/>
		</div>
	);
}
