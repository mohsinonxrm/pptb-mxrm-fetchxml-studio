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
import { useSplitters } from "./hooks/useSplitters";
import { useResultsExport } from "./hooks/useResultsExport";
import { useRootEntityMetadata } from "./hooks/useRootEntityMetadata";
import { useRecordActionAvailability, useRecordActions } from "./hooks/useRecordActions";
import { useQueryExecution } from "./hooks/useQueryExecution";
import { metadataScope, useAdvancedFindOnly } from "../features/fetchxml/state/metadataScope";
import { PreviewTabs } from "../features/fetchxml/ui/RightPane/PreviewTabs";
import type { RelatedEntityColumn } from "../features/fetchxml/ui/RightPane/AddColumnsPanel";
import { EntitySelector } from "../features/fetchxml/ui/Toolbar/EntitySelector";
import { SaveViewButton } from "../features/fetchxml/ui/Toolbar/SaveViewButton";
import { SendToToolButton } from "../features/fetchxml/ui/Toolbar/SendToToolButton";
import {
	isT2TSupported,
	resolveLaunchPrefill,
	viewPrefillNotice,
} from "../features/fetchxml/api/invocation";
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
	whoAmI,
	isDataverseAvailable,
	resolveEntitySet,
	toLoadedViewInfo,
} from "../features/fetchxml/api/pptbClient";
import { SetStateDialog } from "../features/fetchxml/ui/Dialogs/SetStateDialog";
import { CloseRecordsDialog } from "../features/fetchxml/ui/Dialogs/CloseRecordsDialog";
import type { LoadedViewInfo } from "../features/fetchxml/api/pptbClient";
import { generateFetchXml } from "../features/fetchxml/model/fetchxml";
import { debugLog } from "../shared/utils/debug";
import { NON_RETURNING_LINK_TYPES } from "../features/fetchxml/model/fetchxmlRules";
import {
	planRemoveColumns,
	planSortByColumn,
	resolveColumn,
} from "../features/fetchxml/model/queryColumns";
import { relationshipKey } from "../features/fetchxml/model/manyToMany";
import {
	columnFilterConditions,
	type NewCondition,
} from "../features/fetchxml/model/columnFilters";
import { toAttributeTypeInfo } from "../features/fetchxml/model/attributeCapabilities";
import { buildViewLayoutXml } from "../features/fetchxml/model/layoutxml";
import {
	countQuickFindPlaceholders,
	getViewLoadNotice,
	toQuickFindSearchValue,
} from "../features/fetchxml/model/savedViews";
import type { SavedViewDetails } from "../features/fetchxml/ui/Dialogs/SaveViewDialog";
import { QuickFindTestValue } from "../features/fetchxml/ui/RightPane/QuickFindTestValue";
import { SettingsDrawer } from "../features/fetchxml/ui/Settings/SettingsDrawer";
import { usePersistedSettings } from "../shared/hooks/usePersistedSettings";
import { treeSizeFor } from "../features/fetchxml/model/displaySettings";
import { metadataCache } from "../features/fetchxml/state/cache";
import { resetDataverseSessionCache } from "../features/fetchxml/api/sessionCache";
import { nextConnectionScope, type ConnectionScope } from "../shared/hooks/connectionIdentity";

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
	const [connectionScope, setConnectionScope] = useState<ConnectionScope>({
		key: undefined,
		session: 0,
	});
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
				<ToolErrorBoundary
					readFetchXml={() => carryOver.current.readFetchXml?.()}
					onRestart={restart}
				>
					<BuilderProvider key={`${session}-${restarts}`}>
						<AppContent carryOver={carryOver.current} connectionKey={connectionKey} />
					</BuilderProvider>
				</ToolErrorBoundary>
			</ThemeProvider>
		</FluentProvider>
	);
}

/** connectionKey: "" with no connection, undefined until the host answers (usePptbContext) */
function AppContent({
	carryOver,
	connectionKey,
}: {
	carryOver: ConnectionCarryOver;
	connectionKey?: string;
}) {
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
		if (!result.success)
			console.error("Connection change: couldn't reload the query:", result.errors);
		// builder is intentionally omitted — runs once per session (mount)
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	// Lets AppShell read the current query when the connection changes (#66)
	const currentFetchQuery = builder.fetchQuery;
	useEffect(() => {
		carryOver.readFetchXml = () =>
			currentFetchQuery ? generateFetchXml(currentFetchQuery) : undefined;
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

				// A view: loaded as Load view does, with its layout, execution by ID and Save view's
				// overwrite; or a notification saying why it couldn't be opened (#197)
				if (action.kind === "viewUnavailable") {
					console.warn("T2T prefill: the view couldn't be opened:", action);
					await window.toolboxAPI?.utils?.showNotification?.(viewPrefillNotice(action));
					return;
				}
				if (action.kind === "view") {
					let entitySetName: string | undefined;
					try {
						entitySetName = (await resolveEntitySet(action.entityLogicalName))
							?.entitySetName;
						if (!entitySetName)
							throw new Error(
								`${action.entityLogicalName} has no entity set name in metadata.`,
							);
					} catch (error) {
						await window.toolboxAPI?.utils?.showNotification?.(
							viewPrefillNotice({
								kind: "viewUnavailable",
								viewType: action.view.type,
								id: action.view.id,
								error,
							}),
						);
						return;
					}
					applyLoadedViewRef.current(toLoadedViewInfo(action.view, entitySetName));
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

	// Quick find test value (#142): replaces {0} in the executed FetchXML only
	const [quickFindValue, setQuickFindValue] = useState("");
	const [quickFindAddPercent, setQuickFindAddPercent] = useState(true);

	// Settings drawer state
	const [settingsDrawerOpen, setSettingsDrawerOpen] = useState(false);
	const [displaySettings, setDisplaySettings] = usePersistedSettings();
	// The pickers follow "Advanced Find only" (#57). A layout effect, so the scope is set before the
	// children's effects load metadata.
	useLayoutEffect(() => {
		metadataScope.setAdvancedFindOnly(displaySettings.advancedFindOnly);
	}, [displaySettings.advancedFindOnly]);
	const advancedFindOnly = useAdvancedFindOnly();

	const { topHeight, leftPaneWidth, startHorizontalDrag, startVerticalDrag } = useSplitters();

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

	// Record action privileges and Activate / Deactivate / Close availability, checked when the
	// entity changes (#179, #143, #149). Here, so these checks keep their place among the effects.
	const entityLogicalName = builder.fetchQuery?.entity.name;
	const recordActionAvailability = useRecordActionAvailability(entityLogicalName);

	// The query's entity metadata: attributes, the root entity, its relationships (#301)
	const {
		attributeMetadata,
		entityMetadata,
		lookupRelationships,
		oneToManyRelationships,
		isLoadingRelationships,
	} = useRootEntityMetadata({
		fetchQuery: builder.fetchQuery,
		advancedFindOnly,
		loadEntityMetadata,
		loadRelationships,
	});

	// Generate FetchXML from builder state — exactly what the user built, nothing injected (#59).
	const fetchXml = builder.fetchQuery ? generateFetchXml(builder.fetchQuery) : "";
	const quickFindSearchValue = toQuickFindSearchValue(quickFindValue, quickFindAddPercent);

	// Running the query and paging its results (#301). Here, so clearing the results on an entity
	// change keeps its place among the effects.
	const queryExecution = useQueryExecution({
		fetchQuery: builder.fetchQuery,
		fetchXml,
		quickFindSearchValue,
		loadedView: builder.loadedView,
		retrieveAllRecords: builder.retrieveAllRecords,
	});
	const { queryResult, executeError, pagingRef, handleExecuteRef } = queryExecution;

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

	// Quick find views keep {0} in their conditions (#142): a test value is needed to execute them
	const quickFindPlaceholderCount = useMemo(
		() => countQuickFindPlaceholders(fetchXml),
		[fetchXml],
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

	const resultsExport = useResultsExport({
		loadedView: builder.loadedView,
		fetchQuery: builder.fetchQuery,
		columnConfig: builder.columnConfig,
		fetchXml,
		layoutXml,
		queryResult,
		attributeMetadata,
		entityMetadata,
		displaySettings,
	});

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

	// Column filters from the results grid (#23): they change the query, which then runs again.
	// Running waits for the change to commit, so it runs the new query.
	const columnFilters = useMemo(
		() => columnFilterConditions(builder.fetchQuery, builder.columnFilterId),
		[builder.fetchQuery, builder.columnFilterId],
	);
	const [columnFilterRunRequest, setColumnFilterRunRequest] = useState(0);
	useEffect(() => {
		if (columnFilterRunRequest > 0) void handleExecuteRef.current();
	}, [columnFilterRunRequest, handleExecuteRef]);
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

	// The record actions: their dialogs, the selection, and why each action is off (#301)
	const recordActions = useRecordActions({
		entityLogicalName,
		fetchQuery: builder.fetchQuery,
		entityMetadata,
		queryResult,
		executeError,
		pagingRef,
		handleExecuteRef,
		availability: recordActionAvailability,
	});

	/** Loads a view into the tree as Load view does: its layout, execution by ID, Save view's overwrite (#93, #142, #197) */
	const applyLoadedView = (viewInfo: LoadedViewInfo) => {
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
	};
	// The launch prefill runs once, after the first render (#197)
	const applyLoadedViewRef = useRef(applyLoadedView);
	applyLoadedViewRef.current = applyLoadedView;

	return (
		<div className={styles.root} data-app-root>
			{/* Left Pane */}
			<div className={styles.leftPane} data-left-pane style={{ width: `${leftPaneWidth}px` }}>
				{/* Top: Tree View */}
				<div className={styles.leftPaneTop} style={{ height: `${topHeight}%` }}>
					<PaneErrorBoundary
						where="Query scope"
						title="The query scope couldn't be shown"
						resetKey={connectionKey}
					>
						<EntitySelector
							selectedEntity={builder.fetchQuery?.entity.name || null}
							onEntityChange={builder.setEntity}
							onNewQuery={builder.newQuery}
							entityScopeMode={displaySettings.entityScopeMode}
							onEntityScopeModeChange={(entityScopeMode) =>
								setDisplaySettings({ ...displaySettings, entityScopeMode })
							}
							advancedFindOnly={displaySettings.advancedFindOnly}
							onAdvancedFindOnlyChange={(advancedFindOnly) =>
								setDisplaySettings({ ...displaySettings, advancedFindOnly })
							}
							onViewLoad={applyLoadedView}
						/>
					</PaneErrorBoundary>
					{builder.fetchQuery ? (
						<PaneErrorBoundary
							where="Query tree"
							title="The query tree couldn't be shown"
							resetKey={builder.fetchQuery}
						>
							<TreeView
								fetchQuery={builder.fetchQuery}
								size={treeSizeFor(displaySettings.gridDensity)}
								iconStyle={displaySettings.treeIconStyle}
								attributeMetadata={attributeMetadata}
								selectedNodeId={builder.selectedNodeId}
								revealSeq={builder.revealSeq}
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
					onMouseDown={startHorizontalDrag}
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
						<PaneErrorBoundary
							where="Properties panel"
							title="The properties couldn't be shown"
							resetKey={builder.selectedNodeId}
						>
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
				onMouseDown={startVerticalDrag}
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
					isExecuting={queryExecution.isExecuting}
					isLoadingMore={queryExecution.isLoadingMore}
					onExecute={queryExecution.execute}
					onExport={resultsExport.exportToExcel}
					onExportLocal={resultsExport.exportLocal}
					canExport={resultsExport.canExport}
					isExporting={resultsExport.exportStatus.isExporting}
					exportError={resultsExport.exportStatus.error}
					onDismissExportError={resultsExport.dismissExportError}
					executeError={executeError}
					onDismissExecuteError={queryExecution.dismissExecuteError}
					exportDisabledReason={resultsExport.exportDisabledReason}
					onParseToTree={builder.loadFetchXml}
					isCallee={launch.isCallee}
					attributeMetadata={attributeMetadata}
					fetchQuery={builder.fetchQuery}
					primaryIdAttribute={entityMetadata?.PrimaryIdAttribute}
					connectionKey={connectionKey}
					selectedNodeId={builder.selectedNodeId}
					onSelectNode={builder.revealNode}
					onAddOrder={builder.addOrder}
					columnConfig={builder.columnConfig}
					onColumnResize={builder.updateColumnWidth}
					onLoadMore={queryExecution.loadMore}
					pagingNotice={queryExecution.pagingNotice}
					pagingWaitSeconds={queryExecution.pagingWaitSeconds}
					isRetrieveAllInProgress={queryExecution.isRetrieveAllInProgress}
					onStopRetrieveAll={queryExecution.stopRetrieveAll}
					onRetryPaging={queryExecution.retryPaging}
					onDismissPagingNotice={queryExecution.dismissPagingNotice}
					entityDisplayName={
						entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
						entityMetadata?.LogicalName
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
								disabled={queryExecution.isExecuting}
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
							entityDisplayName={
								entityMetadata?.DisplayName?.UserLocalizedLabel?.Label
							}
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
							? attributeMetadata
									?.get(column.container.name)
									?.get(column.attribute.name)
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
					onOpenRecord={recordActions.openRecord}
					onCopyRecordUrl={recordActions.copyRecordUrl}
					onDeleteRecords={recordActions.deleteRecords}
					onBulkDeleteRecords={recordActions.bulkDeleteRecords}
					onActivateRecords={recordActions.activateRecords}
					onDeactivateRecords={recordActions.deactivateRecords}
					activateDisabledReason={recordActions.activateDisabledReason}
					deactivateDisabledReason={recordActions.deactivateDisabledReason}
					closeActions={recordActions.closeActions}
					closeDisabledReason={recordActions.closeDisabledReason}
					onCloseRecords={recordActions.closeRecords}
					onRunSpecificWorkflow={recordActions.runSpecificWorkflow}
					canDelete={recordActions.canDelete}
					canBulkDelete={recordActions.canBulkDelete}
					canRunWorkflow={recordActions.canRunWorkflow}
					onFetchWorkflows={recordActions.fetchWorkflows}
					isAggregateQuery={recordActions.isAggregateQuery}
					hasOneToManyRelationship={recordActions.hasOneToManyRelationship}
					recordActionsDisabledReason={recordActions.recordActionsDisabledReason}
					deleteAllDisabledReason={recordActions.deleteAllDisabledReason}
					onSelectionChange={recordActions.onSelectionChange}
					getSelectedRecordIds={recordActions.getSelectedRecordIds}
					onOpenSettings={() => setSettingsDrawerOpen(true)}
					displaySettings={displaySettings}
					onGridDensityChange={(gridDensity) =>
						setDisplaySettings({ ...displaySettings, gridDensity })
					}
				/>
			</div>
			{/* Delete Confirmation Dialog */}
			<DeleteConfirmDialog
				open={recordActions.deleteDialogState.open}
				recordName={recordActions.deleteDialogState.recordName}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
					entityLogicalName ||
					"record"
				}
				entityLogicalName={entityLogicalName}
				recordCount={recordActions.deleteDialogState.recordIds.length}
				onClose={recordActions.dismissDeleteDialog}
				onConfirm={recordActions.confirmDelete}
				onProgress={recordActions.deleteProgress ?? undefined}
			/>
			{/* Bulk Delete Dialog */}
			<BulkDeleteDialog
				open={recordActions.bulkDeleteDialogState.open}
				recordCount={recordActions.bulkDeleteDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
					entityLogicalName ||
					"records"
				}
				entityLogicalName={entityLogicalName}
				isAllRecords={recordActions.bulkDeleteDialogState.isAllRecords}
				totalViewRecords={recordActions.bulkDeleteDialogState.totalViewRecords}
				moreRecords={queryResult?.moreRecords}
				onClose={recordActions.dismissBulkDeleteDialog}
				onConfirm={recordActions.confirmBulkDelete}
			/>
			{/* Workflow Picker Dialog */}
			<WorkflowPickerDialog
				open={recordActions.workflowDialogState.open}
				recordCount={recordActions.workflowDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
					entityLogicalName ||
					"records"
				}
				entityLogicalName={entityLogicalName}
				preSelectedWorkflow={recordActions.workflowDialogState.preSelectedWorkflow}
				onClose={recordActions.dismissWorkflowDialog}
				onFetchWorkflows={recordActions.fetchWorkflows}
				onExecute={(workflow, onProgress) =>
					recordActions.executeWorkflow(
						workflow,
						recordActions.workflowDialogState.recordIds,
						onProgress,
					)
				}
			/>
			{/* Activate / Deactivate Dialog (#143) */}
			<SetStateDialog
				open={recordActions.setStateDialogState.open}
				action={recordActions.setStateDialogState.action}
				recordIds={recordActions.setStateDialogState.recordIds}
				entityLogicalName={entityLogicalName}
				entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
				primaryIdAttribute={entityMetadata?.PrimaryIdAttribute}
				closeAction={recordActions.setStateDialogState.closeAction}
				onClose={recordActions.dismissSetStateDialog}
				onExecute={recordActions.executeStateChange}
			/>
			{/* Close message Dialog (#149) */}
			<CloseRecordsDialog
				open={recordActions.closeDialogState.open}
				action={recordActions.closeDialogState.action}
				recordIds={recordActions.closeDialogState.recordIds}
				entityLogicalName={entityLogicalName}
				entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
				onClose={recordActions.dismissCloseDialog}
				onExecute={recordActions.executeClose}
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
