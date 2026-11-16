/**
 * Main application shell with Fluent UI theming and layout
 */

import { useState, useEffect, useLayoutEffect, useMemo, useRef } from "react";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import { usePptbContext } from "../shared/hooks/usePptbContext";
import { useLazyMetadata } from "../shared/hooks/useLazyMetadata";
import { useSplitters } from "./hooks/useSplitters";
import { useResultsExport } from "./hooks/useResultsExport";
import { useRootEntityMetadata } from "./hooks/useRootEntityMetadata";
import { useRecordActionAvailability, useRecordActions } from "./hooks/useRecordActions";
import { useQueryExecution } from "./hooks/useQueryExecution";
import { useLaunchPrefill, type ConnectionCarryOver } from "./hooks/useLaunchPrefill";
import { useViewLayout } from "./hooks/useViewLayout";
import { useTreeNotice } from "./hooks/useTreeNotice";
import { TreeNoticeBar } from "./TreeNoticeBar";
import { useStyles } from "./AppShell.styles";
import { RecordActionDialogs } from "./RecordActionDialogs";
import { useResultsGridEdits } from "./hooks/useResultsGridEdits";
import { metadataScope, useAdvancedFindOnly } from "../features/fetchxml/state/metadataScope";
import { PreviewTabs } from "../features/fetchxml/ui/RightPane/PreviewTabs";
import { EntitySelector } from "../features/fetchxml/ui/Toolbar/EntitySelector";
import { SaveViewButton } from "../features/fetchxml/ui/Toolbar/SaveViewButton";
import { SendToToolButton } from "../features/fetchxml/ui/Toolbar/SendToToolButton";
import { isT2TSupported } from "../features/fetchxml/api/invocation";
import { useLaunchContext } from "../shared/hooks/useLaunchContext";
import { useSendToTools } from "../shared/hooks/useSendToTools";
import { TreeView } from "../features/fetchxml/ui/LeftPane/TreeView";
import { PropertiesPanel } from "../features/fetchxml/ui/LeftPane/PropertiesPanel";
import { clearMetadataLoadErrors } from "../features/fetchxml/ui/LeftPane/PropertiesPanel/metadataLoadErrors";
import { BuilderProvider } from "../features/fetchxml/state/builderStore";
import { useBuilder } from "../features/fetchxml/state/builderContext";
import { ThemeProvider } from "../shared/contexts/ThemeContext";
import { NotifyProvider } from "../shared/feedback/NotifyProvider";
import { useNotify } from "../shared/feedback/useNotify";
import { PaneErrorBoundary, ToolErrorBoundary } from "../shared/components/ErrorBoundary";
import { whoAmI, isDataverseAvailable } from "../features/fetchxml/api/pptbClient";
import { generateFetchXml } from "../features/fetchxml/model/fetchxml";
import { debugLog } from "../shared/utils/debug";
import {
	savedViewToast,
	countQuickFindPlaceholders,
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
				{/* One Toaster for the app (#277): above the error boundary, so the fallback's Copy
				    details still toasts; above the remount, so a new session or a restart clears
				    toasts whose actions point at the old builder */}
				<NotifyProvider resetKey={`${session}-${restarts}`}>
					<ToolErrorBoundary
						readFetchXml={() => carryOver.current.readFetchXml?.()}
						onRestart={restart}
					>
						<BuilderProvider key={`${session}-${restarts}`}>
							<AppContent
								carryOver={carryOver.current}
								connectionKey={connectionKey}
							/>
						</BuilderProvider>
					</ToolErrorBoundary>
				</NotifyProvider>
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
	const { notify } = useNotify();
	const { loadAttributes, loadEntityMetadata, loadRelationships } = useLazyMetadata();

	// Tool-to-Tool (T2T) launch context: detected once, shared across the app.
	// When present, another tool launched us with a prefill payload.
	const launch = useLaunchContext();

	// Tools we can send the current query to, discovered via the "fetchxml" capability.
	const { tools: sendToTools } = useSendToTools();

	// Quick find test value (#142): replaces {0} in the executed FetchXML only. Above the launch
	// prefill, which clears it when a view is loaded.
	const [quickFindValue, setQuickFindValue] = useState("");
	const [quickFindAddPercent, setQuickFindAddPercent] = useState(true);

	// The query carried over a connection change, the launch prefill and loading a view (#301)
	// A MessageBar above the tree about the query in it: a launch that couldn't be loaded, a view
	// loaded with changes (#277)
	const treeNotice = useTreeNotice(builder.fetchQuery?.entity.id);
	const { applyLoadedView } = useLaunchPrefill({
		builder,
		carryOver,
		launch,
		onViewApplied: () => setQuickFindValue(""),
		showTreeNotice: treeNotice.show,
	});

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

	// The column config kept in step with the query, and its LayoutXML (#301). Here, so the sync
	// effect keeps its place among the effects.
	const { layoutXml } = useViewLayout({ builder, attributeMetadata, entityMetadata });

	// Quick find views keep {0} in their conditions (#142): a test value is needed to execute them
	const quickFindPlaceholderCount = useMemo(
		() => countQuickFindPlaceholders(fetchXml),
		[fetchXml],
	);

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
		// Says the view was saved, and when a system view isn't published yet (#277, owner decision 5;
		// the Publish action comes with #278's Save menu)
		notify.success(savedViewToast(viewName, details.published));
	};

	// The results grid's edits to the query: column filters, columns and sorting (#301). Here, so the
	// effect that runs the query after a filter change keeps its place among the effects.
	const gridEdits = useResultsGridEdits({ builder, attributeMetadata, handleExecuteRef });

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
							onNewQuery={() => {
								treeNotice.dismiss();
								builder.newQuery();
							}}
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
					<TreeNoticeBar notice={treeNotice.notice} onDismiss={treeNotice.dismiss} />
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
					onReorderColumns={gridEdits.reorderColumns}
					onRemoveColumns={gridEdits.removeColumns}
					onAddColumn={gridEdits.addColumn}
					onAddRelatedColumns={gridEdits.addRelatedColumns}
					columnFilters={gridEdits.columnFilters}
					onAddColumnFilter={gridEdits.addColumnFilter}
					onRemoveColumnFilters={gridEdits.removeColumnFilters}
					onSortChange={gridEdits.sortByColumn}
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
			{/* The record actions' dialogs: Delete, Bulk delete, Run workflow, Activate / Deactivate, Close */}
			<RecordActionDialogs
				recordActions={recordActions}
				entityLogicalName={entityLogicalName}
				entityMetadata={entityMetadata}
				moreRecords={queryResult?.moreRecords}
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
