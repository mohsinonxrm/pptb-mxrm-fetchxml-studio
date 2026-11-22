/**
 * Main application shell with Fluent UI theming and layout
 */

import { useState, useEffect, useLayoutEffect, useMemo, useRef, type ReactNode } from "react";
import {
	AriaLiveAnnouncer,
	FluentProvider,
	mergeClasses,
	webLightTheme,
	webDarkTheme,
} from "@fluentui/react-components";
import { usePptbContext } from "../shared/hooks/usePptbContext";
import { useLazyMetadata } from "../shared/hooks/useLazyMetadata";
import { useBuilderLayout, type PeekPart } from "./hooks/useBuilderLayout";
import { BuilderRail, CollapseButton, PinButton } from "./layout/BuilderRail";
import { PaneSplitter } from "../shared/components/PaneSplitter";
import { InspectorDrawer } from "./layout/InspectorDrawer";
import { SelectionBar } from "./layout/SelectionBar";
import { selectionLabel } from "./layout/selectionLabel";
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
import { AppBar } from "../features/fetchxml/ui/AppBar/AppBar";
import { EditorSyncContext } from "../features/fetchxml/ui/RightPane/editorSync";
import { useEditorSession } from "./hooks/useEditorSession";
import { useRunShortcut } from "./hooks/useRunShortcut";
import { useUndoRedo } from "./hooks/useUndoRedo";
import { useUndoShortcuts } from "./hooks/useUndoShortcuts";
import { isViewDirty } from "../features/fetchxml/model/viewDirty";
import { isT2TSupported } from "../features/fetchxml/api/invocation";
import { useLaunchContext } from "../shared/hooks/useLaunchContext";
import { useSendToTools } from "../shared/hooks/useSendToTools";
import { TreeView } from "../features/fetchxml/ui/LeftPane/TreeView";
import { TreeToolbar } from "../features/fetchxml/ui/LeftPane/TreeToolbar";
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

/** The builder pane and the query tree, which the splitters control (aria-controls) */
const BUILDER_PANE_ID = "fxs-builder-pane";
const QUERY_TREE_ID = "fxs-query-tree";

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
				{/* One live region for screen reader announcements (#192, #186): Undo and Redo say
				    what they did, and Fluent's MessageBars announce themselves through it */}
				<AriaLiveAnnouncer>
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
				</AriaLiveAnnouncer>
			</ThemeProvider>
		</FluentProvider>
	);
}

/** Focus moving into a peek (#234): the selected tree item, else the first field (the Entity picker,
 * or the first property, not a field's info button), else the first control. The Properties peek is
 * searched from its properties, below its Pin. */
function focusInside(container: HTMLElement | null) {
	if (!container) return;
	const target =
		container.querySelector<HTMLElement>('[role="treeitem"][aria-selected="true"]') ??
		container.querySelector<HTMLElement>('[role="treeitem"][tabindex="0"]') ??
		container.querySelector<HTMLElement>(
			'input:not([disabled]):not([type="hidden"]), textarea:not([disabled]), [role="combobox"]:not([disabled]):not([aria-disabled="true"])',
		) ??
		container.querySelector<HTMLElement>('button:not([disabled]), [tabindex="0"]');
	target?.focus();
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

	// The builder pane's sizes, remembered, and where Properties live (#234)
	const layout = useBuilderLayout();
	const editPropertiesRef = useRef<HTMLButtonElement>(null);
	const closeInspector = () => {
		layout.setInspectorOpen(false);
		editPropertiesRef.current?.focus();
	};
	// The drawer has nothing to show once nothing is selected (New, Open view, a deleted node)
	const { setInspectorOpen } = layout;
	const hasSelection = !!builder.selectedNode;
	useEffect(() => {
		if (!hasSelection) setInspectorOpen(false);
	}, [hasSelection, setInspectorOpen]);
	// Collapse, the rail and the peek (#234). Collapse focuses Expand; Expand and Pin focus Collapse;
	// a peek takes focus inside itself, and Escape (or its rail button again) returns it there.
	const collapseRef = useRef<HTMLButtonElement>(null);
	const expandRef = useRef<HTMLButtonElement>(null);
	const railTreeRef = useRef<HTMLButtonElement>(null);
	const railPropertiesRef = useRef<HTMLButtonElement>(null);
	const paneRef = useRef<HTMLElement>(null);
	const treeRegionRef = useRef<HTMLDivElement>(null);
	const peekPropertiesRef = useRef<HTMLDivElement>(null);
	const pendingFocus = useRef<"collapse" | "expand" | "peek" | null>(null);
	const railButton = (part: PeekPart) =>
		part === "tree" ? railTreeRef.current : railPropertiesRef.current;
	const collapsePane = () => {
		pendingFocus.current = "expand";
		layout.collapse();
	};
	const dockPane = () => {
		pendingFocus.current = "collapse";
		layout.dock();
	};
	const togglePeek = (part: PeekPart) => {
		if (layout.peekPart === part) {
			layout.closePeek();
			return;
		}
		pendingFocus.current = "peek";
		layout.openPeek(part);
	};
	useEffect(() => {
		const target = pendingFocus.current;
		if (!target) return;
		pendingFocus.current = null;
		// Docking while the pane is forced collapsed (#296) leaves it collapsed: focus stays on the rail
		if (target === "collapse") (collapseRef.current ?? expandRef.current)?.focus();
		else if (target === "expand") expandRef.current?.focus();
		else
			focusInside(
				layout.peekPart === "props" ? peekPropertiesRef.current : treeRegionRef.current,
			);
	}, [layout.pane, layout.peekPart]);
	// Escape in the peek closes it (unless something inside handled it, like an open dropdown)
	const { closePeek, peekPart } = layout;
	useEffect(() => {
		const peek = paneRef.current;
		if (!peek || !peekPart) return;
		const onKeyDown = (event: KeyboardEvent) => {
			if (event.key !== "Escape" || event.defaultPrevented) return;
			const target = event.target as HTMLElement | null;
			// An open dropdown or menu trigger closes its popup first (an expanded tree item isn't one)
			const popupOpen =
				target?.getAttribute("aria-expanded") === "true" &&
				(target.getAttribute("role") === "combobox" ||
					target.hasAttribute("aria-haspopup"));
			if (popupOpen) return;
			event.preventDefault();
			closePeek();
			railButton(peekPart)?.focus();
		};
		peek.addEventListener("keydown", onKeyDown);
		return () => peek.removeEventListener("keydown", onKeyDown);
	}, [closePeek, peekPart]);
	// When Properties move between the pane and the drawer, focus that was in them would fall to the
	// page: it goes to Edit properties, to Expand when the pane is collapsed, or to the selected tree
	// item when Properties are docked
	const lastHost = useRef(layout.propertiesHost);
	useEffect(() => {
		if (lastHost.current === layout.propertiesHost) return;
		lastHost.current = layout.propertiesHost;
		if (document.activeElement && document.activeElement !== document.body) return;
		if (layout.propertiesHost === "drawer") {
			editPropertiesRef.current?.focus();
			return;
		}
		if (layout.propertiesHost === "peek") {
			expandRef.current?.focus();
			return;
		}
		const tree = document.getElementById(QUERY_TREE_ID);
		const item =
			tree?.querySelector<HTMLElement>('[role="treeitem"][aria-selected="true"]') ??
			tree?.querySelector<HTMLElement>('[role="treeitem"][tabindex="0"]');
		item?.focus();
	}, [layout.propertiesHost]);

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
			const target = builder.viewIdentity?.view;
			const overwritten = target?.id === viewId ? target : null;
			builder.setLoadedView(
				{
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
					// A new personal view is the saver's own (#278)
					ownerId: overwritten?.ownerId ?? details.ownerId,
					ownerName: overwritten?.ownerName,
				},
				// What was saved is the view's baseline for "Unsaved changes" (#278)
				layoutXml,
			);
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

	// The editor session and Run (#278): shared by the app bar and the right pane
	const editorSession = useEditorSession({
		xml: fetchXml,
		onParseToTree: builder.loadFetchXml,
		onRun: queryExecution.execute,
		runDisabledReason:
			quickFindPlaceholderCount > 0 && !quickFindSearchValue
				? "Enter a quick find value first"
				: undefined,
		isRunning: queryExecution.isExecuting,
		isLoadingMore: queryExecution.isLoadingMore,
	});
	useRunShortcut(editorSession.run, editorSession.runDisabledReason);
	// Undo and Redo (#192): the tree toolbar and Ctrl+Z / Ctrl+Y outside text fields, announced
	const undoRedo = useUndoRedo(builder.history);
	useUndoShortcuts({
		...undoRedo,
		restoreFocus: () =>
			treeRegionRef.current
				?.querySelector<HTMLElement>('[role="treeitem"][aria-selected="true"]')
				?.focus(),
	});

	// The view the query is, and whether it has unsaved changes (#278, owner decision Q4): its
	// FetchXML (the editor's in editor mode), columns or widths differ from the view as loaded or saved
	const viewIdentity = builder.viewIdentity;
	// A query counts once there's a tree, or FetchXML in the editor (editor mode works without a tree)
	const isUnsavedWith = (currentFetchXml: string) =>
		(!!builder.fetchQuery || currentFetchXml.trim() !== "") &&
		isViewDirty({
			baselineFetchXml: viewIdentity?.baselineFetchXml ?? null,
			currentFetchXml,
			baselineLayoutXml: viewIdentity?.baselineLayoutXml ?? null,
			currentLayoutXml: layoutXml,
		}).dirty;
	// The badge follows the editor's text a moment after typing stops
	const unsaved = isUnsavedWith(
		editorSession.isEditorModeActive && editorSession.editorText.trim()
			? editorSession.editorText
			: fetchXml,
	);
	// New decides from the editor's text now, so what was typed a moment ago isn't lost unasked
	const hasUnsavedChanges = () => {
		const current = editorSession.getCurrentXml();
		return isUnsavedWith(
			editorSession.isEditorModeActive && current.trim() ? current : fetchXml,
		);
	};
	/** New (#278): back to the first-run state */
	const startNewQuery = () => {
		treeNotice.dismiss();
		builder.newQuery();
		setQuickFindValue("");
		queryExecution.dismissExecuteError();
		editorSession.resetEditor();
	};

	// The builder pane's parts (#234): docked, they're one aside; collapsed, the rail opens either
	// the tree (with the scope) or Properties in a peek over the output
	const scopeSelector = (action: ReactNode) => (
		<PaneErrorBoundary
			where="Query scope"
			title="The query scope couldn't be shown"
			resetKey={connectionKey}
		>
			<EntitySelector
				selectedEntity={builder.fetchQuery?.entity.name || null}
				onEntityChange={builder.setEntity}
				entityScopeMode={displaySettings.entityScopeMode}
				onEntityScopeModeChange={(entityScopeMode) =>
					setDisplaySettings({ ...displaySettings, entityScopeMode })
				}
				advancedFindOnly={displaySettings.advancedFindOnly}
				onAdvancedFindOnlyChange={(advancedFindOnly) =>
					setDisplaySettings({ ...displaySettings, advancedFindOnly })
				}
				headerAction={action}
			/>
		</PaneErrorBoundary>
	);
	const { canUndo, canRedo, undoLabel, redoLabel } = builder.history;
	const treeBody = (
		<>
			<TreeToolbar
				hasQuery={!!builder.fetchQuery}
				canUndo={canUndo}
				canRedo={canRedo}
				undoLabel={undoLabel}
				redoLabel={redoLabel}
				onUndo={undoRedo.undo}
				onRedo={undoRedo.redo}
			/>
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
		</>
	);
	const propertiesPanel = (
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
	);

	return (
		<EditorSyncContext.Provider value={editorSession.syncValue}>
			<div className={styles.shell}>
				<AppBar
					run={{
						onClick: editorSession.run,
						busy: editorSession.isRunning,
						disabledReason: editorSession.runDisabledReason,
					}}
					saveView={{
						fetchXml,
						layoutXml,
						entityLogicalName: builder.fetchQuery?.entity?.name || "",
						entityDisplayName: entityMetadata?.DisplayName?.UserLocalizedLabel?.Label,
						objectTypeCode: entityMetadata?.ObjectTypeCode || 0,
						primaryIdAttribute: entityMetadata?.PrimaryIdAttribute || "",
						// What Save overwrites: the view the query is, kept across edits (#278)
						targetView: builder.viewIdentity?.view ?? null,
						onSaveComplete: handleSaveViewComplete,
						disabled: !fetchXml || !entityMetadata,
					}}
					sendToTool={
						isT2TSupported() && sendToTools.length > 0
							? {
									fetchXml,
									entityLogicalName: builder.fetchQuery?.entity?.name || "",
									tools: sendToTools,
									disabled: !fetchXml || !builder.fetchQuery?.entity?.name,
								}
							: undefined
					}
					isCallee={launch.isCallee}
					getCurrentXml={editorSession.getCurrentXml}
					onOpenSettings={() => setSettingsDrawerOpen(true)}
					viewName={viewIdentity?.name ?? null}
					unsaved={unsaved}
					hasUnsavedChanges={hasUnsavedChanges}
					onNewQuery={startNewQuery}
					openView={{
						entityLogicalName: builder.fetchQuery?.entity?.name,
						entityDisplayName: entityMetadata?.DisplayName?.UserLocalizedLabel?.Label,
						entitySetName: entityMetadata?.EntitySetName,
						currentViewId: viewIdentity?.view?.id,
						onOpen: (viewInfo) => {
							const loaded = applyLoadedView(viewInfo);
							// The editor follows the view: its text was the previous query's, saved or
							// discarded by now. A view that couldn't be loaded leaves everything as it was.
							if (loaded) editorSession.discardEditor();
							return loaded;
						},
					}}
				/>
				<div className={styles.root}>
					{/* The builder pane (#234): docked, the tree then Properties (or, when narrow, a
					    selection bar and the inspector drawer); collapsed, a rail that opens either part in
					    a peek over the output */}
					{layout.pane !== "docked" && (
						<BuilderRail
							peekPart={layout.peekPart}
							peekId={BUILDER_PANE_ID}
							onExpand={dockPane}
							onTogglePeek={togglePeek}
							expandRef={expandRef}
							treeRef={railTreeRef}
							propertiesRef={railPropertiesRef}
						/>
					)}
					{/* One pane in every state, so the tree and the scope keep their state (open
					    branches, the access check): hidden while collapsed, over the output in a peek */}
					<aside
						id={BUILDER_PANE_ID}
						ref={paneRef}
						aria-label={
							layout.pane !== "peek"
								? "Query builder"
								: layout.peekPart === "props"
									? "Properties"
									: "Query tree"
						}
						className={
							layout.pane === "peek"
								? mergeClasses(styles.leftPane, styles.peek)
								: styles.leftPane
						}
						style={{
							width: `${layout.pane === "peek" ? layout.peekWidth : layout.builderWidth}px`,
							display: layout.pane === "collapsed" ? "none" : undefined,
						}}
					>
						<div
							id={QUERY_TREE_ID}
							ref={treeRegionRef}
							className={styles.leftPaneTop}
							style={
								layout.peekPart === "props"
									? { display: "none" }
									: layout.pane === "docked" && layout.propertiesHost === "docked"
										? { height: `${layout.treeHeight}px` }
										: { flex: "1 1 0" }
							}
						>
							{scopeSelector(
								<>
									{layout.pane === "peek" ? (
										<PinButton onClick={dockPane} />
									) : (
										<CollapseButton ref={collapseRef} onClick={collapsePane} />
									)}
								</>,
							)}
							{treeBody}
						</div>
						{layout.pane === "docked" ? (
							layout.propertiesHost === "docked" ? (
								<>
									<PaneSplitter
										orientation="horizontal"
										value={layout.treeHeight}
										min={layout.treeHeightLimits.min}
										max={layout.treeHeightLimits.max}
										defaultValue={layout.treeHeightLimits.defaultValue}
										onChange={layout.setTreeHeight}
										onCommit={layout.commitTreeHeight}
										label="Resize the query tree (size is remembered)"
										description="Drag or use arrow keys; double-click resets."
										controls={QUERY_TREE_ID}
									/>
									<div className={styles.leftPaneBottom}>
										<div className={styles.toolbar}>
											<span style={{ fontWeight: 600 }}>Properties</span>
										</div>
										<div className={styles.propertiesContent}>
											{propertiesPanel}
										</div>
									</div>
								</>
							) : (
								<SelectionBar
									ref={editPropertiesRef}
									label={selectionLabel(builder.selectedNode)}
									hasSelection={!!builder.selectedNode}
									onEdit={() => layout.setInspectorOpen(true)}
								/>
							)
						) : layout.peekPart === "props" ? (
							<div className={styles.leftPaneBottom}>
								<div className={mergeClasses(styles.toolbar, styles.peekBar)}>
									<span style={{ fontWeight: 600 }}>Properties</span>
									<span className={styles.peekBarEnd}>
										<PinButton onClick={dockPane} />
									</span>
								</div>
								<div ref={peekPropertiesRef} className={styles.propertiesContent}>
									{propertiesPanel}
								</div>
							</div>
						) : null}
					</aside>
					{layout.pane === "docked" && (
						<PaneSplitter
							orientation="vertical"
							value={layout.builderWidth}
							min={layout.builderWidthLimits.min}
							max={layout.builderWidthLimits.max}
							defaultValue={layout.builderWidthLimits.defaultValue}
							onChange={layout.setBuilderWidth}
							onCommit={layout.commitBuilderWidth}
							label="Resize the builder pane (size is remembered)"
							description="Drag or use arrow keys; double-click resets."
							controls={BUILDER_PANE_ID}
						/>
					)}
					<main
						aria-label="Query output"
						className={styles.rightPane}
						onPointerDown={layout.pane === "peek" ? layout.closePeek : undefined}
					>
						<PreviewTabs
							bodyOverlay={
								layout.propertiesHost === "drawer" ? (
									<InspectorDrawer
										open={layout.inspectorOpen}
										title={selectionLabel(builder.selectedNode)}
										onClose={closeInspector}
									>
										{propertiesPanel}
									</InspectorDrawer>
								) : null
							}
							xml={fetchXml}
							layoutXml={layoutXml}
							result={queryResult}
							isExecuting={queryExecution.isExecuting}
							isLoadingMore={queryExecution.isLoadingMore}
							session={editorSession}
							onExport={resultsExport.exportToExcel}
							onExportLocal={resultsExport.exportLocal}
							canExport={resultsExport.canExport}
							isExporting={resultsExport.exportStatus.isExporting}
							exportError={resultsExport.exportStatus.error}
							onDismissExportError={resultsExport.dismissExportError}
							onRetryExport={resultsExport.retryExport}
							executeError={executeError}
							onDismissExecuteError={queryExecution.dismissExecuteError}
							exportDisabledReason={resultsExport.exportDisabledReason}
							onParseToTree={builder.loadFetchXml}
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
							deleteDisabledReason={recordActions.deleteDisabledReason}
							workflowDisabledReason={recordActions.workflowDisabledReason}
							canChangeState={recordActions.canChangeState}
							recordActionsDisabledReason={recordActions.recordActionsDisabledReason}
							deleteAllDisabledReason={recordActions.deleteAllDisabledReason}
							onSelectionChange={recordActions.onSelectionChange}
							getSelectedRecordIds={recordActions.getSelectedRecordIds}
							displaySettings={displaySettings}
							onGridDensityChange={(gridDensity) =>
								setDisplaySettings({ ...displaySettings, gridDensity })
							}
						/>
					</main>
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
			</div>
		</EditorSyncContext.Provider>
	);
}
