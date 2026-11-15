/**
 * What the tool loads on its own (#301): the query carried over a connection change (#66), the
 * prefill another tool launched it with (tool-to-tool, #197), and a view loaded as Load view does
 * (`applyLoadedView`: its layout, execution by ID and Save view's overwrite, #93, #142, #197).
 *
 * The prefill is applied once per tool instance, not per connection: `carryOver` lives in AppShell
 * and outlives the remount a connection change causes (#66).
 *
 * AppShell calls the hook where these effects were, so they keep their place among its effects. It
 * passes `launch` in (from useLaunchContext, which stays in AppShell), so the host's launch context
 * is still read once and its effect keeps its place too.
 */
import { useEffect, useRef } from "react";
import { resolveLaunchPrefill, viewPrefillNotice } from "../../features/fetchxml/api/invocation";
import { resolveEntitySet, toLoadedViewInfo } from "../../features/fetchxml/api/pptbClient";
import type { LoadedViewInfo } from "../../features/fetchxml/api/pptbClient";
import { generateFetchXml } from "../../features/fetchxml/model/fetchxml";
import { getViewLoadNotice } from "../../features/fetchxml/model/savedViews";
import type { BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import type { LaunchContextState } from "../../shared/hooks/useLaunchContext";

/**
 * What outlives a connection change (#66): AppContent is remounted for the new connection, so
 * this lives in AppShell.
 */
export interface ConnectionCarryOver {
	/** The launch prefill (T2T) was applied: never apply it again */
	prefillConsumed: boolean;
	/** Reads the current query's FetchXML (set by AppContent) */
	readFetchXml?: () => string | undefined;
	/** The query to load into the remounted AppContent */
	fetchXml?: string;
}

export interface LaunchPrefillInputs {
	/** The query builder (its actions are stable, #65) */
	builder: BuilderContextValue;
	/** AppShell's carry-over, shared by every AppContent of this tool instance */
	carryOver: ConnectionCarryOver;
	/** The launch context from useLaunchContext */
	launch: LaunchContextState;
	/** After a view is loaded: a quick find test value belongs to the view it was entered for (#142) */
	onViewApplied: () => void;
}

export interface LaunchPrefill {
	/** Loads a view into the tree as Load view does (#93, #142, #197) */
	applyLoadedView: (viewInfo: LoadedViewInfo) => void;
}

export function useLaunchPrefill({
	builder,
	carryOver,
	launch,
	onViewApplied,
}: LaunchPrefillInputs): LaunchPrefill {
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
		onViewApplied();
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

	return { applyLoadedView };
}
