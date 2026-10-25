/**
 * Remembers the publisher / solution selection per environment (#25), in toolboxAPI.settings like
 * the display settings (key "fetchxml-studio:scopeSelection", a map keyed by environment URL).
 *
 * Restored once per scope mode, after that mode's publishers / solutions have loaded; ids that no
 * longer exist are dropped silently. Saved on every change after that. Without toolboxAPI.settings
 * (dev server, tests) or a connection, nothing is remembered.
 */

import { useEffect, useRef, useState } from "react";
import { getPrimaryConnection } from "../../features/fetchxml/api/host";
import type { EntityScopeMode } from "../../features/fetchxml/model/displaySettings";
import {
	REMEMBERED_SCOPES_KEY,
	environmentKey,
	rememberedScopeFor,
	restorableIds,
	withRememberedScope,
	type RememberedScope,
} from "../../features/fetchxml/model/scopeSelection";
import { debugLog } from "../utils/debug";

interface RememberedScopeOptions {
	/** The effective scope mode; null while access is still being checked */
	mode: EntityScopeMode | null;
	/** The mode's publishers (publisher-solution) or solutions (solution-only) have loaded */
	ready: boolean;
	/** Publishers that exist (publisher-solution) */
	availablePublisherIds: readonly string[];
	/** Solutions that exist: for the given publishers in publisher-solution mode, all in solution-only */
	availableSolutionIds: (publisherIds: string[]) => string[];
	/** The current selection of the mode */
	publisherIds: readonly string[];
	solutionIds: readonly string[];
	/** Applies a restored selection */
	apply: (scope: RememberedScope) => void;
}

const sameIds = (a: readonly string[], b: readonly string[]) => a.length === b.length && a.every((id, i) => id === b[i]);

const settingsApi = () => (typeof window === "undefined" ? undefined : window.toolboxAPI?.settings);

export interface RememberedScopeState {
	/** The selection is remembered for this environment */
	remembers: boolean;
	/** The mode's remembered selection has been restored (or there was none to restore), so an empty
	 * selection now really is empty: the scope chip's first-run tip waits for this (#227) */
	restored: boolean;
}

export function useRememberedScope({
	mode,
	ready,
	availablePublisherIds,
	availableSolutionIds,
	publisherIds,
	solutionIds,
	apply,
}: RememberedScopeOptions): RememberedScopeState {
	// The stored map and this environment's key, read once: the selector remounts when the
	// connection changes (#66)
	const [loaded, setLoaded] = useState<{ stored: unknown; envKey: string } | null>(null);
	const stored = useRef<unknown>(undefined);
	// Modes whose selection has been restored; saving starts after that, so the empty selection
	// before the restore never overwrites the remembered one
	const [restoredModes, setRestoredModes] = useState<ReadonlySet<EntityScopeMode>>(new Set());

	useEffect(() => {
		let mounted = true;
		const settings = settingsApi();
		if (!settings?.get) {
			setLoaded({ stored: undefined, envKey: "" });
			return;
		}
		Promise.all([settings.get(REMEMBERED_SCOPES_KEY), getPrimaryConnection()])
			.then(([value, connection]) => {
				if (!mounted) return;
				stored.current = value;
				setLoaded({ stored: value, envKey: environmentKey(connection?.url) });
			})
			.catch((err: unknown) => {
				console.warn("Failed to read the remembered publisher / solution selection:", err);
				if (mounted) setLoaded({ stored: undefined, envKey: "" });
			});
		return () => {
			mounted = false;
		};
	}, []);

	// Restore once per mode, when its lists have loaded
	useEffect(() => {
		if (!loaded || !mode || mode === "all" || !ready || restoredModes.has(mode)) return;
		const remembered = rememberedScopeFor(stored.current, loaded.envKey);
		if (remembered) {
			const publishers = mode === "publisher-solution" ? restorableIds(remembered.publisherIds, availablePublisherIds) : [];
			const solutions = restorableIds(remembered.solutionIds, availableSolutionIds(publishers));
			debugLog("entityFilters", "[useRememberedScope] Restoring the selection:", { mode, publishers, solutions });
			if (mode === "solution-only" || publishers.length > 0) apply({ publisherIds: publishers, solutionIds: solutions });
		}
		setRestoredModes((modes) => new Set(modes).add(mode));
	}, [loaded, mode, ready, restoredModes, availablePublisherIds, availableSolutionIds, apply]);

	// Save every change after the restore
	useEffect(() => {
		if (!loaded?.envKey || !mode || mode === "all" || !restoredModes.has(mode)) return;
		const remembered = rememberedScopeFor(stored.current, loaded.envKey);
		const next: RememberedScope = {
			// Solution-only mode has no publisher picker: keep the publishers remembered before
			publisherIds: mode === "publisher-solution" ? [...publisherIds] : (remembered?.publisherIds ?? []),
			solutionIds: [...solutionIds],
		};
		if (remembered ? sameIds(remembered.publisherIds, next.publisherIds) && sameIds(remembered.solutionIds, next.solutionIds) : !next.publisherIds.length && !next.solutionIds.length) {
			return;
		}
		stored.current = withRememberedScope(stored.current, loaded.envKey, next);
		settingsApi()
			?.set?.(REMEMBERED_SCOPES_KEY, stored.current)
			.catch((err: unknown) => console.warn("Failed to remember the publisher / solution selection:", err));
	}, [loaded, mode, restoredModes, publisherIds, solutionIds]);

	return { remembers: !!loaded?.envKey, restored: mode === "all" || (!!mode && restoredModes.has(mode)) };
}
