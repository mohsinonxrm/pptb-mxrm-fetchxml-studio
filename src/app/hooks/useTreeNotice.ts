/**
 * A notice about the query in the tree (#277): a MessageBar above the tree for things the user
 * needs to read where they happened, not as a passing toast (owner decisions 2 and 3):
 * - a launch from another tool that couldn't be loaded (a `viewRef` to a missing view, FetchXML
 *   that doesn't parse): it stays until a query starts or New is pressed;
 * - "View loaded with changes": it stays until the next load or New.
 *
 * A notice belongs to the query it was raised for, by its root entity node's id. Loading a view or
 * FetchXML, Parse to tree and choosing an entity make a new root node, so the notice goes; edits to
 * the query keep it. A launch failure is raised with no query (no root), so the first query clears
 * it. New clears it explicitly, because New on an empty tree leaves no root to compare.
 */
import { useCallback, useEffect, useState } from "react";

export interface TreeNoticeInput {
	intent: "error" | "warning";
	title: string;
	body: string;
}

export interface TreeNotice {
	/** The notice to show, or null */
	notice: TreeNoticeInput | null;
	/** Shows a notice for the query whose root entity node has this id (undefined: no query) */
	show: (notice: TreeNoticeInput, rootId: string | undefined) => void;
	/** Clears it (the bar's Dismiss, New) */
	dismiss: () => void;
}

export function useTreeNotice(currentRootId: string | undefined): TreeNotice {
	const [state, setState] = useState<{ notice: TreeNoticeInput; rootId?: string } | null>(null);

	const show = useCallback(
		(notice: TreeNoticeInput, rootId: string | undefined) => setState({ notice, rootId }),
		[],
	);
	const dismiss = useCallback(() => setState(null), []);

	// Another query is in the tree: the notice was about the one before
	const stale = state !== null && state.rootId !== currentRootId;
	useEffect(() => {
		// Against the root as it is when the update runs, so a notice shown since isn't dropped
		if (stale) setState((s) => (s && s.rootId !== currentRootId ? null : s));
	}, [stale, currentRootId]);

	return { notice: state && !stale ? state.notice : null, show, dismiss };
}
