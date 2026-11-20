/**
 * What Save offers for the view the query is (#278, owner decision Q3), with the checks behind it:
 * the connected user (to tell a shared personal view from the user's own), prvWriteUserQuery and
 * Write on a personal view someone else owns, and the system view privileges. saveOptions (pure,
 * saveTarget.ts) turns the answers into the primary action and the menu.
 *
 * A check that hasn't answered hides the overwrite (so it doesn't flash in and out); one that failed
 * offers it with a warning, and Dataverse decides when the user saves.
 */
import { useEffect, useMemo, useState } from "react";
import {
	checkSavedQueryPrivileges,
	checkUserQueryPrivileges,
	checkUserQueryWriteAccess,
	whoAmI,
} from "../../api/pptbClient";
import { isSharedView } from "../../model/savedViews";
import {
	saveOptions,
	type CheckResult,
	type SaveOptions,
	type SaveTargetView,
} from "../../model/saveTarget";

export type SaveOptionsView = SaveTargetView & { id: string };

interface Answers {
	/** The view these answers are for */
	viewKey: string | null;
	currentUserId?: string;
	canWriteUserQueries?: CheckResult;
	canWriteSystemViews?: CheckResult;
	personalWriteAccess?: CheckResult;
}

const keyOf = (view: SaveOptionsView | null) => (view ? `${view.type}:${view.id}` : null);

export function useSaveOptions(view: SaveOptionsView | null): SaveOptions {
	const viewKey = keyOf(view);
	const [answers, setAnswers] = useState<Answers>({ viewKey: null });

	useEffect(() => {
		if (!view) return;
		let current = true;
		const key = keyOf(view);
		setAnswers({ viewKey: key });
		const set = (patch: Partial<Answers>) => {
			if (current)
				setAnswers((prev) => (prev.viewKey === key ? { ...prev, ...patch } : prev));
		};
		void (async () => {
			const user = await whoAmI().catch(() => null);
			const currentUserId = user?.UserId;
			set({ currentUserId });
			if (view.type === "personal") {
				const privileges = await checkUserQueryPrivileges();
				set({ canWriteUserQueries: privileges.failed ? "unknown" : privileges.canWrite });
				// Write on the view itself matters when someone else owns it (a share, #142)
				if (isSharedView(view, currentUserId) || !currentUserId || !view.ownerId) {
					const access = await checkUserQueryWriteAccess(view.id);
					set({ personalWriteAccess: access === undefined ? "unknown" : access });
				}
			} else {
				const privileges = await checkSavedQueryPrivileges();
				set({ canWriteSystemViews: privileges.failed ? "unknown" : privileges.canWrite });
			}
		})();
		return () => {
			current = false;
		};
		// The checks follow the view, not each render's object
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [viewKey]);

	return useMemo(() => {
		// Answers for another view (just switched) count as not yet answered
		const fresh = answers.viewKey === viewKey ? answers : { viewKey };
		return saveOptions({
			view,
			currentUserId: fresh.currentUserId,
			canWriteUserQueries: fresh.canWriteUserQueries,
			canWriteSystemViews: fresh.canWriteSystemViews,
			personalWriteAccess: fresh.personalWriteAccess,
		});
	}, [view, viewKey, answers]);
}
