/**
 * The entity's views for Open view (#278): system and personal views, and who the user is (to tell
 * their own personal views from ones shared with them, #142). Loaded each time Open view opens, so a
 * view saved a moment ago is listed (#359).
 */
import { useCallback, useEffect, useState } from "react";
import { getAllViews, whoAmI, type SavedView } from "../../api/pptbClient";
import { debugLog } from "../../../../shared/utils/debug";

export interface EntityViews {
	systemViews: SavedView[];
	personalViews: SavedView[];
	currentUserId: string | undefined;
	loading: boolean;
	error: string | null;
	/** Load them again (Try again after an error) */
	reload: () => void;
}

export function useEntityViews(entityLogicalName: string | undefined): EntityViews {
	const [systemViews, setSystemViews] = useState<SavedView[]>([]);
	const [personalViews, setPersonalViews] = useState<SavedView[]>([]);
	const [currentUserId, setCurrentUserId] = useState<string | undefined>(undefined);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);
	const [attempt, setAttempt] = useState(0);

	useEffect(() => {
		if (!entityLogicalName) return;
		let current = true;
		setLoading(true);
		setError(null);
		void Promise.all([getAllViews(entityLogicalName), whoAmI()])
			.then(([views, user]) => {
				if (!current) return;
				setSystemViews(views.systemViews);
				setPersonalViews(views.personalViews);
				setCurrentUserId(user?.UserId);
				debugLog(
					"viewAPI",
					`Open view: ${views.systemViews.length} system and ${views.personalViews.length} personal views of ${entityLogicalName}`,
				);
			})
			.catch((err: unknown) => {
				if (!current) return;
				console.error("Open view: couldn't load the views:", err);
				setError(err instanceof Error ? err.message : String(err));
			})
			.finally(() => {
				if (current) setLoading(false);
			});
		return () => {
			current = false;
		};
	}, [entityLogicalName, attempt]);

	const reload = useCallback(() => setAttempt((n) => n + 1), []);
	return { systemViews, personalViews, currentUserId, loading, error, reload };
}
