/**
 * Hook to determine user's access mode based on privileges
 * Checks WhoAmI and privilege checks on mount
 */

import { useState, useEffect } from "react";
import { getAccessSummary, type AccessSummary } from "../../features/fetchxml/api/pptbClient";

export function useAccessMode() {
	const [accessSummary, setAccessSummary] = useState<AccessSummary | null>(null);
	const [loading, setLoading] = useState(true);
	const [error, setError] = useState<string | null>(null);

	useEffect(() => {
		let mounted = true;

		async function checkAccess() {
			try {
				setLoading(true);
				setError(null);

				const summary = await getAccessSummary();

				if (mounted) {
					setAccessSummary(summary);
				}
			} catch (err) {
				if (mounted) {
					setError(err instanceof Error ? err.message : String(err));
				}
			} finally {
				if (mounted) {
					setLoading(false);
				}
			}
		}

		checkAccess();

		return () => {
			mounted = false;
		};
	}, []);

	return {
		accessSummary,
		loading,
		error,
		// Convenience flags
		fullFilterMode: accessSummary?.fullFilterMode ?? false,
		solutionsOnlyMode: accessSummary?.solutionsOnlyMode ?? false,
		publishersOnlyMode: accessSummary?.publishersOnlyMode ?? false,
		metadataOnlyMode: accessSummary?.metadataOnlyMode ?? false,
		noAccessMode: accessSummary?.noAccessMode ?? false,
	};
}
