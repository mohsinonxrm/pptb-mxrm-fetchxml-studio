/**
 * Hook for Solutions-Only Mode: Solution → Entities
 * Loads all solutions with entities, manages selection, resolves entities
 */

import { useState, useEffect, useCallback } from "react";
import { getAllSolutionsWithEntities, type Solution } from "../../features/fetchxml/api/pptbClient";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { useSolutionEntities } from "./useSolutionEntities";

export function useSolutionFilter() {
	// Solutions
	const [solutions, setSolutions] = useState<Solution[]>([]);
	const [solutionsLoading, setSolutionsLoading] = useState(false);
	const [solutionsError, setSolutionsError] = useState<string | null>(null);
	// The solution request has finished (or came from the cache), for restoring a selection (#25)
	const [solutionsLoaded, setSolutionsLoaded] = useState(false);

	// Selected solutions
	const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

	// Load all solutions on mount (no publisher filter)
	useEffect(() => {
		let mounted = true;
		const abortController = new AbortController();

		async function loadSolutions() {
			// Check cache (use empty array as key for "all solutions")
			const cached = metadataCache.getSolutions([]);
			if (cached) {
				setSolutions(cached);
				setSolutionsLoaded(true);
				return;
			}

			try {
				setSolutionsLoading(true);
				setSolutionsError(null);

				const data = await getAllSolutionsWithEntities();

				if (mounted) {
					setSolutions(data);
					metadataCache.setSolutions([], data);
				}
			} catch (err) {
				if (mounted && !abortController.signal.aborted) {
					setSolutionsError(err instanceof Error ? err.message : String(err));
				}
			} finally {
				if (mounted) {
					setSolutionsLoading(false);
					setSolutionsLoaded(true);
				}
			}
		}

		loadSolutions();

		return () => {
			mounted = false;
			abortController.abort();
		};
	}, []);

	const updateSelectedSolutions = useCallback((ids: string[]) => {
		setSelectedSolutionIds(ids);
	}, []);

	const { entities, entitiesLoading, entitiesError } = useSolutionEntities(selectedSolutionIds);

	return {
		// Solutions
		solutions,
		solutionsLoading,
		solutionsError,
		solutionsLoaded,
		selectedSolutionIds,
		updateSelectedSolutions,

		// Entities
		entities,
		entitiesLoading,
		entitiesError,
	};
}
