/**
 * Hook for Solutions-Only Mode: Solution → Entities
 * Loads all solutions with entities, manages selection, resolves entities
 */

import { useState, useEffect, useCallback } from "react";
import {
	getAllSolutionsWithEntities,
	getSolutionComponents,
	getAdvancedFindEntitiesByNames,
	type Solution,
	type EntityMetadata,
} from "../../features/fetchxml/api/pptbClient";
import { metadataCache } from "../../features/fetchxml/state/cache";

export function useSolutionFilter() {
	// Solutions
	const [solutions, setSolutions] = useState<Solution[]>([]);
	const [solutionsLoading, setSolutionsLoading] = useState(false);
	const [solutionsError, setSolutionsError] = useState<string | null>(null);

	// Selected solutions
	const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

	// Entities
	const [entities, setEntities] = useState<EntityMetadata[]>([]);
	const [entitiesLoading, setEntitiesLoading] = useState(false);
	const [entitiesError, setEntitiesError] = useState<string | null>(null);

	// Load all solutions on mount (no publisher filter)
	useEffect(() => {
		let mounted = true;
		const abortController = new AbortController();

		async function loadSolutions() {
			// Check cache (use empty array as key for "all solutions")
			const cached = metadataCache.getSolutions([]);
			if (cached) {
				setSolutions(cached);
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
				}
			}
		}

		loadSolutions();

		return () => {
			mounted = false;
			abortController.abort();
		};
	}, []);

	// Load entities when solutions are selected
	useEffect(() => {
		if (!selectedSolutionIds.length) {
			setEntities([]);
			setEntitiesError(null);
			return;
		}

		let mounted = true;
		const abortController = new AbortController();

		async function loadEntities() {
			// Check cache
			const cached = metadataCache.getFilteredEntities(selectedSolutionIds);
			if (cached) {
				setEntities(cached);
				return;
			}

			try {
				setEntitiesLoading(true);
				setEntitiesError(null);

				console.log("[useSolutionFilter] Loading entities for solutions:", selectedSolutionIds);

				// Get solution components (entities only)
				const components = await getSolutionComponents(selectedSolutionIds);
				console.log("[useSolutionFilter] Solution components received:", {
					solutionIds: selectedSolutionIds,
					componentCount: components.length,
					components: components.slice(0, 10),
				});

				// Extract unique entity logical names
				const logicalNames = Array.from(
					new Set(components.map((c) => c.msdyn_name).filter(Boolean))
				);

				console.log("[useSolutionFilter] Unique entity logical names extracted:", {
					count: logicalNames.length,
					names: logicalNames,
				});

				// Get EntityDefinitions (AF-valid only)
				const data = await getAdvancedFindEntitiesByNames(logicalNames);
				console.log("[useSolutionFilter] Entity metadata received:", {
					count: data.length,
					entities: data.map((e) => e.LogicalName),
				});

				if (mounted) {
					setEntities(data);
					metadataCache.setFilteredEntities(selectedSolutionIds, data);
				}
			} catch (err) {
				if (mounted && !abortController.signal.aborted) {
					setEntitiesError(err instanceof Error ? err.message : String(err));
				}
			} finally {
				if (mounted) {
					setEntitiesLoading(false);
				}
			}
		}

		loadEntities();

		return () => {
			mounted = false;
			abortController.abort();
		};
	}, [selectedSolutionIds]);

	const updateSelectedSolutions = useCallback((ids: string[]) => {
		setSelectedSolutionIds(ids);
	}, []);

	return {
		// Solutions
		solutions,
		solutionsLoading,
		solutionsError,
		selectedSolutionIds,
		updateSelectedSolutions,

		// Entities
		entities,
		entitiesLoading,
		entitiesError,
	};
}
