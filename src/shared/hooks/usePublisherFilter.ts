/**
 * Hook for Full Filter Mode: Publisher → Solution → Entities
 * Loads publishers with solutions in one call (optimized), manages selection,
 * resolves entities from selected solutions
 */

import { useState, useEffect, useCallback } from "react";
import {
	getPublishersWithSolutions,
	getSolutionComponents,
	getAdvancedFindEntitiesByNames,
	type Solution,
	type EntityMetadata,
	type PublisherWithSolutions,
} from "../../features/fetchxml/api/pptbClient";
import { metadataCache } from "../../features/fetchxml/state/cache";

export function usePublisherFilter() {
	// Publishers with solutions (combined)
	const [publishersWithSolutions, setPublishersWithSolutions] = useState<PublisherWithSolutions[]>(
		[]
	);
	const [publishersLoading, setPublishersLoading] = useState(false);
	const [publishersError, setPublishersError] = useState<string | null>(null);

	// Selected publishers
	const [selectedPublisherIds, setSelectedPublisherIds] = useState<string[]>([]);

	// Solutions (filtered by selected publishers)
	const [solutions, setSolutions] = useState<Solution[]>([]);

	// Selected solutions
	const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

	// Entities
	const [entities, setEntities] = useState<EntityMetadata[]>([]);
	const [entitiesLoading, setEntitiesLoading] = useState(false);
	const [entitiesError, setEntitiesError] = useState<string | null>(null);

	// Load publishers with solutions on mount (optimized single call)
	useEffect(() => {
		let mounted = true;
		const abortController = new AbortController();

		async function loadPublishersWithSolutions() {
			// Check cache
			const cached = metadataCache.getPublishersWithSolutions();
			if (cached) {
				setPublishersWithSolutions(cached);
				return;
			}

			try {
				setPublishersLoading(true);
				setPublishersError(null);

				const data = await getPublishersWithSolutions();

				if (mounted) {
					setPublishersWithSolutions(data);
					metadataCache.setPublishersWithSolutions(data);
				}
			} catch (err) {
				if (mounted && !abortController.signal.aborted) {
					setPublishersError(err instanceof Error ? err.message : String(err));
				}
			} finally {
				if (mounted) {
					setPublishersLoading(false);
				}
			}
		}

		loadPublishersWithSolutions();

		return () => {
			mounted = false;
			abortController.abort();
		};
	}, []);

	// Filter solutions when publishers are selected
	useEffect(() => {
		if (!selectedPublisherIds.length) {
			setSolutions([]);
			return;
		}

		// Filter solutions from the combined response (union from all selected publishers)
		const selectedPublisherSet = new Set(selectedPublisherIds);
		const filteredSolutions = publishersWithSolutions
			.filter((pws) => selectedPublisherSet.has(pws.publisher.publisherid))
			.flatMap((pws) => pws.solutions);

		setSolutions(filteredSolutions);

		// DO NOT clear solution selections - let them persist
	}, [selectedPublisherIds, publishersWithSolutions]);

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

				console.log("[usePublisherFilter] Loading entities for solutions:", selectedSolutionIds);

				// Get solution components (entities only)
				const components = await getSolutionComponents(selectedSolutionIds);
				console.log("[usePublisherFilter] Solution components received:", {
					solutionIds: selectedSolutionIds,
					componentCount: components.length,
					components: components.slice(0, 10),
				});

				// Extract unique entity logical names
				const logicalNames = Array.from(
					new Set(components.map((c) => c.msdyn_name).filter(Boolean))
				);

				console.log("[usePublisherFilter] Unique entity logical names extracted:", {
					count: logicalNames.length,
					names: logicalNames,
				});

				// Get EntityDefinitions (AF-valid only)
				const data = await getAdvancedFindEntitiesByNames(logicalNames);
				console.log("[usePublisherFilter] Entity metadata received:", {
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

	const updateSelectedPublishers = useCallback((ids: string[]) => {
		setSelectedPublisherIds(ids);
	}, []);

	const updateSelectedSolutions = useCallback((ids: string[]) => {
		setSelectedSolutionIds(ids);
	}, []);

	return {
		// Publishers (extracted from combined response)
		publishers: publishersWithSolutions.map((pws) => pws.publisher),
		publishersLoading,
		publishersError,
		selectedPublisherIds,
		updateSelectedPublishers,

		// Solutions (filtered by selected publishers)
		solutions,
		selectedSolutionIds,
		updateSelectedSolutions,

		// Entities
		entities,
		entitiesLoading,
		entitiesError,
	};
}
