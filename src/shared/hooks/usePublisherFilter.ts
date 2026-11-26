/**
 * Hook for Full Filter Mode: Publisher → Solution → Entities
 * Loads publishers with solutions in one call (optimized), manages selection,
 * resolves entities from selected solutions
 */

import { useState, useEffect, useCallback } from "react";
import {
	getPublishersWithSolutions,
	type Solution,
	type PublisherWithSolutions,
} from "../../features/fetchxml/api/pptbClient";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { useSolutionEntities } from "./useSolutionEntities";

export function usePublisherFilter() {
	// Publishers with solutions (combined)
	const [publishersWithSolutions, setPublishersWithSolutions] = useState<
		PublisherWithSolutions[]
	>([]);
	const [publishersLoading, setPublishersLoading] = useState(false);
	const [publishersError, setPublishersError] = useState<string | null>(null);
	// The publisher request has finished (or came from the cache), for restoring a selection (#25)
	const [publishersLoaded, setPublishersLoaded] = useState(false);

	// Selected publishers
	const [selectedPublisherIds, setSelectedPublisherIds] = useState<string[]>([]);

	// Solutions (filtered by selected publishers)
	const [solutions, setSolutions] = useState<Solution[]>([]);

	// Selected solutions
	const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

	// Load publishers with solutions on mount (optimized single call)
	useEffect(() => {
		let mounted = true;
		const abortController = new AbortController();

		async function loadPublishersWithSolutions() {
			// Check cache
			const cached = metadataCache.getPublishersWithSolutions();
			if (cached) {
				setPublishersWithSolutions(cached);
				setPublishersLoaded(true);
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
					setPublishersLoaded(true);
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
			// Clear selected solutions when all publishers are deselected
			setSelectedSolutionIds([]);
			return;
		}

		// Filter solutions from the combined response (union from all selected publishers)
		const selectedPublisherSet = new Set(selectedPublisherIds);
		const filteredSolutions = publishersWithSolutions
			.filter((pws) => selectedPublisherSet.has(pws.publisher.publisherid))
			.flatMap((pws) => pws.solutions);

		setSolutions(filteredSolutions);

		// Validate selected solutions - keep only those still in filtered list
		setSelectedSolutionIds((currentSelected) => {
			if (currentSelected.length === 0) return currentSelected;

			const availableSolutionIds = new Set(filteredSolutions.map((s) => s.solutionid));
			const validSolutionIds = currentSelected.filter((id) => availableSolutionIds.has(id));

			// Only update if the list changed (prevent unnecessary re-renders)
			if (validSolutionIds.length === currentSelected.length) {
				return currentSelected;
			}

			return validSolutionIds;
		});
	}, [selectedPublisherIds, publishersWithSolutions]);

	const updateSelectedPublishers = useCallback((ids: string[]) => {
		setSelectedPublisherIds(ids);
	}, []);

	const updateSelectedSolutions = useCallback((ids: string[]) => {
		setSelectedSolutionIds(ids);
	}, []);

	const { entities, entitiesLoading, entitiesError } = useSolutionEntities(selectedSolutionIds);

	return {
		// Publishers (extracted from combined response)
		publishers: publishersWithSolutions.map((pws) => pws.publisher),
		publishersLoading,
		publishersError,
		publishersLoaded,
		/** Every publisher with its solutions */
		publishersWithSolutions,
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
