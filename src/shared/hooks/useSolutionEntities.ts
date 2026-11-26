/**
 * The entities in the selected solutions, for the Publisher/Solution and Solution scopes. Shared by
 * usePublisherFilter and useSolutionFilter.
 *
 * #358: the solution components used to come back before the startup preload of every entity
 * (useAccessMode), and the entities were then read from a cache that was still empty, so the picker
 * started empty and stayed that way (the empty list was cached too). Now the list waits for that
 * preload (getEntitiesByNames), and a list from before a connection change isn't cached.
 */

import { useEffect, useState } from "react";
import {
	getEntitiesByNames,
	getSolutionComponents,
	type EntityMetadata,
} from "../../features/fetchxml/api/pptbClient";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { debugLog } from "../utils/debug";

export function useSolutionEntities(selectedSolutionIds: string[]) {
	const [entities, setEntities] = useState<EntityMetadata[]>([]);
	const [entitiesLoading, setEntitiesLoading] = useState(false);
	const [entitiesError, setEntitiesError] = useState<string | null>(null);

	useEffect(() => {
		// A new selection starts clean: an earlier request may have stopped mid-way (its cleanup ran
		// before its finally), and its error belongs to the earlier selection
		setEntitiesLoading(false);
		setEntitiesError(null);
		if (!selectedSolutionIds.length) {
			setEntities([]);
			return;
		}

		let mounted = true;

		async function loadEntities() {
			const cached = metadataCache.getFilteredEntities(selectedSolutionIds);
			if (cached) {
				setEntities(cached);
				return;
			}

			// The answer belongs to this connection only (#66)
			const isCurrent = metadataCache.generationGuard();
			try {
				setEntitiesLoading(true);
				setEntitiesError(null);
				debugLog("entityFilters", "Loading entities for solutions:", selectedSolutionIds);

				// Solution components (entities only), cached per solution
				const components = await getSolutionComponents(selectedSolutionIds);
				const logicalNames = Array.from(
					new Set(components.map((c) => c.msdyn_name).filter(Boolean)),
				);
				const data = await getEntitiesByNames(logicalNames);
				debugLog("entityFilters", "Entities in the selected solutions:", {
					components: components.length,
					entities: data.length,
				});

				if (mounted) setEntities(data);
				if (isCurrent()) metadataCache.setFilteredEntities(selectedSolutionIds, data);
			} catch (err) {
				if (mounted) setEntitiesError(err instanceof Error ? err.message : String(err));
			} finally {
				if (mounted) setEntitiesLoading(false);
			}
		}

		loadEntities();

		return () => {
			mounted = false;
		};
	}, [selectedSolutionIds]);

	return { entities, entitiesLoading, entitiesError };
}
