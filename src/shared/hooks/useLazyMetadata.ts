/**
 * React hook for lazy metadata loading with loading states
 */

import { useState, useCallback } from "react";
import type {
	EntityMetadata,
	AttributeMetadata,
	RelationshipMetadata,
} from "../../features/fetchxml/api/pptbClient";
import * as metadataLoader from "../../features/fetchxml/api/dataverseMetadata";
import { metadataCache } from "../../features/fetchxml/state/cache";

interface UseLazyMetadataResult {
	// Loaders
	loadEntities: (advancedFindOnly?: boolean) => Promise<EntityMetadata[]>;
	loadEntityMetadata: (logicalName: string) => Promise<EntityMetadata>;
	loadAttributes: (logicalName: string) => Promise<AttributeMetadata[]>;
	loadRelationships: (logicalName: string) => Promise<{
		oneToMany: RelationshipMetadata[];
		manyToOne: RelationshipMetadata[];
		manyToMany: RelationshipMetadata[];
	}>;

	// State
	isLoading: boolean;
	error: Error | null;

	// Utilities
	clearCache: () => void;
	clearEntityCache: (logicalName: string) => void;
}

export function useLazyMetadata(): UseLazyMetadataResult {
	const [isLoading, setIsLoading] = useState(false);
	const [error, setError] = useState<Error | null>(null);

	const loadEntities = useCallback(async (advancedFindOnly: boolean = true) => {
		setIsLoading(true);
		setError(null);
		try {
			// Check global cache first (preloaded by useAccessMode)
			if (advancedFindOnly) {
				const cached = metadataCache.getAllEntityMetadata();
				if (cached) {
					console.log(
						"[useLazyMetadata] Using preloaded entity metadata:",
						cached.length,
						"entities"
					);
					return cached;
				}
			}

			const entities = await metadataLoader.loadAllEntities(advancedFindOnly);
			return entities;
		} catch (err) {
			const error = err instanceof Error ? err : new Error("Failed to load entities");
			setError(error);
			throw error;
		} finally {
			setIsLoading(false);
		}
	}, []);

	const loadEntityMetadata = useCallback(async (logicalName: string) => {
		setIsLoading(true);
		setError(null);
		try {
			const metadata = await metadataLoader.loadEntityMetadata(logicalName);
			return metadata;
		} catch (err) {
			const error =
				err instanceof Error ? err : new Error(`Failed to load metadata for ${logicalName}`);
			setError(error);
			throw error;
		} finally {
			setIsLoading(false);
		}
	}, []);

	const loadAttributes = useCallback(async (logicalName: string) => {
		setIsLoading(true);
		setError(null);
		try {
			const attributes = await metadataLoader.loadEntityAttributes(logicalName);
			return attributes;
		} catch (err) {
			const error =
				err instanceof Error ? err : new Error(`Failed to load attributes for ${logicalName}`);
			setError(error);
			throw error;
		} finally {
			setIsLoading(false);
		}
	}, []);

	const loadRelationships = useCallback(async (logicalName: string) => {
		setIsLoading(true);
		setError(null);
		try {
			const relationships = await metadataLoader.loadEntityRelationships(logicalName);
			return relationships;
		} catch (err) {
			const error =
				err instanceof Error ? err : new Error(`Failed to load relationships for ${logicalName}`);
			setError(error);
			throw error;
		} finally {
			setIsLoading(false);
		}
	}, []);

	const clearCache = useCallback(() => {
		metadataLoader.clearMetadataCache();
	}, []);

	const clearEntityCache = useCallback((logicalName: string) => {
		metadataLoader.clearEntityCache(logicalName);
	}, []);

	return {
		loadEntities,
		loadEntityMetadata,
		loadAttributes,
		loadRelationships,
		isLoading,
		error,
		clearCache,
		clearEntityCache,
	};
}
