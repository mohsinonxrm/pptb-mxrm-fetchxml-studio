/**
 * Metadata cache with promise memoization
 * Prevents duplicate API calls for the same entity
 */

import type {
	EntityMetadata,
	AttributeMetadata,
	RelationshipMetadata,
	Publisher,
	Solution,
	PublisherWithSolutions,
} from "../api/pptbClient";

interface CachedEntityData {
	metadata?: EntityMetadata;
	attributes?: AttributeMetadata[];
	relationships?: {
		oneToMany: RelationshipMetadata[];
		manyToOne: RelationshipMetadata[];
		manyToMany: RelationshipMetadata[];
	};
}

interface CacheEntry<T> {
	data?: T;
	promise?: Promise<T>;
	timestamp: number;
}

class MetadataCache {
	private entities: Map<string, CacheEntry<CachedEntityData>> = new Map();
	private allEntitiesCache: CacheEntry<EntityMetadata[]> | null = null;
	private allEntityMetadataCache: CacheEntry<EntityMetadata[]> | null = null; // Global AF-valid entities
	private publishersCache: CacheEntry<Publisher[]> | null = null;
	private publishersWithSolutionsCache: CacheEntry<PublisherWithSolutions[]> | null = null;
	private solutionsCache: Map<string, CacheEntry<Solution[]>> = new Map(); // keyed by sorted publisher IDs
	private filteredEntitiesCache: Map<string, CacheEntry<EntityMetadata[]>> = new Map(); // keyed by sorted solution IDs
	private solutionComponentsCache: Map<string, CacheEntry<string[]>> = new Map(); // Entity names per solution ID

	// Cache TTL: 5 minutes
	private readonly TTL = 5 * 60 * 1000;

	/**
	 * Check if cache entry is stale
	 */
	private isStale(timestamp: number): boolean {
		return Date.now() - timestamp > this.TTL;
	}

	/**
	 * Get or create entity cache entry
	 */
	private getEntityEntry(logicalName: string): CacheEntry<CachedEntityData> {
		if (!this.entities.has(logicalName)) {
			this.entities.set(logicalName, {
				data: {},
				timestamp: Date.now(),
			});
		}
		return this.entities.get(logicalName)!;
	}

	/**
	 * Get all entities from cache
	 */
	getAllEntities(): EntityMetadata[] | undefined {
		if (!this.allEntitiesCache || this.isStale(this.allEntitiesCache.timestamp)) {
			return undefined;
		}
		return this.allEntitiesCache.data;
	}

	/**
	 * Set all entities in cache
	 */
	setAllEntities(entities: EntityMetadata[]): void {
		this.allEntitiesCache = {
			data: entities,
			timestamp: Date.now(),
		};
	}

	/**
	 * Get in-flight promise for all entities
	 */
	getAllEntitiesPromise(): Promise<EntityMetadata[]> | undefined {
		return this.allEntitiesCache?.promise;
	}

	/**
	 * Set in-flight promise for all entities
	 */
	setAllEntitiesPromise(promise: Promise<EntityMetadata[]>): void {
		if (!this.allEntitiesCache) {
			this.allEntitiesCache = { timestamp: Date.now() };
		}
		this.allEntitiesCache.promise = promise;
	}

	/**
	 * Clear in-flight promise for all entities
	 */
	clearAllEntitiesPromise(): void {
		if (this.allEntitiesCache) {
			this.allEntitiesCache.promise = undefined;
		}
	}

	/**
	 * Get global entity metadata cache (all AF-valid entities)
	 */
	getAllEntityMetadata(): EntityMetadata[] | undefined {
		if (!this.allEntityMetadataCache || this.isStale(this.allEntityMetadataCache.timestamp)) {
			return undefined;
		}
		return this.allEntityMetadataCache.data;
	}

	/**
	 * Set global entity metadata cache
	 */
	setAllEntityMetadata(entities: EntityMetadata[]): void {
		this.allEntityMetadataCache = {
			data: entities,
			timestamp: Date.now(),
		};
		console.log("[Cache] Global entity metadata cached:", entities.length, "entities");
	}

	/**
	 * Get in-flight promise for global entity metadata
	 */
	getAllEntityMetadataPromise(): Promise<EntityMetadata[]> | undefined {
		return this.allEntityMetadataCache?.promise;
	}

	/**
	 * Set in-flight promise for global entity metadata
	 */
	setAllEntityMetadataPromise(promise: Promise<EntityMetadata[]>): void {
		if (!this.allEntityMetadataCache) {
			this.allEntityMetadataCache = { timestamp: Date.now() };
		}
		this.allEntityMetadataCache.promise = promise;
	}

	/**
	 * Clear in-flight promise for global entity metadata
	 */
	clearAllEntityMetadataPromise(): void {
		if (this.allEntityMetadataCache) {
			this.allEntityMetadataCache.promise = undefined;
		}
	}

	/**
	 * Get entity metadata from cache
	 */
	getEntityMetadata(logicalName: string): EntityMetadata | undefined {
		const entry = this.entities.get(logicalName);
		if (!entry || this.isStale(entry.timestamp)) {
			return undefined;
		}
		return entry.data?.metadata;
	}

	/**
	 * Set entity metadata in cache
	 */
	setEntityMetadata(logicalName: string, metadata: EntityMetadata): void {
		const entry = this.getEntityEntry(logicalName);
		if (!entry.data) entry.data = {};
		entry.data.metadata = metadata;
		entry.timestamp = Date.now();
	}

	/**
	 * Get entity attributes from cache
	 */
	getEntityAttributes(logicalName: string): AttributeMetadata[] | undefined {
		const entry = this.entities.get(logicalName);
		if (!entry || this.isStale(entry.timestamp)) {
			return undefined;
		}
		return entry.data?.attributes;
	}

	/**
	 * Set entity attributes in cache
	 */
	setEntityAttributes(logicalName: string, attributes: AttributeMetadata[]): void {
		const entry = this.getEntityEntry(logicalName);
		if (!entry.data) entry.data = {};
		entry.data.attributes = attributes;
		entry.timestamp = Date.now();
	}

	/**
	 * Get entity relationships from cache
	 */
	getEntityRelationships(logicalName: string):
		| {
				oneToMany: RelationshipMetadata[];
				manyToOne: RelationshipMetadata[];
				manyToMany: RelationshipMetadata[];
		  }
		| undefined {
		const entry = this.entities.get(logicalName);
		if (!entry || this.isStale(entry.timestamp)) {
			return undefined;
		}
		return entry.data?.relationships;
	}

	/**
	 * Set entity relationships in cache
	 */
	setEntityRelationships(
		logicalName: string,
		relationships: {
			oneToMany: RelationshipMetadata[];
			manyToOne: RelationshipMetadata[];
			manyToMany: RelationshipMetadata[];
		}
	): void {
		const entry = this.getEntityEntry(logicalName);
		if (!entry.data) entry.data = {};
		entry.data.relationships = relationships;
		entry.timestamp = Date.now();
	}

	/**
	 * Check if there's an in-flight request for entity attributes
	 */
	getAttributesPromise(logicalName: string): Promise<AttributeMetadata[]> | undefined {
		const entry = this.entities.get(logicalName);
		return entry?.promise as Promise<AttributeMetadata[]> | undefined;
	}

	/**
	 * Set in-flight promise for entity attributes
	 */
	setAttributesPromise(logicalName: string, promise: Promise<AttributeMetadata[]>): void {
		const entry = this.getEntityEntry(logicalName);
		entry.promise = promise as Promise<CachedEntityData>;
	}

	/**
	 * Clear in-flight promise for entity attributes
	 */
	clearAttributesPromise(logicalName: string): void {
		const entry = this.entities.get(logicalName);
		if (entry) {
			entry.promise = undefined;
		}
	}

	/**
	 * Clear all cache
	 */
	clear(): void {
		this.entities.clear();
		this.allEntitiesCache = null;
		this.publishersCache = null;
		this.solutionsCache.clear();
		this.filteredEntitiesCache.clear();
	}

	/**
	 * Clear cache for specific entity
	 */
	clearEntity(logicalName: string): void {
		this.entities.delete(logicalName);
	}

	/**
	 * Get publishers from cache
	 */
	getPublishers(): Publisher[] | undefined {
		if (!this.publishersCache || this.isStale(this.publishersCache.timestamp)) {
			return undefined;
		}
		return this.publishersCache.data;
	}

	/**
	 * Set publishers in cache
	 */
	setPublishers(publishers: Publisher[]): void {
		this.publishersCache = {
			data: publishers,
			timestamp: Date.now(),
		};
	}

	/**
	 * Get publishers with solutions from cache (combined response)
	 */
	getPublishersWithSolutions(): PublisherWithSolutions[] | undefined {
		if (
			!this.publishersWithSolutionsCache ||
			this.isStale(this.publishersWithSolutionsCache.timestamp)
		) {
			return undefined;
		}
		return this.publishersWithSolutionsCache.data;
	}

	/**
	 * Set publishers with solutions in cache (combined response)
	 */
	setPublishersWithSolutions(publishersWithSolutions: PublisherWithSolutions[]): void {
		this.publishersWithSolutionsCache = {
			data: publishersWithSolutions,
			timestamp: Date.now(),
		};
	}

	/**
	 * Get solutions from cache by publisher IDs
	 */
	getSolutions(publisherIds: string[]): Solution[] | undefined {
		const key = publisherIds.sort().join(",");
		const entry = this.solutionsCache.get(key);
		if (!entry || this.isStale(entry.timestamp)) {
			return undefined;
		}
		return entry.data;
	}

	/**
	 * Set solutions in cache by publisher IDs
	 */
	setSolutions(publisherIds: string[], solutions: Solution[]): void {
		const key = publisherIds.sort().join(",");
		this.solutionsCache.set(key, {
			data: solutions,
			timestamp: Date.now(),
		});
	}

	/**
	 * Get filtered entities from cache by solution IDs
	 */
	getFilteredEntities(solutionIds: string[]): EntityMetadata[] | undefined {
		const key = solutionIds.sort().join(",");
		const entry = this.filteredEntitiesCache.get(key);
		if (!entry || this.isStale(entry.timestamp)) {
			return undefined;
		}
		return entry.data;
	}

	/**
	 * Set filtered entities in cache by solution IDs
	 */
	setFilteredEntities(solutionIds: string[], entities: EntityMetadata[]): void {
		const key = solutionIds.sort().join(",");
		this.filteredEntitiesCache.set(key, {
			data: entities,
			timestamp: Date.now(),
		});
	}

	/**
	 * Get solution components (entity names) from cache by solution ID
	 */
	getSolutionComponents(solutionId: string): string[] | undefined {
		const entry = this.solutionComponentsCache.get(solutionId);
		if (!entry || this.isStale(entry.timestamp)) {
			return undefined;
		}
		return entry.data;
	}

	/**
	 * Set solution components (entity names) in cache by solution ID
	 */
	setSolutionComponents(solutionId: string, entityNames: string[]): void {
		this.solutionComponentsCache.set(solutionId, {
			data: entityNames,
			timestamp: Date.now(),
		});
		console.log("[Cache] Solution components cached:", {
			solutionId,
			entityCount: entityNames.length,
		});
	}
}

// Singleton instance
export const metadataCache = new MetadataCache();
