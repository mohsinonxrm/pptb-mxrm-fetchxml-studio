/**
 * Metadata cache with promise memoization
 * Prevents duplicate API calls for the same entity
 */

import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from "../api/pptbClient";

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
	}

	/**
	 * Clear cache for specific entity
	 */
	clearEntity(logicalName: string): void {
		this.entities.delete(logicalName);
	}
}

// Singleton instance
export const metadataCache = new MetadataCache();
