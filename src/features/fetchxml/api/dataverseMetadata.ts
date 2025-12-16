/**
 * Lazy metadata loading with caching and promise memoization
 * Prevents duplicate API calls for the same entity
 */

import * as pptbClient from "../api/pptbClient";
import { metadataCache } from "../state/cache";
import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from "../api/pptbClient";

/**
 * Load all entities (with caching and de-duplication)
 */
export async function loadAllEntities(advancedFindOnly: boolean = true): Promise<EntityMetadata[]> {
	// Check cache first
	const cached = metadataCache.getAllEntities();
	if (cached) {
		return cached;
	}

	// Check if request is already in-flight
	const inFlight = metadataCache.getAllEntitiesPromise();
	if (inFlight) {
		return inFlight;
	}

	// Make new request
	const promise = pptbClient.getAllEntities(advancedFindOnly);
	metadataCache.setAllEntitiesPromise(promise);

	try {
		const entities = await promise;
		metadataCache.setAllEntities(entities);
		metadataCache.clearAllEntitiesPromise();
		return entities;
	} catch (error) {
		console.error("loadAllEntities: API error:", error);
		metadataCache.clearAllEntitiesPromise();
		throw error;
	}
}

/**
 * Load entity metadata (with caching)
 */
export async function loadEntityMetadata(logicalName: string): Promise<EntityMetadata> {
	// Check cache first
	const cached = metadataCache.getEntityMetadata(logicalName);
	if (cached) {
		return cached;
	}

	// Make request
	const metadata = await pptbClient.getEntityMetadata(logicalName);
	metadataCache.setEntityMetadata(logicalName, metadata);
	return metadata;
}

/**
 * Load entity attributes (with caching and de-duplication)
 * @param logicalName - Entity logical name
 * @param advancedFindOnly - Filter to only attributes valid for advanced find (default: true)
 */
export async function loadEntityAttributes(
	logicalName: string,
	advancedFindOnly: boolean = true
): Promise<AttributeMetadata[]> {
	// Check cache first
	const cacheKey = `${logicalName}_${advancedFindOnly}`;
	const cached = metadataCache.getEntityAttributes(cacheKey);
	if (cached) {
		return cached;
	}

	// Check if request is already in-flight
	const inFlight = metadataCache.getAttributesPromise(cacheKey);
	if (inFlight) {
		return inFlight;
	}

	// Make new request
	const promise = pptbClient.getEntityAttributes(logicalName, advancedFindOnly);
	metadataCache.setAttributesPromise(cacheKey, promise);

	try {
		const attributes = await promise;
		metadataCache.setEntityAttributes(cacheKey, attributes);
		metadataCache.clearAttributesPromise(cacheKey);
		return attributes;
	} catch (error) {
		metadataCache.clearAttributesPromise(cacheKey);
		throw error;
	}
}

/**
 * Load a single attribute with full metadata including OptionSet expansion
 * Used for loading picklist/boolean options for value pickers
 * @param entityLogicalName - Entity logical name
 * @param attributeLogicalName - Attribute logical name
 */
export async function loadAttributeWithOptionSet(
	entityLogicalName: string,
	attributeLogicalName: string
): Promise<AttributeMetadata> {
	// Check cache first
	const cacheKey = `${entityLogicalName}_${attributeLogicalName}_optionset`;
	const cached = metadataCache.getEntityAttributes(cacheKey);
	if (cached && cached.length > 0) {
		return cached[0];
	}

	// Make request with OptionSet expansion
	const attribute = await pptbClient.getAttributeWithOptionSet(
		entityLogicalName,
		attributeLogicalName
	);

	// Cache the result
	metadataCache.setEntityAttributes(cacheKey, [attribute]);
	return attribute;
}

/**
 * Load a single attribute with detailed type-specific metadata (MinValue, MaxValue, Precision, Format)
 * Used for numeric and datetime attributes that need validation constraints
 * @param entityLogicalName - Entity logical name
 * @param attributeLogicalName - Attribute logical name
 * @param attributeType - Attribute type (Integer, BigInt, Decimal, Double, DateTime)
 */
export async function loadAttributeDetailedMetadata(
	entityLogicalName: string,
	attributeLogicalName: string,
	attributeType: string
): Promise<AttributeMetadata> {
	// Check cache first
	const cacheKey = `${entityLogicalName}_${attributeLogicalName}_detailed`;
	const cached = metadataCache.getEntityAttributes(cacheKey);
	if (cached && cached.length > 0) {
		return cached[0];
	}

	// Make request with type casting to get detailed properties
	const attribute = await pptbClient.getAttributeDetailedMetadata(
		entityLogicalName,
		attributeLogicalName,
		attributeType
	);

	// Cache the result
	metadataCache.setEntityAttributes(cacheKey, [attribute]);
	return attribute;
}

/**
 * Load entity relationships (with caching)
 * @param logicalName - Entity logical name
 * @param advancedFindOnly - Filter to only relationships valid for advanced find (default: true)
 */
export async function loadEntityRelationships(
	logicalName: string,
	advancedFindOnly: boolean = true
): Promise<{
	oneToMany: RelationshipMetadata[];
	manyToOne: RelationshipMetadata[];
	manyToMany: RelationshipMetadata[];
}> {
	// Check cache first
	const cacheKey = `${logicalName}_${advancedFindOnly}`;
	const cached = metadataCache.getEntityRelationships(cacheKey);
	if (cached) {
		return cached;
	}

	// Make request
	const relationships = await pptbClient.getEntityRelationships(logicalName, advancedFindOnly);
	metadataCache.setEntityRelationships(cacheKey, relationships);
	return relationships;
}

/**
 * Clear all metadata cache
 */
export function clearMetadataCache(): void {
	metadataCache.clear();
}

/**
 * Clear cache for specific entity
 */
export function clearEntityCache(logicalName: string): void {
	metadataCache.clearEntity(logicalName);
}
