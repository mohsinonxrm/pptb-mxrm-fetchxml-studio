/**
 * Lazy metadata loading with caching and promise memoization
 * Prevents duplicate API calls for the same entity
 */

import * as pptbClient from "../api/pptbClient";
import { metadataCache } from "../state/cache";
import { isAttributeInScope, isRelationshipInScope, metadataScope } from "../state/metadataScope";
import { planIntersectEntityJoins, type IntersectRelationship } from "../model/manyToMany";
import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from "../api/pptbClient";

/**
 * Load all entities (cached and de-duplicated; the same list the Publisher/Solution scopes use).
 * Not filtered: the entity picker applies "Advanced Find only" locally, so toggling it is instant.
 */
export async function loadAllEntities(): Promise<EntityMetadata[]> {
	return pptbClient.getAllEntitiesForScopes();
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
 * Every attribute of an entity (cached and de-duplicated), for metadata lookups of attributes
 * already in a query: operators, value inputs, capabilities, grid formatting (#57).
 */
export async function loadAllEntityAttributes(logicalName: string): Promise<AttributeMetadata[]> {
	const cacheKey = `${logicalName}__all`;
	const cached = metadataCache.getEntityAttributes(cacheKey);
	if (cached) {
		return cached;
	}

	const inFlight = metadataCache.getAttributesPromise(cacheKey);
	if (inFlight) {
		return inFlight;
	}

	const promise = pptbClient.getEntityAttributes(logicalName, false);
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
 * The attributes the pickers offer (#57): readable ones, and only those Advanced Find shows while
 * "Advanced Find only" is on.
 * @param advancedFindOnly - Defaults to the current setting
 */
export async function loadEntityAttributes(
	logicalName: string,
	advancedFindOnly: boolean = metadataScope.advancedFindOnly,
): Promise<AttributeMetadata[]> {
	const attributes = await loadAllEntityAttributes(logicalName);
	return attributes.filter((attribute) => isAttributeInScope(attribute, advancedFindOnly));
}

/**
 * Load a single attribute with full metadata including OptionSet expansion
 * Used for loading picklist/boolean options for value pickers
 * @param entityLogicalName - Entity logical name
 * @param attributeLogicalName - Attribute logical name
 */
export async function loadAttributeWithOptionSet(
	entityLogicalName: string,
	attributeLogicalName: string,
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
		attributeLogicalName,
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
	attributeType: string,
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
		attributeType,
	);

	// Cache the result
	metadataCache.setEntityAttributes(cacheKey, [attribute]);
	return attribute;
}

type EntityRelationships = {
	oneToMany: RelationshipMetadata[];
	manyToOne: RelationshipMetadata[];
	manyToMany: RelationshipMetadata[];
};

/**
 * An intersect entity's many-to-many relationship becomes a many-to-one join to each side
 * (planIntersectEntityJoins, #57): joining "through" the intersect entity from itself makes no
 * sense, and Dataverse gives it no many-to-one relationships for those attributes.
 */
async function withIntersectEntityJoins(logicalName: string, relationships: EntityRelationships): Promise<EntityRelationships> {
	const own = relationships.manyToMany.filter((r) => r.IntersectEntityName === logicalName);
	if (!own.length) return relationships;
	const targets = [...new Set(own.flatMap((r) => [r.Entity1LogicalName, r.Entity2LogicalName]).filter((e): e is string => !!e))];
	const keys = new Map<string, string>();
	await Promise.all(
		targets.map(async (target) => {
			try {
				keys.set(target, (await loadEntityMetadata(target)).PrimaryIdAttribute);
			} catch {
				// No primary key, no join to that side
			}
		}),
	);
	const joins = planIntersectEntityJoins(logicalName, own as IntersectRelationship[], (entity) => keys.get(entity));
	return {
		oneToMany: relationships.oneToMany,
		manyToOne: [...relationships.manyToOne, ...joins],
		manyToMany: relationships.manyToMany.filter((r) => r.IntersectEntityName !== logicalName),
	};
}

/**
 * The relationships the pickers offer (#57): all of them, or only those Advanced Find shows while
 * "Advanced Find only" is on. Every relationship is fetched once and cached; an intersect entity
 * also gets a join to each side of its many-to-many relationship.
 * @param advancedFindOnly - Defaults to the current setting
 */
export async function loadEntityRelationships(
	logicalName: string,
	advancedFindOnly: boolean = metadataScope.advancedFindOnly,
): Promise<EntityRelationships> {
	const cacheKey = `${logicalName}__all`;
	let relationships = metadataCache.getEntityRelationships(cacheKey);
	if (!relationships) {
		relationships = await withIntersectEntityJoins(logicalName, await pptbClient.getEntityRelationships(logicalName, false));
		metadataCache.setEntityRelationships(cacheKey, relationships);
	}
	const inScope = (r: RelationshipMetadata) => isRelationshipInScope(r, advancedFindOnly);
	return {
		oneToMany: relationships.oneToMany.filter(inScope),
		manyToOne: relationships.manyToOne.filter(inScope),
		manyToMany: relationships.manyToMany.filter(inScope),
	};
}

/**
 * Whether the entity has a hierarchical relationship (cached)
 */
export async function loadIsHierarchical(logicalName: string): Promise<boolean> {
	const cached = metadataCache.getEntityHierarchical(logicalName);
	if (cached !== undefined) return cached;
	const isHierarchical = await pptbClient.getIsHierarchical(logicalName);
	metadataCache.setEntityHierarchical(logicalName, isHierarchical);
	return isHierarchical;
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
