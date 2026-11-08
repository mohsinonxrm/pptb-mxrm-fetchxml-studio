/**
 * The metadata of the query's entities (#301): every attribute of the root entity and its
 * link-entities, the root entity's metadata, and the root entity's relationships for the Add Columns
 * panel. They load again when the root or the set of entities changes, or when "Advanced Find only"
 * changes (#57); responses for entities the query no longer has are dropped (#54). A failed load is
 * logged and leaves that part empty.
 *
 * The caches these loads go through (metadataCache, the metadata load errors) are cleared by AppShell
 * when the connection changes, before this hook's component mounts again (#66, #201).
 */
import { useEffect, useMemo, useState } from "react";
import { loadAllEntityAttributes } from "../../features/fetchxml/api/dataverseMetadata";
import type {
	AttributeMetadata,
	EntityMetadata,
	RelationshipMetadata,
} from "../../features/fetchxml/api/pptbClient";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import { getQueryEntities } from "../../features/fetchxml/model/treeUtils";

export interface RootEntityMetadataInputs {
	/** The builder's query; its root <entity> and link-entities decide what loads */
	fetchQuery: FetchNode | null;
	/** The "Advanced Find only" setting; the relationships follow it, so a change reloads them (#57) */
	advancedFindOnly: boolean;
	/** The root entity's metadata loader (useLazyMetadata) */
	loadEntityMetadata: (logicalName: string) => Promise<EntityMetadata>;
	/** The root entity's relationships loader (useLazyMetadata) */
	loadRelationships: (logicalName: string) => Promise<{
		oneToMany: RelationshipMetadata[];
		manyToOne: RelationshipMetadata[];
	}>;
}

export interface RootEntityMetadata {
	/** Every attribute's metadata, per entity of the query: Map<entity, Map<attribute, metadata>> */
	attributeMetadata: Map<string, Map<string, AttributeMetadata>>;
	/** The root entity's metadata, or null while it loads, after a failure, or with no root entity */
	entityMetadata: EntityMetadata | null;
	/** The root entity's many-to-one relationships */
	lookupRelationships: RelationshipMetadata[];
	/** The root entity's one-to-many relationships */
	oneToManyRelationships: RelationshipMetadata[];
	isLoadingRelationships: boolean;
}

export function useRootEntityMetadata({
	fetchQuery,
	advancedFindOnly,
	loadEntityMetadata,
	loadRelationships,
}: RootEntityMetadataInputs): RootEntityMetadata {
	// Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>>
	const [attributeMetadata, setAttributeMetadata] = useState<
		Map<string, Map<string, AttributeMetadata>>
	>(new Map());
	// State for entity metadata (for layoutxml generation)
	const [loadedEntityMetadata, setLoadedEntityMetadata] = useState<EntityMetadata | null>(null);
	// State for lookup relationships (many-to-one) and one-to-many relationships
	const [lookupRelationships, setLookupRelationships] = useState<RelationshipMetadata[]>([]);
	const [oneToManyRelationships, setOneToManyRelationships] = useState<RelationshipMetadata[]>(
		[],
	);
	const [isLoadingRelationships, setIsLoadingRelationships] = useState(false);

	// Root entity + all link-entity names. The root always comes from the query's <entity>
	// element, never from a position in a sorted list (#54).
	const queryEntities = useMemo(() => getQueryEntities(fetchQuery), [fetchQuery]);
	const rootEntityName = queryEntities.rootEntityName;

	// Only trust root metadata that belongs to the current root entity. While a new root's
	// metadata is loading (or after a stale response), this is null rather than another table's.
	const entityMetadata =
		loadedEntityMetadata && loadedEntityMetadata.LogicalName === rootEntityName
			? loadedEntityMetadata
			: null;

	// Load attribute metadata for all entities in the FetchXML query
	// This includes the root entity and all link-entities
	useEffect(() => {
		const { rootEntityName, entityNames } = queryEntities;
		if (!rootEntityName) {
			setAttributeMetadata(new Map());
			setLoadedEntityMetadata(null);
			return;
		}

		// Ignore responses that arrive after the query's entities have changed.
		let cancelled = false;

		// Metadata of every attribute of the query's entities, for looking up the attributes in the
		// query (grid formatting, columns, export); not filtered by "Advanced Find only" (#57)
		Promise.all(
			entityNames.map(async (entityName) => {
				try {
					const attributes = await loadAllEntityAttributes(entityName);
					const attrMap = new Map<string, AttributeMetadata>();
					attributes.forEach((attr) => {
						attrMap.set(attr.LogicalName, attr);
					});
					return { entityName, attrMap };
				} catch (error) {
					console.error(`Failed to load attributes for ${entityName}:`, error);
					return { entityName, attrMap: new Map<string, AttributeMetadata>() };
				}
			}),
		).then((results) => {
			if (cancelled) return;
			const multiEntityMap = new Map<string, Map<string, AttributeMetadata>>();
			results.forEach(({ entityName, attrMap }) => {
				multiEntityMap.set(entityName, attrMap);
			});
			setAttributeMetadata(multiEntityMap);
		});

		// Load root entity metadata (ObjectTypeCode, PrimaryIdAttribute, EntitySetName, ...)
		loadEntityMetadata(rootEntityName)
			.then((entity) => {
				if (!cancelled) setLoadedEntityMetadata(entity);
			})
			.catch((error) => {
				console.error("Failed to load entity metadata:", error);
				if (!cancelled) setLoadedEntityMetadata(null);
			});

		// Load relationships (many-to-one and one-to-many) for the Add Columns panel, per the
		// "Advanced Find only" setting (#57)
		setIsLoadingRelationships(true);
		loadRelationships(rootEntityName)
			.then((relationships) => {
				if (cancelled) return;
				// Store lookup relationships (many-to-one)
				setLookupRelationships(relationships.manyToOne);
				// Store one-to-many relationships for 1-N column support
				setOneToManyRelationships(relationships.oneToMany);
			})
			.catch((error) => {
				console.error("Failed to load relationships:", error);
				if (cancelled) return;
				setLookupRelationships([]);
				setOneToManyRelationships([]);
			})
			.finally(() => {
				if (!cancelled) setIsLoadingRelationships(false);
			});

		return () => {
			cancelled = true;
		};
		// queryEntities.key changes whenever the root or the set of entities changes.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [queryEntities.key, loadEntityMetadata, loadRelationships, advancedFindOnly]);

	return {
		attributeMetadata,
		entityMetadata,
		lookupRelationships,
		oneToManyRelationships,
		isLoadingRelationships,
	};
}
