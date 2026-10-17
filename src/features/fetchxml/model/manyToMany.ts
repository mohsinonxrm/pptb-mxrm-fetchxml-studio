/**
 * Many-to-many joins (#85). A many-to-many relationship joins through its intersect entity with two
 * nested link-entities, as the docs show:
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables#many-to-many-relationships
 *
 *   <entity name="systemuser">
 *     <link-entity name="teammembership" from="systemuserid" to="systemuserid" intersect="true">
 *       <link-entity name="team" from="teamid" to="teamid" alias="team" />
 *     </link-entity>
 *   </entity>
 *
 * Verified live (2026-10-15) on every many-to-many relationship with data in the sandbox, from both
 * sides (system, custom and self-referential): the join returns exactly the intersect entity's pairs.
 */

/** ManyToManyRelationshipMetadata fields the join needs */
export interface ManyToManyRelationship {
	Entity1LogicalName: string;
	Entity2LogicalName: string;
	Entity1IntersectAttribute: string;
	Entity2IntersectAttribute: string;
	IntersectEntityName: string;
}

export interface ManyToManyJoin {
	/** Outer link-entity: the intersect entity (intersect="true") */
	intersect: { name: string; from: string; to: string };
	/** Nested link-entity: the related entity */
	target: { name: string; from: string; to: string };
}

/**
 * The two link-entities that join `parentEntity` to the other side of `relationship`.
 * - intersect: from = the intersect entity's attribute for the parent's side, to = the parent's key;
 * - target: from = the related entity's key, to = the intersect entity's attribute for its side.
 * A self-referential relationship (Entity1 = Entity2) goes from the Entity1 side to the Entity2 side.
 */
export function planManyToManyJoin(
	parentEntity: string,
	parentPrimaryKey: string,
	relationship: ManyToManyRelationship,
	targetPrimaryKey: (targetEntity: string) => string,
): ManyToManyJoin | undefined {
	const fromEntity1 = relationship.Entity1LogicalName === parentEntity;
	if (!fromEntity1 && relationship.Entity2LogicalName !== parentEntity) return undefined;
	const parentAttribute = fromEntity1 ? relationship.Entity1IntersectAttribute : relationship.Entity2IntersectAttribute;
	const target = fromEntity1 ? relationship.Entity2LogicalName : relationship.Entity1LogicalName;
	const targetAttribute = fromEntity1 ? relationship.Entity2IntersectAttribute : relationship.Entity1IntersectAttribute;
	return {
		intersect: { name: relationship.IntersectEntityName, from: parentAttribute, to: parentPrimaryKey },
		target: { name: target, from: targetPrimaryKey(target), to: targetAttribute },
	};
}

/** The relationship fields a join from an intersect entity needs (RelationshipMetadata subset) */
export interface IntersectRelationship extends ManyToManyRelationship {
	SchemaName: string;
	MetadataId: string;
	IsValidForAdvancedFind?: boolean;
}

/** A join from an intersect entity to one side of its many-to-many relationship (#57) */
export interface IntersectEntityJoin {
	SchemaName: string;
	RelationshipType: "ManyToOneRelationship";
	MetadataId: string;
	IsValidForAdvancedFind?: boolean;
	/** The intersect entity */
	ReferencingEntity: string;
	/** Its attribute for this side */
	ReferencingAttribute: string;
	/** This side of the many-to-many relationship */
	ReferencedEntity: string;
	/** This side's primary key */
	ReferencedAttribute: string;
	/** Which side (Entity1 / Entity2) of the many-to-many relationship this join goes to */
	IntersectSide: 1 | 2;
}

/**
 * Joins from an intersect entity (systemuserroles, teammembership, …) to each side of the
 * many-to-many relationship it implements (#57). Dataverse gives intersect entities no many-to-one
 * relationships for these attributes (they're plain Uniqueidentifiers), so without these an
 * intersect entity can't be joined to either side from the relationship picker.
 *
 *   <entity name="systemuserroles">
 *     <link-entity name="role" from="roleid" to="roleid" />
 *   </entity>
 *
 * The attribute for each side is the relationship's Entity1/Entity2IntersectAttribute and the
 * target attribute is that side's primary key. Verified live (2026-10-16) on every intersect entity
 * in the sandbox (141, 302 joins): every join runs, except on the few intersect entities that can't
 * be queried at all (RetrieveMultiple isn't supported), which fail with or without the join.
 * Relationships of other entities, and sides whose primary key is unknown, give no join.
 */
export function planIntersectEntityJoins(
	intersectEntity: string,
	relationships: readonly IntersectRelationship[],
	primaryKeyOf: (entity: string) => string | undefined,
): IntersectEntityJoin[] {
	const joins: IntersectEntityJoin[] = [];
	for (const relationship of relationships) {
		if (relationship.IntersectEntityName !== intersectEntity) continue;
		for (const side of [1, 2] as const) {
			const target = side === 1 ? relationship.Entity1LogicalName : relationship.Entity2LogicalName;
			const attribute = side === 1 ? relationship.Entity1IntersectAttribute : relationship.Entity2IntersectAttribute;
			const primaryKey = primaryKeyOf(target);
			if (!target || !attribute || !primaryKey) continue;
			joins.push({
				SchemaName: relationship.SchemaName,
				RelationshipType: "ManyToOneRelationship",
				MetadataId: `${relationship.MetadataId}:${side}`,
				IsValidForAdvancedFind: relationship.IsValidForAdvancedFind,
				ReferencingEntity: intersectEntity,
				ReferencingAttribute: attribute,
				ReferencedEntity: target,
				ReferencedAttribute: primaryKey,
				IntersectSide: side,
			});
		}
	}
	return joins;
}

/**
 * A key that tells relationships apart in lists: the schema name, plus the side for the two joins an
 * intersect entity's many-to-many relationship gives (they share the schema name, #57).
 */
export function relationshipKey(relationship: { SchemaName: string; IntersectSide?: 1 | 2 }): string {
	return relationship.IntersectSide ? `${relationship.SchemaName}#${relationship.IntersectSide}` : relationship.SchemaName;
}
