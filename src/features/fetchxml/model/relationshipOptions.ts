/**
 * The relationship picker's options (#25): the related entity on the first line, the lookup and
 * schema name on the second, grouped many-to-one, one-to-many, many-to-many.
 */

import type { RelationshipMetadata } from "../api/pptbClient";
import { relationshipKey } from "./manyToMany";

export type RelationshipKind = "1N" | "N1" | "NN";

export interface RelationshipOption {
	/** Unique in the list: an intersect entity's two joins share the schema name (#57) */
	id: string;
	schemaName: string;
	/** The selected value's text: "{Related entity} via {lookup}" */
	displayText: string;
	/** First line: "{Related entity display} ({logical})" */
	primaryText: string;
	/** Second line: "via {Lookup display} ({lookup logical}) · {schema name}" */
	secondaryText: string;
	/** Everything the search matches: entity, lookup (display and logical) and schema name */
	searchText: string;
	referencedEntity: string;
	referencingEntity: string;
	referencedAttribute: string;
	referencingAttribute: string;
	type: RelationshipKind;
	category: RelationshipCategory;
	/** N:N only: the intersect entity the join goes through (#85) */
	intersectEntity?: string;
}

export const RELATIONSHIP_CATEGORIES = ["Many-to-one (N:1)", "One-to-many (1:N)", "Many-to-many (N:N)"] as const;
export type RelationshipCategory = (typeof RELATIONSHIP_CATEGORIES)[number];

const CATEGORY: Record<RelationshipKind, RelationshipCategory> = {
	N1: "Many-to-one (N:1)",
	"1N": "One-to-many (1:N)",
	NN: "Many-to-many (N:N)",
};

/** Display names from the cached metadata; undefined when not known */
export interface RelationshipLabels {
	entity: (logicalName: string) => string | undefined;
	attribute: (entityLogicalName: string, attributeLogicalName: string) => string | undefined;
}

/** "Display (logical)", or the logical name alone when there's no other display name */
const named = (logical: string, display: string | undefined) => (display && display !== logical ? `${display} (${logical})` : logical);

function option(
	rel: RelationshipMetadata,
	type: RelationshipKind,
	related: string,
	via: { logical: string; display?: string },
	labels: RelationshipLabels,
): Omit<RelationshipOption, "referencedEntity" | "referencingEntity" | "referencedAttribute" | "referencingAttribute"> {
	const relatedDisplay = labels.entity(related);
	const secondaryText = via.logical ? `via ${named(via.logical, via.display)} · ${rel.SchemaName}` : rel.SchemaName;
	return {
		id: relationshipKey(rel),
		schemaName: rel.SchemaName,
		displayText: via.logical ? `${relatedDisplay || related} via ${via.display || via.logical}` : relatedDisplay || related,
		primaryText: named(related, relatedDisplay),
		secondaryText,
		searchText: [relatedDisplay, related, via.display, via.logical, rel.SchemaName].filter(Boolean).join(" "),
		type,
		category: CATEGORY[type],
	};
}

/**
 * The options for an entity's relationships. One-to-many: the lookup is on the related entity;
 * many-to-one: on this entity; many-to-many: the join goes through the intersect entity.
 */
export function buildRelationshipOptions(
	entityLogicalName: string,
	data: { oneToMany: RelationshipMetadata[]; manyToOne: RelationshipMetadata[]; manyToMany: RelationshipMetadata[] },
	labels: RelationshipLabels,
): RelationshipOption[] {
	const options: RelationshipOption[] = [];
	const endpoints = (rel: RelationshipMetadata) => ({
		referencedEntity: rel.ReferencedEntity,
		referencingEntity: rel.ReferencingEntity,
		referencedAttribute: rel.ReferencedAttribute,
		referencingAttribute: rel.ReferencingAttribute,
	});

	// N:1 (this entity is the referencing / child one)
	for (const rel of data.manyToOne) {
		const lookup = rel.ReferencingAttribute || "";
		options.push({
			...option(rel, "N1", rel.ReferencedEntity, { logical: lookup, display: lookup ? labels.attribute(rel.ReferencingEntity, lookup) : undefined }, labels),
			...endpoints(rel),
		});
	}

	// 1:N (this entity is the referenced / parent one)
	for (const rel of data.oneToMany) {
		const lookup = rel.ReferencingAttribute || "";
		options.push({
			...option(rel, "1N", rel.ReferencingEntity, { logical: lookup, display: lookup ? labels.attribute(rel.ReferencingEntity, lookup) : undefined }, labels),
			...endpoints(rel),
		});
	}

	// N:N: Entity1 / Entity2, through the intersect entity
	for (const rel of data.manyToMany) {
		const entity1 = rel.Entity1LogicalName || "";
		const entity2 = rel.Entity2LogicalName || "";
		const related = entity1 === entityLogicalName ? entity2 : entity1;
		const intersect = rel.IntersectEntityName || "";
		options.push({
			...option(rel, "NN", related, { logical: intersect, display: intersect ? labels.entity(intersect) : undefined }, labels),
			referencedEntity: entity2, // Store entity2 as "referenced" for compatibility
			referencingEntity: entity1, // Store entity1 as "referencing" for compatibility
			referencedAttribute: rel.Entity2IntersectAttribute || "",
			referencingAttribute: rel.Entity1IntersectAttribute || "",
			intersectEntity: rel.IntersectEntityName,
		});
	}

	return options;
}

/** The options the search matches, any part of either line */
export function filterRelationshipOptions(options: RelationshipOption[], query: string): RelationshipOption[] {
	const q = query.trim().toLowerCase();
	if (!q) return options;
	return options.filter((o) => o.searchText.toLowerCase().includes(q));
}
