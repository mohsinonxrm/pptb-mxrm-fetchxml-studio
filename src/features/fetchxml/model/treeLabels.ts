/**
 * Labels and counts for the query tree (#24, #25): link-entity labels that name the lookup, the
 * "(no attributes)" hint, and how many items a delete also removes.
 */

import type { EntityNode, FilterNode, LinkEntityNode } from "./nodes";
import { NON_RETURNING_LINK_TYPES } from "./fetchxmlRules";

/**
 * The lookup a link-entity joins on (#25): the parent's `to` attribute for many-to-one, the related
 * entity's `from` attribute for one-to-many. Without a relationship type, a `to` that is the
 * parent's primary key means one-to-many. Empty when neither attribute is set.
 */
export function linkEntityLookup(link: Pick<LinkEntityNode, "from" | "to" | "relationshipType">, parentPrimaryId?: string): string {
	const from = link.from ?? "";
	const to = link.to ?? "";
	if (link.relationshipType === "N1") return to || from;
	if (link.relationshipType === "1N") return from || to;
	if (parentPrimaryId && to && to === parentPrimaryId) return from || to;
	return to || from;
}

/**
 * The tree label of a link-entity (#25): with an alias, "{alias} ({link-type})"; without one,
 * "{name} via {lookup} ({link-type})", or "{name} ({link-type})" when no attribute is set.
 */
export function linkEntityLabel(
	link: Pick<LinkEntityNode, "name" | "alias" | "from" | "to" | "linkType" | "relationshipType">,
	parentPrimaryId?: string,
): string {
	if (link.alias) return `${link.alias} (${link.linkType})`;
	const lookup = linkEntityLookup(link, parentPrimaryId);
	return lookup ? `${link.name} via ${lookup} (${link.linkType})` : `${link.name} (${link.linkType})`;
}

/**
 * Show "(no attributes)" after a link-entity (#24): it returns no attributes although its link type
 * could. Not for exists / in / any / not any / all / not all (they can't return any), nor for an
 * intersect link-entity (it returns none on purpose).
 */
export function showsNoAttributesHint(link: Pick<LinkEntityNode, "linkType" | "attributes" | "allAttributes" | "intersect">): boolean {
	if (NON_RETURNING_LINK_TYPES.has(link.linkType) || link.intersect) return false;
	return link.attributes.length === 0 && !link.allAttributes?.enabled;
}

/** Every item inside a node, at any depth (#24: a delete removes them too) */
export function countDescendants(node: EntityNode | LinkEntityNode | FilterNode): number {
	if (node.type === "filter") {
		return (
			node.conditions.length +
			node.subfilters.reduce((sum, f) => sum + 1 + countDescendants(f), 0) +
			(node.links ?? []).reduce((sum, l) => sum + 1 + countDescendants(l), 0)
		);
	}
	return (
		(node.allAttributes?.enabled ? 1 : 0) +
		node.attributes.length +
		node.orders.length +
		node.filters.reduce((sum, f) => sum + 1 + countDescendants(f), 0) +
		node.links.reduce((sum, l) => sum + 1 + countDescendants(l), 0)
	);
}

/** "This also deletes the item inside it." / "…the N items inside it." (#24) */
export function deleteConfirmationText(descendants: number): string {
	return descendants === 1 ? "This also deletes the item inside it." : `This also deletes the ${descendants} items inside it.`;
}
