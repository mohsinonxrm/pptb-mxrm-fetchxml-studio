/**
 * Which entities, attributes and relationships the pickers offer (#57).
 *
 * - "Advanced Find only" (the setting, on by default) keeps what Advanced Find shows:
 *   IsValidForAdvancedFind, "whether the item is shown in Advanced Find" (Microsoft Learn). It
 *   isn't a query restriction: hidden items query fine (verified live: 347 of 421 hidden entities,
 *   1,028 of 1,053 hidden attributes, all 1,593 hidden relationships), which is what lets
 *   FetchXML Studio build queries Advanced Find can't.
 * - Attributes with IsValidForRead = false are never offered: they can't be retrieved, and every
 *   query that uses one fails (0x80041a08, 25 of 25).
 *
 * Only the pickers are scoped. Metadata lookups for attributes already in a query (operators,
 * value inputs, grid formatting) use the full list, so a pasted query keeps working either way.
 */
import { useSyncExternalStore } from "react";
import type { AttributeMetadata, EntityMetadata, RelationshipMetadata } from "../api/pptbClient";

export function isEntityInScope(entity: Pick<EntityMetadata, "IsValidForAdvancedFind">, advancedFindOnly: boolean): boolean {
	return !advancedFindOnly || entity.IsValidForAdvancedFind !== false;
}

export function isAttributeInScope(
	attribute: Pick<AttributeMetadata, "IsValidForAdvancedFind" | "IsValidForRead">,
	advancedFindOnly: boolean,
): boolean {
	if (attribute.IsValidForRead === false) return false;
	return !advancedFindOnly || attribute.IsValidForAdvancedFind?.Value !== false;
}

export function isRelationshipInScope(relationship: Pick<RelationshipMetadata, "IsValidForAdvancedFind">, advancedFindOnly: boolean): boolean {
	return !advancedFindOnly || relationship.IsValidForAdvancedFind !== false;
}

// The current setting, readable outside React (metadata loaders, callbacks); AppShell keeps it in
// sync with the persisted display settings.
let advancedFindOnly = true;
const listeners = new Set<() => void>();

export const metadataScope = {
	get advancedFindOnly(): boolean {
		return advancedFindOnly;
	},
	setAdvancedFindOnly(value: boolean): void {
		if (value === advancedFindOnly) return;
		advancedFindOnly = value;
		listeners.forEach((listener) => listener());
	},
	subscribe(listener: () => void): () => void {
		listeners.add(listener);
		return () => listeners.delete(listener);
	},
};

/** The "Advanced Find only" setting; components re-render (and reload their pickers) when it changes. */
export function useAdvancedFindOnly(): boolean {
	return useSyncExternalStore(metadataScope.subscribe, () => metadataScope.advancedFindOnly);
}
