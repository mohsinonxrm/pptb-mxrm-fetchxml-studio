/**
 * The view's layout (#301): keeps the column config in step with the query, and builds the
 * LayoutXML that Save view and the export use (#93).
 *
 * AppShell calls the hook where the sync effect was, so the effect keeps its place among its
 * effects.
 */
import { useEffect, useMemo } from "react";
import type { AttributeMetadata, EntityMetadata } from "../../features/fetchxml/api/pptbClient";
import { buildViewLayoutXml } from "../../features/fetchxml/model/layoutxml";
import type { BuilderContextValue } from "../../features/fetchxml/state/builderContext";

export interface ViewLayoutInputs {
	/** The query builder (its actions are stable, #65) */
	builder: BuilderContextValue;
	/** Attribute metadata by entity logical name (useRootEntityMetadata) */
	attributeMetadata: Map<string, Map<string, AttributeMetadata>>;
	/** The root entity's metadata, or null until it loads */
	entityMetadata: EntityMetadata | null;
}

export interface ViewLayout {
	/** The LayoutXML for the column config; empty without a query or the root entity's metadata */
	layoutXml: string;
}

export function useViewLayout({
	builder,
	attributeMetadata,
	entityMetadata,
}: ViewLayoutInputs): ViewLayout {
	// Sync layout with FetchXML when needed (e.g., after adding/removing attributes)
	// Also sync when columnConfig is null but there are attributes (ensures columns button works)
	const { syncLayoutWithFetchXml } = builder; // stable: builder actions are created once (#65)
	useEffect(() => {
		const hasAttributes = (builder.fetchQuery?.entity?.attributes?.length ?? 0) > 0;
		const needsSync = builder.layoutNeedsSync || (hasAttributes && !builder.columnConfig);

		if (needsSync && builder.fetchQuery) {
			// Build attribute type map from metadata for better column widths
			// Use root entity's attribute metadata
			const attributeTypeMap = new Map<string, string>();
			const rootEntityName = builder.fetchQuery.entity?.name;
			if (attributeMetadata && rootEntityName) {
				const rootEntityAttrs = attributeMetadata.get(rootEntityName);
				rootEntityAttrs?.forEach((attr, name) => {
					attributeTypeMap.set(name, attr.AttributeType || "");
				});
			}
			syncLayoutWithFetchXml(attributeTypeMap, entityMetadata?.PrimaryIdAttribute);
		}
	}, [
		builder.layoutNeedsSync,
		builder.fetchQuery,
		builder.columnConfig,
		attributeMetadata,
		entityMetadata,
		syncLayoutWithFetchXml,
	]);

	// Generate LayoutXML from column config for saving
	const layoutXml = useMemo(() => {
		if (!builder.fetchQuery || !entityMetadata) {
			return "";
		}
		// object, row id and jump from the root Entity's metadata (#93)
		return buildViewLayoutXml(builder.columnConfig, builder.fetchQuery, {
			objectTypeCode: entityMetadata.ObjectTypeCode,
			primaryIdAttribute: entityMetadata.PrimaryIdAttribute,
			primaryNameAttribute: entityMetadata.PrimaryNameAttribute,
		});
	}, [builder.columnConfig, builder.fetchQuery, entityMetadata]);

	return { layoutXml };
}
