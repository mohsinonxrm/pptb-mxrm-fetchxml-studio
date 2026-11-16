/**
 * The results grid's edits to the query (#301): column filters (#23), which run the query again;
 * reordering, removing and adding columns, related columns through link-entities; and sorting by a
 * column (#114).
 *
 * AppShell calls the hook where the column filters were, so the effect that runs the query after a
 * filter change keeps its place among its effects.
 */
import { useCallback, useEffect, useMemo, useState } from "react";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";
import type { SyncedQuery } from "./useQueryExecution";
import type { BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import type { LayoutXmlConfig } from "../../features/fetchxml/model/layoutxml";
import type { ConditionNode } from "../../features/fetchxml/model/nodes";
import type { RelatedEntityColumn } from "../../features/fetchxml/ui/RightPane/AddColumnsPanel";
import type { SortChangeData } from "../../features/fetchxml/ui/RightPane/ResultsGrid";
import { NON_RETURNING_LINK_TYPES } from "../../features/fetchxml/model/fetchxmlRules";
import {
	planRemoveColumns,
	planSortByColumn,
	resolveColumn,
} from "../../features/fetchxml/model/queryColumns";
import { relationshipKey } from "../../features/fetchxml/model/manyToMany";
import {
	columnFilterConditions,
	type NewCondition,
} from "../../features/fetchxml/model/columnFilters";
import { toAttributeTypeInfo } from "../../features/fetchxml/model/attributeCapabilities";

export interface ResultsGridEditsInputs {
	/** The query builder (its actions are stable, #65) */
	builder: BuilderContextValue;
	/** Attribute metadata by entity logical name (useRootEntityMetadata), for sorting */
	attributeMetadata: Map<string, Map<string, AttributeMetadata>>;
	/** Runs the query (useQueryExecution); a column filter change runs it again */
	handleExecuteRef: { readonly current: (synced?: SyncedQuery) => Promise<void> };
}

export interface ResultsGridEdits {
	/** The column filter conditions in the query */
	columnFilters: ConditionNode[];
	/** Adds a column filter, then runs the query */
	addColumnFilter: (condition: NewCondition) => void;
	/** Removes column filters, then runs the query */
	removeColumnFilters: (conditionIds: string[]) => void;
	/** Puts the column config's columns in this order */
	reorderColumns: (columns: LayoutXmlConfig["columns"]) => void;
	/** Removes the columns' attributes from the query */
	removeColumns: (columnKeys: string[]) => void;
	/** Adds an attribute to the root entity */
	addColumn: (attributeName: string) => void;
	/** Adds related entities' attributes, through a link-entity per relationship */
	addRelatedColumns: (relatedColumns: RelatedEntityColumn[]) => void;
	/** Sorts the query by a column, or says why it can't */
	sortByColumn: (data: SortChangeData) => void;
}

export function useResultsGridEdits({
	builder,
	attributeMetadata,
	handleExecuteRef,
}: ResultsGridEditsInputs): ResultsGridEdits {
	const { fetchQuery, columnConfig } = builder;
	// stable: builder actions are created once (#65)
	const {
		addColumnFilter: addFilterCondition,
		removeColumnFilter,
		clearColumnFilters,
		setColumnConfig,
		removeNode,
		addAttributeByName,
		addLinkEntityWithConfig,
		setSort,
	} = builder;

	// Column filters from the results grid (#23): they change the query, which then runs again.
	// Running waits for the change to commit, so it runs the new query.
	const columnFilters = useMemo(
		() => columnFilterConditions(fetchQuery, builder.columnFilterId),
		[fetchQuery, builder.columnFilterId],
	);
	const [columnFilterRunRequest, setColumnFilterRunRequest] = useState(0);
	useEffect(() => {
		if (columnFilterRunRequest > 0) void handleExecuteRef.current();
	}, [columnFilterRunRequest, handleExecuteRef]);
	const addColumnFilter = useCallback(
		(condition: NewCondition) => {
			addFilterCondition(condition);
			setColumnFilterRunRequest((n) => n + 1);
		},
		[addFilterCondition],
	);
	const removeColumnFilters = useCallback(
		(conditionIds: string[]) => {
			if (conditionIds.length === 0) return;
			if (conditionIds.length === columnFilters.length) clearColumnFilters();
			else conditionIds.forEach((id) => removeColumnFilter(id));
			setColumnFilterRunRequest((n) => n + 1);
		},
		[columnFilters.length, clearColumnFilters, removeColumnFilter],
	);

	const reorderColumns = useCallback(
		(columns: LayoutXmlConfig["columns"]) => {
			// Set the column config with new order
			if (columnConfig) {
				setColumnConfig({
					...columnConfig,
					columns,
				});
			}
		},
		[columnConfig, setColumnConfig],
	);

	const removeColumns = useCallback(
		(columnKeys: string[]) => {
			// Each column maps to its attribute (root or link-entity, aliased or not); sorts on a
			// removed alias and outer link-entities left empty go too (queryColumns.ts, #114)
			if (!fetchQuery) return;
			const plan = planRemoveColumns(fetchQuery, columnKeys);
			plan.removeNodeIds.forEach((id) => removeNode(id));
		},
		[fetchQuery, removeNode],
	);

	const addColumn = useCallback(
		(attributeName: string) => {
			// Add the attribute to the root entity
			if (fetchQuery?.entity.id) {
				addAttributeByName(fetchQuery.entity.id, attributeName);
			}
		},
		[fetchQuery, addAttributeByName],
	);

	const addRelatedColumns = useCallback(
		(relatedColumns: RelatedEntityColumn[]) => {
			// Add related entity columns - this requires creating link-entities
			if (!fetchQuery?.entity.id) return;

			// Group by relationship to create link-entities
			const byRelationship = new Map<string, RelatedEntityColumn[]>();
			for (const col of relatedColumns) {
				// An intersect entity's two joins share a schema name (#57)
				const key = relationshipKey(col.relationship);
				if (!byRelationship.has(key)) {
					byRelationship.set(key, []);
				}
				byRelationship.get(key)!.push(col);
			}

			// For each relationship, find or create link-entity and add attributes
			for (const [, cols] of byRelationship) {
				const rel = cols[0].relationship;
				const relType = cols[0].relationshipType;

				let fromAttr: string;
				let toAttr: string;
				let relatedEntity: string;
				let linkType: "inner" | "outer";
				let relTypeForBuilder: "N1" | "1N";

				if (relType === "N1") {
					// For N:1 (Many-to-One) lookup relationships:
					// - ReferencingAttribute = FK on current entity (e.g., primarycontactid on account)
					// - ReferencedAttribute = PK on related entity (e.g., contactid on contact)
					// In FetchXML link-entity:
					// - from = attribute on the LINKED entity (the PK)
					// - to = attribute on the PARENT entity (the FK)
					fromAttr = rel.ReferencedAttribute; // PK on linked (contact.contactid)
					toAttr = rel.ReferencingAttribute; // FK on parent (account.primarycontactid)
					relatedEntity = rel.ReferencedEntity; // contact
					linkType = "outer"; // N:1 typically use outer join to not filter parent records
					relTypeForBuilder = "N1";
				} else {
					// For 1:N (One-to-Many) relationships:
					// - ReferencingEntity = the related entity (the "many" side that has the FK)
					// - ReferencedEntity = the current entity (the "one" side)
					// - ReferencingAttribute = FK on the related entity
					// - ReferencedAttribute = PK on the current entity
					// In FetchXML link-entity:
					// - from = FK on the LINKED entity (the FK to parent)
					// - to = PK on the PARENT entity
					fromAttr = rel.ReferencingAttribute; // FK on linked entity
					toAttr = rel.ReferencedAttribute; // PK on parent (e.g., accountid)
					relatedEntity = rel.ReferencingEntity; // The related entity (e.g., contact)
					linkType = "outer"; // 1:N use outer join to include parent records without children
					relTypeForBuilder = "1N";
				}

				// Reuse a link-entity for this relationship, unless its type returns no attributes
				// (exists / in / any / not any / all / not all reject them: 0x80041121)
				const existingLinkEntity = fetchQuery.entity.links.find(
					(le) =>
						le.from === fromAttr &&
						le.to === toAttr &&
						le.name === relatedEntity &&
						!NON_RETURNING_LINK_TYPES.has(le.linkType),
				);

				let linkEntityId: string;

				if (existingLinkEntity) {
					linkEntityId = existingLinkEntity.id;
				} else {
					// Create the link-entity with config - returns the ID immediately
					linkEntityId = addLinkEntityWithConfig(
						fetchQuery.entity.id,
						relatedEntity,
						fromAttr,
						toAttr,
						linkType,
						relTypeForBuilder,
					);
				}

				// Add attributes to the link-entity using the ID
				for (const col of cols) {
					addAttributeByName(linkEntityId, col.attribute.LogicalName);
				}
			}
		},
		[fetchQuery, addLinkEntityWithConfig, addAttributeByName],
	);

	const sortByColumn = useCallback(
		(data: SortChangeData) => {
			// The column maps to the right <order>: alias in aggregate queries, entityname for
			// link-entity columns, nothing for types Dataverse can't sort (queryColumns.ts, #114)
			if (!fetchQuery) return;
			const column = resolveColumn(fetchQuery, data.columnId);
			const metadata = column
				? attributeMetadata?.get(column.container.name)?.get(column.attribute.name)
				: undefined;
			const plan = planSortByColumn(
				fetchQuery,
				data.columnId,
				metadata ? toAttributeTypeInfo(metadata) : undefined,
			);
			if ("reason" in plan) {
				void window.toolboxAPI?.utils?.showNotification?.({
					title: "Can't sort by this column",
					body: plan.reason,
					type: "warning",
				});
				return;
			}
			setSort(plan, data.direction === "descending", data.isMultiSort);
		},
		[fetchQuery, attributeMetadata, setSort],
	);

	return {
		columnFilters,
		addColumnFilter,
		removeColumnFilters,
		reorderColumns,
		removeColumns,
		addColumn,
		addRelatedColumns,
		sortByColumn,
	};
}
