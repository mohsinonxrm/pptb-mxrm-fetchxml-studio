/**
 * TypeScript type definitions for FetchXML Builder tree nodes
 * Based on FetchXML XSD schema with all supported attributes
 */

export type NodeId = string;

/**
 * Root fetch node - top level query configuration
 */
export interface FetchNode {
	id: NodeId;
	type: "fetch";
	entity: EntityNode;
	options: {
		aggregate?: boolean;
		distinct?: boolean;
		top?: number;
		count?: number;
		page?: number;
		pagingCookie?: string;
		returnTotalRecordCount?: boolean;
		noLock?: boolean;
		utcOffset?: number;
		latematerialize?: boolean;
		// Query optimization hints
		hints?: {
			forceOrder?: boolean;
			disableRowGoal?: boolean;
			enableOptimizerHotfixes?: boolean;
			loopJoin?: boolean;
			mergeJoin?: boolean;
			hashJoin?: boolean;
			noPerformanceSpool?: boolean;
			enableHistAmendmentForAscKeys?: boolean;
		};
	};
}

/**
 * Entity node - the root or linked table being queried
 */
export interface EntityNode {
	id: NodeId;
	type: "entity";
	name: string; // logical name (e.g., 'account', 'contact')
	enablePrefiltering?: boolean;
	prefilterParameterName?: string;
	allAttributes?: AllAttributesNode;
	attributes: AttributeNode[];
	orders: OrderNode[];
	filters: FilterNode[];
	links: LinkEntityNode[];
}

/**
 * All-Attributes node - select all columns from entity
 */
export interface AllAttributesNode {
	id: NodeId;
	type: "all-attributes";
	enabled: boolean; // toggle to include/exclude all attributes
}

/**
 * Attribute node - individual column selection
 */
export interface AttributeNode {
	id: NodeId;
	type: "attribute";
	name: string; // logical attribute name
	alias?: string;
	groupby?: boolean;
	aggregate?: "sum" | "count" | "countcolumn" | "min" | "max" | "avg" | "rowaggregate";
	usertimezone?: boolean;
	dategrouping?: "day" | "week" | "month" | "quarter" | "year" | "fiscal-period" | "fiscal-year";
}

/**
 * Order node - sorting specification
 */
export interface OrderNode {
	id: NodeId;
	type: "order";
	attribute: string; // logical attribute name
	descending?: boolean;
	entityname?: string; // for link-entity attributes, references the link alias
}

/**
 * Filter node - logical grouping of conditions
 */
export interface FilterNode {
	id: NodeId;
	type: "filter";
	conjunction: "and" | "or"; // filter type
	hint?: "union"; // optimization hint for multi-table filters
	conditions: ConditionNode[];
	subfilters: FilterNode[]; // nested filters
}

/**
 * Condition node - individual filter criteria
 */
export interface ConditionNode {
	id: NodeId;
	type: "condition";
	attribute: string; // logical attribute name
	operator: OperatorType;
	value?: string | number | boolean | string[] | null;
	aggregate?: "sum" | "count" | "countcolumn" | "min" | "max" | "avg"; // for aggregate filtering
	entityname?: string; // for link-entity filters, references the entity/link alias
	valueof?: string; // for cross-table column comparisons (references another attribute)
}

/**
 * Link-Entity node - join to related table
 */
export interface LinkEntityNode {
	id: NodeId;
	type: "link-entity";
	name: string; // related entity logical name
	from: string; // attribute on related entity
	to: string; // attribute on current entity
	linkType: LinkType;
	alias?: string;
	intersect?: boolean; // for N:N relationships
	visible?: boolean; // include in query results
	relationshipType?: "1N" | "N1" | "NN"; // metadata hint for UI
	allAttributes?: AllAttributesNode;
	attributes: AttributeNode[];
	orders: OrderNode[];
	filters: FilterNode[];
	links: LinkEntityNode[]; // nested link-entities (max 15 total depth)
}

/**
 * Link types supported by FetchXML
 */
export type LinkType =
	| "inner" // Inner join (default)
	| "outer" // Left outer join
	| "any" // Exists with ANY match in related
	| "not any" // NOT EXISTS with any match
	| "all" // ALL records match
	| "not all" // NOT ALL records match
	| "exists" // EXISTS (semi-join)
	| "in" // IN (similar to exists)
	| "matchfirstrowusingcrossapply"; // CROSS APPLY for first match

/**
 * FetchXML operators (60+ supported)
 * Applicability depends on attribute type
 */
export type OperatorType =
	// Comparison operators (all types)
	| "eq"
	| "ne"
	| "gt"
	| "ge"
	| "lt"
	| "le"
	// Null checks (all types)
	| "null"
	| "not-null"
	// String operators
	| "like"
	| "not-like"
	| "begins-with"
	| "not-begin-with"
	| "ends-with"
	| "not-end-with"
	// Collection operators
	| "in"
	| "not-in"
	| "between"
	| "not-between"
	| "contain-values"
	| "not-contain-values" // for multi-select picklist
	// Date relative operators
	| "yesterday"
	| "today"
	| "tomorrow"
	| "last-seven-days"
	| "last-week"
	| "this-week"
	| "next-week"
	| "next-seven-days"
	| "last-month"
	| "this-month"
	| "next-month"
	| "last-year"
	| "this-year"
	| "next-year"
	| "on"
	| "on-or-before"
	| "on-or-after"
	// Date range operators (require integer value for 'x')
	| "last-x-hours"
	| "next-x-hours"
	| "last-x-days"
	| "next-x-days"
	| "last-x-weeks"
	| "next-x-weeks"
	| "last-x-months"
	| "next-x-months"
	| "last-x-years"
	| "next-x-years"
	| "olderthan-x-minutes"
	| "olderthan-x-hours"
	| "olderthan-x-days"
	| "olderthan-x-weeks"
	| "olderthan-x-months"
	| "olderthan-x-years"
	// Fiscal period operators
	| "this-fiscal-year"
	| "this-fiscal-period"
	| "next-fiscal-year"
	| "next-fiscal-period"
	| "last-fiscal-year"
	| "last-fiscal-period"
	| "last-x-fiscal-years"
	| "last-x-fiscal-periods"
	| "next-x-fiscal-years"
	| "next-x-fiscal-periods"
	| "in-fiscal-year"
	| "in-fiscal-period"
	| "in-fiscal-period-and-year"
	| "in-or-before-fiscal-period-and-year"
	| "in-or-after-fiscal-period-and-year"
	// User/Business unit operators (for Lookup to systemuser/businessunit)
	| "eq-userid"
	| "ne-userid"
	| "eq-userteams"
	| "eq-useroruserteams"
	| "eq-useroruserhierarchy"
	| "eq-useroruserhierarchyandteams"
	| "eq-businessid"
	| "ne-businessid"
	| "eq-userlanguage"
	// Hierarchy operators (for hierarchical lookups)
	| "under"
	| "eq-or-under"
	| "not-under"
	| "above"
	| "eq-or-above";

/**
 * Value node - for operators that accept multiple values (e.g., 'in')
 * Not a tree node, but used in FetchXML generation
 */
export interface ValueNode {
	value: string | number | boolean;
}
