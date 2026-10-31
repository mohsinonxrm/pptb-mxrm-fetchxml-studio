/**
 * TypeScript type definitions for FetchXML Studio tree nodes
 * Based on FetchXML XSD schema with all supported attributes
 */

export type NodeId = string;

/**
 * Attributes the tree has no dedicated field for, kept verbatim so Parse to Tree never loses them
 * (#90): unknown attributes and namespaced attributes (Dataverse ignores both), namespace
 * declarations such as xmlns:generator, and legacy attributes from Advanced Find / saved views
 * (version, output-format, mapping, uiname, uitype, …). Keys are qualified attribute names.
 */
export type ExtraAttributes = Record<string, string>;

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
		/** latematerialize: break the query into parts and reassemble the results */
		latematerialize?: boolean;
		/** aggregatelimit: per-query aggregate record limit (1–50,000) */
		aggregateLimit?: number;
		/** datasource: "retained" (long-term retention, documented) or "bin" (accepted by Dataverse) */
		dataSource?: string;
		/** useraworderby: sort choice columns by integer value instead of label */
		useRawOrderBy?: boolean;
		/**
		 * options: SQL query hints, kept exactly as written (e.g. ["HashJoin", "DisableRowGoal"]).
		 * Documented names: ForceOrder, DisableRowGoal, EnableOptimizerHotfixes, LoopJoin, MergeJoin,
		 * HashJoin, NO_PERFORMANCE_SPOOL, ENABLE_HIST_AMENDMENT_FOR_ASC_KEYS.
		 */
		queryHints?: string[];
	};
	extraAttributes?: ExtraAttributes;
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
	extraAttributes?: ExtraAttributes;
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
	aggregate?: "sum" | "count" | "countcolumn" | "min" | "max" | "avg";
	usertimezone?: boolean;
	dategrouping?: "day" | "week" | "month" | "quarter" | "year" | "fiscal-period" | "fiscal-year";
	/** distinct: with aggregate="countcolumn", count only unique values */
	distinct?: boolean;
	/** rowaggregate: "CountChildren" returns the number of hierarchical child records (casing kept) */
	rowAggregate?: string;
	extraAttributes?: ExtraAttributes;
}

/**
 * Order node - sorting specification
 */
export interface OrderNode {
	id: NodeId;
	type: "order";
	/** Logical attribute name. Used by non-aggregate queries (Dataverse rejects alias there). */
	attribute?: string;
	/** Attribute alias to sort by. Required by aggregate queries (Dataverse rejects attribute there). */
	alias?: string;
	descending?: boolean;
	entityname?: string; // for link-entity attributes, references the link alias
	extraAttributes?: ExtraAttributes;
}

/**
 * Filter node - logical grouping of conditions
 */
export interface FilterNode {
	id: NodeId;
	type: "filter";
	conjunction: "and" | "or"; // filter type
	hint?: "union"; // optimization hint for multi-table filters
	/** isquickfindfields: run as a quick find query */
	isQuickFindFields?: boolean;
	/** overridequickfindrecordlimitenabled: apply the quick find record limit */
	overrideQuickFindRecordLimitEnabled?: boolean;
	/** overridequickfindrecordlimitdisabled: bypass the quick find record limit */
	overrideQuickFindRecordLimitDisabled?: boolean;
	conditions: ConditionNode[];
	subfilters: FilterNode[]; // nested filters
	links: LinkEntityNode[]; // link-entities inside filter (for any/not any/all/not all)
	extraAttributes?: ExtraAttributes;
}

/**
 * Condition node - individual filter criteria
 */
export interface ConditionNode {
	id: NodeId;
	type: "condition";
	attribute: string; // logical attribute name
	operator: OperatorType;
	value?: string | number | boolean | (string | number)[] | null; // arrays for in/not-in operators
	aggregate?: "sum" | "count" | "countcolumn" | "min" | "max" | "avg"; // for aggregate filtering
	entityname?: string; // for link-entity filters, references the entity/link alias
	valueof?: string; // for cross-table column comparisons (references another attribute)
	/**
	 * Attributes on each `<value>` element (uiname, uitype, … as saved views write them), kept
	 * verbatim by the value's text (#181). Changing or removing a value drops its own attributes.
	 */
	valueAttributes?: Record<string, ExtraAttributes>;
	extraAttributes?: ExtraAttributes;
}

/**
 * Link-Entity node - join to related table
 */
export interface LinkEntityNode {
	id: NodeId;
	type: "link-entity";
	name: string; // related entity logical name
	/** Attribute on the related entity. Empty = omitted (Dataverse infers it for system N:N only). */
	from: string;
	/** Attribute on the parent entity. Empty = omitted (Dataverse infers it for system N:N only). */
	to: string;
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
	extraAttributes?: ExtraAttributes;
}

/**
 * Link types supported by FetchXML
 */
export type LinkType =
	| "inner" // Inner join (default)
	| "outer" // Left outer join
	| "any" // Related records exist that match the link's filter
	| "not any" // No related record matches the link's filter
	| "all" // Related records exist, none match the link's filter (use with an inverted filter)
	| "not all" // Same as any, despite the name
	| "exists" // Inner variant using EXISTS; returns no link attributes
	| "in" // Inner variant using IN; returns no link attributes
	| "matchfirstrowusingcrossapply" // Inner variant matching one related row
	| "natural"; // Undocumented, but accepted by Dataverse (verified)

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
