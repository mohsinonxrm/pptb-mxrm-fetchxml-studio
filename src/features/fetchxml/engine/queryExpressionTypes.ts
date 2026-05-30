/**
 * TypeScript interfaces mirroring the Dataverse Web API JSON schema
 * for QueryExpression, as returned by the FetchXmlToQueryExpression function.
 *
 * Reference:
 *   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/fetchxmltoqueryexpression
 *   https://learn.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.query
 */

// ─────────────────────────────────────────────────────────────────────────────
// Core QueryExpression structure
// ─────────────────────────────────────────────────────────────────────────────

export interface QeQueryExpression {
	"@odata.type"?: string;
	EntityName: string;
	ColumnSet: QeColumnSet;
	Criteria?: QeFilterExpression | null;
	LinkEntities?: QeLinkEntity[];
	Orders?: QeOrderExpression[];
	PageInfo?: QePagingInfo | null;
	TopCount?: number | null;
	Distinct?: boolean;
	NoLock?: boolean;
}

export interface QeColumnSet {
	AllColumns: boolean;
	Columns?: string[];
	AttributeExpressions?: QeXrmAttributeExpression[];
}

export interface QeXrmAttributeExpression {
	AttributeName: string;
	Alias?: string | null;
	/** XrmAggregateType — numeric or string name (e.g. "None", "Sum", "Count") */
	AggregateType?: number | string;
	/** XrmDateTimeGrouping — numeric or string name (e.g. "None", "Month", "Year") */
	DateTimeGrouping?: number | string;
	HasGroupBy?: boolean;
}

export interface QeFilterExpression {
	/** LogicalOperator — numeric (0=And, 1=Or) or string name ("And", "Or") */
	FilterOperator: number | string;
	Conditions?: QeConditionExpression[];
	Filters?: QeFilterExpression[];
	AnyAllFilterLinkEntity?: QeAnyAllFilterLinkEntity | null;
}

export interface QeConditionExpression {
	EntityName?: string | null;
	AttributeName: string;
	/** ConditionOperator — numeric or string name (e.g. "Equal", "EqualUserId") */
	Operator: number | string;
	Values?: unknown[];
	/** When true, Values contains attribute names to compare against (not literal values) */
	CompareColumns?: boolean;
}

export interface QeAnyAllFilterLinkEntity {
	LinkFromEntityName: string;
	LinkToEntityName: string;
	LinkFromAttributeName: string;
	LinkToAttributeName: string;
	/** JoinOperator — numeric or string name */
	JoinOperator: number | string;
	EntityAlias?: string | null;
	Columns?: QeColumnSet | null;
	LinkCriteria?: QeFilterExpression | null;
	LinkEntities?: QeLinkEntity[];
	Orders?: QeOrderExpression[];
}

export interface QeLinkEntity {
	LinkFromEntityName: string;
	LinkToEntityName: string;
	LinkFromAttributeName: string;
	LinkToAttributeName: string;
	/** JoinOperator — numeric or string name — default 0/"Inner" */
	JoinOperator?: number | string;
	EntityAlias?: string | null;
	Columns?: QeColumnSet | null;
	LinkCriteria?: QeFilterExpression | null;
	LinkEntities?: QeLinkEntity[];
	Orders?: QeOrderExpression[];
}

export interface QeOrderExpression {
	AttributeName: string;
	/** OrderType — numeric or string name (0/"Ascending", 1/"Descending") */
	OrderType?: number;
	Alias?: string | null;
	EntityName?: string | null;
}

export interface QePagingInfo {
	Count?: number;
	PageNumber?: number;
	ReturnTotalRecordCount?: boolean;
	PagingCookie?: string | null;
}

// ─────────────────────────────────────────────────────────────────────────────
// Enum value lookup maps (numeric value → C# enum member name)
// ─────────────────────────────────────────────────────────────────────────────

export const LOGICAL_OPERATOR: Record<number, string> = {
	0: "And",
	1: "Or",
};

/** JoinOperator enum values */
export const JOIN_OPERATOR: Record<number, string> = {
	0: "Inner",
	1: "LeftOuter",
	2: "Natural",
	3: "MatchFirstRowUsingCrossApply",
	4: "In",
	5: "Exists",
	6: "Any",
	7: "NotAny",
	8: "All",
	9: "NotAll",
};

/** OrderType enum values */
export const ORDER_TYPE: Record<number, string> = {
	0: "Ascending",
	1: "Descending",
};

/** XrmAggregateType enum values */
export const XRM_AGGREGATE_TYPE: Record<number, string> = {
	0: "None",
	1: "Count",
	2: "CountColumn",
	3: "Sum",
	4: "Avg",
	5: "Min",
	6: "Max",
};

/** XrmDateTimeGrouping enum values */
export const XRM_DATETIME_GROUPING: Record<number, string> = {
	0: "None",
	1: "Day",
	2: "Week",
	3: "Month",
	4: "Quarter",
	5: "Year",
	6: "FiscalPeriod",
	7: "FiscalYear",
};

/**
 * ConditionOperator enum values.
 * Reference: https://learn.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.query.conditionoperator
 */
export const CONDITION_OPERATOR: Record<number, string> = {
	0: "Equal",
	1: "NotEqual",
	2: "GreaterThan",
	3: "LessThan",
	4: "GreaterEqual",
	5: "LessEqual",
	6: "Like",
	7: "NotLike",
	8: "In",
	9: "NotIn",
	10: "Between",
	11: "NotBetween",
	12: "Null",
	13: "NotNull",
	14: "Yesterday",
	15: "Today",
	16: "Tomorrow",
	17: "Last7Days",
	18: "Next7Days",
	19: "LastWeek",
	20: "ThisWeek",
	21: "NextWeek",
	22: "LastMonth",
	23: "ThisMonth",
	24: "NextMonth",
	25: "On",
	26: "OnOrBefore",
	27: "OnOrAfter",
	28: "LastYear",
	29: "ThisYear",
	30: "NextYear",
	31: "LastXHours",
	32: "NextXHours",
	33: "LastXDays",
	34: "NextXDays",
	35: "LastXWeeks",
	36: "NextXWeeks",
	37: "LastXMonths",
	38: "NextXMonths",
	39: "OlderThanXMonths",
	40: "OlderThanXYears",
	41: "OlderThanXWeeks",
	42: "OlderThanXDays",
	43: "OlderThanXHours",
	44: "OlderThanXMinutes",
	45: "LastXYears",
	46: "NextXYears",
	47: "EqualUserId",
	48: "NotEqualUserId",
	49: "EqualBusinessId",
	50: "NotEqualBusinessId",
	51: "ChildOf",
	52: "Mask",
	53: "NotMask",
	54: "MasksSelect",
	55: "Contains",
	56: "DoesNotContain",
	57: "EqualUserLanguage",
	58: "NotOn",
	59: "OlderThanXMinutes", // alias — same numeric range as 44 per some API versions
	60: "Next",
	61: "EqualUserOrUserTeams",
	62: "EqualUserTeams",
	63: "EqualUserOrUserHierarchy",
	64: "EqualUserOrUserHierarchyAndTeams",
	65: "Under",
	66: "NotUnder",
	67: "UnderOrEqual",
	68: "Above",
	69: "AboveOrEqual",
	70: "Begins",
	71: "ContainValues",
	72: "DoesNotContainValues",
	73: "EqualRoleBusinessId",
};
