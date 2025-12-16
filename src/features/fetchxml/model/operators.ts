/**
 * FetchXML operator definitions and filtering based on attribute type
 */

import type { OperatorType } from "./nodes";

export interface OperatorDefinition {
	value: OperatorType;
	label: string;
	requiresValue: boolean; // Does it need any value?
	requiresTwoValues?: boolean; // Does it need exactly TWO values? (e.g., between, in-fiscal-period-and-year)
	requiresMultipleValues?: boolean; // Does it accept multiple values? (e.g., in, contain-values)
	valueType?: "string" | "number" | "date" | "guid" | "boolean"; // What type of value(s)?
	applicableTypes: string[]; // AttributeType values
}

/**
 * All available FetchXML operators with their metadata
 */
export const ALL_OPERATORS: OperatorDefinition[] = [
	// Comparison operators (most types)
	{
		value: "eq",
		label: "Equals",
		requiresValue: true,
		applicableTypes: [
			"String",
			"Integer",
			"BigInt",
			"Decimal",
			"Double",
			"Money",
			"DateTime",
			"Boolean",
			"Picklist",
			"State",
			"Status",
			"Lookup",
			"Customer",
			"Owner",
			"Uniqueidentifier",
		],
	},
	{
		value: "ne",
		label: "Not Equals",
		requiresValue: true,
		applicableTypes: [
			"String",
			"Integer",
			"BigInt",
			"Decimal",
			"Double",
			"Money",
			"DateTime",
			"Boolean",
			"Picklist",
			"State",
			"Status",
			"Lookup",
			"Customer",
			"Owner",
			"Uniqueidentifier",
		],
	},
	{
		value: "lt",
		label: "Less Than",
		requiresValue: true,
		applicableTypes: ["Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "le",
		label: "Less Than or Equal",
		requiresValue: true,
		applicableTypes: ["Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "gt",
		label: "Greater Than",
		requiresValue: true,
		applicableTypes: ["Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "ge",
		label: "Greater Than or Equal",
		requiresValue: true,
		applicableTypes: ["Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},

	// String operators
	{
		value: "like",
		label: "Like (pattern)",
		requiresValue: true,
		applicableTypes: ["String", "Memo"],
	},
	{
		value: "not-like",
		label: "Not Like",
		requiresValue: true,
		applicableTypes: ["String", "Memo"],
	},
	{
		value: "begins-with",
		label: "Begins With",
		requiresValue: true,
		applicableTypes: ["String", "Memo"],
	},
	{
		value: "ends-with",
		label: "Ends With",
		requiresValue: true,
		applicableTypes: ["String", "Memo"],
	},
	{
		value: "not-begin-with",
		label: "Does Not Begin With",
		requiresValue: true,
		applicableTypes: ["String", "Memo"],
	},
	{
		value: "not-end-with",
		label: "Does Not End With",
		requiresValue: true,
		applicableTypes: ["String", "Memo"],
	},

	// List operators
	{
		value: "in",
		label: "In (list)",
		requiresValue: true,
		requiresMultipleValues: true,
		valueType: "string",
		applicableTypes: [
			"String",
			"Integer",
			"BigInt",
			"Decimal",
			"Double",
			"Money",
			"Picklist",
			"State",
			"Status",
			"Lookup",
			"Customer",
			"Owner",
			"Uniqueidentifier",
		],
	},
	{
		value: "not-in",
		label: "Not In (list)",
		requiresValue: true,
		requiresMultipleValues: true,
		valueType: "string",
		applicableTypes: [
			"String",
			"Integer",
			"BigInt",
			"Decimal",
			"Double",
			"Money",
			"Picklist",
			"State",
			"Status",
			"Lookup",
			"Customer",
			"Owner",
			"Uniqueidentifier",
		],
	},
	{
		value: "between",
		label: "Between (range)",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "number",
		applicableTypes: ["Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "not-between",
		label: "Not Between",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "number",
		applicableTypes: ["Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},

	// Null operators (all types)
	{
		value: "null",
		label: "Is Null",
		requiresValue: false,
		applicableTypes: [
			"String",
			"Memo",
			"Integer",
			"BigInt",
			"Decimal",
			"Double",
			"Money",
			"DateTime",
			"Boolean",
			"Picklist",
			"State",
			"Status",
			"Lookup",
			"Customer",
			"Owner",
			"Uniqueidentifier",
		],
	},
	{
		value: "not-null",
		label: "Is Not Null",
		requiresValue: false,
		applicableTypes: [
			"String",
			"Memo",
			"Integer",
			"BigInt",
			"Decimal",
			"Double",
			"Money",
			"DateTime",
			"Boolean",
			"Picklist",
			"State",
			"Status",
			"Lookup",
			"Customer",
			"Owner",
			"Uniqueidentifier",
		],
	},

	// Choice (multi-select picklist) operators
	{
		value: "contain-values",
		label: "Contains Values",
		requiresValue: true,
		requiresMultipleValues: true,
		valueType: "number",
		applicableTypes: ["Picklist", "State", "Status"],
	},
	{
		value: "not-contain-values",
		label: "Not Contain Values",
		requiresValue: true,
		requiresMultipleValues: true,
		valueType: "number",
		applicableTypes: ["Picklist", "State", "Status"],
	},

	// Date/Time operators (relative dates)
	{ value: "yesterday", label: "Yesterday", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "today", label: "Today", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "tomorrow", label: "Tomorrow", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "last-week", label: "Last Week", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "this-week", label: "This Week", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "next-week", label: "Next Week", requiresValue: false, applicableTypes: ["DateTime"] },
	{
		value: "last-seven-days",
		label: "Last Seven Days",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{ value: "last-month", label: "Last Month", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "this-month", label: "This Month", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "next-month", label: "Next Month", requiresValue: false, applicableTypes: ["DateTime"] },
	{
		value: "next-seven-days",
		label: "Next Seven Days",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{ value: "last-year", label: "Last Year", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "this-year", label: "This Year", requiresValue: false, applicableTypes: ["DateTime"] },
	{ value: "next-year", label: "Next Year", requiresValue: false, applicableTypes: ["DateTime"] },
	{
		value: "last-x-hours",
		label: "Last X Hours",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-hours",
		label: "Next X Hours",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-days",
		label: "Last X Days",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-days",
		label: "Next X Days",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-weeks",
		label: "Last X Weeks",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-weeks",
		label: "Next X Weeks",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-months",
		label: "Last X Months",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-months",
		label: "Next X Months",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-years",
		label: "Last X Years",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-years",
		label: "Next X Years",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{ value: "on", label: "On (specific date)", requiresValue: true, applicableTypes: ["DateTime"] },
	{
		value: "on-or-before",
		label: "On or Before",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "on-or-after",
		label: "On or After",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-minutes",
		label: "Older Than X Minutes",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-hours",
		label: "Older Than X Hours",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-days",
		label: "Older Than X Days",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-weeks",
		label: "Older Than X Weeks",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-months",
		label: "Older Than X Months",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-years",
		label: "Older Than X Years",
		requiresValue: true,
		applicableTypes: ["DateTime"],
	},

	// Fiscal period operators
	{
		value: "this-fiscal-year",
		label: "This Fiscal Year",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{
		value: "this-fiscal-period",
		label: "This Fiscal Period",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-fiscal-year",
		label: "Next Fiscal Year",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-fiscal-period",
		label: "Next Fiscal Period",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-fiscal-year",
		label: "Last Fiscal Year",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-fiscal-period",
		label: "Last Fiscal Period",
		requiresValue: false,
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-fiscal-year",
		label: "In Fiscal Year",
		requiresValue: true,
		valueType: "number",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-fiscal-period",
		label: "In Fiscal Period",
		requiresValue: true,
		valueType: "number",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-fiscal-period-and-year",
		label: "In Fiscal Period and Year",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "number",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-or-before-fiscal-period-and-year",
		label: "In or Before Fiscal Period and Year",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "number",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-or-after-fiscal-period-and-year",
		label: "In or After Fiscal Period and Year",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "number",
		applicableTypes: ["DateTime"],
	},

	// User/Team context operators (Lookup, Owner types)
	{
		value: "eq-userid",
		label: "Equals Current User",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
	},
	{
		value: "ne-userid",
		label: "Not Equals Current User",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
	},
	{
		value: "eq-businessid",
		label: "Equals Current Business Unit",
		requiresValue: false,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
	{
		value: "ne-businessid",
		label: "Not Equals Current Business Unit",
		requiresValue: false,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
	{
		value: "eq-userteams",
		label: "Equals User Teams",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
	},
	{
		value: "eq-useroruserteams",
		label: "Equals User or User Teams",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
	},
	{
		value: "eq-useroruserhierarchy",
		label: "Equals User or User Hierarchy",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
	},
	{
		value: "eq-useroruserhierarchyandteams",
		label: "Equals User or User Hierarchy and Teams",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
	},
	{
		value: "eq-userlanguage",
		label: "Equals User Language",
		requiresValue: false,
		applicableTypes: ["Integer"],
	},

	// Special operators
	{
		value: "under",
		label: "Under (hierarchy)",
		requiresValue: true,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
	{
		value: "not-under",
		label: "Not Under (hierarchy)",
		requiresValue: true,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
	{
		value: "above",
		label: "Above (hierarchy)",
		requiresValue: true,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
	{
		value: "eq-or-under",
		label: "Equals or Under (hierarchy)",
		requiresValue: true,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
	{
		value: "eq-or-above",
		label: "Equals or Above (hierarchy)",
		requiresValue: true,
		applicableTypes: ["Lookup", "Uniqueidentifier"],
	},
];

/**
 * Get operators applicable to a specific attribute type
 * @param attributeType - The AttributeType from metadata (e.g., 'String', 'Integer', 'DateTime')
 * @returns Array of operator definitions applicable to this type
 */
export function getOperatorsForAttributeType(
	attributeType: string | undefined
): OperatorDefinition[] {
	if (!attributeType) {
		// No type info - return all operators as fallback
		return ALL_OPERATORS;
	}

	// Normalize type name (handle both 'String' and 'StringType' formats)
	const normalizedType = attributeType.replace(/Type$/, "");

	// Filter operators that support this attribute type
	const filtered = ALL_OPERATORS.filter((op) =>
		op.applicableTypes.some((type) => type.toLowerCase() === normalizedType.toLowerCase())
	);

	// If no matches found (unknown type), return all as fallback
	return filtered.length > 0 ? filtered : ALL_OPERATORS;
}

/**
 * Check if an operator requires a value input
 * @param operator - The operator to check
 * @returns true if the operator requires a value, false otherwise
 */
export function operatorRequiresValue(operator: OperatorType): boolean {
	const definition = ALL_OPERATORS.find((op) => op.value === operator);
	return definition?.requiresValue ?? true; // Default to requiring value if not found
}

/**
 * Get common operators (subset for simpler UIs)
 * These are the most frequently used operators across all types
 */
export function getCommonOperators(): OperatorDefinition[] {
	const commonValues: OperatorType[] = [
		"eq",
		"ne",
		"lt",
		"le",
		"gt",
		"ge",
		"like",
		"begins-with",
		"in",
		"not-in",
		"between",
		"null",
		"not-null",
		"today",
		"yesterday",
		"tomorrow",
		"this-week",
		"this-month",
		"this-year",
		"last-x-days",
		"next-x-days",
		"eq-userid",
		"ne-userid",
	];

	return ALL_OPERATORS.filter((op) => commonValues.includes(op.value));
}
