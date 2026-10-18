/**
 * FetchXML operator definitions and filtering based on attribute metadata.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators
 *
 * Which operators each attribute offers is the documented data types per operator, checked against a
 * live matrix of every operator × every Advanced Find attribute type (2026-10-15, #87). Combinations
 * Dataverse rejects aren't offered (e.g. contain-values on single-choice columns, user / team operators
 * on Customer lookups, hierarchy operators on anything but the primary key).
 *
 * #95: the reference's "String" type covers multi-line text (Memo) too, and the choices (multi-select)
 * doc lists eq, ne, in and not-in besides contain-values / not-contain-values; all of them pass in the
 * #87 matrix (Memo: address1_composite; multi-select: appaction.clienttype), so Memo and multi-select
 * choices now offer them. Operator × type matrix test:
 * operators.matrix.test.ts.
 * https://learn.microsoft.com/power-apps/developer/data-platform/multi-select-picklist#query-data-from-choices
 */

import type { ConditionNode, OperatorType } from "./nodes";

const HIERARCHY_DOCS = "https://learn.microsoft.com/power-apps/developer/data-platform/query-hierarchical-data";
const CHOICES_DOCS = "https://learn.microsoft.com/power-apps/developer/data-platform/multi-select-picklist#query-data-from-choices";
/** query-hierarchical-data: "a default limit of 100 recursions" for above / under and their variants */
const HIERARCHY_NOTE =
	"Only for the primary key of an entity with a hierarchical relationship. Dataverse follows up to 100 levels.";

export interface OperatorDefinition {
	value: OperatorType;
	label: string;
	requiresValue: boolean; // Does it need any value?
	requiresTwoValues?: boolean; // Does it need exactly TWO values? (e.g., between, in-fiscal-period-and-year)
	requiresMultipleValues?: boolean; // Does it accept multiple values? (e.g., in, contain-values)
	/**
	 * What value(s) it takes, where that isn't the attribute's own type:
	 * - integer: a whole number above 0 in the value attribute (verified: 0, negative, decimal or a date
	 *   fail with 0x80040203 / 0x80041139);
	 * - year: a four-digit fiscal year (verified: "23" fails);
	 * - fiscalPeriodYear: two <value> elements, period then year.
	 */
	valueType?: "string" | "number" | "date" | "guid" | "boolean" | "integer" | "year" | "fiscalPeriodYear";
	applicableTypes: string[]; // AttributeType values ("MultiSelectPicklist" for multi-select choices)
	/** Hierarchy operators: only the primary key of a hierarchical entity (verified: 0x80040203 otherwise) */
	hierarchyPrimaryKeyOnly?: boolean;
	/** Needs hours or minutes: not for date-only behaviour (verified: 0x80048415) */
	needsTime?: boolean;
	/** Short help shown with the operator in the condition editor (#95), from the linked docs */
	description?: string;
	/** The Learn page behind description (the operators reference otherwise) */
	docsUrl?: string;
}

/**
 * All available FetchXML operators with their metadata
 */
export const ALL_OPERATORS: OperatorDefinition[] = [
	// Comparison operators (most types). eq / ne / in / not-in on Memo: "String" in the reference, verified
	// on address1_composite (#87 matrix); on multi-select choices: listed in the choices doc, verified (#87 matrix, #95)
	{
		value: "eq",
		label: "Equals",
		requiresValue: true,
		applicableTypes: [
			"String",
			"Memo",
			"MultiSelectPicklist",
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
			"Virtual",
		],
	},
	{
		value: "ne",
		label: "Not Equals",
		requiresValue: true,
		applicableTypes: [
			"String",
			"Memo",
			"MultiSelectPicklist",
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
		applicableTypes: ["String", "Memo", "Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "le",
		label: "Less Than or Equal",
		requiresValue: true,
		applicableTypes: ["String", "Memo", "Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "gt",
		label: "Greater Than",
		requiresValue: true,
		applicableTypes: ["String", "Memo", "Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},
	{
		value: "ge",
		label: "Greater Than or Equal",
		requiresValue: true,
		applicableTypes: ["String", "Memo", "Integer", "BigInt", "Decimal", "Double", "Money", "DateTime"],
	},

	// String operators
	{
		value: "like",
		label: "Like (pattern)",
		requiresValue: true,
		applicableTypes: ["String", "Memo", "Virtual"],
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
			"Memo",
			"MultiSelectPicklist",
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
		// The reference lists Number only; String, choices (choices doc) and the others are verified (#87)
		applicableTypes: [
			"String",
			"Memo",
			"MultiSelectPicklist",
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
		// Virtual (the "...name" label attributes), Image, File, ManagedProperty and EntityName only
		// appear with "Advanced Find only" off; null / not-null verified on each (#57)
		applicableTypes: [
			"Virtual",
			"Image",
			"File",
			"ManagedProperty",
			"EntityName",
			"MultiSelectPicklist",
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
			"Virtual",
			"Image",
			"File",
			"ManagedProperty",
			"EntityName",
			"MultiSelectPicklist",
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

	// Choice (multi-select picklist) operators. The choices doc: they rely on full-text indexing, so new
	// records can take a few seconds to match.
	{
		value: "contain-values",
		label: "Contains Values",
		requiresValue: true,
		requiresMultipleValues: true,
		valueType: "number",
		applicableTypes: ["MultiSelectPicklist"],
		description:
			"Matches records where any of these options is selected. Uses full-text indexing, so new records can take a few seconds to match.",
		docsUrl: CHOICES_DOCS,
	},
	{
		value: "not-contain-values",
		label: "Not Contain Values",
		requiresValue: true,
		requiresMultipleValues: true,
		valueType: "number",
		applicableTypes: ["MultiSelectPicklist"],
		description:
			"Matches records where none of these options is selected. Uses full-text indexing, so new records can take a few seconds to match.",
		docsUrl: CHOICES_DOCS,
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
		valueType: "integer",
		needsTime: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-hours",
		label: "Next X Hours",
		requiresValue: true,
		valueType: "integer",
		needsTime: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-days",
		label: "Last X Days",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-days",
		label: "Next X Days",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-weeks",
		label: "Last X Weeks",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-weeks",
		label: "Next X Weeks",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-months",
		label: "Last X Months",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-months",
		label: "Next X Months",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-years",
		label: "Last X Years",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-years",
		label: "Next X Years",
		requiresValue: true,
		valueType: "integer",
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
		valueType: "integer",
		needsTime: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-hours",
		label: "Older Than X Hours",
		requiresValue: true,
		valueType: "integer",
		needsTime: true,
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-days",
		label: "Older Than X Days",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-weeks",
		label: "Older Than X Weeks",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-months",
		label: "Older Than X Months",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "olderthan-x-years",
		label: "Older Than X Years",
		requiresValue: true,
		valueType: "integer",
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
		valueType: "year",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-fiscal-period",
		label: "In Fiscal Period",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-fiscal-years",
		label: "Last X Fiscal Years",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "last-x-fiscal-periods",
		label: "Last X Fiscal Periods",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-fiscal-years",
		label: "Next X Fiscal Years",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "next-x-fiscal-periods",
		label: "Next X Fiscal Periods",
		requiresValue: true,
		valueType: "integer",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-fiscal-period-and-year",
		label: "In Fiscal Period and Year",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "fiscalPeriodYear",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-or-before-fiscal-period-and-year",
		label: "In or Before Fiscal Period and Year",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "fiscalPeriodYear",
		applicableTypes: ["DateTime"],
	},
	{
		value: "in-or-after-fiscal-period-and-year",
		label: "In or After Fiscal Period and Year",
		requiresValue: true,
		requiresTwoValues: true,
		valueType: "fiscalPeriodYear",
		applicableTypes: ["DateTime"],
	},

	// User/Team context operators (Lookup, Owner types). The reference types eq-userteams /
	// eq-useroruserteams as Owner, but they're verified on user lookups (createdby) and Uniqueidentifier
	// too, so those stay (#95: not narrowed). eq-businessid / ne-businessid on Owner are accepted but
	// can't match (an owner is a user or team, never a business unit), so they aren't offered there.
	{
		value: "eq-userid",
		label: "Equals Current User",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Uniqueidentifier"],
	},
	{
		value: "ne-userid",
		label: "Not Equals Current User",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Uniqueidentifier"],
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
		applicableTypes: ["Lookup", "Owner", "Uniqueidentifier"],
	},
	{
		value: "eq-useroruserteams",
		label: "Equals User or User Teams",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Uniqueidentifier"],
	},
	{
		value: "eq-useroruserhierarchy",
		label: "Equals User or User Hierarchy",
		requiresValue: false,
		// Customer: verified (#87 matrix), unlike the other user operators (0x8004034b)
		applicableTypes: ["Lookup", "Owner", "Customer", "Uniqueidentifier"],
		// query-hierarchical-data: hierarchy security operators, depth from the Hierarchy Depth setting
		description:
			"When hierarchy security is used: the current user or their reporting hierarchy, up to the Hierarchy Depth setting.",
		docsUrl: HIERARCHY_DOCS,
	},
	{
		value: "eq-useroruserhierarchyandteams",
		label: "Equals User or User Hierarchy and Teams",
		requiresValue: false,
		applicableTypes: ["Lookup", "Owner", "Uniqueidentifier"],
		description:
			"When hierarchy security is used: the current user and their teams, or their reporting hierarchy and its teams, up to the Hierarchy Depth setting.",
		docsUrl: HIERARCHY_DOCS,
	},
	{
		value: "eq-userlanguage",
		label: "Equals User Language",
		requiresValue: false,
		// The reference: for Language attributes (Integer with the Language format)
		applicableTypes: ["Integer"],
		description: "For a language attribute: matches the current user's UI language.",
		docsUrl: "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators#eq-userlanguage",
	},

	// Hierarchy operators: query-hierarchical-data ("only apply for the table relationship specifically
	// defined as a hierarchical relationship"); the primary key only (verified, #87)
	{
		value: "under",
		label: "Under (hierarchy)",
		requiresValue: true,
		valueType: "guid",
		applicableTypes: ["Uniqueidentifier"],
		hierarchyPrimaryKeyOnly: true,
		description: HIERARCHY_NOTE,
		docsUrl: HIERARCHY_DOCS,
	},
	{
		value: "not-under",
		label: "Not Under (hierarchy)",
		requiresValue: true,
		valueType: "guid",
		applicableTypes: ["Uniqueidentifier"],
		hierarchyPrimaryKeyOnly: true,
		description: HIERARCHY_NOTE,
		docsUrl: HIERARCHY_DOCS,
	},
	{
		value: "above",
		label: "Above (hierarchy)",
		requiresValue: true,
		valueType: "guid",
		applicableTypes: ["Uniqueidentifier"],
		hierarchyPrimaryKeyOnly: true,
		description: HIERARCHY_NOTE,
		docsUrl: HIERARCHY_DOCS,
	},
	{
		value: "eq-or-under",
		label: "Equals or Under (hierarchy)",
		requiresValue: true,
		valueType: "guid",
		applicableTypes: ["Uniqueidentifier"],
		hierarchyPrimaryKeyOnly: true,
		description: HIERARCHY_NOTE,
		docsUrl: HIERARCHY_DOCS,
	},
	{
		value: "eq-or-above",
		label: "Equals or Above (hierarchy)",
		requiresValue: true,
		valueType: "guid",
		applicableTypes: ["Uniqueidentifier"],
		hierarchyPrimaryKeyOnly: true,
		description: HIERARCHY_NOTE,
		docsUrl: HIERARCHY_DOCS,
	},
];

/** The attribute metadata operator filtering needs (AttributeMetadata subset). */
export interface OperatorAttributeInfo {
	AttributeType?: string;
	AttributeTypeName?: { Value?: string };
	IsPrimaryId?: boolean;
	/** DateTimeAttributeMetadata.DateTimeBehavior.Value */
	DateTimeBehavior?: { Value?: string };
}

/**
 * Operators for an attribute, from its metadata:
 * - multi-select choices (AttributeTypeName MultiSelectPicklistType) get the multi-select operators;
 * - hierarchy operators only for the primary key, and not when the entity is known not to be
 *   hierarchical (isHierarchical false);
 * - operators that need hours / minutes aren't offered for date-only behaviour (a date-only format
 *   with user-local behaviour accepts them, verified).
 * Without metadata every operator is offered, because the type can't be determined.
 */
export function getOperatorsForAttribute(
	attribute: OperatorAttributeInfo | undefined,
	isHierarchical?: boolean,
): OperatorDefinition[] {
	if (!attribute?.AttributeType) return ALL_OPERATORS;
	// Virtual covers several types: AttributeTypeName tells multi-select choices, images and files
	// apart from the "...name" label attributes (VirtualType)
	const typeName = attribute.AttributeTypeName?.Value;
	const byTypeName = !!typeName && (attribute.AttributeType === "Virtual" || typeName === "MultiSelectPicklistType");
	const type = (byTypeName ? typeName : attribute.AttributeType)
		.replace(/Type$/, "")
		.toLowerCase();
	const dateOnly = attribute.DateTimeBehavior?.Value === "DateOnly";
	return ALL_OPERATORS.filter((op) => {
		if (op.hierarchyPrimaryKeyOnly) return !!attribute.IsPrimaryId && isHierarchical !== false;
		if (op.needsTime && dateOnly) return false;
		return op.applicableTypes.some((t) => t.toLowerCase() === type);
	});
}

/**
 * Get operators applicable to an AttributeType (no other metadata).
 * @deprecated Use getOperatorsForAttribute, which also handles multi-select, primary keys and
 * date-only behaviour.
 */
export function getOperatorsForAttributeType(attributeType: string | undefined): OperatorDefinition[] {
	return getOperatorsForAttribute(attributeType ? { AttributeType: attributeType } : undefined);
}

/** The shape of value an operator takes. */
export type OperatorValueShape = "none" | "single" | "integer" | "year" | "two" | "fiscalPeriodYear" | "multiple";

export function operatorValueShape(operator: string): OperatorValueShape {
	const op = ALL_OPERATORS.find((o) => o.value === operator);
	if (!op) return "single";
	if (!op.requiresValue) return "none";
	if (op.valueType === "fiscalPeriodYear") return "fiscalPeriodYear";
	if (op.valueType === "integer") return "integer";
	if (op.valueType === "year") return "year";
	if (op.requiresTwoValues) return "two";
	if (op.requiresMultipleValues) return "multiple";
	return "single";
}

/**
 * The value to keep when the operator changes: kept when the shape is the same (and still fits),
 * otherwise cleared, so the condition never holds a value Dataverse rejects (#87).
 */
export function adaptValueForOperator(
	previousOperator: string,
	nextOperator: string,
	value: ConditionNode["value"],
): ConditionNode["value"] {
	const before = operatorValueShape(previousOperator);
	const after = operatorValueShape(nextOperator);
	if (after === "none") return undefined;
	if (before !== after) return undefined;
	return value;
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

/**
 * Every condition operator Dataverse accepts: the documented operators (all selectable since #87)
 * plus the deprecated "neq" (verified: works like "ne"). Shared by the parser and the editor's
 * intellisense. https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators
 */
export const KNOWN_OPERATORS: ReadonlySet<string> = new Set<string>([...ALL_OPERATORS.map((o) => o.value), "neq"]);
