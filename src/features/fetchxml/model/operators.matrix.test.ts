/**
 * Operator × attribute type matrix (#95): the exact operators FetchXML Studio offers for each attribute
 * type it supports. Each row follows the operators reference's "By Data type" tables
 * (https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators#by-data-type),
 * plus combinations verified live in the #87 matrix (2026-10-15, every operator × every Advanced Find
 * attribute type) and the #57 checks (2026-10-16). Combinations Dataverse rejects aren't offered.
 */

import { describe, expect, it } from "vitest";
import { ALL_OPERATORS, getOperatorsForAttribute, type OperatorAttributeInfo } from "./operators";

const sorted = (ops: readonly string[]) => [...ops].sort();
const offered = (attr: OperatorAttributeInfo | undefined, hierarchical?: boolean) =>
	sorted(getOperatorsForAttribute(attr, hierarchical).map((o) => o.value));

const NULL = ["null", "not-null"];
// String data: also the reference's eq / ne / in for String; not-in verified (#87)
const STRING = ["eq", "ne", "gt", "ge", "lt", "le", "like", "not-like", "begins-with", "not-begin-with", "ends-with", "not-end-with", "in", "not-in", ...NULL];
// Number data (eq-userlanguage is for Language attributes, an Integer format)
const NUMBER = ["eq", "ne", "gt", "ge", "lt", "le", "in", "not-in", "between", "not-between", ...NULL];
// Datetime data: operators that need hours or minutes aren't for DateOnly behaviour (Datetime data note)
const TIME = ["last-x-hours", "next-x-hours", "olderthan-x-minutes", "olderthan-x-hours"];
const DATE = [
	"eq", "ne", "gt", "ge", "lt", "le", "between", "not-between", ...NULL,
	"on", "on-or-before", "on-or-after", "yesterday", "today", "tomorrow",
	"last-week", "this-week", "next-week", "last-seven-days", "next-seven-days",
	"last-month", "this-month", "next-month", "last-year", "this-year", "next-year",
	"last-x-days", "next-x-days", "last-x-weeks", "next-x-weeks", "last-x-months", "next-x-months",
	"last-x-years", "next-x-years", "olderthan-x-days", "olderthan-x-weeks", "olderthan-x-months", "olderthan-x-years",
	"this-fiscal-year", "this-fiscal-period", "next-fiscal-year", "next-fiscal-period", "last-fiscal-year", "last-fiscal-period",
	"in-fiscal-year", "in-fiscal-period", "last-x-fiscal-years", "last-x-fiscal-periods", "next-x-fiscal-years", "next-x-fiscal-periods",
	"in-fiscal-period-and-year", "in-or-before-fiscal-period-and-year", "in-or-after-fiscal-period-and-year",
	...TIME,
];
// Choice data (single choice): contain-values / not-contain-values fail on Picklist (0x80048415, #87)
const CHOICE = ["eq", "ne", "in", "not-in", ...NULL];
// Unique Identifier data, plus eq-userteams / eq-useroruserteams (Owner data; verified on user lookups and
// Uniqueidentifier, #87) and the hierarchy security operators (Hierarchical data; verified, #87)
const USER = ["eq-userid", "ne-userid", "eq-userteams", "eq-useroruserteams", "eq-useroruserhierarchy", "eq-useroruserhierarchyandteams"];
const GUID = ["eq", "ne", "in", "not-in", ...NULL, ...USER, "eq-businessid", "ne-businessid"];
// Hierarchical data: query-hierarchical-data, the primary key of a hierarchical entity only (#87)
const HIERARCHY = ["under", "not-under", "above", "eq-or-under", "eq-or-above"];

const MULTI_SELECT: OperatorAttributeInfo = { AttributeType: "Virtual", AttributeTypeName: { Value: "MultiSelectPicklistType" } };

const rows: [string, OperatorAttributeInfo, string[], boolean?][] = [
	// String data
	["String", { AttributeType: "String" }, STRING],
	// String data: multi-line text is a string; eq / ne / in / not-in verified on address1_composite (#87, #95)
	["Memo", { AttributeType: "Memo" }, STRING],
	// Number data
	["Integer", { AttributeType: "Integer" }, [...NUMBER, "eq-userlanguage"]],
	["BigInt", { AttributeType: "BigInt" }, NUMBER],
	["Decimal", { AttributeType: "Decimal" }, NUMBER],
	["Double", { AttributeType: "Double" }, NUMBER],
	["Money", { AttributeType: "Money" }, NUMBER],
	// Datetime data
	["DateTime (user local)", { AttributeType: "DateTime", DateTimeBehavior: { Value: "UserLocal" } }, DATE],
	["DateTime (date only)", { AttributeType: "DateTime", DateTimeBehavior: { Value: "DateOnly" } }, DATE.filter((o) => !TIME.includes(o))],
	// Not in the reference's types; eq / ne / null / not-null verified (#87)
	["Boolean", { AttributeType: "Boolean" }, ["eq", "ne", ...NULL]],
	// Choice data
	["Picklist", { AttributeType: "Picklist" }, CHOICE],
	["State", { AttributeType: "State" }, CHOICE],
	["Status", { AttributeType: "Status" }, CHOICE],
	// Choice data + multi-select-picklist#query-data-from-choices (eq, ne, in, not-in, null, not-null,
	// contain-values, not-contain-values); all verified on appaction.clienttype (#87, #95)
	["Choices (multi-select)", MULTI_SELECT, ["contain-values", "not-contain-values", ...CHOICE]],
	// Unique Identifier data
	["Lookup", { AttributeType: "Lookup" }, GUID],
	["Uniqueidentifier", { AttributeType: "Uniqueidentifier" }, GUID],
	// Owner data (eq, ne, in, eq-userteams, eq-useroruserteams); eq-userid / ne-userid and the hierarchy
	// security operators verified (#87). eq-businessid passes but an owner is never a business unit.
	["Owner", { AttributeType: "Owner" }, ["eq", "ne", "in", "not-in", ...NULL, ...USER]],
	// Customer: user / team operators fail (0x8004034b) except eq-useroruserhierarchy (verified, #87)
	["Customer", { AttributeType: "Customer" }, ["eq", "ne", "in", "not-in", ...NULL, "eq-useroruserhierarchy"]],
	// Hierarchical data: under / above and variants on the primary key of a hierarchical entity
	["Primary key, hierarchical entity", { AttributeType: "Uniqueidentifier", IsPrimaryId: true }, [...GUID, ...HIERARCHY], true],
	["Primary key, hierarchy unknown", { AttributeType: "Uniqueidentifier", IsPrimaryId: true }, [...GUID, ...HIERARCHY], undefined],
	["Primary key, not hierarchical", { AttributeType: "Uniqueidentifier", IsPrimaryId: true }, GUID, false],
	// A hierarchical lookup (parentaccountid) rejects the hierarchy operators (0x80040203, #87)
	["Lookup on a hierarchical entity", { AttributeType: "Lookup" }, GUID, true],
	// Advanced Find only off (#57): label attributes compare as text; the rest take null / not-null only
	["Virtual (label)", { AttributeType: "Virtual", AttributeTypeName: { Value: "VirtualType" } }, ["eq", "like", ...NULL]],
	["Image", { AttributeType: "Virtual", AttributeTypeName: { Value: "ImageType" } }, NULL],
	["File", { AttributeType: "Virtual", AttributeTypeName: { Value: "FileType" } }, NULL],
	["ManagedProperty", { AttributeType: "ManagedProperty", AttributeTypeName: { Value: "ManagedPropertyType" } }, NULL],
	["EntityName", { AttributeType: "EntityName", AttributeTypeName: { Value: "EntityNameType" } }, NULL],
	// Every condition on a party list fails (0x80040203, #57)
	["PartyList", { AttributeType: "PartyList", AttributeTypeName: { Value: "PartyListType" } }, []],
];

describe("operator × attribute type matrix (#95)", () => {
	it.each(rows)("%s", (_, attr, expected, hierarchical) => {
		expect(offered(attr, hierarchical)).toEqual(sorted(expected));
	});

	it("offers every operator when the attribute's type is unknown", () => {
		expect(offered(undefined)).toEqual(sorted(ALL_OPERATORS.map((o) => o.value)));
	});

	it("every operator is offered for at least one attribute type", () => {
		const reachable = new Set(rows.flatMap(([, attr, , hierarchical]) => offered(attr, hierarchical)));
		expect(ALL_OPERATORS.map((o) => o.value).filter((op) => !reachable.has(op))).toEqual([]);
	});
});

/**
 * Every operator the reference documents for a data type is offered for the matching attribute types,
 * apart from the exceptions listed (each verified to fail, or to be meaningless, in the #87 matrix).
 */
describe("documented operators per data type are offered (#95)", () => {
	const byDataType: Record<string, { ops: string[]; attrs: [OperatorAttributeInfo, boolean?][]; except?: Record<string, string> }> = {
		// "Choice data"
		Choice: {
			ops: ["contain-values", "eq", "in", "ne", "not-contain-values", "not-null", "null"],
			attrs: [[{ AttributeType: "Picklist" }], [{ AttributeType: "State" }], [{ AttributeType: "Status" }], [MULTI_SELECT]],
			except: { "contain-values": "single choice: 0x80048415", "not-contain-values": "single choice: 0x80048415" },
		},
		// "Datetime data"
		Datetime: {
			ops: DATE,
			attrs: [[{ AttributeType: "DateTime" }]],
		},
		// "Number data" (eq-userlanguage: Language attributes, which are Integer)
		Number: {
			ops: ["between", "eq", "ge", "gt", "in", "le", "lt", "ne", "not-between", "not-in", "not-null", "null"],
			attrs: [[{ AttributeType: "Integer" }], [{ AttributeType: "BigInt" }], [{ AttributeType: "Decimal" }], [{ AttributeType: "Double" }], [{ AttributeType: "Money" }]],
		},
		// "Owner data"
		Owner: {
			ops: ["eq-useroruserteams", "eq-userteams", "eq", "in", "ne"],
			attrs: [[{ AttributeType: "Owner" }]],
		},
		// "String data": String and multi-line text (Memo)
		String: {
			ops: ["begins-with", "ends-with", "eq", "ge", "gt", "in", "le", "like", "lt", "ne", "not-begin-with", "not-end-with", "not-like", "not-null", "null"],
			attrs: [[{ AttributeType: "String" }], [{ AttributeType: "Memo" }]],
		},
		// "Unique Identifier data"
		"Unique Identifier": {
			ops: ["eq-businessid", "eq-userid", "eq", "in", "ne-businessid", "ne-userid", "ne", "not-null", "null"],
			attrs: [[{ AttributeType: "Uniqueidentifier" }], [{ AttributeType: "Lookup" }]],
		},
		// "Hierarchical data" (query-hierarchical-data): the primary key of a hierarchical entity
		Hierarchical: {
			ops: ["above", "eq-or-above", "eq-or-under", "eq-useroruserhierarchy", "eq-useroruserhierarchyandteams", "not-under", "under"],
			attrs: [[{ AttributeType: "Uniqueidentifier", IsPrimaryId: true }, true]],
		},
	};

	for (const [dataType, { ops, attrs, except = {} }] of Object.entries(byDataType)) {
		it(dataType, () => {
			for (const [attr, hierarchical] of attrs) {
				const list = offered(attr, hierarchical);
				const isMulti = attr.AttributeTypeName?.Value === "MultiSelectPicklistType";
				const missing = ops.filter((op) => !list.includes(op) && (isMulti || !except[op]));
				expect(missing, `${attr.AttributeTypeName?.Value ?? attr.AttributeType}`).toEqual([]);
			}
		});
	}
});

describe("operator help text (#95)", () => {
	const op = (value: string) => ALL_OPERATORS.find((o) => o.value === value);

	it.each([...HIERARCHY, "eq-useroruserhierarchy", "eq-useroruserhierarchyandteams"])("%s links to query-hierarchical-data", (value) => {
		expect(op(value)?.description).toBeTruthy();
		expect(op(value)?.docsUrl).toBe("https://learn.microsoft.com/power-apps/developer/data-platform/query-hierarchical-data");
	});

	it("the hierarchy security operators say they need hierarchy security", () => {
		expect(op("eq-useroruserhierarchy")?.description).toMatch(/hierarchy security/);
		expect(op("eq-useroruserhierarchyandteams")?.description).toMatch(/hierarchy security/);
	});

	it.each(["contain-values", "not-contain-values"])("%s links to the choices doc", (value) => {
		expect(op(value)?.description).toMatch(/full-text index/);
		expect(op(value)?.docsUrl).toContain("multi-select-picklist#query-data-from-choices");
	});
});
