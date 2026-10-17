import { describe, expect, it } from "vitest";
import {
	ALL_OPERATORS,
	adaptValueForOperator,
	getOperatorsForAttribute,
	getOperatorsForAttributeType,
	operatorRequiresValue,
	operatorValueShape,
} from "./operators";

describe("operatorRequiresValue", () => {
	it.each(["null", "not-null", "today", "eq-userid", "this-fiscal-year"] as const)(
		"%s needs no value",
		(op) => {
			expect(operatorRequiresValue(op)).toBe(false);
		},
	);

	it.each(["eq", "like", "in", "between", "last-x-days", "under"] as const)(
		"%s needs a value",
		(op) => {
			expect(operatorRequiresValue(op)).toBe(true);
		},
	);
});

describe("ALL_OPERATORS", () => {
	it("has no duplicate operator values", () => {
		const values = ALL_OPERATORS.map((o) => o.value);
		expect(new Set(values).size).toBe(values.length);
	});

	it("uses the documented spelling for the negated begins/ends-with operators", () => {
		const values = new Set(ALL_OPERATORS.map((o) => o.value));
		expect(values.has("not-begin-with")).toBe(true);
		expect(values.has("not-end-with")).toBe(true);
	});
});

describe("getOperatorsForAttributeType", () => {
	it("offers string operators for String attributes", () => {
		const ops = getOperatorsForAttributeType("String").map((o) => o.value);
		expect(ops).toContain("like");
		expect(ops).toContain("begins-with");
	});

	it("does not offer string operators for Integer attributes", () => {
		const ops = getOperatorsForAttributeType("Integer").map((o) => o.value);
		expect(ops).not.toContain("like");
		expect(ops).toContain("between");
	});
});

// #87: operators per attribute from the live operator × attribute-type matrix (2026-10-15)

describe("every documented operator is selectable (#87)", () => {
	it.each(["last-x-fiscal-years", "last-x-fiscal-periods", "next-x-fiscal-years", "next-x-fiscal-periods"])("%s", (op) => {
		expect(ALL_OPERATORS.map((o) => o.value)).toContain(op);
	});
});

describe("getOperatorsForAttribute (#87)", () => {
	const ops = (attr: Parameters<typeof getOperatorsForAttribute>[0], hierarchical?: boolean) =>
		getOperatorsForAttribute(attr, hierarchical).map((o) => o.value);

	it("contain-values only for multi-select choices (picklist: 0x80048415)", () => {
		expect(ops({ AttributeType: "Picklist" })).not.toContain("contain-values");
		const multi = ops({ AttributeType: "Virtual", AttributeTypeName: { Value: "MultiSelectPicklistType" } });
		expect(multi).toEqual(expect.arrayContaining(["contain-values", "not-contain-values", "null", "not-null"]));
		expect(multi).not.toContain("like");
	});

	it("user / team operators aren't offered for Customer lookups (0x8004034b)", () => {
		const customer = ops({ AttributeType: "Customer" });
		for (const op of ["eq-userid", "ne-userid", "eq-userteams", "eq-useroruserteams"]) expect(customer).not.toContain(op);
		expect(ops({ AttributeType: "Owner" })).toContain("eq-userid");
	});

	it("hierarchy operators only for the primary key of a hierarchical entity", () => {
		expect(ops({ AttributeType: "Lookup" }, true)).not.toContain("under");
		expect(ops({ AttributeType: "Uniqueidentifier", IsPrimaryId: true }, true)).toContain("under");
		expect(ops({ AttributeType: "Uniqueidentifier", IsPrimaryId: true }, false)).not.toContain("under");
	});

	it("hour / minute operators are left out for date-only behaviour, not for a date-only format", () => {
		expect(ops({ AttributeType: "DateTime", DateTimeBehavior: { Value: "DateOnly" } })).not.toContain("last-x-hours");
		expect(ops({ AttributeType: "DateTime", DateTimeBehavior: { Value: "UserLocal" } })).toContain("olderthan-x-minutes");
		expect(ops({ AttributeType: "DateTime", DateTimeBehavior: { Value: "DateOnly" } })).toContain("last-x-days");
	});

	it("strings get gt / ge / lt / le (documented)", () => {
		expect(ops({ AttributeType: "String" })).toEqual(expect.arrayContaining(["gt", "ge", "lt", "le"]));
	});
});

describe("value shapes (#87)", () => {
	it.each([
		["last-x-days", "integer"],
		["olderthan-x-minutes", "integer"],
		["in-fiscal-period", "integer"],
		["next-x-fiscal-years", "integer"],
		["in-fiscal-year", "year"],
		["in-fiscal-period-and-year", "fiscalPeriodYear"],
		["between", "two"],
		["in", "multiple"],
		["eq", "single"],
		["today", "none"],
	])("%s takes %s", (op, shape) => expect(operatorValueShape(op)).toBe(shape));

	it("keeps the value only when the shape stays the same", () => {
		expect(adaptValueForOperator("eq", "ne", "x")).toBe("x");
		expect(adaptValueForOperator("on", "last-x-days", "2024-01-01")).toBeUndefined();
		expect(adaptValueForOperator("in", "eq", ["a", "b"])).toBeUndefined();
		expect(adaptValueForOperator("last-x-days", "next-x-weeks", 7)).toBe(7);
		expect(adaptValueForOperator("eq", "null", "x")).toBeUndefined();
	});
});

describe("getOperatorsForAttribute: attribute types shown with Advanced Find only off (#57)", () => {
	// Verified on the sandbox (2026-10-16) against hidden attributes of 15 entities
	const ops = (AttributeType: string, typeName: string) =>
		getOperatorsForAttribute({ AttributeType, AttributeTypeName: { Value: typeName } }).map((o) => o.value);

	it("Virtual \"...name\" label attributes: null, not-null, eq and like", () => {
		expect(ops("Virtual", "VirtualType").sort()).toEqual(["eq", "like", "not-null", "null"]);
	});

	it.each([
		["Virtual", "ImageType"],
		["Virtual", "FileType"],
		["ManagedProperty", "ManagedPropertyType"],
		["EntityName", "EntityNameType"],
	])("%s (%s): null and not-null only (eq / like fail: 0x80040203)", (type, typeName) => {
		expect(ops(type, typeName).sort()).toEqual(["not-null", "null"]);
	});

	it("PartyList: none (every condition fails: 0x80040203)", () => {
		expect(ops("PartyList", "PartyListType")).toEqual([]);
	});

	it("multi-select choices (AttributeType Virtual) keep their own operators", () => {
		expect(ops("Virtual", "MultiSelectPicklistType")).toContain("contain-values");
		expect(ops("Virtual", "MultiSelectPicklistType")).not.toContain("like");
	});
});
