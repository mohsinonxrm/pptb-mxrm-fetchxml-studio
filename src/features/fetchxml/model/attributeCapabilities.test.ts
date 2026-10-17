/**
 * Attribute capability matrix (#90b). Each row is what Dataverse did for an Advanced Find attribute
 * of that type (912-query live matrix + org-wide scan, 2026-10-13); see the design doc §9A.
 */
import { describe, expect, it } from "vitest";
import {
	getAttributeCapabilities,
	getAttributeMetadataIssues,
	getDataSourceIssue,
	getOrderMetadataIssue,
	getUnsupportedFunctionReason,
	getValueOfCompatibleTypes,
	toAttributeTypeInfo,
} from "./attributeCapabilities";

const ALL = ["count", "countcolumn", "sum", "avg", "min", "max"];

describe("getAttributeCapabilities", () => {
	it.each([
		// typeName, isLogical, order, groupby, functions
		["StringType", false, true, true, ["count", "countcolumn"]],
		["MemoType", false, true, true, ["count", "countcolumn"]],
		["MemoType", true, true, false, []], // composite: address1_composite
		["IntegerType", false, true, true, ALL],
		["BigIntType", false, true, true, ALL],
		["DecimalType", false, true, true, ALL],
		["DoubleType", false, true, true, ALL],
		["MoneyType", false, true, true, ["sum", "avg", "min", "max"]],
		["DateTimeType", false, true, true, ["count", "countcolumn", "min", "max"]],
		["BooleanType", false, true, true, ["count", "countcolumn"]],
		["PicklistType", false, true, true, ["count", "countcolumn"]],
		["StateType", false, true, true, ["count", "countcolumn"]],
		["StatusType", false, true, true, ["count", "countcolumn"]],
		["LookupType", false, true, true, ["count", "countcolumn"]],
		["CustomerType", false, true, true, ["count", "countcolumn"]],
		["OwnerType", false, true, true, ["count", "countcolumn"]],
		["UniqueidentifierType", false, true, true, ["count", "countcolumn"]],
		["MultiSelectPicklistType", false, false, false, []],
		["ImageType", false, false, false, []],
	])("%s (logical=%s)", (typeName, isLogical, order, groupby, functions) => {
		const caps = getAttributeCapabilities({ typeName, isLogical });
		expect(caps.order).toBe(order);
		expect(caps.groupby).toBe(groupby);
		expect([...caps.functions].sort()).toEqual([...functions].sort());
		if (!order || !groupby || functions.length === 0) expect(caps.reason).toBeTruthy();
	});

	it("DateTime group by needs a date grouping", () => {
		expect(getAttributeCapabilities({ typeName: "DateTimeType" }).groupbyNeedsDateGrouping).toBe(true);
		expect(getAttributeCapabilities({ typeName: "IntegerType" }).groupbyNeedsDateGrouping).toBe(false);
	});

	it("reads AttributeTypeName, falling back to AttributeType", () => {
		expect(toAttributeTypeInfo({ AttributeType: "Virtual", AttributeTypeName: { Value: "MultiSelectPicklistType" } }).typeName).toBe("MultiSelectPicklistType");
		expect(toAttributeTypeInfo({ AttributeType: "Money" }).typeName).toBe("MoneyType");
	});
});

describe("editor messages", () => {
	it("explains why Money can't be counted", () => {
		expect(getUnsupportedFunctionReason("count", { typeName: "MoneyType" })).toMatch(/0x80060888/);
		expect(getUnsupportedFunctionReason("sum", { typeName: "MoneyType" })).toBeUndefined();
	});

	it("flags unsupported aggregates, group by and missing date grouping in aggregate queries", () => {
		const ctx = { aggregateQuery: true };
		expect(getAttributeMetadataIssues({ id: "1", type: "attribute", name: "revenue", alias: "r", aggregate: "count" }, { typeName: "MoneyType" }, ctx)).toHaveLength(1);
		expect(getAttributeMetadataIssues({ id: "1", type: "attribute", name: "name", alias: "n", aggregate: "sum" }, { typeName: "StringType" }, ctx)[0]).toMatch(/0x8004112f/);
		expect(getAttributeMetadataIssues({ id: "1", type: "attribute", name: "createdon", alias: "c", groupby: true }, { typeName: "DateTimeType" }, ctx)[0]).toMatch(/0x80041135/);
		expect(getAttributeMetadataIssues({ id: "1", type: "attribute", name: "createdon", alias: "c", groupby: true, dategrouping: "year" }, { typeName: "DateTimeType" }, ctx)).toEqual([]);
	});

	it("doesn't block when metadata is unknown", () => {
		expect(getAttributeMetadataIssues({ id: "1", type: "attribute", name: "x", alias: "x", aggregate: "sum" }, undefined, { aggregateQuery: true })).toEqual([]);
		expect(getOrderMetadataIssue(undefined)).toBeUndefined();
	});

	it("count child records needs a hierarchical entity and the primary key", () => {
		const attr = { id: "1", type: "attribute" as const, name: "accountid", alias: "k", rowAggregate: "CountChildren" };
		expect(getAttributeMetadataIssues(attr, undefined, { aggregateQuery: false, isHierarchical: false })[0]).toMatch(/0x80047020/);
		expect(getAttributeMetadataIssues({ ...attr, name: "parentaccountid" }, undefined, { aggregateQuery: false, isHierarchical: true, primaryIdAttribute: "accountid" })[0]).toMatch(/0x80040203/);
		expect(getAttributeMetadataIssues(attr, undefined, { aggregateQuery: false, isHierarchical: true, primaryIdAttribute: "accountid" })).toEqual([]);
	});

	it("orders on multi-select option sets and retained data without retention are flagged", () => {
		expect(getOrderMetadataIssue({ typeName: "MultiSelectPicklistType" })).toBeTruthy();
		expect(getOrderMetadataIssue({ typeName: "PicklistType" })).toBeUndefined();
		expect(getDataSourceIssue("retained", { logicalName: "account", isRetentionEnabled: false })).toMatch(/IsRetentionEnabled/);
		expect(getDataSourceIssue("retained", { logicalName: "account", isRetentionEnabled: true })).toBeUndefined();
		expect(getDataSourceIssue("retained", { logicalName: "account" })).toBeUndefined();
	});
});

describe("getValueOfCompatibleTypes (valueof type matrix run live, 2026-10-15)", () => {
	it.each([
		// accepted by Dataverse
		["Integer", "Picklist", true],
		["State", "Status", true],
		["EntityName", "Picklist", true],
		["Money", "Decimal", true],
		["Lookup", "Uniqueidentifier", true],
		["Customer", "Owner", true],
		["Memo", "String", true],
		["DateTime", "DateTime", true],
		["EntityName", "Status", true],
		["ManagedProperty", "Boolean", true],
		["BigInt", "BigInt", true],
		["Double", "Double", true],
		// rejected with 0x80040203
		["Integer", "BigInt", false],
		["Double", "Decimal", false],
		["Boolean", "Integer", false],
		["String", "Uniqueidentifier", false],
		["String", "State", false],
		["Integer", "Decimal", false],
		["Integer", "Double", false],
	])("%s compared to %s: %s", (a, b, ok) => {
		expect(getValueOfCompatibleTypes(a).includes(b)).toBe(ok);
		expect(getValueOfCompatibleTypes(b).includes(a)).toBe(ok);
	});
	it("multi-select and unverified types are never offered", () => {
		expect(getValueOfCompatibleTypes("Virtual")).toEqual([]);
		expect(getValueOfCompatibleTypes("PartyList")).toEqual([]);
		expect(getValueOfCompatibleTypes(undefined)).toEqual([]);
	});
});

describe("label attributes (VirtualType, shown with Advanced Find only off, #57)", () => {
	it("sort, but can't be grouped or aggregated (verified: 0x8004e01d / 0x80044150)", () => {
		const caps = getAttributeCapabilities({ typeName: "VirtualType" });
		expect(caps).toMatchObject({ order: true, groupby: false, functions: [] });
		expect(caps.reason).toBeTruthy();
	});

	it("images and files still can't be sorted, grouped or aggregated", () => {
		expect(getAttributeCapabilities({ typeName: "ImageType" }).order).toBe(false);
		expect(getAttributeCapabilities({ typeName: "FileType" }).order).toBe(false);
	});
});
