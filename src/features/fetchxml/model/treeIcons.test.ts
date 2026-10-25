import { describe, expect, it } from "vitest";
import type { AttributeMetadata } from "../api/pptbClient";
import {
	ATTRIBUTE_TYPE_ICONS,
	NODE_ICONS,
	UNKNOWN_ATTRIBUTE_ICON,
	attributeIcon,
	normalizeTreeIconStyle,
	sourceTypeBadge,
} from "./treeIcons";

const attr = (typeName: string | undefined, extra: Partial<AttributeMetadata> = {}): AttributeMetadata =>
	({
		LogicalName: "x",
		SchemaName: "X",
		MetadataId: "00000000-0000-0000-0000-000000000000",
		AttributeType: "String",
		AttributeTypeName: typeName ? { Value: typeName } : undefined,
		...extra,
	}) as AttributeMetadata;

// Every AttributeTypeDisplayName on Microsoft Learn's "Column definitions" page
const LEARN_TYPES = [
	"BooleanType", "EntityNameType", "MultiSelectPicklistType", "PicklistType", "StateType", "StatusType",
	"CalendarRulesType", "PartyListType", "DateTimeType", "FileType", "ImageType", "ManagedPropertyType",
	"BigIntType", "DecimalType", "DoubleType", "IntegerType", "MoneyType", "CustomerType", "LookupType",
	"OwnerType", "MemoType", "StringType", "UniqueidentifierType", "VirtualType",
];

describe("tree icons (#230)", () => {
	it("has an icon for every column type Microsoft Learn lists", () => {
		for (const t of LEARN_TYPES) {
			expect(ATTRIBUTE_TYPE_ICONS[t], t).toBeDefined();
			expect(attributeIcon(attr(t)), t).not.toBe(UNKNOWN_ATTRIBUTE_ICON);
		}
		expect(Object.keys(ATTRIBUTE_TYPE_ICONS).sort()).toEqual([...LEARN_TYPES].sort());
	});

	it("maps the common types to their icons", () => {
		expect(attributeIcon(attr("StringType"))).toBe("TextT");
		expect(attributeIcon(attr("MoneyType"))).toBe("Money");
		expect(attributeIcon(attr("LookupType"))).toBe("TableSearch");
		expect(attributeIcon(attr("PicklistType"))).toBe("RadioButton");
		expect(attributeIcon(attr("MultiSelectPicklistType"))).toBe("MultiselectLtr");
		expect(attributeIcon(attr("BooleanType"))).toBe("ToggleLeft");
		expect(attributeIcon(attr("DateTimeType"))).toBe("CalendarClock");
	});

	it("gives the primary key a key, and other unique identifiers two keys", () => {
		expect(attributeIcon(attr("UniqueidentifierType", { IsPrimaryId: true }))).toBe("Key");
		expect(attributeIcon(attr("UniqueidentifierType"))).toBe("KeyMultiple");
	});

	it("falls back while metadata loads or for an unknown type", () => {
		expect(attributeIcon(undefined)).toBe(UNKNOWN_ATTRIBUTE_ICON);
		expect(attributeIcon(attr("SomeFutureType", { AttributeType: "SomeFuture" }))).toBe(UNKNOWN_ATTRIBUTE_ICON);
	});

	it("uses AttributeType when AttributeTypeName is missing", () => {
		expect(attributeIcon(attr(undefined, { AttributeType: "Money" }))).toBe("Money");
	});

	it("badges calculated, rollup, formula and prompt columns", () => {
		expect(sourceTypeBadge(attr("MoneyType", { SourceType: 0 }))).toBeUndefined();
		expect(sourceTypeBadge(attr("MoneyType", { SourceType: 1 }))).toBe("Calculated");
		expect(sourceTypeBadge(attr("MoneyType", { SourceType: 2 }))).toBe("Rollup");
		expect(sourceTypeBadge(attr("StringType", { SourceType: 3 }))).toBe("Formula");
		expect(sourceTypeBadge(attr("StringType", { SourceType: 4 }))).toBe("Prompt");
		expect(sourceTypeBadge(attr("StringType", { SourceType: null }))).toBeUndefined();
		expect(sourceTypeBadge(undefined)).toBeUndefined();
	});

	it("colours nodes by role, and every node type has its own icon", () => {
		expect(NODE_ICONS.entity.role).toBe(NODE_ICONS.linkEntity.role);
		expect(NODE_ICONS.filter.role).toBe(NODE_ICONS.condition.role);
		const icons = Object.values(NODE_ICONS).map((n) => n.icon);
		expect(new Set(icons).size).toBe(icons.length);
		// Fetch and Entity, and Filter and Condition, used to share an icon
		expect(NODE_ICONS.fetch.icon).not.toBe(NODE_ICONS.entity.icon);
		expect(NODE_ICONS.filter.icon).not.toBe(NODE_ICONS.condition.icon);
	});

	it("defaults the icon style to tinted", () => {
		expect(normalizeTreeIconStyle("color")).toBe("color");
		for (const v of ["tinted", "tiles", undefined, null, 1]) expect(normalizeTreeIconStyle(v)).toBe("tinted");
	});
});
