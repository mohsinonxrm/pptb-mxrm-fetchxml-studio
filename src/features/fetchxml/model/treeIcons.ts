/**
 * Which icon and colour each query tree node gets (#230). Colour says what the node does (five
 * roles); the icon says what it is. Attributes show their Dataverse column type, from the
 * AttributeTypeName FXS already loads. Pure, so it can be tested; TreeNodeIcon draws it.
 *
 * Column types: https://learn.microsoft.com/power-apps/developer/data-platform/entity-attribute-metadata#column-types
 * SourceType: https://learn.microsoft.com/power-apps/developer/data-platform/specialized-columns
 */
import type { AttributeMetadata } from "../api/pptbClient";

/** What a node does: decides its colour */
export type NodeRole = "query" | "table" | "column" | "criteria" | "sort";

/** The Fluent icon names (without size and style) the tree uses */
export type TreeIconName =
	| "Code"
	| "Table"
	| "TableLink"
	| "ColumnTriple"
	| "Filter"
	| "MathSymbols"
	| "ArrowSortUpLines"
	| "ArrowSortDownLines"
	| "TextT"
	| "TextAlignLeft"
	| "NumberSymbol"
	| "DecimalArrowRight"
	| "Money"
	| "ToggleLeft"
	| "CalendarClock"
	| "RadioButton"
	| "MultiselectLtr"
	| "Status"
	| "TagCircle"
	| "TableSearch"
	| "BuildingPeople"
	| "PeopleTeam"
	| "PeopleList"
	| "Key"
	| "KeyMultiple"
	| "Tag"
	| "Image"
	| "Attach"
	| "ShieldLock"
	| "CalendarSettings"
	| "Shapes"
	| "TextColumnOne";

export interface TreeIcon {
	role: NodeRole;
	icon: TreeIconName;
}

/** Icon per AttributeTypeName (every type Microsoft Learn lists) */
export const ATTRIBUTE_TYPE_ICONS: Record<string, TreeIconName> = {
	StringType: "TextT",
	MemoType: "TextAlignLeft",
	IntegerType: "NumberSymbol",
	BigIntType: "NumberSymbol",
	DecimalType: "DecimalArrowRight",
	DoubleType: "DecimalArrowRight",
	MoneyType: "Money",
	BooleanType: "ToggleLeft",
	DateTimeType: "CalendarClock",
	PicklistType: "RadioButton",
	MultiSelectPicklistType: "MultiselectLtr",
	StateType: "Status",
	StatusType: "TagCircle",
	LookupType: "TableSearch",
	CustomerType: "BuildingPeople",
	OwnerType: "PeopleTeam",
	PartyListType: "PeopleList",
	UniqueidentifierType: "KeyMultiple",
	EntityNameType: "Tag",
	ImageType: "Image",
	FileType: "Attach",
	ManagedPropertyType: "ShieldLock",
	CalendarRulesType: "CalendarSettings",
	VirtualType: "Shapes",
};

/** While metadata loads, or for a type not in the list */
export const UNKNOWN_ATTRIBUTE_ICON: TreeIconName = "TextColumnOne";

/** The attribute's column type icon; the primary key gets the key */
export function attributeIcon(meta: AttributeMetadata | undefined): TreeIconName {
	if (!meta) return UNKNOWN_ATTRIBUTE_ICON;
	if (meta.IsPrimaryId) return "Key";
	// AttributeTypeName is the current property; AttributeType + "Type" covers metadata without it
	const typeName = meta.AttributeTypeName?.Value ?? (meta.AttributeType ? `${meta.AttributeType}Type` : undefined);
	return (typeName && ATTRIBUTE_TYPE_ICONS[typeName]) || UNKNOWN_ATTRIBUTE_ICON;
}

/** A badge for calculated, rollup, formula and prompt columns (SourceType 1–4), or undefined */
export function sourceTypeBadge(meta: AttributeMetadata | undefined): string | undefined {
	switch (meta?.SourceType) {
		case 1:
			return "Calculated";
		case 2:
			return "Rollup";
		case 3:
			return "Formula";
		case 4:
			return "Prompt";
		default:
			return undefined;
	}
}

/** The fixed icons of the other node types */
export const NODE_ICONS = {
	fetch: { role: "query", icon: "Code" },
	entity: { role: "table", icon: "Table" },
	linkEntity: { role: "table", icon: "TableLink" },
	allAttributes: { role: "column", icon: "ColumnTriple" },
	filter: { role: "criteria", icon: "Filter" },
	condition: { role: "criteria", icon: "MathSymbols" },
	orderAscending: { role: "sort", icon: "ArrowSortUpLines" },
	orderDescending: { role: "sort", icon: "ArrowSortDownLines" },
} as const satisfies Record<string, TreeIcon>;

/** How tree icons are drawn: one-colour line icons (default) or Fluent's multicolour Color icons */
export type TreeIconStyle = "tinted" | "color";

export function normalizeTreeIconStyle(value: unknown): TreeIconStyle {
	return value === "color" ? "color" : "tinted";
}
