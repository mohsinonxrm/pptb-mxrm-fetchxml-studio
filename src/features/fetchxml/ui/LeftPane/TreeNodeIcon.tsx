/**
 * Draws a query tree node's icon (#230): the role colour from Fluent palette tokens, in the style the
 * user chose in Settings.
 * - Tinted (default): the Regular icon in the role colour. One solid colour per system icon, as
 *   Fluent's iconography asks; it follows Windows high contrast because it draws with currentColor.
 * - Color: Fluent's multicolour Color icon where one exists, otherwise the Filled icon in the role
 *   colour. Color icons don't follow Windows high contrast, so under forced-colors the Filled icon
 *   shows instead, and each one gets its own idPrefix so the SVG ids don't collide (both as the
 *   @fluentui/react-icons README advises; it recommends avoiding Color icons, so they're opt-in).
 */
import { useId } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import {
	Code20Regular,
	Code20Filled,
	Code20Color,
	Table20Regular,
	Table20Filled,
	Table20Color,
	TableLink20Regular,
	TableLink20Filled,
	ColumnTriple20Regular,
	ColumnTriple20Filled,
	Filter20Regular,
	Filter20Filled,
	MathSymbols20Regular,
	MathSymbols20Filled,
	ArrowSortUpLines20Regular,
	ArrowSortUpLines20Filled,
	ArrowSortDownLines20Regular,
	ArrowSortDownLines20Filled,
	TextT20Regular,
	TextT20Filled,
	TextAlignLeft20Regular,
	TextAlignLeft20Filled,
	NumberSymbol20Regular,
	NumberSymbol20Filled,
	DecimalArrowRight20Regular,
	DecimalArrowRight20Filled,
	Money20Regular,
	Money20Filled,
	ToggleLeft20Regular,
	ToggleLeft20Filled,
	CalendarClock20Regular,
	CalendarClock20Filled,
	CalendarClock20Color,
	RadioButton20Regular,
	RadioButton20Filled,
	MultiselectLtr20Regular,
	MultiselectLtr20Filled,
	Status20Regular,
	Status20Filled,
	TagCircle20Regular,
	TagCircle20Filled,
	TableSearch20Regular,
	TableSearch20Filled,
	BuildingPeople20Regular,
	BuildingPeople20Filled,
	BuildingPeople20Color,
	PeopleTeam20Regular,
	PeopleTeam20Filled,
	PeopleTeam20Color,
	PeopleList20Regular,
	PeopleList20Filled,
	PeopleList20Color,
	Key20Regular,
	Key20Filled,
	KeyMultiple20Regular,
	KeyMultiple20Filled,
	Tag20Regular,
	Tag20Filled,
	Image20Regular,
	Image20Filled,
	Image20Color,
	Attach20Regular,
	Attach20Filled,
	ShieldLock20Regular,
	ShieldLock20Filled,
	CalendarSettings20Regular,
	CalendarSettings20Filled,
	Shapes20Regular,
	Shapes20Filled,
	TextColumnOne20Regular,
	TextColumnOne20Filled,
	type FluentIcon,
} from "@fluentui/react-icons";
import type { NodeRole, TreeIconName, TreeIconStyle } from "../../model/treeIcons";

const ICONS: Record<TreeIconName, { regular: FluentIcon; filled: FluentIcon; color?: FluentIcon }> = {
	Code: { regular: Code20Regular, filled: Code20Filled, color: Code20Color },
	Table: { regular: Table20Regular, filled: Table20Filled, color: Table20Color },
	TableLink: { regular: TableLink20Regular, filled: TableLink20Filled },
	ColumnTriple: { regular: ColumnTriple20Regular, filled: ColumnTriple20Filled },
	Filter: { regular: Filter20Regular, filled: Filter20Filled },
	MathSymbols: { regular: MathSymbols20Regular, filled: MathSymbols20Filled },
	ArrowSortUpLines: { regular: ArrowSortUpLines20Regular, filled: ArrowSortUpLines20Filled },
	ArrowSortDownLines: { regular: ArrowSortDownLines20Regular, filled: ArrowSortDownLines20Filled },
	TextT: { regular: TextT20Regular, filled: TextT20Filled },
	TextAlignLeft: { regular: TextAlignLeft20Regular, filled: TextAlignLeft20Filled },
	NumberSymbol: { regular: NumberSymbol20Regular, filled: NumberSymbol20Filled },
	DecimalArrowRight: { regular: DecimalArrowRight20Regular, filled: DecimalArrowRight20Filled },
	Money: { regular: Money20Regular, filled: Money20Filled },
	ToggleLeft: { regular: ToggleLeft20Regular, filled: ToggleLeft20Filled },
	CalendarClock: { regular: CalendarClock20Regular, filled: CalendarClock20Filled, color: CalendarClock20Color },
	RadioButton: { regular: RadioButton20Regular, filled: RadioButton20Filled },
	MultiselectLtr: { regular: MultiselectLtr20Regular, filled: MultiselectLtr20Filled },
	Status: { regular: Status20Regular, filled: Status20Filled },
	TagCircle: { regular: TagCircle20Regular, filled: TagCircle20Filled },
	TableSearch: { regular: TableSearch20Regular, filled: TableSearch20Filled },
	BuildingPeople: { regular: BuildingPeople20Regular, filled: BuildingPeople20Filled, color: BuildingPeople20Color },
	PeopleTeam: { regular: PeopleTeam20Regular, filled: PeopleTeam20Filled, color: PeopleTeam20Color },
	PeopleList: { regular: PeopleList20Regular, filled: PeopleList20Filled, color: PeopleList20Color },
	Key: { regular: Key20Regular, filled: Key20Filled },
	KeyMultiple: { regular: KeyMultiple20Regular, filled: KeyMultiple20Filled },
	Tag: { regular: Tag20Regular, filled: Tag20Filled },
	Image: { regular: Image20Regular, filled: Image20Filled, color: Image20Color },
	Attach: { regular: Attach20Regular, filled: Attach20Filled },
	ShieldLock: { regular: ShieldLock20Regular, filled: ShieldLock20Filled },
	CalendarSettings: { regular: CalendarSettings20Regular, filled: CalendarSettings20Filled },
	Shapes: { regular: Shapes20Regular, filled: Shapes20Filled },
	TextColumnOne: { regular: TextColumnOne20Regular, filled: TextColumnOne20Filled },
};

/** Role colours: Fluent palette tokens, with tuned light, dark and high-contrast values */
const ROLE_COLOR: Record<NodeRole, string> = {
	query: tokens.colorBrandForeground1,
	table: tokens.colorPaletteBlueForeground2,
	column: tokens.colorPaletteTealForeground2,
	criteria: tokens.colorPaletteDarkOrangeForeground2,
	sort: tokens.colorPaletteGrapeForeground2,
};

const useStyles = makeStyles({
	colorWrap: { display: "inline-flex" },
	colorOnly: { "@media (forced-colors: active)": { display: "none" } },
	highContrastOnly: { display: "none", "@media (forced-colors: active)": { display: "inline" } },
});

export function TreeNodeIcon({ role, icon, iconStyle }: { role: NodeRole; icon: TreeIconName; iconStyle: TreeIconStyle }) {
	const styles = useStyles();
	const id = useId();
	const { regular: Regular, filled: Filled, color: Color } = ICONS[icon];
	const color = ROLE_COLOR[role];
	if (iconStyle === "tinted") return <Regular style={{ color }} aria-hidden />;
	if (!Color) return <Filled style={{ color }} aria-hidden />;
	return (
		<span className={styles.colorWrap} aria-hidden>
			<Color idPrefix={`${id}-`} className={styles.colorOnly} />
			<Filled className={styles.highContrastOnly} />
		</span>
	);
}
