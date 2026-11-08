/**
 * Command and tab shapes for the shared overflow bars (#232): OverflowToolbar and OverflowTabList.
 * A bar describes its commands as data; the component renders each one on the bar and, once it
 * overflows, as the matching entry in the More menu, so both stay in step.
 */
import type { ReactElement, ReactNode } from "react";
import type { FluentIcon } from "@fluentui/react-icons";

/** How a command renders: on the bar, and in More once it overflows */
export type BarCommandKind =
	/** ToolbarButton → MenuItem */
	| "button"
	/** MenuButton → submenu MenuItem */
	| "menu"
	/** SplitButton → submenu whose first item is the primary action */
	| "split"
	/** ToolbarToggleButton → MenuItemCheckbox; with `radioName`, ToolbarRadioButton → MenuItemRadio */
	| "toggle";

interface BarCommandBase {
	/** Unique in the bar; also the OverflowItem id */
	id: string;
	/** Visible text (sentence case), and the accessible name unless `ariaLabel` is set */
	label: string;
	/** A 20 px bundleIcon (`bundleIcon(X20Filled, X20Regular)`) */
	icon?: FluentIcon;
	/** Commands in different groups get a divider between them, on the bar and in More */
	group?: string;
	/** "end": after More, never overflows (e.g. Settings) */
	placement?: "end";
	kind: BarCommandKind;
	/** Default "subtle". One primary per bar. */
	appearance?: "primary" | "secondary" | "subtle";
	/** Icon only on the bar, with a label Tooltip ("Label" or "Label: reason") */
	iconOnly?: boolean;
	/** Destructive: danger text on the bar and in More */
	danger?: boolean;
	/** Running: a spinner replaces the icon, and the command can't be pressed (but stays focusable) */
	busy?: boolean;
	/** button: the action. split: the primary action. toggle: called to flip `checked` (or pick this radio). */
	onClick?: () => void;
	/** menu / split: a `<MenuList>` body */
	menu?: ReactElement;
	/** menu / split: open state changes (e.g. to load the menu's items lazily) */
	onMenuOpenChange?: (open: boolean) => void;
	/** toggle: on or off (for a radio, whether this one is the chosen one) */
	checked?: boolean;
	/** toggle: makes it a radio; toggles sharing a radioName are one group */
	radioName?: string;
	/** Why it can't be pressed. Sets disabledFocusable and shows the reason in a Tooltip, on the bar and in More. */
	disabledReason?: string;
	/** A description Tooltip for an enabled, labelled command (e.g. its shortcut) */
	hint?: string;
	/** Overrides the accessible name (keep the visible label in it) */
	ariaLabel?: string;
}

/**
 * A command in an OverflowToolbar. `priority` (higher stays on the bar longer, default 0) and
 * `pinned` (never overflows) are mutually exclusive, as in Fluent's OverflowItemProps.
 */
export type BarCommand = BarCommandBase &
	({ priority?: number; pinned?: undefined } | { pinned: true; priority?: undefined });

/** A tab in an OverflowTabList */
export interface BarTab {
	value: string;
	/** Visible text, and the More menu entry */
	label: string;
	/** A bundleIcon sized to the tab list (20 px for medium, 16 px for small) */
	icon?: FluentIcon;
	/** Extra content after the label inside the tab, such as a count badge */
	content?: ReactNode;
}
