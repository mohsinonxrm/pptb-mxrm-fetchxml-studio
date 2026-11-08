/**
 * The shared overflow command bar (#232): commands drop into "More commands" by priority, More keeps
 * their order, reasons and submenus, pinned and end commands stay, and focus follows an overflowed
 * command to More. jsdom has no layout, so widths come from stubOverflowWidths (ui.ts).
 */
import { act, createElement, useState } from "react";
import {
	Dialog,
	DialogBody,
	DialogSurface,
	DialogTitle,
	MenuItem,
	MenuList,
} from "@fluentui/react-components";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../testing/ui";
import { OverflowToolbar, type OverflowToolbarProps } from "./OverflowToolbar";
import type { BarCommand } from "./types";

// Every item is 80 px, More 40 px, a divider 24 px; the medium toolbar keeps 24 px free
const PADDING = 24;
/** A container wide enough for `px` of items */
const fits = (px: number) => PADDING + px;

const button = (id: string, priority: number, extra: Partial<BarCommand> = {}): BarCommand =>
	({ id, label: id, kind: "button", priority, onClick: () => {}, ...extra }) as BarCommand;

const renderBar = (commands: BarCommand[], extra: Partial<OverflowToolbarProps> = {}) =>
	render(createElement(OverflowToolbar, { ariaLabel: "Test commands", commands, ...extra }));

const toolbar = () => document.querySelector<HTMLElement>('[role="toolbar"]');
/** A bar control by its visible text (a SplitButton's primary action for a split) */
const onBar = (label: string) =>
	[...(toolbar()?.querySelectorAll<HTMLElement>("button") ?? [])].find(
		(b) => b.textContent?.trim() === label,
	);
const overflowed = (element: Element | undefined | null) =>
	!!element?.closest("[data-overflowing]");
const visibleOnBar = (labels: string[]) => labels.filter((l) => !overflowed(onBar(l)));
const more = () =>
	document.querySelector<HTMLButtonElement>(
		'[role="toolbar"] button[aria-haspopup="menu"][aria-label="More commands"]',
	);
/** The open menus, outermost first (Fluent portals each one into the body) */
const menus = () => [...document.querySelectorAll<HTMLElement>('[role="menu"]')];
const itemsOf = (menu: HTMLElement | undefined) => [
	...(menu?.querySelectorAll<HTMLElement>('[role^="menuitem"]') ?? []),
];
const textsOf = (menu: HTMLElement | undefined) => itemsOf(menu).map((i) => i.textContent?.trim());
const item = (menu: HTMLElement | undefined, label: string) =>
	itemsOf(menu).find((i) => i.textContent?.trim() === label);
/** The text of the element an aria-describedby points at */
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};
/** Shown on screen: not overflowed, and no display:none from Griffel's classes (jsdom applies them) */
const shown = (element: Element | null | undefined) =>
	!!element && !overflowed(element) && getComputedStyle(element).display !== "none";

beforeEach(() => {
	// Fluent's positioning warns in jsdom
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("OverflowToolbar layout", () => {
	it("shows no More when everything fits", async () => {
		stubOverflowWidths({ container: fits(4 * 80) });
		await renderBar([button("A", 4), button("B", 3), button("C", 2), button("D", 1)]);
		expect(visibleOnBar(["A", "B", "C", "D"])).toEqual(["A", "B", "C", "D"]);
		expect(more()).toBeNull();
	});

	it("narrowing hides the lowest priority first, and More lists the hidden ones in bar order", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		const commands = [
			button("Open", 50),
			button("Delete", 20),
			button("Copy URL", 30),
			button("Export", 10),
		];
		await renderBar(commands);
		expect(more()).toBeNull();

		// Room for 3 items and More: Export (10) goes first
		await layout.resize({ container: fits(3 * 80 + 40) });
		expect(visibleOnBar(["Open", "Delete", "Copy URL", "Export"])).toEqual([
			"Open",
			"Delete",
			"Copy URL",
		]);

		// Room for 2 and More: then Delete (20)
		await layout.resize({ container: fits(2 * 80 + 40) });
		expect(visibleOnBar(["Open", "Delete", "Copy URL", "Export"])).toEqual([
			"Open",
			"Copy URL",
		]);

		await click(more());
		// Bar order (Delete before Export), not priority order
		expect(textsOf(menus()[0])).toEqual(["Delete", "Export"]);

		// Wider again: everything comes back and More goes
		await layout.resize({ container: 1000 });
		expect(visibleOnBar(["Open", "Delete", "Copy URL", "Export"])).toHaveLength(4);
		expect(more()).toBeNull();
	});

	it("never overflows pinned and end commands", async () => {
		stubOverflowWidths({ container: 10 });
		await renderBar([
			button("Run", 0, { pinned: true, priority: undefined, appearance: "primary" }),
			button("Save", 60),
			button("Send to", 20),
			{
				id: "settings",
				label: "Settings",
				kind: "button",
				iconOnly: true,
				placement: "end",
				onClick: () => {},
			},
		]);
		expect(overflowed(onBar("Run"))).toBe(false);
		expect(overflowed(onBar("Save"))).toBe(true);
		expect(overflowed(onBar("Send to"))).toBe(true);
		const settings = toolbar()?.querySelector('button[aria-label="Settings"]');
		expect(settings).not.toBeNull();
		expect(overflowed(settings)).toBe(false);
		// Settings sits after More
		const order = [...(toolbar()?.querySelectorAll("button") ?? [])];
		expect(order.indexOf(settings as HTMLButtonElement)).toBeGreaterThan(
			order.indexOf(more() as HTMLButtonElement),
		);

		await click(more());
		expect(textsOf(menus()[0])).toEqual(["Save", "Send to"]);
	});

	it("hides a divider with its group, whichever side overflows", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		const commands = [
			button("A1", 50, { group: "a" }),
			button("A2", 50, { group: "a" }),
			button("B1", 10, { group: "b" }),
			button("B2", 10, { group: "b" }),
		];
		await renderBar(commands);
		const divider = () => toolbar()?.querySelector('[role="separator"]');
		expect(shown(divider())).toBe(true);

		// Group b (lower priority, after the divider) all in More
		await layout.resize({ container: fits(2 * 80 + 24 + 40) });
		expect(visibleOnBar(["A1", "A2", "B1", "B2"])).toEqual(["A1", "A2"]);
		expect(shown(divider())).toBe(false);

		// Now group a (before the divider) is the lower priority one
		cleanup();
		const layout2 = stubOverflowWidths({ container: 1000 });
		await renderBar([
			button("A1", 10, { group: "a" }),
			button("A2", 10, { group: "a" }),
			button("B1", 50, { group: "b" }),
			button("B2", 50, { group: "b" }),
		]);
		await layout2.resize({ container: fits(2 * 80 + 24 + 40) });
		expect(visibleOnBar(["A1", "A2", "B1", "B2"])).toEqual(["B1", "B2"]);
		expect(shown(divider())).toBe(false);
	});

	it("separates groups in More with a divider", async () => {
		// Room for Open, the divider after its group, and More
		stubOverflowWidths({ container: fits(80 + 24 + 40) });
		await renderBar([
			button("Open", 50, { group: "a" }),
			button("Copy", 20, { group: "a" }),
			button("Delete", 10, { group: "b", danger: true }),
		]);
		await click(more());
		const menu = menus()[0];
		expect(textsOf(menu)).toEqual(["Copy", "Delete"]);
		// Copy, divider, Delete
		// (Tabster's focus dummies aside)
		const children = [...(menu?.children ?? [])].filter(
			(c) => !c.hasAttribute("data-tabster-dummy"),
		);
		expect(
			children.map((c) =>
				(c.getAttribute("role") ?? "").startsWith("menuitem")
					? c.textContent?.trim()
					: c.classList.contains("fui-MenuDivider")
						? "|"
						: "?",
			),
		).toEqual(["Copy", "|", "Delete"]);
	});
});

describe("OverflowToolbar commands in More", () => {
	it("keeps a disabled command focusable with its reason, on the bar and in More", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		const onDelete = vi.fn();
		await renderBar([
			button("Open", 50),
			button("Delete", 10, {
				onClick: onDelete,
				disabledReason: "Select a row to delete it",
			}),
		]);

		const onBarDelete = onBar("Delete");
		expect(onBarDelete?.getAttribute("aria-disabled")).toBe("true");
		expect(onBarDelete?.hasAttribute("disabled")).toBe(false);
		expect(description(onBarDelete)).toBe("Select a row to delete it");
		await click(onBarDelete);
		expect(onDelete).not.toHaveBeenCalled();

		await layout.resize({ container: fits(80 + 40) });
		await click(more());
		const inMore = item(menus()[0], "Delete");
		expect(inMore?.getAttribute("aria-disabled")).toBe("true");
		// Still reachable with the arrow keys
		expect(inMore?.getAttribute("tabindex")).not.toBeNull();
		expect(description(inMore)).toBe("Select a row to delete it");
		await click(inMore);
		expect(onDelete).not.toHaveBeenCalled();
	});

	it("labels an icon-only disabled command 'Label: reason'", async () => {
		stubOverflowWidths({ container: 1000 });
		await renderBar([
			{
				id: "copy",
				label: "Copy",
				kind: "button",
				iconOnly: true,
				priority: 1,
				disabledReason: "There is no LayoutXML to copy",
			},
		]);
		const copy = toolbar()?.querySelector(
			'button[aria-label="Copy: There is no LayoutXML to copy"]',
		);
		expect(copy?.getAttribute("aria-disabled")).toBe("true");
	});

	it("runs an enabled command from More", async () => {
		stubOverflowWidths({ container: fits(80 + 40) });
		const onExport = vi.fn();
		await renderBar([button("Open", 50), button("Export", 10, { onClick: onExport })]);
		await click(more());
		await click(item(menus()[0], "Export"));
		expect(onExport).toHaveBeenCalledTimes(1);
	});

	it("turns a menu into a submenu, and a split into a submenu with its primary action first", async () => {
		stubOverflowWidths({ container: fits(80 + 40) });
		const onSave = vi.fn();
		const onExcel = vi.fn();
		const onMenuOpenChange = vi.fn();
		await renderBar([
			button("Open", 50),
			{
				id: "export",
				label: "Export",
				kind: "menu",
				priority: 10,
				onMenuOpenChange,
				menu: createElement(
					MenuList,
					null,
					createElement(MenuItem, { onClick: onExcel }, "Excel"),
					createElement(MenuItem, null, "CSV"),
				),
			},
			{
				id: "save",
				label: "Save",
				kind: "split",
				priority: 10,
				onClick: onSave,
				menu: createElement(MenuList, null, createElement(MenuItem, null, "Save as…")),
			},
		]);
		expect(overflowed(onBar("Export"))).toBe(true);
		expect(overflowed(onBar("Save"))).toBe(true);

		await click(more());
		const moreMenu = menus()[0];
		expect(textsOf(moreMenu)).toEqual(["Export", "Save"]);
		expect(item(moreMenu, "Export")?.getAttribute("aria-haspopup")).toBe("menu");
		expect(item(moreMenu, "Save")?.getAttribute("aria-haspopup")).toBe("menu");

		await click(item(moreMenu, "Export"));
		expect(onMenuOpenChange).toHaveBeenLastCalledWith(true);
		expect(textsOf(menus()[1])).toEqual(["Excel", "CSV"]);
		await click(item(menus()[1], "Excel"));
		expect(onExcel).toHaveBeenCalledTimes(1);

		await click(more());
		await click(item(menus()[0], "Save"));
		const saveMenu = menus()[menus().length - 1];
		expect(textsOf(saveMenu)).toEqual(["Save", "Save as…"]);
		await click(item(saveMenu, "Save"));
		expect(onSave).toHaveBeenCalledTimes(1);
	});

	it("turns a toggle into a checkbox item and a radio into a radio item", async () => {
		stubOverflowWidths({ container: fits(80 + 40) });
		const onWrap = vi.fn();
		const onWarnings = vi.fn();
		await renderBar([
			button("Check", 50),
			{
				id: "wrap",
				label: "Wrap",
				kind: "toggle",
				priority: 10,
				checked: false,
				onClick: onWrap,
			},
			{
				id: "all",
				label: "All",
				kind: "toggle",
				radioName: "show",
				priority: 5,
				checked: true,
				onClick: () => {},
			},
			{
				id: "warnings",
				label: "Warnings",
				kind: "toggle",
				radioName: "show",
				priority: 5,
				checked: false,
				onClick: onWarnings,
			},
		]);
		await click(more());
		const menu = menus()[0];
		expect(item(menu, "Wrap")?.getAttribute("role")).toBe("menuitemcheckbox");
		expect(item(menu, "Wrap")?.getAttribute("aria-checked")).toBe("false");
		expect(item(menu, "All")?.getAttribute("role")).toBe("menuitemradio");
		expect(item(menu, "All")?.getAttribute("aria-checked")).toBe("true");
		await click(item(menu, "Warnings"));
		expect(onWarnings).toHaveBeenCalledTimes(1);
		// A radio closes the menu; open it again
		await click(more());
		await click(item(menus()[0], "Wrap"));
		expect(onWrap).toHaveBeenCalledTimes(1);
	});
});

describe("OverflowToolbar More button", () => {
	it("is an icon-only ToolbarButton inside the toolbar, named 'More commands'", async () => {
		stubOverflowWidths({ container: fits(80 + 40) });
		await renderBar([button("Open", 50), button("Export", 10)]);
		const moreButton = more();
		expect(moreButton).not.toBeNull();
		expect(toolbar()?.contains(moreButton as Node)).toBe(true);
		expect(moreButton?.textContent?.trim()).toBe("");
	});

	it("takes another name when the bar asks", async () => {
		stubOverflowWidths({ container: fits(80 + 40) });
		await renderBar([button("Open", 50), button("Export", 10)], { moreLabel: "More actions" });
		expect(toolbar()?.querySelector('button[aria-label="More actions"]')).not.toBeNull();
	});

	it("receives focus when the focused command overflows", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		await renderBar([button("Open", 50), button("Copy", 30), button("Export", 10)]);
		const exportButton = onBar("Export");
		await act(async () => exportButton?.focus());
		expect(document.activeElement).toBe(exportButton);

		await layout.resize({ container: fits(2 * 80 + 40) });
		expect(overflowed(exportButton)).toBe(true);
		expect(document.activeElement).toBe(more());

		// More already showing: the next command to overflow hands focus to it straight away
		const copyButton = onBar("Copy");
		await act(async () => copyButton?.focus());
		await layout.resize({ container: fits(80 + 40) });
		expect(overflowed(copyButton)).toBe(true);
		expect(document.activeElement).toBe(more());
	});

	it("leaves focus alone when a command that isn't focused overflows", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		await renderBar([button("Open", 50), button("Copy", 30), button("Export", 10)]);
		const openButton = onBar("Open");
		await act(async () => openButton?.focus());
		await layout.resize({ container: fits(2 * 80 + 40) });
		expect(document.activeElement).toBe(openButton);
	});

	it("doesn't take focus back after it left the toolbar for the page", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		await renderBar([button("Open", 50), button("Copy", 30), button("Export", 10)]);
		const exportButton = onBar("Export");
		await act(async () => exportButton?.focus());
		// A click on empty space: focus goes to the body, the button still showing
		await act(async () => exportButton?.blur());
		await layout.resize({ container: fits(2 * 80 + 40) });
		expect(overflowed(exportButton)).toBe(true);
		expect(document.activeElement).toBe(document.body);
	});
});

describe("OverflowToolbar commands that change", () => {
	it("still overflows a command after it gains a disabled reason", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		await renderBar([button("Open", 50), button("Delete", 10)]);
		// A Tooltip now wraps Delete, so it's a new element; Overflow must measure that one
		await renderBar([
			button("Open", 50),
			button("Delete", 10, { disabledReason: "Select a row to delete it" }),
		]);
		await layout.resize({ container: fits(80 + 40) });
		expect(overflowed(onBar("Delete"))).toBe(true);
		await click(more());
		expect(item(menus()[0], "Delete")?.getAttribute("aria-disabled")).toBe("true");

		// And back without the reason, wider then narrower again
		await renderBar([button("Open", 50), button("Delete", 10)]);
		await layout.resize({ container: 1000 });
		expect(overflowed(onBar("Delete"))).toBe(false);
		await layout.resize({ container: fits(80 + 40) });
		expect(overflowed(onBar("Delete"))).toBe(true);
	});

	it("keeps focus on a command that turns busy with a reason", async () => {
		stubOverflowWidths({ container: 1000 });
		const run = (extra: Partial<BarCommand> = {}) =>
			({
				...button("Run", 0),
				pinned: true,
				priority: undefined,
				appearance: "primary",
				...extra,
			}) as BarCommand;
		await renderBar([run(), button("Save", 10)]);
		await act(async () => onBar("Run")?.focus());
		expect(document.activeElement).toBe(onBar("Run"));

		await renderBar([
			run({ busy: true, disabledReason: "The query is running" }),
			button("Save", 10),
		]);
		const busyRun = onBar("Run");
		expect(document.activeElement).toBe(busyRun);
		expect(busyRun?.getAttribute("aria-disabled")).toBe("true");
		expect(description(busyRun)).toBe("The query is running");

		await renderBar([run(), button("Save", 10)]);
		expect(document.activeElement).toBe(onBar("Run"));
	});
});

describe("OverflowToolbar details", () => {
	it("keeps room for More when moreExtra makes it always show", async () => {
		// Room for exactly 2 items, but More is always there: only 1 fits beside it
		stubOverflowWidths({ container: fits(2 * 80) });
		await renderBar([button("Open", 50), button("Export", 10)], {
			moreExtra: createElement(MenuItem, null, "Settings…"),
		});
		expect(visibleOnBar(["Open", "Export"])).toEqual(["Open"]);
		await click(more());
		expect(textsOf(menus()[0])).toEqual(["Export", "Settings…"]);
	});

	it("doesn't pick a disabled radio", async () => {
		stubOverflowWidths({ container: 1000 });
		const onWarnings = vi.fn();
		await renderBar([
			{
				id: "all",
				label: "All",
				kind: "toggle",
				radioName: "show",
				priority: 5,
				checked: true,
				onClick: () => {},
			},
			{
				id: "warnings",
				label: "Warnings",
				kind: "toggle",
				radioName: "show",
				priority: 5,
				checked: false,
				onClick: onWarnings,
				disabledReason: "No warnings",
			},
		]);
		const warnings = onBar("Warnings");
		expect(warnings?.getAttribute("aria-disabled")).toBe("true");
		// ToolbarRadioButton reports the click to the Toolbar even when disabledFocusable
		await click(warnings);
		expect(onWarnings).not.toHaveBeenCalled();
	});

	it("keeps a toggle whose id is also a radio group's name apart from that group", async () => {
		stubOverflowWidths({ container: 1000 });
		const onToggle = vi.fn();
		const onRadio = vi.fn();
		await renderBar([
			{
				id: "show",
				label: "Show",
				kind: "toggle",
				priority: 5,
				checked: false,
				onClick: onToggle,
			},
			{
				id: "errors",
				label: "Errors",
				kind: "toggle",
				radioName: "show",
				priority: 5,
				checked: false,
				onClick: onRadio,
			},
		]);
		await click(onBar("Show"));
		expect(onToggle).toHaveBeenCalledTimes(1);
		expect(onRadio).not.toHaveBeenCalled();
		await click(onBar("Errors"));
		expect(onRadio).toHaveBeenCalledTimes(1);
	});

	it("names an icon-only command with its ariaLabel and keeps the reason", async () => {
		stubOverflowWidths({ container: 1000 });
		await renderBar([
			{
				id: "copy",
				label: "Copy",
				ariaLabel: "Copy LayoutXML",
				kind: "button",
				iconOnly: true,
				priority: 1,
				disabledReason: "There is no LayoutXML to copy",
			},
		]);
		expect(
			toolbar()?.querySelector(
				'button[aria-label="Copy LayoutXML: There is no LayoutXML to copy"]',
			),
		).not.toBeNull();
	});

	it("keeps one divider between a and c when group b is all in More", async () => {
		stubOverflowWidths({ container: fits(2 * 80 + 24 + 40) });
		await renderBar([
			button("A", 50, { group: "a" }),
			button("B", 10, { group: "b" }),
			button("C", 50, { group: "c" }),
		]);
		expect(visibleOnBar(["A", "B", "C"])).toEqual(["A", "C"]);
		const shownDividers = [...(toolbar()?.querySelectorAll('[role="separator"]') ?? [])].filter(
			shown,
		);
		expect(shownDividers).toHaveLength(1);
		// It sits between A and C
		const a = onBar("A") as HTMLElement;
		const c = onBar("C") as HTMLElement;
		expect(
			a.compareDocumentPosition(shownDividers[0]) & Node.DOCUMENT_POSITION_FOLLOWING,
		).toBeTruthy();
		expect(
			c.compareDocumentPosition(shownDividers[0]) & Node.DOCUMENT_POSITION_PRECEDING,
		).toBeTruthy();
	});
});

/** A bar whose "Export" command opens a modal Fluent Dialog, as the app's bar commands do */
function BarWithDialog({ barOnly }: { barOnly?: boolean }) {
	const [open, setOpen] = useState(false);
	const commands = [
		button("Open", 50),
		button("Export", barOnly ? 50 : 10, { onClick: () => setOpen(true) }),
	];
	return createElement(
		"div",
		null,
		createElement(OverflowToolbar, { ariaLabel: "Test commands", commands }),
		createElement(Dialog, {
			modalType: "modal",
			open,
			onOpenChange: (_: unknown, data: { open: boolean }) => setOpen(data.open),
			children: createElement(
				DialogSurface,
				null,
				createElement(DialogBody, null, createElement(DialogTitle, null, "Export dialog")),
			),
		}),
	);
}

const dialogOpen = () => !!document.querySelector('[role="dialog"]');
/** Escape on the focused element, then lets Fluent's dialog unmount and Tabster's async restore (a 0 ms timer) run */
async function closeWithEscape() {
	await act(async () => {
		(document.activeElement ?? document.body).dispatchEvent(
			new KeyboardEvent("keydown", { key: "Escape", bubbles: true }),
		);
	});
	await act(async () => {
		await new Promise((resolve) => setTimeout(resolve, 50));
	});
}

describe("OverflowToolbar focus return", () => {
	it("returns focus to a bar command after the dialog it opened closes", async () => {
		stubOverflowWidths({ container: 1000 });
		await render(createElement(BarWithDialog, { barOnly: true }));
		const exportButton = onBar("Export");
		await act(async () => exportButton?.focus());
		await click(exportButton);
		expect(dialogOpen()).toBe(true);
		expect(document.activeElement).not.toBe(exportButton);

		await closeWithEscape();
		expect(dialogOpen()).toBe(false);
		expect(document.activeElement).toBe(exportButton);
	});

	it("returns focus to More after a dialog opened from an overflowed command closes", async () => {
		stubOverflowWidths({ container: fits(80 + 40) });
		await render(createElement(BarWithDialog));
		expect(overflowed(onBar("Export"))).toBe(true);
		const moreButton = more();
		await act(async () => moreButton?.focus());
		await click(moreButton);
		await click(item(menus()[0], "Export"));
		expect(dialogOpen()).toBe(true);
		expect(menus()).toHaveLength(0);

		await closeWithEscape();
		expect(dialogOpen()).toBe(false);
		expect(document.activeElement).toBe(moreButton);
	});
});
