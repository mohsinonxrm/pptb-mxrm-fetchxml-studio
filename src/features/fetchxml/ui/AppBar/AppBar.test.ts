/**
 * The app bar (#278): Run | Save · Send to · Return FetchXML … More | Settings on the shared
 * overflow bar. Run is first, the one primary button, and never overflows; nor does Settings. More
 * always has Copy FetchXML. Unavailable commands keep their reason. jsdom has no layout, so widths
 * come from stubOverflowWidths.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	cleanup,
	click,
	flush,
	installHost,
	removeHost,
	render,
	stubOverflowWidths,
	toasts,
} from "../../../../shared/testing/ui";
import { AppBar, type AppBarProps } from "./AppBar";
import { NO_QUERY_REASON } from "../Toolbar/useSaveViewCommand";
import { EditorSyncContext } from "../RightPane/editorSync";

const XML = "<fetch><entity name='account'/></fetch>";
const tools = [{ id: "other", name: "Other" }];

const run = (extra: Partial<AppBarProps["run"]> = {}) => ({
	onClick: vi.fn(),
	busy: false,
	...extra,
});

const saveView = (extra: Partial<AppBarProps["saveView"]> = {}): AppBarProps["saveView"] => ({
	fetchXml: XML,
	layoutXml: "",
	entityLogicalName: "account",
	objectTypeCode: 1,
	primaryIdAttribute: "accountid",
	targetView: null,
	...extra,
});

const barProps = (props: Partial<AppBarProps> = {}): AppBarProps => ({
	run: run(),
	saveView: saveView(),
	getCurrentXml: () => XML,
	viewName: null,
	unsaved: false,
	onNewQuery: () => {},
	...props,
});

const renderBar = (props: Partial<AppBarProps> = {}) =>
	render(createElement(AppBar, barProps(props)));

const bar = () => document.querySelector<HTMLElement>('[role="toolbar"]');
const overflowed = (element: Element | null | undefined) =>
	!!element?.closest("[data-overflowing]");
/** A bar control by its accessible name or text (a split's primary action for a split) */
const onBar = (name: string) =>
	[...(bar()?.querySelectorAll<HTMLElement>("button") ?? [])].find(
		(b) => b.getAttribute("aria-label") === name || b.textContent?.trim() === name,
	);
const shownOnBar = (names: string[]) => names.filter((n) => onBar(n) && !overflowed(onBar(n)));
/** More is icon-only: its label Tooltip names it */
const moreByTooltip = () =>
	bar()?.querySelector<HTMLButtonElement>('button[aria-label="More commands"]');
const menus = () => [...document.querySelectorAll<HTMLElement>('[role="menu"]')];
const menuItems = (menu: HTMLElement | undefined) => [
	...(menu?.querySelectorAll<HTMLElement>('[role^="menuitem"]') ?? []),
];
const menuTexts = (menu: HTMLElement | undefined) =>
	menuItems(menu).map((i) => i.textContent?.trim());
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};
/** A control's name: its aria-label, its label Tooltip, or its text */
const nameOf = (b: HTMLElement) => {
	const labelledBy = b.getAttribute("aria-labelledby");
	return (
		b.getAttribute("aria-label") ??
		(labelledBy ? document.getElementById(labelledBy)?.textContent : undefined) ??
		b.textContent?.trim()
	);
};

beforeEach(() => {
	// Fluent's positioning warns in jsdom
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

describe("the app bar's commands", () => {
	it("are Run, Save, Send to tool and Return FetchXML, then Settings, with Run the one primary", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({
			isCallee: true,
			sendToTool: { fetchXml: XML, entityLogicalName: "account", tools },
		});
		// A split's menu half shares its primary action's name; count the split once
		const names = [...(bar()?.querySelectorAll<HTMLElement>("button") ?? [])]
			.filter((b) => !b.classList.contains("fui-SplitButton__menuButton"))
			.map(nameOf)
			.filter(Boolean);
		expect(names).toEqual([
			"Run",
			"New",
			"Save",
			"Send to tool",
			"Return FetchXML",
			"More commands",
			"Settings",
		]);
		expect(bar()?.getAttribute("aria-label")).toBe("Query commands");
	});

	it("keeps Run and Settings when narrow, and moves the rest into More in bar order", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar({
			isCallee: true,
			sendToTool: { fetchXml: XML, entityLogicalName: "account", tools },
		});
		await layout.resize({ container: 200 });
		expect(shownOnBar(["Run", "Save", "Send to tool", "Return FetchXML"])).toEqual(["Run"]);
		expect(overflowed(moreByTooltip())).toBe(false);
		await click(moreByTooltip());
		expect(menuTexts(menus()[0])).toEqual([
			"New",
			"Save",
			"Send to tool",
			"Return FetchXML",
			"Copy FetchXML",
		]);
	});

	it("always has Copy FetchXML in More, off without a query", async () => {
		stubOverflowWidths({ container: 3000 });
		installHost();
		await renderBar({ getCurrentXml: () => "" });
		await click(moreByTooltip());
		const copy = menuItems(menus()[0]).find((i) => i.textContent?.trim() === "Copy FetchXML");
		expect(copy?.getAttribute("aria-disabled")).toBe("true");
	});

	it("keeps an unavailable Run and Save focusable, with the reason", async () => {
		stubOverflowWidths({ container: 3000 });
		const onClick = vi.fn();
		await renderBar({
			run: run({ disabledReason: NO_QUERY_REASON, onClick }),
			saveView: saveView({ disabled: true }),
		});
		const runButton = onBar("Run");
		expect(runButton?.getAttribute("aria-disabled")).toBe("true");
		expect(description(runButton)).toBe("Choose an entity first");
		await click(runButton);
		expect(onClick).not.toHaveBeenCalled();
		const save = onBar("Save");
		expect(save?.getAttribute("aria-disabled")).toBe("true");
		expect(description(save)).toBe("Choose an entity first");
	});

	it("runs, and says Ctrl+Enter runs it too", async () => {
		stubOverflowWidths({ container: 3000 });
		const onClick = vi.fn();
		await renderBar({ run: run({ onClick }) });
		expect(description(onBar("Run"))).toBe("Run query (Ctrl+Enter)");
		await click(onBar("Run"));
		expect(onClick).toHaveBeenCalledOnce();
	});

	it("is busy while more rows load, even without a reason, and says its shortcut", async () => {
		stubOverflowWidths({ container: 3000 });
		const onClick = vi.fn();
		await renderBar({ run: run({ busy: true, onClick }) });
		expect(onBar("Run")?.getAttribute("aria-disabled")).toBe("true");
		await click(onBar("Run"));
		expect(onClick).not.toHaveBeenCalled();
		await renderBar();
		expect(onBar("Run")?.getAttribute("aria-keyshortcuts")).toBe("Control+Enter");
	});

	it("puts Settings at the far end", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar();
		const divider = [...(bar()?.querySelectorAll('[role="separator"]') ?? [])].at(-1);
		expect(divider && getComputedStyle(divider).marginInlineStart).toBe("auto");
	});

	it("can't run again while a query runs", async () => {
		stubOverflowWidths({ container: 3000 });
		const onClick = vi.fn();
		await renderBar({
			run: run({ busy: true, disabledReason: "Running the query…", onClick }),
		});
		expect(onBar("Run")?.getAttribute("aria-disabled")).toBe("true");
		expect(description(onBar("Run"))).toBe("Running the query…");
		await click(onBar("Run"));
		expect(onClick).not.toHaveBeenCalled();
	});

	it("opens Settings, which is icon only", async () => {
		stubOverflowWidths({ container: 3000 });
		const onOpenSettings = vi.fn();
		await renderBar({ onOpenSettings });
		const settings = [...(bar()?.querySelectorAll<HTMLElement>("button") ?? [])].find(
			(b) => nameOf(b) === "Settings",
		);
		expect(settings?.textContent?.trim()).toBe("");
		await click(settings);
		expect(onOpenSettings).toHaveBeenCalledOnce();
	});

	it("shows Save's own menu in More, without a second Save", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar();
		await layout.resize({ container: 200 });
		await click(moreByTooltip());
		const save = menuItems(menus()[0]).find((i) => i.textContent?.trim() === "Save");
		await click(save);
		// Save's own menu (the rules are tested in saveRules.test.ts), once
		expect(menuTexts(menus()[1])).toEqual(["Save as new view…"]);
	});
});

describe("the view's identity", () => {
	const identity = () =>
		bar()?.parentElement?.querySelector<HTMLElement>(".fui-Body1Strong")?.textContent;
	const badge = () =>
		[...document.querySelectorAll<HTMLElement>(".fui-Badge")].find(
			(b) => b.textContent === "Unsaved changes",
		);

	it("shows the view's name, or New query, with Unsaved changes when there are some", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar();
		expect(identity()).toBe("New query");
		expect(badge()).toBeUndefined();
		await renderBar({ viewName: "Active accounts", unsaved: true });
		expect(identity()).toBe("Active accounts");
		expect(badge()).toBeTruthy();
	});
});

describe("New", () => {
	const dialog = () =>
		document.querySelector<HTMLElement>('[role="alertdialog"], [role="dialog"]');
	const dialogButton = (label: string) =>
		[...(dialog()?.querySelectorAll<HTMLButtonElement>("button") ?? [])].find(
			(b) => b.textContent?.trim() === label,
		);

	it("starts a new query straight away when nothing is unsaved", async () => {
		stubOverflowWidths({ container: 3000 });
		const onNewQuery = vi.fn();
		await renderBar({ onNewQuery });
		await click(onBar("New"));
		expect(onNewQuery).toHaveBeenCalledOnce();
		expect(dialog()).toBeNull();
	});

	it("asks first when there are unsaved changes: Cancel keeps the query", async () => {
		stubOverflowWidths({ container: 3000 });
		const onNewQuery = vi.fn();
		await renderBar({ onNewQuery, unsaved: true, viewName: "Active accounts" });
		await click(onBar("New"));
		expect(dialog()?.textContent).toContain(`Your changes to "Active accounts" aren't saved.`);
		await click(dialogButton("Cancel"));
		expect(onNewQuery).not.toHaveBeenCalled();
	});

	it("Discard changes starts new and says what it discarded", async () => {
		stubOverflowWidths({ container: 3000 });
		const onNewQuery = vi.fn();
		await renderBar({ onNewQuery, unsaved: true, viewName: "Active accounts" });
		await click(onBar("New"));
		await click(dialogButton("Discard changes"));
		expect(onNewQuery).toHaveBeenCalledOnce();
		expect(toasts()).toEqual([
			expect.stringContaining('Discarded changes to "Active accounts"'),
		]);
	});

	it("Save and start new opens the save; cancelling it keeps the query", async () => {
		stubOverflowWidths({ container: 3000 });
		const onNewQuery = vi.fn();
		await renderBar({ onNewQuery, unsaved: true });
		await click(onBar("New"));
		expect(dialog()?.textContent).toContain("Your query isn't saved.");
		await click(dialogButton("Save and start new"));
		await flush(5);
		expect(dialog()?.textContent).toContain("Save as new view");
		await click(dialogButton("Cancel"));
		await flush(5);
		expect(onNewQuery).not.toHaveBeenCalled();
		// A later save that isn't from New doesn't start a new query either
	});
});

describe("which commands show", () => {
	it("shows Return FetchXML only when another tool opened FetchXML Studio", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar();
		expect(onBar("Return FetchXML")).toBeUndefined();
		await renderBar({ isCallee: true });
		expect(onBar("Return FetchXML")).toBeTruthy();
	});

	it("shows Send to tool only when a tool takes FetchXML", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ sendToTool: { fetchXml: XML, entityLogicalName: "account", tools: [] } });
		expect(onBar("Send to tool")).toBeUndefined();
		await renderBar({ sendToTool: { fetchXml: XML, entityLogicalName: "account", tools } });
		expect(onBar("Send to tool")?.getAttribute("aria-disabled")).toBeNull();
	});

	it("keeps Send to tool focusable with its reason before there's a query", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({
			sendToTool: { fetchXml: "", entityLogicalName: "", tools, disabled: true },
		});
		const send = onBar("Send to tool");
		expect(send?.getAttribute("aria-disabled")).toBe("true");
		expect(description(send)).toBe("Choose an entity first");
	});

	it("offers Save and Send to tool with an empty tree when the editor has FetchXML", async () => {
		stubOverflowWidths({ container: 3000 });
		await render(
			createElement(
				EditorSyncContext.Provider,
				{
					value: {
						sync: async () => ({ status: "unchanged" as const }),
						editorHasContent: true,
					},
				},
				createElement(
					AppBar,
					barProps({
						saveView: saveView({ disabled: true }),
						sendToTool: { fetchXml: "", entityLogicalName: "", tools, disabled: true },
					}),
				),
			),
		);
		expect(onBar("Save")?.getAttribute("aria-disabled")).toBeNull();
		expect(onBar("Send to tool")?.getAttribute("aria-disabled")).toBeNull();
	});
});

describe("dividers", () => {
	it("never shows two in a row, or one first, as commands overflow", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar({
			isCallee: true,
			sendToTool: { fetchXml: XML, entityLogicalName: "account", tools },
		});
		// What shows on the bar, in order: "|" for a divider, a name for a control
		const sequence = () =>
			[...(bar()?.children ?? [])]
				.filter(
					(c) =>
						!c.hasAttribute("data-overflowing") &&
						getComputedStyle(c).display !== "none",
				)
				.map((c) => {
					if (
						c.getAttribute("role") === "separator" ||
						c.querySelector('[role="separator"]')
					)
						return "|";
					const b = (
						c.matches("button") ? c : c.querySelector("button")
					) as HTMLElement | null;
					return b ? (nameOf(b) ?? "") : "";
				})
				.filter(Boolean);
		for (const width of [700, 600, 520, 480, 440, 400, 360, 300]) {
			await layout.resize({ container: width });
			const shown = sequence();
			const doubled = shown.some((s, i) => s === "|" && shown[i + 1] === "|");
			expect({ width, shown, doubled }).toEqual({ width, shown, doubled: false });
			expect(shown[0]).not.toBe("|");
		}
	});
});
