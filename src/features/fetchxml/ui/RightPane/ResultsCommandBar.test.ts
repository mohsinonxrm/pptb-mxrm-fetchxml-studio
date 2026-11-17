/**
 * The results command bar on the shared overflow bar (#232): D's priorities decide what moves into
 * More, Stop never does, every unavailable command keeps its reason (no "Record actions unavailable"
 * info buttons), and commands the user lacks the privilege for are hidden (owner decision 1).
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { ResultsCommandBar, type CommandBarProps } from "./ResultsCommandBar";

const renderBar = (props: Partial<CommandBarProps> = {}) =>
	render(
		createElement(ResultsCommandBar, {
			selectedCount: 1,
			loadedCount: 10,
			onOpen: () => {},
			onCopyUrl: () => {},
			onActivate: () => {},
			onDeactivate: () => {},
			onDelete: () => {},
			canDelete: true,
			canBulkDelete: true,
			canRunWorkflow: true,
			onDensityChange: () => {},
			...props,
		}),
	);

const bar = () =>
	document.querySelector<HTMLElement>('[role="toolbar"][aria-label="Record actions"]');
const overflowed = (element: Element | null | undefined) =>
	!!element?.closest("[data-overflowing]");
/** A bar control by its accessible name or visible text */
const onBar = (name: string) =>
	[...(bar()?.querySelectorAll<HTMLElement>("button") ?? [])].find(
		(b) => b.getAttribute("aria-label") === name || b.textContent?.trim() === name,
	);
const names = () =>
	[...(bar()?.querySelectorAll<HTMLElement>("button") ?? [])]
		.filter((b) => !overflowed(b))
		.map((b) => b.getAttribute("aria-label") ?? b.textContent?.trim())
		.filter(Boolean);
const more = () => bar()?.querySelector<HTMLButtonElement>('button[aria-label="More commands"]');
const menus = () => [...document.querySelectorAll<HTMLElement>('[role="menu"]')];
const menuTexts = (menu: HTMLElement | undefined) =>
	[...(menu?.querySelectorAll<HTMLElement>('[role^="menuitem"]') ?? [])].map((i) =>
		i.textContent?.trim(),
	);
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("the results command bar", () => {
	it("keeps today's commands and order", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar();
		expect(names()).toEqual([
			"Open record",
			"Copy record URL",
			"Activate selected records",
			"Deactivate selected records",
			"Delete selected records",
			"Run workflow on selected records",
			"Copy",
			"Export to Excel",
			"Edit columns",
			"Medium",
		]);
		expect(more()).toBeNull();
	});

	it("moves commands into More by D's priorities, and lists them in bar order", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar();
		// Room for three commands, their dividers and More: Open 50, Delete 45, Copy URL 40 stay
		await layout.resize({ container: 24 + 3 * 80 + 2 * 24 + 40 });
		expect(names()).toEqual([
			"Open record",
			"Copy record URL",
			"Delete selected records",
			"More commands",
		]);
		await click(more());
		expect(menuTexts(menus()[0])).toEqual([
			"Activate",
			"Deactivate",
			"Run workflow",
			"Copy",
			"Export to Excel",
			"Edit columns",
			"Medium",
		]);
	});

	it("never moves Stop into More while Retrieve all runs", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar({ isRetrieveAllInProgress: true, onStopRetrieveAll: () => {} });
		await layout.resize({ container: 24 + 80 + 40 });
		expect(names()).toEqual(["Stop", "More commands"]);
	});

	it("shows Delete's own menu in More, without a second Delete", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar({ selectedCount: 2 });
		await layout.resize({ container: 24 + 80 + 40 });
		await click(more());
		const del = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? []),
		].find((i) => i.textContent?.trim() === "Delete");
		await click(del);
		expect(menuTexts(menus()[1])).toEqual(["Delete (2)", "Bulk delete (job)"]);
	});
});

describe("unavailable commands say why", () => {
	it("no 'Record actions unavailable' button: each command carries the reason", async () => {
		stubOverflowWidths({ container: 3000 });
		const reason =
			"Aggregate query results don't contain record IDs, so record actions aren't available.";
		await renderBar({ recordActionsDisabledReason: reason });
		expect(onBar("Record actions unavailable")).toBeUndefined();
		for (const name of [
			"Open record",
			"Copy record URL",
			"Activate selected records",
			"Deactivate selected records",
			"Delete selected records",
			"Run workflow on selected records",
		]) {
			expect(onBar(name)?.getAttribute("aria-disabled")).toBe("true");
			expect(description(onBar(name))).toBe(reason);
		}
	});

	it("Delete all says why it's off, without a 'Delete all unavailable' button (#174)", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ selectedCount: 0, deleteAllDisabledReason: "Run the query first." });
		expect(onBar("Delete all unavailable")).toBeUndefined();
		const del = onBar("Delete all records from view");
		expect(del?.textContent?.trim()).toBe("Delete all");
		expect(del?.getAttribute("aria-disabled")).toBe("true");
		expect(description(del)).toBe("Run the query first.");
	});

	it("Open asks for exactly one row", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ selectedCount: 2 });
		expect(description(onBar("Open record"))).toBe("Select only one row to open it");
	});

	it("Export is busy while it exports", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ isExporting: true });
		const exp = onBar("Export to Excel");
		expect(exp?.textContent?.trim()).toBe("Exporting...");
		expect(exp?.getAttribute("aria-disabled")).toBe("true");
	});

	it("Close asks for rows", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({
			selectedCount: 0,
			closeActions: [{ key: "resolve", label: "Resolve case" }],
			onCloseAction: () => {},
		});
		const close = onBar("Close selected records");
		expect(close?.getAttribute("aria-disabled")).toBe("true");
		expect(description(close)).toBe("Select rows to close them");
	});

	it("a blocked Delete can't be reached: not its menu half, not from More", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		const onDelete = vi.fn();
		const onBulkDelete = vi.fn();
		const blocked = "The query has a one-to-many or many-to-many link-entity, so rows repeat.";
		await renderBar({
			selectedCount: 2,
			onDelete,
			onBulkDelete,
			deleteDisabledReason: blocked,
		});
		const split = onBar("Delete selected records")?.closest(".fui-SplitButton");
		for (const half of split?.querySelectorAll("button") ?? []) await click(half);
		expect(menus()).toHaveLength(0);

		await layout.resize({ container: 24 + 80 + 40 });
		await click(more());
		const del = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? []),
		].find((i) => i.textContent?.trim() === "Delete");
		expect(del?.getAttribute("aria-disabled")).toBe("true");
		await click(del);
		expect(menus()).toHaveLength(1);
		expect(onDelete).not.toHaveBeenCalled();
		expect(onBulkDelete).not.toHaveBeenCalled();
	});

	it("Export says why while it exports", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ isExporting: true });
		expect(description(onBar("Export to Excel"))).toBe("Export in progress...");
	});

	it("Copy has nothing to copy without rows", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ loadedCount: 0 });
		expect(description(onBar("Copy"))).toBe("There are no rows to copy");
	});
});

describe("commands the user lacks the privilege for are hidden", () => {
	it("Delete and Run workflow", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ canDelete: false, canRunWorkflow: false });
		expect(onBar("Delete selected records")).toBeUndefined();
		expect(onBar("Run workflow on selected records")).toBeUndefined();
	});

	it("Activate, Deactivate and Close without Write", async () => {
		stubOverflowWidths({ container: 3000 });
		await renderBar({ canChangeState: false });
		expect(onBar("Activate selected records")).toBeUndefined();
		expect(onBar("Deactivate selected records")).toBeUndefined();
		await renderBar({
			canChangeState: false,
			closeActions: [{ key: "resolve", label: "Resolve case" }],
			onCloseAction: () => {},
		});
		expect(onBar("Close selected records")).toBeUndefined();
	});

	it("Close takes Deactivate's place, and goes with no item the user can use", async () => {
		stubOverflowWidths({ container: 3000 });
		const onCloseAction = vi.fn();
		await renderBar({
			closeActions: [{ key: "resolve", label: "Resolve case" }],
			onCloseAction,
		});
		expect(onBar("Deactivate selected records")).toBeUndefined();
		await click(onBar("Close selected records"));
		await click(menus()[0]?.querySelector('[role="menuitem"]'));
		expect(onCloseAction).toHaveBeenCalledWith("resolve");

		await renderBar({ closeActions: [], onCloseAction });
		expect(onBar("Close selected records")).toBeUndefined();
		expect(onBar("Deactivate selected records")).toBeUndefined();
	});
});
