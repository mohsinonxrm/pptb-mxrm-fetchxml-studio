/**
 * Buttons outside the command bars that can't be used say why and stay focusable (#279): the
 * Advisories fix, the column drawers' Add, Apply and Move up / down, so moving a column to the top
 * keeps focus on Move up.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	click,
	cleanup,
	expectDisabledWithReason,
	render,
	stubOverflowWidths,
} from "../../../../shared/testing/ui";
import { AdvisoriesPanel } from "./AdvisoriesPanel";
import { EditColumnsPanel } from "./EditColumnsPanel";
import { AddColumnsPanel } from "./AddColumnsPanel";
import type { Advisory } from "../../model/advisories";

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});
afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

const buttons = (label: string) =>
	[...document.querySelectorAll<HTMLElement>("button")].filter(
		(b) => b.getAttribute("aria-label") === label || b.textContent?.trim() === label,
	);

describe("Advisories: the fix", () => {
	it("says the editor's FetchXML isn't in the tree yet", async () => {
		stubOverflowWidths({ container: 2000 });
		const advisory: Advisory = {
			id: "a",
			level: "advice",
			title: "Paging needs a unique order",
			detail: "",
			location: "account",
			doc: { label: "Docs", href: "https://learn.microsoft.com" },
			source: "fxs",
			fix: { kind: "add-root-order", attribute: "accountid" },
		};
		await render(
			createElement(AdvisoriesPanel, {
				advisories: [advisory],
				noQuery: false,
				editorInvalid: false,
				inTree: false,
				checking: false,
				stale: false,
				added: 0,
				onDismissCheckError: () => {},
				onCheck: () => {},
				onFix: () => {},
			}),
		);
		expectDisabledWithReason(
			buttons("Add accountid as the last order")[0],
			"The editor's FetchXML isn't in the tree yet. Use Parse to tree first.",
		);
	});
});

describe("the column drawers", () => {
	const columns = [
		{ name: "name", width: 200 },
		{ name: "telephone1", width: 100 },
	];

	it("Edit columns: Move up at the top and Move down at the bottom say why", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(EditColumnsPanel, {
				open: true,
				columns,
				onClose: () => {},
				onApply: () => {},
				onAddColumns: () => {},
			}),
		);
		expectDisabledWithReason(buttons("Move up")[0], "It's already the first column");
		expectDisabledWithReason(buttons("Move down")[1], "It's already the last column");
	});

	it("Edit columns: moving a column to the top keeps focus on its Move up", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(EditColumnsPanel, {
				open: true,
				columns,
				onClose: () => {},
				onApply: () => {},
				onAddColumns: () => {},
			}),
		);
		const up = buttons("Move up")[1];
		up.focus();
		await click(up);
		// telephone1 is first now: its Move up is the same element, focused, and says why
		expect(buttons("Move up")[0]).toBe(up);
		expect(document.activeElement).toBe(up);
		expectDisabledWithReason(up, "It's already the first column");
	});

	it("Edit columns: Apply asks to keep a column", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(EditColumnsPanel, {
				open: true,
				columns: [{ name: "name", width: 200 }],
				onClose: () => {},
				onApply: () => {},
				onAddColumns: () => {},
			}),
		);
		await click(buttons("Remove column")[0]);
		expectDisabledWithReason(buttons("Apply")[0], "Keep at least one column");
	});

	it("Add columns: Add asks for columns", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(AddColumnsPanel, {
				open: true,
				entityDisplayName: "Account",
				onClose: () => {},
				onApply: () => {},
			}),
		);
		expectDisabledWithReason(buttons("Add")[0], "Select columns to add");
	});
});
