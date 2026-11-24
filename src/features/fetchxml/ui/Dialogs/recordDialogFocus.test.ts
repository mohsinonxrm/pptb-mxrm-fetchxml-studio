/**
 * First focus in the record dialogs (#279, owner): Cancel, the safe action, also when Fluent's own
 * first focus (Tabster, which needs layout) would pick something else: Bulk delete's I understand
 * box. jsdom has no layout, so it's faked here and put back after this file.
 */
import { act, createElement } from "react";
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { BulkDeleteDialog } from "./BulkDeleteDialog";
import { cleanup, installHost, removeHost, render } from "../../../../shared/testing/ui";

const originalOffsetParent = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "offsetParent");
const originalRect = HTMLElement.prototype.getBoundingClientRect;
beforeAll(() => {
	Object.defineProperty(HTMLElement.prototype, "offsetParent", {
		configurable: true,
		get: () => document.body,
	});
	HTMLElement.prototype.getBoundingClientRect = () =>
		({ x: 0, y: 0, top: 0, left: 0, right: 100, bottom: 20, width: 100, height: 20, toJSON: () => ({}) }) as DOMRect;
});
afterAll(() => {
	if (originalOffsetParent)
		Object.defineProperty(HTMLElement.prototype, "offsetParent", originalOffsetParent);
	HTMLElement.prototype.getBoundingClientRect = originalRect;
});
afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

const settle = () => act(async () => void (await new Promise((r) => setTimeout(r, 80))));

describe("record dialogs' first focus", () => {
	it("Bulk delete of all records: Cancel, not the I understand box Fluent would pick", async () => {
		installHost();
		await render(
			createElement(BulkDeleteDialog, {
				open: true,
				recordCount: 0,
				entityDisplayName: "Contact",
				entityLogicalName: "contact",
				isAllRecords: true,
				onClose: () => {},
				onConfirm: async () => ({ asyncOperationId: "job", jobUrl: "" }),
			}),
		);
		await settle();
		expect(document.activeElement?.textContent?.trim()).toBe("Cancel");
	});
});
