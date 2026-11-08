/** The shared "Discard unsaved changes?" dialog (#278, #192): wording, order, focus, saving state. */
import { act, createElement } from "react";
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { DiscardChangesDialog, type DiscardChangesDialogProps } from "./DiscardChangesDialog";
import {
	button,
	cleanup,
	click,
	dialogText,
	flush,
	isDisabled,
	openDialog,
	removeHost,
	render,
} from "../../../../shared/testing/ui";

// jsdom has no layout, so Tabster (which Fluent uses to find the first focusable) sees nothing as visible.
// The stubs are put back after this file so no other test inherits a fake layout.
const originalOffsetParent = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "offsetParent");
const originalRect = HTMLElement.prototype.getBoundingClientRect;
beforeAll(() => {
	Object.defineProperty(HTMLElement.prototype, "offsetParent", {
		configurable: true,
		get: () => document.body,
	});
	HTMLElement.prototype.getBoundingClientRect = () => ({
		x: 0,
		y: 0,
		top: 0,
		left: 0,
		right: 100,
		bottom: 20,
		width: 100,
		height: 20,
		toJSON: () => ({}),
	});
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

const props = (over: Partial<DiscardChangesDialogProps> = {}): DiscardChangesDialogProps => ({
	open: true,
	intent: "new",
	viewName: "Active contacts",
	onSave: () => {},
	onDiscard: () => {},
	onCancel: () => {},
	...over,
});

/** Fluent moves initial focus through Tabster on a timer */
const settle = () =>
	act(async () => {
		await new Promise((r) => setTimeout(r, 50));
	});

const labels = () =>
	[...(openDialog()?.querySelectorAll("button") ?? [])].map((b) => b.textContent?.trim());

describe("Discard unsaved changes dialog", () => {
	it("says what a new query clears, for a loaded view", async () => {
		await render(createElement(DiscardChangesDialog, props()));
		expect(dialogText()).toContain("Discard unsaved changes?");
		expect(dialogText()).toContain(
			'Your changes to "Active contacts" aren\'t saved. A new query clears the tree, the FetchXML and the results.',
		);
	});

	it("says what opening replaces, naming both views", async () => {
		await render(
			createElement(
				DiscardChangesDialog,
				props({ intent: "open", targetViewName: "My open leads" }),
			),
		);
		expect(dialogText()).toContain("Discard unsaved changes?");
		expect(dialogText()).toContain(
			'Your changes to "Active contacts" aren\'t saved. Opening "My open leads" replaces the tree, the FetchXML and the results.',
		);
	});

	it("says 'Your query isn't saved.' when no view is loaded", async () => {
		await render(createElement(DiscardChangesDialog, props({ viewName: null })));
		expect(dialogText()).toContain("Your query isn't saved. A new query clears");
		expect(dialogText()).not.toContain("Your changes to");
		await render(
			createElement(
				DiscardChangesDialog,
				props({ viewName: null, intent: "open", targetViewName: "Leads" }),
			),
		);
		expect(dialogText()).toContain('Your query isn\'t saved. Opening "Leads" replaces');
	});

	it("orders the buttons primary, Discard, Cancel, labelled by intent", async () => {
		await render(createElement(DiscardChangesDialog, props()));
		expect(labels()).toEqual(["Save and start new", "Discard changes", "Cancel"]);
		await render(
			createElement(DiscardChangesDialog, props({ intent: "open", targetViewName: "Leads" })),
		);
		expect(labels()).toEqual(["Save and open", "Discard changes", "Cancel"]);
	});

	it("puts initial focus on the primary button", async () => {
		await render(createElement(DiscardChangesDialog, props()));
		await settle();
		expect(document.activeElement).toBe(button("Save and start new"));
	});

	it("calls each handler once", async () => {
		const onSave = vi.fn();
		const onDiscard = vi.fn();
		const onCancel = vi.fn();
		const counts = () => [
			onSave.mock.calls.length,
			onDiscard.mock.calls.length,
			onCancel.mock.calls.length,
		];
		await render(createElement(DiscardChangesDialog, props({ onSave, onDiscard, onCancel })));
		await click(button("Save and start new"));
		expect(counts()).toEqual([1, 0, 0]);
		await click(button("Discard changes"));
		expect(counts()).toEqual([1, 1, 0]);
		await click(button("Cancel"));
		expect(counts()).toEqual([1, 1, 1]);
	});

	it("treats Escape as Cancel", async () => {
		const onCancel = vi.fn();
		await render(createElement(DiscardChangesDialog, props({ onCancel })));
		await act(async () => {
			document.activeElement?.dispatchEvent(
				new KeyboardEvent("keydown", { key: "Escape", bubbles: true }),
			);
		});
		await flush();
		expect(onCancel).toHaveBeenCalledTimes(1);
	});

	it("while saving: primary busy and focusable, Discard disabled, Cancel enabled", async () => {
		const onSave = vi.fn();
		const onDiscard = vi.fn();
		const onCancel = vi.fn();
		await render(
			createElement(
				DiscardChangesDialog,
				props({ saving: true, onSave, onDiscard, onCancel }),
			),
		);
		const primary = button("Saving…");
		expect(primary).toBeDefined();
		expect(primary?.querySelector(".fui-Spinner")).not.toBeNull();
		expect(isDisabled(primary)).toBe(true);
		expect(primary?.disabled).toBe(false); // disabledFocusable: still focusable
		await settle();
		expect(document.activeElement).toBe(primary);
		const discard = button("Discard changes");
		expect(isDisabled(discard)).toBe(true);
		expect(discard?.disabled).toBe(false);
		await click(primary);
		await click(discard);
		expect([onSave.mock.calls.length, onDiscard.mock.calls.length]).toEqual([0, 0]);
		expect(isDisabled(button("Cancel"))).toBe(false);
		await click(button("Cancel"));
		expect(onCancel).toHaveBeenCalledTimes(1);
	});

	it("explains why Discard is disabled while saving", async () => {
		await render(createElement(DiscardChangesDialog, props({ saving: true })));
		const discard = button("Discard changes");
		await act(async () => {
			discard?.focus();
		});
		await flush();
		expect(document.body.textContent).toContain("Wait for the save to finish");
	});
});
