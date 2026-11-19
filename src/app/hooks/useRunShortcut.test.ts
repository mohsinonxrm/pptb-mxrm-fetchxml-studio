/**
 * Ctrl+Enter runs the query (#278): from anywhere but Monaco, a button or menu item, or while a
 * dialog or menu is open; not when Run has a reason. It's handled in the capture phase and stopped
 * there, so a tree branch doesn't also expand or collapse. A closed Monaco Find widget (role="dialog",
 * aria-hidden) doesn't block it.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { hasOpenPopup, useRunShortcut } from "./useRunShortcut";

function Probe({ run, reason }: { run: () => void; reason?: string }) {
	useRunShortcut(run, reason);
	return createElement(
		"div",
		null,
		createElement("div", { role: "treeitem", tabIndex: 0, id: "tree" }, "account"),
		createElement("input", { id: "field" }),
		createElement("button", { id: "button" }, "Run anyway"),
		createElement(
			"div",
			{ className: "monaco-editor", id: "monaco" },
			createElement("textarea"),
		),
	);
}

const ctrlEnter = async (target: Element | null, extra: KeyboardEventInit = {}) => {
	let event!: KeyboardEvent;
	await act(async () => {
		event = new KeyboardEvent("keydown", {
			key: "Enter",
			ctrlKey: true,
			bubbles: true,
			cancelable: true,
			...extra,
		});
		target?.dispatchEvent(event);
	});
	return event;
};
const byId = (id: string) => document.getElementById(id);

afterEach(() => {
	cleanup();
	for (const el of document.querySelectorAll("[data-test-popup]")) el.remove();
});

describe("Ctrl+Enter", () => {
	it("runs from a tree item and a field, and the tree doesn't see the key", async () => {
		const run = vi.fn();
		await render(createElement(Probe, { run }));
		const seenByTree = vi.fn();
		byId("tree")?.addEventListener("keydown", seenByTree);
		const event = await ctrlEnter(byId("tree"));
		expect(run).toHaveBeenCalledOnce();
		expect(event.defaultPrevented).toBe(true);
		expect(seenByTree).not.toHaveBeenCalled();
		await ctrlEnter(byId("field"));
		expect(run).toHaveBeenCalledTimes(2);
	});

	it("doesn't run in Monaco, on a button, when repeated, or with plain Enter", async () => {
		const run = vi.fn();
		await render(createElement(Probe, { run }));
		await ctrlEnter(byId("monaco")?.querySelector("textarea") ?? null);
		await ctrlEnter(byId("button"));
		await ctrlEnter(byId("tree"), { repeat: true });
		await ctrlEnter(byId("tree"), { ctrlKey: false });
		expect(run).not.toHaveBeenCalled();
	});

	it("doesn't run when Run has a reason", async () => {
		const run = vi.fn();
		await render(createElement(Probe, { run, reason: "Choose an entity first" }));
		await ctrlEnter(byId("tree"));
		expect(run).not.toHaveBeenCalled();
	});

	it("doesn't run while a dialog is open, but a closed Find widget or a hidden one doesn't count", async () => {
		const run = vi.fn();
		await render(createElement(Probe, { run }));
		const dialog = document.createElement("div");
		dialog.setAttribute("role", "dialog");
		dialog.dataset.testPopup = "";
		document.body.appendChild(dialog);
		expect(hasOpenPopup()).toBe(true);
		await ctrlEnter(byId("tree"));
		expect(run).not.toHaveBeenCalled();

		// Monaco's Find widget (role="dialog") stays in the editor after it closes
		dialog.remove();
		const find = document.createElement("div");
		find.setAttribute("role", "dialog");
		find.setAttribute("aria-hidden", "true");
		byId("monaco")?.appendChild(find);
		expect(hasOpenPopup()).toBe(false);
		await ctrlEnter(byId("tree"));
		expect(run).toHaveBeenCalledOnce();

		// A hidden one doesn't count either
		const hiddenDialog = document.createElement("div");
		hiddenDialog.setAttribute("role", "dialog");
		hiddenDialog.hidden = true;
		hiddenDialog.dataset.testPopup = "";
		document.body.appendChild(hiddenDialog);
		expect(hasOpenPopup()).toBe(false);
	});
});
