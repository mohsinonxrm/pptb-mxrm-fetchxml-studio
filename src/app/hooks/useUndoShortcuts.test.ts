/**
 * Ctrl+Z undoes, Ctrl+Y and Ctrl+Shift+Z redo (#192): from a tree row, a button or the page; not in
 * a text field or Monaco, not while a dialog or menu is open, not when something already handled the
 * key. Undoing from the tree puts focus back on the selected row when the focused row went away.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { useUndoShortcuts, type UndoShortcutHandlers } from "./useUndoShortcuts";

function Probe(handlers: UndoShortcutHandlers) {
	useUndoShortcuts(handlers);
	return createElement(
		"div",
		null,
		createElement(
			"div",
			{ role: "tree", id: "tree" },
			createElement("div", { role: "treeitem", tabIndex: 0, id: "row" }, "filter"),
			createElement("div", { role: "treeitem", tabIndex: -1, id: "other" }, "account"),
		),
		createElement("input", { id: "field" }),
		createElement("input", { id: "check", type: "checkbox" }),
		createElement("textarea", { id: "area" }),
		createElement("button", { id: "button" }, "Undo"),
		createElement(
			"div",
			{ className: "monaco-editor", id: "monaco" },
			createElement("textarea"),
		),
	);
}

const press = async (target: Element | null, init: KeyboardEventInit) => {
	let event!: KeyboardEvent;
	await act(async () => {
		event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
		target?.dispatchEvent(event);
	});
	return event;
};
const ctrlZ = (target: Element | null, extra: KeyboardEventInit = {}) =>
	press(target, { key: "z", ctrlKey: true, ...extra });
const byId = (id: string) => document.getElementById(id);
const frames = () => act(() => new Promise<void>((resolve) => setTimeout(resolve, 60)));

const setup = async (extra: Partial<UndoShortcutHandlers> = {}) => {
	const handlers = { undo: vi.fn(), redo: vi.fn(), restoreFocus: vi.fn(), ...extra };
	await render(createElement(Probe, handlers));
	return handlers;
};

afterEach(() => {
	cleanup();
	for (const el of document.querySelectorAll("[data-test-popup]")) el.remove();
});

describe("Ctrl+Z / Ctrl+Y", () => {
	it("undo from a tree row, a button, a checkbox and the page, and take the key", async () => {
		const h = await setup();
		const event = await ctrlZ(byId("row"));
		expect(event.defaultPrevented).toBe(true);
		await ctrlZ(byId("button"));
		await ctrlZ(byId("check"));
		await ctrlZ(document.body);
		expect(h.undo).toHaveBeenCalledTimes(4);
		expect(h.redo).not.toHaveBeenCalled();
	});

	it("redo with Ctrl+Y and Ctrl+Shift+Z (capital Z, as Shift sends it)", async () => {
		const h = await setup();
		await press(byId("row"), { key: "y", ctrlKey: true });
		await press(byId("row"), { key: "Z", ctrlKey: true, shiftKey: true });
		expect(h.redo).toHaveBeenCalledTimes(2);
		expect(h.undo).not.toHaveBeenCalled();
	});

	it("repeat while held, as undo does elsewhere", async () => {
		const h = await setup();
		await ctrlZ(byId("row"), { repeat: true });
		expect(h.undo).toHaveBeenCalledOnce();
	});

	it("leave text fields, textareas and Monaco their own undo", async () => {
		const h = await setup();
		const event = await ctrlZ(byId("field"));
		await ctrlZ(byId("area"));
		await ctrlZ(byId("monaco")?.querySelector("textarea") ?? null);
		await press(byId("field"), { key: "y", ctrlKey: true });
		expect(h.undo).not.toHaveBeenCalled();
		expect(h.redo).not.toHaveBeenCalled();
		expect(event.defaultPrevented).toBe(false);
	});

	it("do nothing while a dialog or a menu is open", async () => {
		const h = await setup();
		for (const role of ["dialog", "alertdialog", "menu"]) {
			const popup = document.createElement("div");
			popup.setAttribute("role", role);
			popup.setAttribute("data-test-popup", "");
			document.body.appendChild(popup);
			await ctrlZ(byId("row"));
			popup.remove();
		}
		expect(h.undo).not.toHaveBeenCalled();
		await ctrlZ(byId("row"));
		expect(h.undo).toHaveBeenCalledOnce();
	});

	it("do nothing when the focused control already handled the key, or while composing", async () => {
		const h = await setup();
		const handled = (e: Event) => e.preventDefault();
		byId("row")?.addEventListener("keydown", handled);
		await ctrlZ(byId("row"));
		byId("row")?.removeEventListener("keydown", handled);
		await ctrlZ(byId("row"), { isComposing: true });
		expect(h.undo).not.toHaveBeenCalled();
	});

	it("ignore other keys: plain Z, Ctrl+Alt+Z, Ctrl+Shift+Y", async () => {
		const h = await setup();
		await press(byId("row"), { key: "z" });
		await ctrlZ(byId("row"), { altKey: true });
		await press(byId("row"), { key: "Y", ctrlKey: true, shiftKey: true });
		expect(h.undo).not.toHaveBeenCalled();
		expect(h.redo).not.toHaveBeenCalled();
	});
});

describe("focus after a key undo from the tree", () => {
	it("goes back to the tree when the focused row went away", async () => {
		const h = await setup({
			undo: vi.fn(() => byId("row")?.remove()),
		});
		byId("row")?.focus();
		await ctrlZ(byId("row"));
		await frames();
		expect(h.restoreFocus).toHaveBeenCalledOnce();
	});

	it("isn't moved when the key came from outside the tree, even with focus on the page", async () => {
		const h = await setup();
		await ctrlZ(document.body);
		await frames();
		expect(h.undo).toHaveBeenCalledOnce();
		expect(h.restoreFocus).not.toHaveBeenCalled();
	});

	it("stays put when the focused row is still there, or the key came from outside the tree", async () => {
		const h = await setup();
		byId("row")?.focus();
		await ctrlZ(byId("row"));
		byId("button")?.focus();
		await ctrlZ(byId("button"));
		await frames();
		expect(h.undo).toHaveBeenCalledTimes(2);
		expect(h.restoreFocus).not.toHaveBeenCalled();
	});
});
