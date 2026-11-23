/**
 * The tool's keys (#186), except Undo / Redo (useShortcuts.undo.test.ts):
 *  - Ctrl+Enter runs (#278) from anywhere, now the FetchXML editor too, but not on a button or menu
 *    item, while a modal dialog or menu is open, or while Run has a reason. It's stopped in the
 *    capture phase, so a tree branch doesn't also expand and Monaco doesn't insert a line.
 *  - Ctrl+S saves, in fields and the editor too, once per press.
 *  - Ctrl+/ shows the shortcuts, outside text fields and the editor (Monaco toggles a comment).
 *  - Ctrl+F6 / Ctrl+Shift+F6 move between areas, in fields and the editor too.
 * A closed Monaco Find widget, a hidden dialog and the non-modal Properties drawer (#385) don't
 * count as a popup in front.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { hasOpenPopup, useShortcuts, type ShortcutHandlers } from "./useShortcuts";
import { RUNS_ITSELF_ATTRIBUTE } from "../shortcuts/registry";

function Probe(handlers: ShortcutHandlers) {
	useShortcuts(handlers);
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
		// The FetchXML editor runs Ctrl+Enter itself (Monaco's action), and has a Find box
		createElement(
			"div",
			{ className: "monaco-editor", id: "fetchxml", [RUNS_ITSELF_ATTRIBUTE]: "" },
			createElement("textarea", { id: "fetchxmlText" }),
			createElement(
				"div",
				{ className: "find-widget" },
				createElement("textarea", { id: "find" }),
			),
		),
	);
}

const setup = async (extra: Partial<ShortcutHandlers> = {}) => {
	const handlers = {
		run: vi.fn(),
		save: vi.fn(),
		undo: vi.fn(),
		redo: vi.fn(),
		showShortcuts: vi.fn(),
		nextArea: vi.fn(),
		previousArea: vi.fn(),
		...extra,
	};
	await render(createElement(Probe, handlers));
	return handlers;
};

const press = async (target: Element | null, init: KeyboardEventInit) => {
	let event!: KeyboardEvent;
	await act(async () => {
		event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
		target?.dispatchEvent(event);
	});
	return event;
};
const ctrlEnter = (target: Element | null, extra: KeyboardEventInit = {}) =>
	press(target, { key: "Enter", ctrlKey: true, ...extra });
const byId = (id: string) => document.getElementById(id);
const monaco = () => byId("monaco")?.querySelector("textarea") ?? null;
const addPopup = (attrs: Record<string, string>) => {
	const el = document.createElement("div");
	for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
	el.dataset.testPopup = "";
	document.body.appendChild(el);
	return el;
};

afterEach(() => {
	cleanup();
	for (const el of document.querySelectorAll("[data-test-popup]")) el.remove();
});

describe("Ctrl+Enter", () => {
	it("runs from a tree item, a field and a read-only editor (LayoutXML), and they don't see the key", async () => {
		const h = await setup();
		const seenByTree = vi.fn();
		const seenByEditor = vi.fn();
		byId("tree")?.addEventListener("keydown", seenByTree);
		monaco()?.addEventListener("keydown", seenByEditor);
		const event = await ctrlEnter(byId("tree"));
		expect(event.defaultPrevented).toBe(true);
		await ctrlEnter(byId("field"));
		await ctrlEnter(monaco());
		expect(h.run).toHaveBeenCalledTimes(3);
		expect(seenByTree).not.toHaveBeenCalled();
		expect(seenByEditor).not.toHaveBeenCalled();
	});

	it("doesn't run on a button or with plain Enter", async () => {
		const h = await setup();
		const onButton = await ctrlEnter(byId("button"));
		await ctrlEnter(byId("tree"), { ctrlKey: false });
		expect(h.run).not.toHaveBeenCalled();
		expect(onButton.defaultPrevented).toBe(false);
	});

	it("held down, or while Run has a reason, runs nothing but still takes the key", async () => {
		const h = await setup({ runDisabledReason: "The query is running" });
		const seenByTree = vi.fn();
		byId("tree")?.addEventListener("keydown", seenByTree);
		const event = await ctrlEnter(byId("tree"));
		expect(event.defaultPrevented).toBe(true);
		expect(h.run).not.toHaveBeenCalled();
		const ready = await setup();
		const held = await ctrlEnter(byId("tree"), { repeat: true });
		expect(held.defaultPrevented).toBe(true);
		expect(ready.run).not.toHaveBeenCalled();
		expect(seenByTree).not.toHaveBeenCalled();
	});

	it("leaves the FetchXML editor's own Run action and Monaco's Find box their Ctrl+Enter", async () => {
		const h = await setup();
		const inEditor = await ctrlEnter(byId("fetchxmlText"));
		const inFind = await ctrlEnter(byId("find"));
		expect(h.run).not.toHaveBeenCalled();
		expect(inEditor.defaultPrevented).toBe(false);
		expect(inFind.defaultPrevented).toBe(false);
	});

	it("doesn't run while a dialog is open, but a closed Find widget or a hidden one doesn't count", async () => {
		const h = await setup();
		const dialog = addPopup({ role: "dialog" });
		expect(hasOpenPopup()).toBe(true);
		await ctrlEnter(byId("tree"));
		expect(h.run).not.toHaveBeenCalled();

		// Monaco's Find widget (role="dialog") stays in the editor after it closes
		dialog.remove();
		const find = document.createElement("div");
		find.setAttribute("role", "dialog");
		find.setAttribute("aria-hidden", "true");
		byId("monaco")?.appendChild(find);
		expect(hasOpenPopup()).toBe(false);
		await ctrlEnter(byId("tree"));
		expect(h.run).toHaveBeenCalledOnce();

		// A hidden one doesn't count either
		const hidden = addPopup({ role: "dialog" });
		hidden.hidden = true;
		expect(hasOpenPopup()).toBe(false);
	});

	it("runs beside a non-modal dialog, like the Properties drawer (#385)", async () => {
		const h = await setup();
		const drawer = addPopup({ role: "dialog", "aria-modal": "false" });
		expect(hasOpenPopup()).toBe(false);
		await ctrlEnter(byId("tree"));
		expect(h.run).toHaveBeenCalledOnce();
		// A modal one still holds it
		drawer.setAttribute("aria-modal", "true");
		expect(hasOpenPopup()).toBe(true);
	});
});

describe("Ctrl+S", () => {
	it("saves from the tree, a field and the editor, and takes the key", async () => {
		const h = await setup();
		const seenByEditor = vi.fn();
		monaco()?.addEventListener("keydown", seenByEditor);
		const event = await press(byId("tree"), { key: "s", ctrlKey: true });
		expect(event.defaultPrevented).toBe(true);
		await press(byId("field"), { key: "s", ctrlKey: true });
		await press(monaco(), { key: "S", ctrlKey: true });
		expect(h.save).toHaveBeenCalledTimes(3);
		expect(seenByEditor).not.toHaveBeenCalled();
	});

	it("saves once per press, and not behind a dialog (the Save dialog itself, say)", async () => {
		const h = await setup();
		await press(byId("tree"), { key: "s", ctrlKey: true, repeat: true });
		addPopup({ role: "dialog" });
		await press(byId("tree"), { key: "s", ctrlKey: true });
		expect(h.save).not.toHaveBeenCalled();
	});

	it("isn't Ctrl+Shift+S or Ctrl+Alt+S", async () => {
		const h = await setup();
		await press(byId("tree"), { key: "S", ctrlKey: true, shiftKey: true });
		await press(byId("tree"), { key: "s", ctrlKey: true, altKey: true });
		expect(h.save).not.toHaveBeenCalled();
	});
});

describe("Ctrl+/", () => {
	it("works on a layout where / needs Shift (German, French)", async () => {
		const h = await setup();
		await press(byId("tree"), { key: "/", code: "Digit7", ctrlKey: true, shiftKey: true });
		expect(h.showShortcuts).toHaveBeenCalledOnce();
	});

	it("shows the shortcuts from the tree, a button or the page", async () => {
		const h = await setup();
		const event = await press(byId("tree"), { key: "/", ctrlKey: true });
		expect(event.defaultPrevented).toBe(true);
		await press(byId("button"), { key: "/", ctrlKey: true });
		await press(document.body, { key: "/", ctrlKey: true });
		expect(h.showShortcuts).toHaveBeenCalledTimes(3);
		expect(h.undo).not.toHaveBeenCalled();
	});

	it("leaves text fields and the editor alone (Monaco toggles a comment), and waits behind a dialog", async () => {
		const h = await setup();
		await press(byId("field"), { key: "/", ctrlKey: true });
		await press(monaco(), { key: "/", ctrlKey: true });
		await press(byId("tree"), { key: "/", ctrlKey: true, repeat: true });
		addPopup({ role: "dialog" });
		await press(byId("tree"), { key: "/", ctrlKey: true });
		expect(h.showShortcuts).not.toHaveBeenCalled();
	});

	it("waits for a control that handled the key itself", async () => {
		const h = await setup();
		const handled = (e: Event) => e.preventDefault();
		byId("tree")?.addEventListener("keydown", handled);
		await press(byId("tree"), { key: "/", ctrlKey: true });
		expect(h.showShortcuts).not.toHaveBeenCalled();
	});
});

describe("Ctrl+F6 / Ctrl+Shift+F6", () => {
	it("move to the next and previous area from anywhere, the editor too", async () => {
		const h = await setup();
		const event = await press(byId("tree"), { key: "F6", ctrlKey: true });
		expect(event.defaultPrevented).toBe(true);
		await press(byId("field"), { key: "F6", ctrlKey: true });
		await press(monaco(), { key: "F6", ctrlKey: true, shiftKey: true });
		expect(h.nextArea).toHaveBeenCalledTimes(2);
		expect(h.previousArea).toHaveBeenCalledOnce();
	});

	it("stay inside a modal dialog, and plain F6 is nothing", async () => {
		const h = await setup();
		await press(byId("tree"), { key: "F6" });
		addPopup({ role: "dialog" });
		await press(byId("tree"), { key: "F6", ctrlKey: true });
		expect(h.nextArea).not.toHaveBeenCalled();
	});
});

describe("the handlers", () => {
	it("are the latest ones after a re-render", async () => {
		const first = await setup();
		const next = await setup();
		await press(byId("tree"), { key: "s", ctrlKey: true });
		expect(next.save).toHaveBeenCalledOnce();
		expect(first.save).not.toHaveBeenCalled();
	});
});
