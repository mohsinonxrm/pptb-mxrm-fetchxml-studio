import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import {
	PPTB_RESERVED,
	SHORTCUTS,
	formatKeys,
	getEntry,
	isEditableTarget,
	matches,
	normalizeCombo,
	type KeyEventLike,
	type ShortcutPlatform,
} from "./registry";

const ev = (key: string, mods: Partial<Omit<KeyEventLike, "key">> = {}): KeyEventLike => ({
	key,
	ctrlKey: false,
	metaKey: false,
	shiftKey: false,
	altKey: false,
	...mods,
});

describe("matches", () => {
	const win: ShortcutPlatform = "other";
	const mac: ShortcutPlatform = "mac";

	it("matches Ctrl on Windows and Cmd on a Mac, never the other one", () => {
		for (const [id, key] of [
			["run", "Enter"],
			["save", "s"],
			["shortcuts", "/"],
			["undo", "z"],
			["redo", "y"],
		] as const) {
			expect(matches(ev(key, { ctrlKey: true }), id, win), id).toBe(true);
			expect(matches(ev(key, { metaKey: true }), id, win), `${id} meta on win`).toBe(false);
			expect(matches(ev(key, { metaKey: true }), id, mac), `${id} on mac`).toBe(true);
			expect(matches(ev(key, { ctrlKey: true }), id, mac), `${id} ctrl on mac`).toBe(false);
			expect(matches(ev(key), id, win), `${id} bare`).toBe(false);
		}
	});

	it("compares event.key without case", () => {
		expect(matches(ev("S", { ctrlKey: true }), "save", win)).toBe(true);
		expect(matches(ev("Z", { ctrlKey: true, shiftKey: true }), "redo", win)).toBe(true);
		expect(matches(ev("z", { ctrlKey: true, shiftKey: true }), "redo", win)).toBe(true);
	});

	it("needs Shift exactly: Ctrl+Z is undo, Ctrl+Shift+Z is redo, Ctrl+Shift+Enter is nothing", () => {
		expect(matches(ev("z", { ctrlKey: true, shiftKey: true }), "undo", win)).toBe(false);
		expect(matches(ev("z", { ctrlKey: true }), "redo", win)).toBe(false);
		expect(matches(ev("Enter", { ctrlKey: true, shiftKey: true }), "run", win)).toBe(false);
		expect(matches(ev("y", { ctrlKey: true, shiftKey: true }), "redo", win)).toBe(false);
	});

	it("rejects an extra Alt", () => {
		expect(matches(ev("s", { ctrlKey: true, altKey: true }), "save", win)).toBe(false);
	});

	it("treats / and ? as different keys", () => {
		expect(matches(ev("/", { ctrlKey: true }), "shortcuts", win)).toBe(true);
		expect(matches(ev("?", { ctrlKey: true, shiftKey: true }), "shortcuts", win)).toBe(false);
		expect(matches(ev("?", { ctrlKey: true }), "shortcuts", win)).toBe(false);
	});

	it("matches the documented keys too", () => {
		expect(matches(ev("F10", { shiftKey: true }), "openMenu", win)).toBe(true);
		expect(matches(ev("ContextMenu"), "openMenu", win)).toBe(true);
		expect(matches(ev("F", { shiftKey: true, altKey: true }), "formatXml", win)).toBe(true);
		expect(matches(ev("Escape"), "closePopup", win)).toBe(true);
		expect(matches(ev("Delete"), "removeNode", win)).toBe(true);
	});

	it("matches the model-driven app keys, with Ctrl+F6 and Ctrl+Shift+F6 exact", () => {
		expect(matches(ev("F6", { ctrlKey: true }), "nextArea", win)).toBe(true);
		expect(matches(ev("F6", { ctrlKey: true, shiftKey: true }), "nextArea", win)).toBe(false);
		expect(matches(ev("F6", { ctrlKey: true, shiftKey: true }), "previousArea", win)).toBe(true);
		expect(matches(ev("F6", { ctrlKey: true }), "previousArea", win)).toBe(false);
		expect(matches(ev("F6"), "nextArea", win)).toBe(false);
		expect(matches(ev("F6", { metaKey: true }), "nextArea", mac)).toBe(true);
		expect(matches(ev("F6", { ctrlKey: true }), "nextArea", mac)).toBe(false);
	});

	it("matches F2 with no modifiers only", () => {
		expect(matches(ev("F2"), "toasts", win)).toBe(true);
		expect(matches(ev("F2"), "toasts", mac)).toBe(true);
		for (const mods of [{ ctrlKey: true }, { metaKey: true }, { shiftKey: true }, { altKey: true }]) {
			expect(matches(ev("F2", mods), "toasts", win), JSON.stringify(mods)).toBe(false);
		}
	});

	it("throws for an unknown id", () => {
		expect(() => matches(ev("a"), "nope", win)).toThrow();
	});
});

describe("isEditableTarget", () => {
	const make = (html: string) => {
		const host = document.createElement("div");
		host.innerHTML = html;
		return host;
	};

	it("is true for text inputs, textareas, contenteditable and Monaco", () => {
		const host = make(
			'<input id="a"><input id="n" type="number"><textarea id="t"></textarea><div id="c" contenteditable="true"><b id="cc"></b></div><div class="monaco-editor"><span id="m"></span></div>',
		);
		for (const id of ["a", "n", "t", "c", "cc", "m"]) expect(isEditableTarget(host.querySelector("#" + id)), id).toBe(true);
	});

	it("is false for buttons, checkboxes, plain elements, contenteditable=false and null", () => {
		const host = make('<button id="b"></button><input id="k" type="checkbox"><div id="d"></div><div contenteditable="false"><i id="f"></i></div>');
		for (const id of ["b", "k", "d", "f"]) expect(isEditableTarget(host.querySelector("#" + id)), id).toBe(false);
		expect(isEditableTarget(null)).toBe(false);
		expect(isEditableTarget(document)).toBe(false);
	});

	it("goes with the registry: run and save work in fields, undo, redo and the dialog do not", () => {
		expect(getEntry("run").inEditable && getEntry("save").inEditable).toBe(true);
		for (const id of ["nextArea", "previousArea", "toasts"]) expect(getEntry(id).inEditable, id).toBe(true);
		expect(getEntry("undo").inEditable || getEntry("redo").inEditable || getEntry("shortcuts").inEditable).toBe(false);
	});
});

describe("registry shape", () => {
	const globals = SHORTCUTS.filter((s) => s.scope === "global");

	it("has the decided global entries", () => {
		expect(globals.map((s) => s.id)).toEqual(["run", "save", "shortcuts", "undo", "redo", "nextArea", "previousArea", "toasts"]);
	});

	it("has unique ids and one aria combination per display key", () => {
		expect(new Set(SHORTCUTS.map((s) => s.id)).size).toBe(SHORTCUTS.length);
		for (const s of SHORTCUTS) {
			expect(s.aria.split(" ").map(normalizeCombo), s.id).toEqual(s.keys.map(normalizeCombo));
		}
	});

	it("has no duplicate key combination within a scope", () => {
		for (const scope of ["global", "documented"] as const) {
			const combos = SHORTCUTS.filter((s) => s.scope === scope).flatMap((s) => s.keys.map(normalizeCombo));
			expect(new Set(combos).size, scope).toBe(combos.length);
		}
	});

	it("has no global entry on a PPTB-reserved combination", () => {
		const reserved = new Set(PPTB_RESERVED.map(normalizeCombo));
		for (const s of globals) for (const k of s.keys) expect(reserved.has(normalizeCombo(k)), `${s.id} ${k}`).toBe(false);
	});

	it("does not include F5", () => {
		expect(SHORTCUTS.flatMap((s) => s.keys)).not.toContain("F5");
	});
});

describe("formatKeys", () => {
	it("shows Ctrl on Windows and the command symbol on a Mac", () => {
		expect(formatKeys(getEntry("run"), "other")).toEqual(["Ctrl+Enter"]);
		expect(formatKeys(getEntry("run"), "mac")).toEqual(["⌘+Enter"]);
		expect(formatKeys(getEntry("redo"), "other")).toEqual(["Ctrl+Y", "Ctrl+Shift+Z"]);
		expect(formatKeys(getEntry("redo"), "mac")).toEqual(["⌘+Y", "⌘+Shift+Z"]);
		expect(formatKeys(getEntry("formatXml"), "mac")).toEqual(["Shift+Option+F"]);
		expect(formatKeys(getEntry("closePopup"), "mac")).toEqual(["Escape"]);
	});
});

describe("README", () => {
	const readme = readFileSync(resolve(__dirname, "../../../README.md"), "utf8");
	const heading = readme.search(/^#{2,4} .*Keyboard shortcuts/m);
	const after = heading < 0 ? "" : readme.slice(heading).replace(/^[^\n]*\n/, "");
	const body = after.split(/\n#{1,4} /)[0];

	it("has a Keyboard shortcuts section", () => {
		expect(heading).toBeGreaterThanOrEqual(0);
	});

	it("lists every entry's label and keys", () => {
		for (const s of SHORTCUTS) {
			expect(body, s.label).toContain(s.label);
			for (const k of s.keys) expect(body, `${s.id} ${k}`).toContain("`" + k + "`");
		}
	});

	it("no longer claims the old Ctrl+Enter line", () => {
		expect(readme).not.toContain("to execute query, copy XML");
	});
});
