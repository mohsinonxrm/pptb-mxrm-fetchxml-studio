/**
 * The Keyboard shortcuts dialog (#186): every entry of the shortcut table, the global ones under
 * "Anywhere in the tool" and the local ones under "Where focus is", keys in <kbd>, several keys
 * joined by "or", and Close.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { button, cleanup, click, openDialog, render } from "../../../../shared/testing/ui";
import { SHORTCUTS } from "../../../../app/shortcuts/registry";
import { ShortcutsDialog } from "./ShortcutsDialog";

afterEach(() => cleanup());

const sections = () =>
	[...(openDialog()?.querySelectorAll("section") ?? [])].map((section) => ({
		title: section.querySelector("h3")?.textContent,
		rows: [...section.querySelectorAll("dt")].map((dt) => ({
			label: dt.textContent,
			keys: [...(dt.nextElementSibling?.querySelectorAll("kbd") ?? [])].map(
				(k) => k.textContent,
			),
			text: dt.nextElementSibling?.textContent,
		})),
	}));

describe("ShortcutsDialog", () => {
	it("lists the shortcut table in two sections", async () => {
		await render(createElement(ShortcutsDialog, { open: true, onClose: () => {} }));
		expect(openDialog()?.textContent).toContain("Keyboard shortcuts");
		const [anywhere, local] = sections();
		expect(anywhere.title).toBe("Anywhere in the tool");
		expect(local.title).toBe("Where focus is");
		expect(anywhere.rows.map((r) => r.label)).toEqual(
			SHORTCUTS.filter((s) => s.scope === "global").map((s) => s.label),
		);
		expect(local.rows.map((r) => r.label)).toEqual(
			SHORTCUTS.filter((s) => s.scope === "documented").map((s) => s.label),
		);
	});

	it("shows each key in a kbd, and several keys joined by 'or'", async () => {
		await render(createElement(ShortcutsDialog, { open: true, onClose: () => {} }));
		const rows = sections().flatMap((s) => s.rows);
		expect(rows.find((r) => r.label === "Run query")?.keys).toEqual(["Ctrl+Enter"]);
		const redo = rows.find((r) => r.label === "Redo");
		expect(redo?.keys).toEqual(["Ctrl+Y", "Ctrl+Shift+Z"]);
		expect(redo?.text).toBe("Ctrl+YorCtrl+Shift+Z");
	});

	it("closes with Close", async () => {
		const onClose = vi.fn();
		await render(createElement(ShortcutsDialog, { open: true, onClose }));
		await click(button("Close"));
		expect(onClose).toHaveBeenCalledOnce();
	});

	it("renders nothing while closed", async () => {
		await render(createElement(ShortcutsDialog, { open: false, onClose: () => {} }));
		expect(openDialog()).toBeUndefined();
	});
});
