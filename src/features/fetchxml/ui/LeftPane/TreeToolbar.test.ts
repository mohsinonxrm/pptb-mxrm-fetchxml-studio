/**
 * The tree toolbar's Undo and Redo (#192): icon-only, named by the step and the key ("Undo: Delete
 * filter (Ctrl+Z)"), focusable with a reason when there's nothing to undo or redo.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render } from "../../../../shared/testing/ui";
import { TreeToolbar, type TreeToolbarProps } from "./TreeToolbar";

afterEach(() => cleanup());

const show = (props: Partial<TreeToolbarProps> = {}) => {
	const all: TreeToolbarProps = {
		hasQuery: true,
		canUndo: false,
		canRedo: false,
		undoLabel: null,
		redoLabel: null,
		onUndo: vi.fn(),
		onRedo: vi.fn(),
		...props,
	};
	return render(createElement(TreeToolbar, all)).then(() => all);
};
const toolbar = () => document.querySelector<HTMLElement>('[role="toolbar"]');
/** The commands, not the More button the toolbar keeps for when they overflow */
const buttons = () => [
	...(toolbar()?.querySelectorAll<HTMLButtonElement>("button:not([aria-haspopup])") ?? []),
];
const named = (start: string) =>
	buttons().find((b) => b.getAttribute("aria-label")?.startsWith(start));

describe("TreeToolbar", () => {
	it("is a toolbar named for the tree, with Undo then Redo", async () => {
		await show();
		expect(toolbar()?.getAttribute("aria-label")).toBe("Query tree commands");
		expect(buttons().map((b) => b.getAttribute("aria-label"))).toEqual([
			"Undo: Nothing to undo",
			"Redo: Nothing to redo",
		]);
	});

	it("with nothing to undo or redo, the buttons stay focusable and do nothing", async () => {
		const props = await show();
		const undo = named("Undo")!;
		expect(undo.getAttribute("aria-disabled")).toBe("true");
		expect(undo.disabled).toBe(false);
		// No key to advertise while it does nothing
		expect(undo.hasAttribute("aria-keyshortcuts")).toBe(false);
		await click(undo);
		await click(named("Redo"));
		expect(props.onUndo).not.toHaveBeenCalled();
		expect(props.onRedo).not.toHaveBeenCalled();
	});

	it("names the step and the key, and runs it", async () => {
		const props = await show({
			canUndo: true,
			canRedo: true,
			undoLabel: "Delete filter",
			redoLabel: "Add condition",
		});
		const undo = named("Undo")!;
		const redo = named("Redo")!;
		expect(undo.getAttribute("aria-label")).toBe("Undo: Delete filter (Ctrl+Z)");
		expect(redo.getAttribute("aria-label")).toBe("Redo: Add condition (Ctrl+Y)");
		expect(undo.getAttribute("aria-keyshortcuts")).toBe("Control+Z");
		expect(redo.getAttribute("aria-keyshortcuts")).toBe("Control+Y Control+Shift+Z");
		expect(undo.getAttribute("aria-disabled")).not.toBe("true");
		await click(undo);
		await click(redo);
		expect(props.onUndo).toHaveBeenCalledOnce();
		expect(props.onRedo).toHaveBeenCalledOnce();
	});

	it("is hidden before there's a query, unless there's a step to undo or redo", async () => {
		await show({ hasQuery: false });
		expect(toolbar()).toBeNull();
		cleanup();
		await show({ hasQuery: false, canRedo: true, redoLabel: "Change entity" });
		expect(named("Redo")?.getAttribute("aria-label")).toBe("Redo: Change entity (Ctrl+Y)");
		cleanup();
		await show({ hasQuery: false, canUndo: true, undoLabel: "Parse to tree" });
		expect(named("Undo")?.getAttribute("aria-label")).toBe("Undo: Parse to tree (Ctrl+Z)");
	});
});
