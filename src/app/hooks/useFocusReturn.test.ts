/**
 * Focus goes back to whatever opened a dialog once it closes (#186): the element focused before the
 * first popup of a run (a menu and the dialog its item opens count as one), unless it's gone or
 * hidden; a non-modal drawer isn't a popup.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { useFocusReturn } from "./useFocusReturn";

function Probe() {
	useFocusReturn();
	return createElement(
		"div",
		null,
		createElement("div", { role: "treeitem", tabIndex: 0, id: "row" }, "account"),
		createElement("input", { id: "field", type: "checkbox" }),
		createElement("button", { id: "toolbar" }, "Copy"),
	);
}

const byId = (id: string) => document.getElementById(id) as HTMLElement;
const focus = (el: HTMLElement) => act(async () => el.focus());
const frames = () => act(() => new Promise<void>((resolve) => setTimeout(resolve, 80)));
const popup = async (role: string, attrs: Record<string, string> = {}) => {
	const el = document.createElement("div");
	el.setAttribute("role", role);
	for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
	el.dataset.testPopup = "";
	const inside = document.createElement("button");
	el.appendChild(inside);
	await act(async () => {
		document.body.appendChild(el);
		inside.focus();
	});
	await frames();
	return el;
};
/** Fluent's restore: focus on the last toolbar button, when the dialog goes */
const close = async (el: HTMLElement) => {
	await act(async () => {
		el.remove();
		byId("toolbar").focus();
	});
	await frames();
};

afterEach(() => {
	cleanup();
	for (const el of document.querySelectorAll("[data-test-popup]")) el.remove();
});

describe("useFocusReturn", () => {
	it("puts focus back on the element that had it before the dialog", async () => {
		await render(createElement(Probe));
		await focus(byId("toolbar"));
		await focus(byId("row"));
		const dialog = await popup("dialog");
		await close(dialog);
		expect(document.activeElement).toBe(byId("row"));
	});

	it("counts a menu and the dialog its item opens as one: back to what opened the menu", async () => {
		await render(createElement(Probe));
		await focus(byId("field"));
		const menu = await popup("menu");
		const dialog = await popup("alertdialog");
		await act(async () => menu.remove());
		await frames();
		expect(document.activeElement).not.toBe(byId("field"));
		await close(dialog);
		expect(document.activeElement).toBe(byId("field"));
	});

	it("leaves focus alone when what opened the dialog is gone or hidden", async () => {
		await render(createElement(Probe));
		await focus(byId("row"));
		const dialog = await popup("dialog");
		await act(async () => byId("row").remove());
		await close(dialog);
		expect(document.activeElement).toBe(byId("toolbar"));

		await focus(byId("field"));
		const second = await popup("dialog");
		await act(async () => {
			byId("field").style.display = "none";
		});
		await close(second);
		expect(document.activeElement).toBe(byId("toolbar"));
	});

	it("doesn't treat a non-modal drawer as a dialog", async () => {
		await render(createElement(Probe));
		await focus(byId("row"));
		const drawer = await popup("dialog", { "aria-modal": "false" });
		await focus(byId("toolbar"));
		await act(async () => drawer.remove());
		await frames();
		expect(document.activeElement).toBe(byId("toolbar"));
	});

	it("keeps what opened the dialog when focus moves outside while it's still open", async () => {
		await render(createElement(Probe));
		await focus(byId("row"));
		const dialog = await popup("dialog");
		// Something focuses an element outside while the dialog stays open; the dialog changes
		await focus(byId("field"));
		await act(async () => dialog.appendChild(document.createElement("span")));
		await frames();
		await close(dialog);
		expect(document.activeElement).toBe(byId("row"));
	});

	it("waits when one dialog closes as the next opens, and returns once both are gone", async () => {
		await render(createElement(Probe));
		await focus(byId("row"));
		const first = await popup("dialog");
		let second!: HTMLElement;
		// The first goes, and the second comes before two frames have passed
		await act(async () => first.remove());
		await act(async () => {
			second = document.createElement("div");
			second.setAttribute("role", "dialog");
			second.dataset.testPopup = "";
			const inside = document.createElement("button");
			second.appendChild(inside);
			document.body.appendChild(second);
			inside.focus();
		});
		await frames();
		// Still in the second dialog: nothing pulled focus out of it
		expect(second.contains(document.activeElement)).toBe(true);
		await close(second);
		expect(document.activeElement).toBe(byId("row"));
	});

	it("comes back to a field inside the non-modal Properties drawer", async () => {
		await render(createElement(Probe));
		const drawer = document.createElement("div");
		drawer.setAttribute("role", "dialog");
		drawer.setAttribute("aria-modal", "false");
		drawer.dataset.testPopup = "";
		const field = document.createElement("input");
		drawer.appendChild(field);
		await act(async () => document.body.appendChild(drawer));
		await focus(field);
		await close(await popup("alertdialog"));
		expect(document.activeElement).toBe(field);
	});

	it("falls back to the row that held the control when the control is gone", async () => {
		await render(createElement(Probe));
		const row = byId("row");
		const inline = document.createElement("button");
		row.appendChild(inline);
		await focus(inline);
		// The button is gone before the dialog is even seen (Fluent unmounts a row's actions on blur)
		const dialog = await act(async () => {
			const el = document.createElement("div");
			el.setAttribute("role", "dialog");
			el.dataset.testPopup = "";
			const inside = document.createElement("button");
			el.appendChild(inside);
			document.body.appendChild(el);
			inside.focus();
			inline.remove();
			return el;
		});
		await frames();
		await close(dialog);
		expect(document.activeElement).toBe(row);
	});

	it("comes back to nothing when focus had fallen to the page", async () => {
		await render(createElement(Probe));
		await focus(byId("row"));
		await act(async () => byId("row").blur());
		await act(() => new Promise<void>((resolve) => setTimeout(resolve, 10)));
		await close(await popup("dialog"));
		expect(document.activeElement).toBe(byId("toolbar"));
	});

	it("starts over for the next dialog", async () => {
		await render(createElement(Probe));
		await focus(byId("row"));
		await close(await popup("dialog"));
		await focus(byId("field"));
		await close(await popup("dialog"));
		expect(document.activeElement).toBe(byId("field"));
	});
});
