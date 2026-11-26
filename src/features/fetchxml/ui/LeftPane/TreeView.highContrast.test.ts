/**
 * The selected tree node in Windows high contrast (#361). jsdom has no forced colours, so these tests
 * take each part of the rendered selected row and collect the `@media (forced-colors: active)` rules
 * Griffel inserted that match it: the label (and anything inside it) opts out of the Canvas backplate
 * it vanished on, and every part reads HighlightText on Highlight, hover included.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TreeView } from "./TreeView";
import { parseFetchXml } from "../../model/fetchxmlParser";
import type { FetchNode, NodeId } from "../../model/nodes";

let root: Root;
let container: HTMLDivElement;

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	Element.prototype.scrollIntoView = () => {};
	container = document.createElement("div");
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

const show = (fetchQuery: FetchNode, selectedNodeId: NodeId) =>
	act(async () => {
		root.render(
			createElement(
				FluentProvider,
				{ theme: webLightTheme },
				createElement(TreeView, {
					fetchQuery,
					selectedNodeId,
					onNodeSelect: () => {},
					onAddAttribute: () => {},
					onAddAllAttributes: () => {},
					onAddOrder: () => {},
					onAddFilter: () => {},
					onAddSubfilter: () => {},
					onAddCondition: () => {},
					onAddLinkEntity: () => {},
					onRemoveNode: () => {},
					onUpdateAttributes: () => {},
				}),
			),
		);
	});

const parse = (xml: string) => parseFetchXml(xml).fetchNode!;

/** The selected row's layout (the element with the selected style) */
const selectedLayout = () => {
	const layout = document.querySelector<HTMLElement>(
		'[role="treeitem"][aria-selected="true"] > .fui-TreeItemLayout',
	);
	expect(layout).not.toBeNull();
	return layout!;
};

/** Every style rule inside `@media (forced-colors: active)` */
function forcedColorRules(): CSSStyleRule[] {
	const rules: CSSStyleRule[] = [];
	for (const sheet of [...document.styleSheets])
		for (const rule of [...sheet.cssRules])
			if (rule instanceof CSSMediaRule && rule.conditionText.includes("forced-colors"))
				for (const inner of [...rule.cssRules])
					if (inner instanceof CSSStyleRule) rules.push(inner);
	return rules;
}

/**
 * The forced-colours declarations that apply to the element, merged in sheet order (values in
 * lowercase, as Griffel writes system colours). With `hover`, the row's `:hover` rules count too.
 */
function forced(element: Element, { hover = false } = {}): Record<string, string> {
	const out: Record<string, string> = {};
	for (const rule of forcedColorRules()) {
		const selector = hover ? rule.selectorText.replace(/:hover/g, "") : rule.selectorText;
		if (!hover && rule.selectorText.includes(":hover")) continue;
		let matches = false;
		try {
			matches = element.matches(selector);
		} catch {
			continue; // pseudo-elements and the like
		}
		if (!matches) continue;
		for (const name of [...rule.style]) {
			const priority = rule.style.getPropertyPriority(name);
			out[name] = (
				rule.style.getPropertyValue(name) + (priority ? ` !${priority}` : "")
			).toLowerCase();
		}
	}
	return out;
}

describe("the selected tree node in high contrast (#361)", () => {
	const ACCOUNT = parse(
		`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`,
	);

	it("is HighlightText on Highlight, and stays so while hovered", async () => {
		await show(ACCOUNT, ACCOUNT.entity.id);
		const row = forced(selectedLayout());
		expect(row["background-color"]).toBe("highlight");
		expect(row.color).toBe("highlighttext");
		const hovered = forced(selectedLayout(), { hover: true });
		expect(hovered["background-color"]).toBe("highlight");
		expect(hovered.color).toBe("highlighttext");
	});

	it("draws the label without the Canvas backplate, in HighlightText", async () => {
		await show(ACCOUNT, ACCOUNT.entity.id);
		const label = forced(selectedLayout().querySelector(".fui-TreeItemLayout__main")!);
		expect(label["forced-color-adjust"]).toBe("none");
		expect(label.color).toBe("highlighttext");
	});

	it("draws the chevron and the node icon in HighlightText, over a tinted icon's inline colour", async () => {
		await show(ACCOUNT, ACCOUNT.entity.id);
		const layout = selectedLayout();
		expect(forced(layout.querySelector(".fui-TreeItemLayout__expandIcon")!).color).toBe(
			"highlighttext !important",
		);
		expect(forced(layout.querySelector(".fui-TreeItemLayout__iconBefore svg")!).color).toBe(
			"highlighttext !important",
		);
	});

	it("draws the '(no attributes)' hint in HighlightText too, not its own grey (review)", async () => {
		const query = parse(
			`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="contactid" to="primarycontactid" alias="pc" /></entity></fetch>`,
		);
		await show(query, query.entity.links[0].id);
		const main = selectedLayout().querySelector(".fui-TreeItemLayout__main")!;
		const hint = [...main.querySelectorAll("span")].find((s) => s.textContent?.trim());
		expect(hint?.textContent).toMatch(/attribute/i);
		const declared = forced(hint!);
		expect(declared["forced-color-adjust"]).toBe("none");
		expect(declared.color).toBe("highlighttext");
	});

	it("draws Add and Delete in HighlightText, and leaves their borders to Fluent's focus and hover styles", async () => {
		await show(ACCOUNT, ACCOUNT.entity.id);
		// Fluent renders a row's actions only while it's hovered or focused
		const item = selectedLayout().parentElement!;
		await act(async () => {
			item.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
			item.dispatchEvent(new FocusEvent("focus", { bubbles: false }));
			item.focus();
		});
		const buttons = [...selectedLayout().querySelectorAll(".actions button")];
		expect(buttons.length).toBeGreaterThan(0);
		for (const button of buttons) {
			const declared = forced(button);
			expect(declared.color).toBe("highlighttext");
			expect(Object.keys(declared).filter((k) => k.startsWith("border"))).toEqual([]);
		}
	});
});
