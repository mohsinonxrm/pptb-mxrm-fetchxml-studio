/**
 * The selected tree node in Windows high contrast (#361). jsdom has no forced colours, so these tests
 * read the rules Griffel inserts for the selected row: under `@media (forced-colors: active)` the
 * label opts out of the Canvas backplate (it read black on black) and every part of the row reads
 * HighlightText on Highlight, hover included.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TreeView } from "./TreeView";
import { parseFetchXml } from "../../model/fetchxmlParser";

let root: Root;
let container: HTMLDivElement;

const query = parseFetchXml(
	`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`,
).fetchNode!;

beforeEach(async () => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	Element.prototype.scrollIntoView = () => {};
	container = document.createElement("div");
	document.body.appendChild(container);
	root = createRoot(container);
	await act(async () => {
		root.render(
			createElement(
				FluentProvider,
				{ theme: webLightTheme },
				createElement(TreeView, {
					fetchQuery: query,
					selectedNodeId: query.entity.id,
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
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

/** Style rules inside `@media (forced-colors: active)` that apply to the selected row's layout */
function forcedColorRules(): CSSStyleRule[] {
	const layout = document.querySelector<HTMLElement>(
		'[role="treeitem"][aria-selected="true"] > .fui-TreeItemLayout',
	);
	expect(layout).not.toBeNull();
	const own = [...layout!.classList].filter((c) => !c.startsWith("fui-"));
	const rules: CSSStyleRule[] = [];
	for (const sheet of [...document.styleSheets]) {
		for (const rule of [...sheet.cssRules]) {
			if (!(rule instanceof CSSMediaRule) || !rule.conditionText.includes("forced-colors"))
				continue;
			for (const inner of [...rule.cssRules]) {
				if (
					inner instanceof CSSStyleRule &&
					own.some((c) => inner.selectorText.includes(`.${c}`))
				)
					rules.push(inner);
			}
		}
	}
	return rules;
}

/** The declarations of the forced-colours rules whose selector matches, merged (values in lowercase, as Griffel writes system colours) */
function declarations(match: (selector: string) => boolean): Record<string, string> {
	const out: Record<string, string> = {};
	for (const rule of forcedColorRules().filter((r) => match(r.selectorText))) {
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
	it("is HighlightText on Highlight", () => {
		const row = declarations((s) => !s.includes(" ") && !s.includes(":"));
		expect(row["background-color"]).toBe("highlight");
		expect(row.color).toBe("highlighttext");
	});

	it("keeps the highlight while hovered", () => {
		const hover = declarations((s) => s.includes(":hover") && !s.includes(" "));
		expect(hover["background-color"]).toBe("highlight");
		expect(hover.color).toBe("highlighttext");
	});

	it("draws the label without the Canvas backplate, in HighlightText", () => {
		const label = declarations((s) => s.includes(".fui-TreeItemLayout__main"));
		expect(label["forced-color-adjust"]).toBe("none");
		expect(label.color).toBe("highlighttext");
	});

	it("draws the chevron and the node icon in HighlightText, over a tinted icon's inline colour", () => {
		const icons = declarations(
			(s) =>
				s.includes(".fui-TreeItemLayout__expandIcon") &&
				s.includes(".fui-TreeItemLayout__iconBefore svg"),
		);
		expect(icons.color).toBe("highlighttext !important");
	});

	it("draws Add and Delete in HighlightText", () => {
		const actions = declarations((s) => s.includes(".actions button"));
		expect(actions.color).toBe("highlighttext");
		expect(actions["border-top-color"]).toBe("highlighttext");
		expect(actions["border-left-color"]).toBe("highlighttext");
	});
});
