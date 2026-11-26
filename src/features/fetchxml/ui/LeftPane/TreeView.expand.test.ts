/**
 * The tree's root after a query replaces the open one (#375): Open view and Parse to tree bring a
 * query with new node ids, and its root entity used to stay collapsed. It opens as on a fresh launch;
 * a root the user collapsed stays collapsed while they edit (the ids don't change).
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

const show = (fetchQuery: FetchNode, selectedNodeId: NodeId | null = null) =>
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
const ACCOUNT = `<fetch><entity name="account"><attribute name="name" /><filter><condition attribute="statecode" operator="eq" value="0" /></filter><order attribute="name" /></entity></fetch>`;
const CONTACT = `<fetch><entity name="contact"><attribute name="fullname" /><order attribute="fullname" /></entity></fetch>`;

const row = (id: NodeId) =>
	container.querySelector<HTMLElement>(`[data-fui-tree-item-value="${CSS.escape(id)}"]`);
const expanded = (id: NodeId) => row(id)?.getAttribute("aria-expanded");
/** Collapse a row the way a click on its chevron does */
const collapse = (id: NodeId) =>
	act(async () => {
		row(id)
			?.querySelector<HTMLElement>(
				":scope > .fui-TreeItemLayout .fui-TreeItemLayout__expandIcon",
			)
			?.click();
	});

describe("the tree's root after a query replaces the open one (#375)", () => {
	it("opens the root entity of a query that replaces the open one (Open view, Parse to tree)", async () => {
		const first = parse(ACCOUNT);
		await show(first);
		expect(expanded(first.entity.id)).toBe("true");

		const replaced = parse(CONTACT);
		await show(replaced, replaced.entity.id);
		expect(expanded(replaced.id)).toBe("true");
		expect(expanded(replaced.entity.id)).toBe("true");
		// Its nodes show straight away
		expect(container.textContent).toContain("fullname");
	});

	it("opens it even when the replacing query is the same FetchXML parsed again", async () => {
		const first = parse(ACCOUNT);
		await show(first);
		await collapse(first.entity.id);
		expect(expanded(first.entity.id)).toBe("false");
		const again = parse(ACCOUNT);
		await show(again);
		expect(expanded(again.entity.id)).toBe("true");
	});

	it("opens a new root entity in the same query", async () => {
		const first = parse(ACCOUNT);
		await show(first);
		const contact = parse(CONTACT).entity;
		await show({ ...first, entity: contact });
		expect(expanded(contact.id)).toBe("true");
	});

	it("keeps a root the user collapsed collapsed while they edit (same ids)", async () => {
		const first = parse(ACCOUNT);
		await show(first);
		await collapse(first.entity.id);
		expect(expanded(first.entity.id)).toBe("false");
		// An edit: a new FetchNode object with the same ids
		const edited: FetchNode = {
			...first,
			options: { ...first.options, top: 10 },
			entity: { ...first.entity, attributes: [...first.entity.attributes] },
		};
		await show(edited);
		expect(expanded(first.entity.id)).toBe("false");
	});
});
