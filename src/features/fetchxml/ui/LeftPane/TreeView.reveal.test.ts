/**
 * Advisories "Go to node" on the node that is already selected must still reveal it: the tree
 * re-runs expand + scrollIntoView when revealSeq changes, not only when the selection does.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TreeView } from "./TreeView";
import { parseFetchXml } from "../../model/fetchxmlParser";

let root: Root;
let container: HTMLDivElement;
const scrollIntoView = vi.fn();

const query = parseFetchXml(
	`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`,
).fetchNode!;

const render = (revealSeq: number) =>
	act(async () => {
		root.render(
			createElement(
				FluentProvider,
				{ theme: webLightTheme },
				createElement(TreeView, {
					fetchQuery: query,
					selectedNodeId: query.entity.id,
					revealSeq,
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
		await Promise.resolve();
	});

const settle = () =>
	act(async () => {
		await vi.advanceTimersByTimeAsync(300);
	});

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	vi.useFakeTimers();
	scrollIntoView.mockClear();
	Element.prototype.scrollIntoView = scrollIntoView;
	container = document.createElement("div");
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	vi.useRealTimers();
});

describe("TreeView reveal", () => {
	it("scrolls the selected node into view again when a reveal is requested for it", async () => {
		await render(0);
		await settle();
		expect(scrollIntoView).toHaveBeenCalledTimes(1);

		// Same selection, same tree: nothing more happens
		await render(0);
		await settle();
		expect(scrollIntoView).toHaveBeenCalledTimes(1);

		// Go to node on the already-selected node
		await render(1);
		await settle();
		expect(scrollIntoView).toHaveBeenCalledTimes(2);
	});
});
