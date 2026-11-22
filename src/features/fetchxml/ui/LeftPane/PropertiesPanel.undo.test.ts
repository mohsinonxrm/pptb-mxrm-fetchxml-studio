/**
 * Property edits that change several nodes are one undo step (#192): an attribute's alias rename
 * with the orders that sort by it, and turning aggregate off with its clean-up plan.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { PropertiesPanel } from "./PropertiesPanel";
import { BuilderProvider } from "../../state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../state/builderContext";
import type { FetchNode } from "../../model/nodes";
import {
	button,
	cleanup,
	click,
	flush,
	installHost,
	removeHost,
	render,
	typeInto,
} from "../../../../shared/testing/ui";

afterEach(() => {
	cleanup();
	removeHost();
});

let builder: BuilderContextValue;
function Harness() {
	const b = useBuilder();
	builder = b;
	return createElement(PropertiesPanel, {
		selectedNode: b.selectedNode,
		fetchQuery: b.fetchQuery,
		onNodeUpdate: b.updateNode,
	});
}

async function loadAndSelect(xml: string, pick: (fetch: FetchNode) => { id: string }) {
	installHost();
	await render(createElement(BuilderProvider, null, createElement(Harness)));
	await act(async () => builder.loadFetchXml(xml, { baseline: true }));
	await act(async () => builder.selectNode(pick(builder.fetchQuery!).id));
	await flush(3);
}
const inputWithValue = (value: string) =>
	[...document.querySelectorAll<HTMLInputElement>("input")].find((i) => i.value === value);

describe("one undo step for a property edit that changes several nodes", () => {
	it("an attribute's alias rename, with the order that sorts by it", async () => {
		await loadAndSelect(
			`<fetch aggregate="true"><entity name="account">
				<attribute name="name" alias="n" groupby="true" />
				<order alias="n" />
			</entity></fetch>`,
			(f) => f.entity.attributes[0],
		);
		expect(builder.history.canUndo).toBe(false);
		await typeInto(inputWithValue("n"), "nm");
		await typeInto(inputWithValue("nm"), "nme");
		expect(builder.fetchQuery!.entity.attributes[0].alias).toBe("nme");
		expect(builder.fetchQuery!.entity.orders[0].alias).toBe("nme");
		// Typing merges, and the order follows in the same step
		expect(builder.history.undoLabel).toBe("Edit attribute");
		await act(async () => builder.history.undo());
		expect(builder.fetchQuery!.entity.attributes[0].alias).toBe("n");
		expect(builder.fetchQuery!.entity.orders[0].alias).toBe("n");
		expect(builder.history.canUndo).toBe(false);
	});

	it("turning aggregate off and updating the attributes", async () => {
		await loadAndSelect(
			`<fetch aggregate="true"><entity name="account">
				<attribute name="accountid" alias="total" aggregate="count" />
				<attribute name="name" alias="n" groupby="true" />
			</entity></fetch>`,
			(f) => f,
		);
		const checkbox = [
			...document.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'),
		].find(
			(c) => c.closest("label, div")?.textContent?.includes("Aggregate query") && c.checked,
		);
		await click(checkbox);
		await click(button("Turn off and update", document));
		const fetch = builder.fetchQuery!;
		expect(fetch.options.aggregate).toBeUndefined();
		expect(fetch.entity.attributes.some((a) => a.aggregate || a.groupby)).toBe(false);
		expect(builder.history.undoLabel).toBe("Turn off aggregate");
		await act(async () => builder.history.undo());
		const restored = builder.fetchQuery!;
		expect(restored.options.aggregate).toBe(true);
		expect(restored.entity.attributes.map((a) => a.aggregate ?? a.groupby)).toEqual([
			"count",
			true,
		]);
		expect(builder.history.canUndo).toBe(false);
	});
});
