/**
 * The provider's undo API (#192): useBuilder().history follows the steps, its functions and the
 * builder's actions keep their identity (#65), transactions group dispatches, a baseline load
 * clears the history, and an Undo toast's step is undone only while it's the last one.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "../../../shared/testing/ui";
import { BuilderProvider } from "./builderStore";
import { useBuilder, type BuilderContextValue } from "./builderContext";

afterEach(() => cleanup());

let builder: BuilderContextValue;
function Probe() {
	builder = useBuilder();
	return null;
}
const show = () => render(createElement(BuilderProvider, null, createElement(Probe)));
const run = (fn: () => void) => act(async () => fn());
const XML =
	"<fetch><entity name='account'><attribute name='name'/><filter><condition attribute='name' operator='eq' value='a'/></filter></entity></fetch>";

describe("useBuilder().history", () => {
	it("follows the steps: labels, can undo / redo, the last step's id", async () => {
		await show();
		expect(builder.history).toMatchObject({
			canUndo: false,
			canRedo: false,
			undoLabel: null,
			redoLabel: null,
			lastEntryId: null,
		});
		await run(() => builder.loadFetchXml(XML));
		expect(builder.history.undoLabel).toBe("Parse to tree");
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		expect(builder.history.undoLabel).toBe("Add filter");
		expect(builder.history.lastEntryId).toEqual(expect.any(Number));
		await run(() => builder.history.undo());
		expect(builder.history).toMatchObject({
			canUndo: true,
			canRedo: true,
			undoLabel: "Parse to tree",
			redoLabel: "Add filter",
		});
		await run(() => builder.history.redo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(2);
	});

	it("keeps the actions and history functions the same across renders, undo and redo", async () => {
		await show();
		const first = builder;
		await run(() => builder.loadFetchXml(XML));
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => builder.history.undo());
		await run(() => builder.history.redo());
		for (const key of ["addFilter", "updateNode", "loadFetchXml", "selectNode"] as const)
			expect(builder[key]).toBe(first[key]);
		for (const key of ["undo", "redo", "undoIfLast", "transaction"] as const)
			expect(builder.history[key]).toBe(first.history[key]);
	});

	it("groups a transaction's dispatches into one step", async () => {
		await show();
		await run(() => builder.loadFetchXml(XML, { baseline: true }));
		const entityId = builder.fetchQuery!.entity.id;
		await run(() =>
			builder.history.transaction("Add two filters", () => {
				builder.addFilter(entityId);
				builder.addFilter(entityId);
			}),
		);
		expect(builder.fetchQuery!.entity.filters).toHaveLength(3);
		expect(builder.history.undoLabel).toBe("Add two filters");
		await run(() => builder.history.undo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(1);
		expect(builder.history.canUndo).toBe(false);
	});

	it("closes a transaction whose function throws", async () => {
		await show();
		await run(() => builder.loadFetchXml(XML, { baseline: true }));
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => {
			expect(() =>
				builder.history.transaction("Broken", () => {
					builder.addFilter(entityId);
					throw new Error("boom");
				}),
			).toThrow("boom");
		});
		// Still one step, and undo works (no transaction left open)
		expect(builder.history.undoLabel).toBe("Broken");
		await run(() => builder.history.undo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(1);
	});

	it("a baseline load clears it; New clears it", async () => {
		await show();
		await run(() => builder.loadFetchXml(XML));
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => builder.loadFetchXml(XML, { baseline: true }));
		expect(builder.history.canUndo).toBe(false);
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => builder.newQuery());
		expect(builder.history.canUndo).toBe(false);
		expect(builder.fetchQuery).toBeNull();
	});

	it("undoes an Undo toast's step only while it's the last one", async () => {
		await show();
		await run(() => builder.loadFetchXml(XML, { baseline: true }));
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => builder.addFilter(entityId));
		const id = builder.history.lastEntryId!;
		await run(() => builder.addOrder(entityId, "name"));
		await run(() => builder.history.undoIfLast(id));
		expect(builder.fetchQuery!.entity.orders).toHaveLength(1);
		await run(() => builder.history.undo());
		await run(() => builder.history.undoIfLast(id));
		expect(builder.fetchQuery!.entity.filters).toHaveLength(1);
	});
});
