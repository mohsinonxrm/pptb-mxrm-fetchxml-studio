/**
 * Undo and Redo with an announcement (#192): "Undone: {step}" / "Redone: {step}", polite and in one
 * batch; nothing happens or is announced with nothing to undo or redo; the functions keep their
 * identity (#65).
 */
import { act, createElement, type ReactNode } from "react";
import { AnnounceProvider } from "@fluentui/react-components";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { BuilderProvider } from "../../features/fetchxml/state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import { useUndoRedo } from "./useUndoRedo";

afterEach(() => cleanup());

let builder: BuilderContextValue;
let undoRedo: ReturnType<typeof useUndoRedo>;
function Probe() {
	builder = useBuilder();
	undoRedo = useUndoRedo(builder.history);
	return null;
}
const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";

const show = async () => {
	const announce = vi.fn();
	const tree = (children: ReactNode) =>
		createElement(AnnounceProvider, { value: { announce } }, children);
	await render(tree(createElement(BuilderProvider, null, createElement(Probe))));
	await act(async () => builder.loadFetchXml(XML, { baseline: true }));
	return announce;
};
const run = (fn: () => void) => act(async () => fn());

describe("useUndoRedo", () => {
	it("undoes and redoes, announcing the step politely", async () => {
		const announce = await show();
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => undoRedo.undo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(0);
		expect(announce).toHaveBeenLastCalledWith("Undone: Add filter", {
			polite: true,
			batchId: "fxs-undo-redo",
		});
		await run(() => undoRedo.redo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(1);
		expect(announce).toHaveBeenLastCalledWith("Redone: Add filter", {
			polite: true,
			batchId: "fxs-undo-redo",
		});
	});

	it("names each step in turn", async () => {
		const announce = await show();
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => builder.addFilter(entityId));
		await run(() => builder.addOrder(entityId, "name"));
		await run(() => {
			undoRedo.undo();
		});
		await run(() => {
			undoRedo.undo();
		});
		expect(announce.mock.calls.map((c) => c[0])).toEqual([
			"Undone: Add order",
			"Undone: Add filter",
		]);
	});

	it("does nothing, and says nothing, with nothing to undo or redo", async () => {
		const announce = await show();
		const before = builder.fetchQuery;
		await run(() => undoRedo.undo());
		await run(() => undoRedo.redo());
		expect(builder.fetchQuery).toBe(before);
		expect(announce).not.toHaveBeenCalled();
	});

	it("keeps the same functions across steps", async () => {
		await show();
		const first = undoRedo;
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => undoRedo.undo());
		expect(undoRedo.undo).toBe(first.undo);
		expect(undoRedo.redo).toBe(first.redo);
	});
});
