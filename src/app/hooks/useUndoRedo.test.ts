/**
 * Undo and Redo with an announcement (#192): "Undone: {step}" / "Redone: {step}" for a polite live
 * region, once the presses stop; cleared on each press so the same text twice is read twice;
 * nothing happens or is announced with nothing to undo or redo; the functions keep their identity
 * (#65).
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { BuilderProvider } from "../../features/fetchxml/state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import { ANNOUNCE_DELAY_MS, CLEAR_AFTER_MS, useUndoRedo, type UndoRedo } from "./useUndoRedo";

afterEach(() => cleanup());

let builder: BuilderContextValue;
let undoRedo: UndoRedo;
function Probe() {
	builder = useBuilder();
	undoRedo = useUndoRedo(builder.history);
	return null;
}
const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";

const show = async () => {
	await render(createElement(BuilderProvider, null, createElement(Probe)));
	await act(async () => builder.loadFetchXml(XML, { baseline: true }));
};
const run = (fn: () => void) => act(async () => fn());
const wait = (ms: number) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));
const settle = () => wait(ANNOUNCE_DELAY_MS + 50);

describe("useUndoRedo", () => {
	it("undoes and redoes, announcing the step once the presses stop", async () => {
		await show();
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => undoRedo.undo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(0);
		expect(undoRedo.announcement).toBe("");
		await settle();
		expect(undoRedo.announcement).toBe("Undone: Add filter");
		await run(() => undoRedo.redo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(1);
		expect(undoRedo.announcement).toBe("");
		await settle();
		expect(undoRedo.announcement).toBe("Redone: Add filter");
	});

	it("announces only the last step of quick presses in a row", async () => {
		await show();
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => builder.addFilter(entityId));
		await run(() => builder.addOrder(entityId, "name"));
		await run(() => undoRedo.undo());
		await wait(ANNOUNCE_DELAY_MS / 3);
		expect(undoRedo.announcement).toBe("");
		await run(() => undoRedo.undo());
		// When the first press would have spoken, the second has put it off
		await wait((ANNOUNCE_DELAY_MS * 3) / 4);
		expect(undoRedo.announcement).toBe("");
		await settle();
		expect(undoRedo.announcement).toBe("Undone: Add filter");
		expect(builder.fetchQuery!.entity.orders).toHaveLength(0);
	});

	it("clears the text on each press, so the same step twice is read twice", async () => {
		await show();
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => builder.addFilter(entityId));
		await run(() => builder.addFilter(entityId));
		await run(() => undoRedo.undo());
		await settle();
		expect(undoRedo.announcement).toBe("Undone: Add filter");
		await run(() => undoRedo.undo());
		expect(undoRedo.announcement).toBe("");
		await settle();
		expect(undoRedo.announcement).toBe("Undone: Add filter");
	});

	it("clears the text a while after it was read", async () => {
		vi.useFakeTimers();
		try {
			await show();
			await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
			await run(() => undoRedo.undo());
			await act(async () => vi.advanceTimersByTime(ANNOUNCE_DELAY_MS));
			expect(undoRedo.announcement).toBe("Undone: Add filter");
			await act(async () => vi.advanceTimersByTime(CLEAR_AFTER_MS - 10));
			expect(undoRedo.announcement).toBe("Undone: Add filter");
			await act(async () => vi.advanceTimersByTime(10));
			expect(undoRedo.announcement).toBe("");
		} finally {
			vi.useRealTimers();
		}
	});

	it("does nothing, and says nothing, with nothing to undo or redo", async () => {
		await show();
		const before = builder.fetchQuery;
		await run(() => undoRedo.undo());
		await run(() => undoRedo.redo());
		await settle();
		expect(builder.fetchQuery).toBe(before);
		expect(undoRedo.announcement).toBe("");
	});

	it("keeps the same functions across steps", async () => {
		await show();
		const first = undoRedo;
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => undoRedo.undo());
		await settle();
		expect(undoRedo.undo).toBe(first.undo);
		expect(undoRedo.redo).toBe(first.redo);
	});
});
