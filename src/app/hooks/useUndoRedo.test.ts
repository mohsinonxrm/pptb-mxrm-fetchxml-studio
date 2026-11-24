/**
 * Undo and Redo with an announcement (#192): "Undone: {step}" / "Redone: {step}" in the app
 * announcer's polite region (#186), once the presses stop; the same step twice is read twice;
 * nothing happens or is announced with nothing to undo or redo; the functions keep their identity
 * (#65). The announcer's own clearing is tested in Announcer.test.ts.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { BuilderProvider } from "../../features/fetchxml/state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import { ANNOUNCE_DELAY_MS, useUndoRedo, type UndoRedo } from "./useUndoRedo";
import { AnnouncerProvider } from "../../shared/feedback/Announcer";
import { FLUSH_MS } from "../../shared/feedback/useAnnouncer";

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
	await render(
		createElement(
			AnnouncerProvider,
			null,
			createElement(BuilderProvider, null, createElement(Probe)),
		),
	);
	await act(async () => builder.loadFetchXml(XML, { baseline: true }));
};
const run = (fn: () => void) => act(async () => fn());
const wait = (ms: number) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));
const settle = () => wait(ANNOUNCE_DELAY_MS + FLUSH_MS + 60);
/** What the announcer's polite region says now */
const said = () => document.querySelector('[data-fxs-announcer="polite"]')?.textContent ?? "";

describe("useUndoRedo", () => {
	it("undoes and redoes, announcing the step once the presses stop", async () => {
		await show();
		await run(() => builder.addFilter(builder.fetchQuery!.entity.id));
		await run(() => undoRedo.undo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(0);
		expect(said()).toBe("");
		await settle();
		expect(said()).toBe("Undone: Add filter");
		await run(() => undoRedo.redo());
		expect(builder.fetchQuery!.entity.filters).toHaveLength(1);
		await settle();
		expect(said()).toBe("Redone: Add filter");
	});

	it("announces only the last step of quick presses in a row", async () => {
		await show();
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => builder.addFilter(entityId));
		await run(() => builder.addOrder(entityId, "name"));
		await run(() => undoRedo.undo());
		await wait(ANNOUNCE_DELAY_MS / 3);
		expect(said()).toBe("");
		await run(() => undoRedo.undo());
		// When the first press would have spoken, the second has put it off
		await wait((ANNOUNCE_DELAY_MS * 3) / 4);
		expect(said()).toBe("");
		await settle();
		expect(said()).toBe("Undone: Add filter");
		expect(builder.fetchQuery!.entity.orders).toHaveLength(0);
	});

	it("reads the same step twice: the text is cleared before the second reading", async () => {
		await show();
		const entityId = builder.fetchQuery!.entity.id;
		await run(() => builder.addFilter(entityId));
		await run(() => builder.addFilter(entityId));
		await run(() => undoRedo.undo());
		await settle();
		expect(said()).toBe("Undone: Add filter");
		await run(() => undoRedo.undo());
		// Once the presses stop, the announcer clears the text, then reads it again
		await wait(ANNOUNCE_DELAY_MS + FLUSH_MS / 3);
		expect(said()).toBe("");
		await settle();
		expect(said()).toBe("Undone: Add filter");
	});

	it("does nothing, and says nothing, with nothing to undo or redo", async () => {
		await show();
		const before = builder.fetchQuery;
		await run(() => undoRedo.undo());
		await run(() => undoRedo.redo());
		await settle();
		expect(builder.fetchQuery).toBe(before);
		expect(said()).toBe("");
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
