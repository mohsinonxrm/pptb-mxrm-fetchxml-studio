/**
 * A save that waits for the synced tree before its dialog opens (#102), and how it ends (#373): it
 * opens once the new tree and its entity metadata arrive; a newer save replaces it; and if the tree
 * moves on to other FetchXML first, it's dropped and reported as a cancel, so nothing waits for it.
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, flush, installHost, removeHost, render } from "../../../../shared/testing/ui";
import { EditorSyncContext, type EditorSyncResult } from "../RightPane/editorSync";
import { parseFetchXml } from "../../model/fetchxmlParser";
import { useSaveViewCommand, type SaveViewCommandOptions } from "./useSaveViewCommand";

const SYNCED = "<fetch><entity name='nosuchentity'><attribute name='name'/></entity></fetch>";
const OTHER = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";

let latest: ReturnType<typeof useSaveViewCommand>;
function Probe(props: SaveViewCommandOptions) {
	latest = useSaveViewCommand(props);
	return latest.dialog;
}

const base = (extra: Partial<SaveViewCommandOptions> = {}): SaveViewCommandOptions => ({
	fetchXml: OTHER,
	layoutXml: "",
	entityLogicalName: "account",
	objectTypeCode: 1,
	primaryIdAttribute: "accountid",
	targetView: null,
	...extra,
});

let syncResult: EditorSyncResult;
const show = (props: SaveViewCommandOptions) =>
	render(
		createElement(
			EditorSyncContext.Provider,
			{ value: { sync: async () => syncResult, editorHasContent: true } },
			createElement(Probe, props),
		),
	);
const dialogOpen = () => !!document.querySelector('[role="dialog"]');
const press = async () => {
	await act(async () => {
		latest.command.onClick?.();
	});
	await flush(5);
};

beforeEach(() => {
	installHost();
	vi.spyOn(console, "warn").mockImplementation(() => {});
	syncResult = {
		status: "synced",
		fetchXml: SYNCED,
		fetchQuery: parseFetchXml(SYNCED).fetchNode!,
	};
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

describe("a save waiting for the synced tree (#373)", () => {
	it("says when a save starts, before anything else", async () => {
		const onSaveStart = vi.fn();
		syncResult = { status: "unchanged" };
		await show(base({ onSaveStart }));
		await press();
		expect(onSaveStart).toHaveBeenCalledOnce();
		expect(dialogOpen()).toBe(true);
	});

	it("opens once the tree and its entity's metadata have arrived", async () => {
		const onCancel = vi.fn();
		await show(base({ onCancel }));
		await press();
		// The tree has the synced FetchXML, its metadata hasn't come yet
		await show(
			base({
				onCancel,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 0,
				primaryIdAttribute: "",
			}),
		);
		expect(dialogOpen()).toBe(false);
		await show(
			base({
				onCancel,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 5,
				primaryIdAttribute: "nosuchentityid",
			}),
		);
		expect(dialogOpen()).toBe(true);
		expect(onCancel).not.toHaveBeenCalled();
	});

	it("is dropped, as a cancel, when the tree moves on before the metadata comes", async () => {
		const onCancel = vi.fn();
		await show(base({ onCancel }));
		await press();
		await show(
			base({
				onCancel,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 0,
				primaryIdAttribute: "",
			}),
		);
		// Parse to tree with other FetchXML: the tree moved on
		await show(base({ onCancel, fetchXml: OTHER }));
		expect(onCancel).toHaveBeenCalledOnce();
		// The synced FetchXML coming back later, with its metadata, doesn't open a dialog nobody asked for
		await show(
			base({
				onCancel,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 5,
				primaryIdAttribute: "nosuchentityid",
			}),
		);
		expect(dialogOpen()).toBe(false);
	});

	it("isn't dropped while the tree hasn't reached the synced FetchXML yet", async () => {
		const onCancel = vi.fn();
		await show(base({ onCancel }));
		await press();
		// Still the old tree for a render or two
		await show(base({ onCancel, fetchXml: OTHER }));
		expect(onCancel).not.toHaveBeenCalled();
		await show(
			base({
				onCancel,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 5,
				primaryIdAttribute: "nosuchentityid",
			}),
		);
		expect(dialogOpen()).toBe(true);
	});

	it("gives way to a newer save", async () => {
		const onSaveStart = vi.fn();
		await show(base({ onSaveStart }));
		await press();
		await show(
			base({
				onSaveStart,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 0,
				primaryIdAttribute: "",
			}),
		);
		// A later Save of the tree as it is: it opens at once
		syncResult = { status: "unchanged" };
		await press();
		expect(onSaveStart).toHaveBeenCalledTimes(2);
		expect(dialogOpen()).toBe(true);
		// Closed without saving; then the entity's details arrive late: the earlier, replaced save
		// doesn't open the dialog again
		await act(async () => {
			document
				.querySelector<HTMLElement>('[role="dialog"]')
				?.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
		});
		await flush(5);
		expect(dialogOpen()).toBe(false);
		await show(
			base({
				onSaveStart,
				fetchXml: SYNCED,
				entityLogicalName: "nosuchentity",
				objectTypeCode: 5,
				primaryIdAttribute: "nosuchentityid",
			}),
		);
		expect(dialogOpen()).toBe(false);
	});
});
