/**
 * Run's reason in the editor session (#278): no query, an empty editor in editor mode, the quick
 * find value (only outside editor mode, #142), and a query that's running or loading more rows.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import {
	EMPTY_EDITOR_REASON,
	RUNNING_REASON,
	useEditorSession,
	type EditorSession,
	type EditorSessionInputs,
} from "./useEditorSession";
import { NO_QUERY_REASON } from "../../features/fetchxml/ui/Toolbar/useSaveViewCommand";

let session: EditorSession;
function Probe(props: EditorSessionInputs) {
	session = useEditorSession(props);
	return null;
}
const show = (props: Partial<EditorSessionInputs> = {}) =>
	render(createElement(Probe, { xml: "<fetch/>", ...props }));

afterEach(() => cleanup());

describe("Run's reason", () => {
	it("is none for a query, and 'Choose an entity first' without one", async () => {
		await show();
		expect(session.runDisabledReason).toBeUndefined();
		await show({ xml: "" });
		expect(session.runDisabledReason).toBe(NO_QUERY_REASON);
	});

	it("is the quick find reason outside editor mode only", async () => {
		await show({ runDisabledReason: "Enter a quick find value first" });
		expect(session.runDisabledReason).toBe("Enter a quick find value first");
		await act(async () => session.onEditorStateChange(true, "<fetch/>"));
		expect(session.runDisabledReason).toBeUndefined();
	});

	it("asks for FetchXML when the editor is empty in editor mode", async () => {
		await show();
		await act(async () => session.onEditorStateChange(true, "  "));
		expect(session.runDisabledReason).toBe(EMPTY_EDITOR_REASON);
		await act(async () => session.onEditorStateChange(true, "<fetch/>"));
		expect(session.runDisabledReason).toBeUndefined();
	});

	it("says it's running while a query runs or more rows load, and Run shows busy", async () => {
		await show({ isRunning: true });
		expect(session.runDisabledReason).toBe(RUNNING_REASON);
		expect(session.isRunning).toBe(true);
		await show({ isLoadingMore: true });
		expect(session.runDisabledReason).toBe(RUNNING_REASON);
		expect(session.isRunning).toBe(true);
		await show();
		expect(session.isRunning).toBe(false);
	});
});
