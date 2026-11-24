/**
 * The MessageBar above the tree (#277): a notice belongs to the query it was raised for, by its root
 * node, and Dismiss clears it.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { button, cleanup, click, render } from "../../shared/testing/ui";
import { useTreeNotice, type TreeNotice, type TreeNoticeInput } from "./useTreeNotice";
import { TreeNoticeBar } from "../TreeNoticeBar";

afterEach(() => cleanup());

let latest: TreeNotice;
function Probe({ rootId }: { rootId: string | undefined }) {
	latest = useTreeNotice(rootId);
	return createElement(TreeNoticeBar, { notice: latest.notice, onDismiss: latest.dismiss });
}
const show = (rootId: string | undefined) => render(createElement(Probe, { rootId }));

const changed: TreeNoticeInput = {
	intent: "warning",
	title: "View loaded with changes",
	body: 'The tree can\'t represent everything in "My accounts".',
};
const bars = () =>
	[...document.querySelectorAll(".fui-MessageBar")].map((b) =>
		(b.textContent ?? "").replace(/\s+/g, " ").trim(),
	);

describe("useTreeNotice", () => {
	it("shows a notice for the query in the tree, and keeps it while that query is edited", async () => {
		await show("root-1");
		await act(async () => latest.show(changed, "root-1"));
		expect(latest.notice).toEqual(changed);
		// The title ends with a visually hidden full stop, so it isn't read into the text (BarTitle, #186)
		expect(bars()).toEqual([
			'View loaded with changes. The tree can\'t represent everything in "My accounts".',
		]);
		await show("root-1");
		expect(latest.notice).toEqual(changed);
	});

	it("goes when another query is in the tree, and doesn't come back", async () => {
		await show("root-1");
		await act(async () => latest.show(changed, "root-1"));
		await show("root-2");
		expect(latest.notice).toBeNull();
		expect(bars()).toEqual([]);
		await show("root-1");
		expect(latest.notice).toBeNull();
	});

	it("a notice with no query (a launch failure) goes when a query starts", async () => {
		await show(undefined);
		await act(async () =>
			latest.show({ intent: "error", title: "Couldn't open the view", body: "x" }, undefined),
		);
		expect(latest.notice?.title).toBe("Couldn't open the view");
		await show("root-1");
		expect(latest.notice).toBeNull();
		await show(undefined);
		expect(latest.notice).toBeNull();
	});

	it("Dismiss clears it", async () => {
		await show("root-1");
		await act(async () => latest.show(changed, "root-1"));
		await click(document.querySelector('[aria-label="Dismiss"]'));
		expect(latest.notice).toBeNull();
		expect(bars()).toEqual([]);
		expect(button("Dismiss")).toBeUndefined();
	});
});
