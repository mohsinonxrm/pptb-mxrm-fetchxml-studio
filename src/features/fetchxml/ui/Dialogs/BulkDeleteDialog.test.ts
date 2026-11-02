/**
 * The Bulk delete dialog stays mounted for the whole session; its default job name is made when
 * it opens, for the entity and time of that job, not when the tool started (#182).
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { BulkDeleteDialog, type BulkDeleteDialogProps } from "./BulkDeleteDialog";
import { defaultBulkDeleteJobName } from "./bulkDeleteJobName";

let root: Root;
let container: HTMLDivElement;

const props = (over: Partial<BulkDeleteDialogProps>): BulkDeleteDialogProps => ({
	open: false,
	recordCount: 3,
	entityDisplayName: "Account",
	entityLogicalName: "account",
	onClose: () => {},
	onConfirm: async () => ({ asyncOperationId: "x", jobUrl: "" }),
	...over,
});

const render = (over: Partial<BulkDeleteDialogProps>) =>
	act(async () => {
		root.render(createElement(FluentProvider, { theme: webLightTheme }, createElement(BulkDeleteDialog, props(over))));
		await Promise.resolve();
	});

const jobNameInput = () => {
	const label = [...document.querySelectorAll("label")].find((l) => /job name/i.test(l.textContent ?? ""));
	const input = label?.htmlFor ? (document.getElementById(label.htmlFor) as HTMLInputElement | null) : null;
	return input ?? (document.querySelector("input[type=text], input:not([type])") as HTMLInputElement | null);
};

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	vi.useFakeTimers({ toFake: ["Date"] });
	vi.spyOn(console, "warn").mockImplementation(() => {});
	container = document.createElement("div");
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	vi.useRealTimers();
	vi.restoreAllMocks();
});

describe("BulkDeleteDialog job name (#182)", () => {
	it("is made when the dialog opens, for the entity of that job", async () => {
		const toolStart = new Date(2026, 10, 2, 8, 0, 0);
		vi.setSystemTime(toolStart);
		await render({ open: false, entityDisplayName: "Account" });

		const opened = new Date(2026, 10, 2, 9, 45, 0);
		vi.setSystemTime(opened);
		await render({ open: true, entityDisplayName: "Contact" });
		expect(jobNameInput()?.value).toBe(defaultBulkDeleteJobName("Contact", opened));

		// Closed and opened again later: a fresh name
		await render({ open: false, entityDisplayName: "Contact" });
		const reopened = new Date(2026, 10, 2, 10, 15, 0);
		vi.setSystemTime(reopened);
		await render({ open: true, entityDisplayName: "Contact" });
		expect(jobNameInput()?.value).toBe(defaultBulkDeleteJobName("Contact", reopened));
	});
});
