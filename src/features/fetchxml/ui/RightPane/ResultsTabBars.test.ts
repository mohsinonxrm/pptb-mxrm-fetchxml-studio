/**
 * The Results tab's bars (#277): Export failed offers Try again only when it can help, and the
 * group is most severe first.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { button, cleanup, click, render } from "../../../../shared/testing/ui";
import { ResultsTabBars, type ResultsTabBarsProps } from "./ResultsTabBars";

afterEach(() => cleanup());

const titles = () =>
	[...document.querySelectorAll(".fui-MessageBarTitle")].map((t) => t.textContent?.trim());
const show = (props: Partial<ResultsTabBarsProps>) =>
	render(createElement(ResultsTabBars, { rowCount: 0, ...props }));

describe("ResultsTabBars", () => {
	it("nothing to say: no bars at all", async () => {
		await show({});
		expect(document.querySelector(".fui-MessageBarGroup")).toBeNull();
	});

	it("an export that failed offers Try again, which repeats it", async () => {
		const onRetryExport = vi.fn();
		await show({
			exportError: { kind: "local", message: "Disk full", retryable: true },
			onRetryExport,
		});
		expect(titles()).toEqual(["Export failed"]);
		await click(button("Try again", document));
		expect(onRetryExport).toHaveBeenCalledTimes(1);
	});

	it("an export that couldn't start doesn't offer Try again", async () => {
		await show({
			exportError: {
				kind: "server",
				message: "Export to Excel requires a saved view.",
				retryable: false,
			},
			onRetryExport: vi.fn(),
		});
		expect(titles()).toEqual(["Export failed"]);
		expect(button("Try again", document)).toBeUndefined();
	});

	it("while exporting, the failure from before isn't shown", async () => {
		await show({
			isExporting: true,
			exportError: { kind: "local", message: "Disk full", retryable: true },
		});
		expect(titles()).toEqual(["Exporting"]);
	});

	it("Dismiss on the failed run", async () => {
		const onDismissExecuteError = vi.fn();
		await show({ executeError: "0x80040203: bad value", onDismissExecuteError });
		await click(document.querySelector('[aria-label="Dismiss"]'));
		expect(onDismissExecuteError).toHaveBeenCalledTimes(1);
	});

	it("most severe first: the failed run, the failed export, paging, exporting", async () => {
		await show({
			executeError: "boom",
			exportError: { kind: "local", message: "Disk full", retryable: true },
			pagingNotice: { kind: "error", message: "Throttled", page: 2 },
			rowCount: 50,
		});
		expect(titles()).toEqual([
			"The query didn't run",
			"Export failed",
			"Couldn't load all records",
		]);
	});
});
