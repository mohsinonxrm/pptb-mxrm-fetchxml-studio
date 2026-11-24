/**
 * The Results tab's bars (#277): Export failed offers Try again only when it can help, and the
 * group is most severe first.
 */
import { createElement } from "react";
import { AnnounceProvider } from "@fluentui/react-components";
import { afterEach, describe, expect, it, vi } from "vitest";
import { button, cleanup, click, render } from "../../../../shared/testing/ui";
import { ResultsTabBars, type ResultsTabBarsProps } from "./ResultsTabBars";

afterEach(() => cleanup());

/** The titles as shown: without BarTitle's visually hidden full stop (#186) */
const titles = () =>
	[...document.querySelectorAll(".fui-MessageBarTitle")].map((t) =>
		t.textContent?.trim().replace(/\.$/, ""),
	);
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

	describe("the paging bar for screen readers (#186 PR 3)", () => {
		const heard = vi.fn();
		afterEach(() => heard.mockReset());
		const showHeard = (props: Partial<ResultsTabBarsProps>) =>
			render(
				createElement(
					AnnounceProvider,
					{ value: { announce: heard } },
					createElement(ResultsTabBars, { rowCount: 50, ...props }),
				),
			);

		it("Stop is read politely: the user's own doing", async () => {
			await showHeard({ pagingNotice: { kind: "stopped", message: "Stopped at 50 rows." } });
			expect(heard).toHaveBeenCalledOnce();
			expect(heard.mock.calls[0][1]).toMatchObject({ alert: false, polite: true });
		});

		it("the limit after a stop is a new bar, so it's read too, assertively", async () => {
			await showHeard({ pagingNotice: { kind: "stopped", message: "Stopped at 50 rows." } });
			await showHeard({ pagingNotice: { kind: "limit", message: "100,000 rows." } });
			expect(heard).toHaveBeenCalledTimes(2);
			expect(heard.mock.calls[1][0]).toContain("100,000 rows.");
			expect(heard.mock.calls[1][1]).toMatchObject({ alert: true });
		});

		it("a failed page after the limit is read too, though it's as urgent", async () => {
			await showHeard({ pagingNotice: { kind: "limit", message: "100,000 rows." } });
			await showHeard({ pagingNotice: { kind: "error", message: "Throttled", page: 3 } });
			expect(heard).toHaveBeenCalledTimes(2);
			expect(heard.mock.calls[1][0]).toContain("Throttled");
		});
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
