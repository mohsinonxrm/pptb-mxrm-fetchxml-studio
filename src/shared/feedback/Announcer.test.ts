/**
 * The app's announcer (#186 PR 3): a polite and an assertive live region; close messages read as
 * one, a batch replaces its waiting message, the same text is read again, the text clears later;
 * Fluent's MessageBar announces through it at its own politeness, except under SilentAnnouncements.
 */
import { act, createElement, useState, type ReactNode } from "react";
import {
	AnnounceProvider,
	Button,
	Dialog,
	DialogBody,
	DialogSurface,
	DialogTitle,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
} from "@fluentui/react-components";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, flush, render } from "../testing/ui";
import { AnnouncementScope, AnnouncerProvider, SilentAnnouncements } from "./Announcer";
import { CLEAR_AFTER_MS, FLUSH_MS, useAnnouncer, type Announce } from "./useAnnouncer";

let announce: Announce;
function Probe() {
	announce = useAnnouncer();
	return null;
}
const show = (children?: ReactNode) =>
	render(createElement(AnnouncerProvider, null, createElement(Probe), children));
const polite = () => document.querySelector('[data-fxs-announcer="polite"]')?.textContent;
const assertive = () => document.querySelector('[data-fxs-announcer="assertive"]')?.textContent;
const wait = (ms: number) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));
const flushed = () => wait(FLUSH_MS + 30);

afterEach(() => {
	cleanup();
	vi.useRealTimers();
});

describe("AnnouncerProvider", () => {
	it("reads a message politely by default, and assertively when asked", async () => {
		await show();
		await act(async () => announce("Running query…"));
		await act(async () => announce("Couldn't run", { politeness: "assertive" }));
		await flushed();
		expect(polite()).toBe("Running query…");
		expect(assertive()).toBe("Couldn't run");
	});

	it("reads messages that come close together as one, in order, and a batch's last one", async () => {
		await show();
		await act(async () => {
			announce("Undone: Add filter", { batchId: "undo" });
			announce("50 rows loaded.");
			announce("Undone: Add order", { batchId: "undo" });
		});
		await flushed();
		expect(polite()).toBe("50 rows loaded. Undone: Add order");
	});

	it("clears before each reading, so the same text is read twice; and clears a while later", async () => {
		vi.useFakeTimers();
		await show();
		await act(async () => announce("Undone: Add filter"));
		await act(async () => vi.advanceTimersByTime(FLUSH_MS));
		expect(polite()).toBe("Undone: Add filter");
		await act(async () => announce("Undone: Add filter"));
		expect(polite()).toBe("");
		await act(async () => vi.advanceTimersByTime(FLUSH_MS));
		expect(polite()).toBe("Undone: Add filter");
		await act(async () => vi.advanceTimersByTime(CLEAR_AFTER_MS));
		expect(polite()).toBe("");
	});

	it("ignores an empty message", async () => {
		await show();
		await act(async () => announce("   "));
		await flushed();
		expect(polite()).toBe("");
	});

	it("lets Fluent's MessageBar announce itself: info politely, a warning assertively", async () => {
		await show(
			createElement(
				"div",
				null,
				createElement(
					MessageBar,
					{ intent: "info" },
					createElement(MessageBarBody, null, "The query changed."),
				),
				createElement(
					MessageBar,
					{ intent: "warning" },
					createElement(MessageBarBody, null, "Might be rejected."),
				),
			),
		);
		await flushed();
		expect(polite()).toContain("The query changed.");
		expect(assertive()).toContain("Might be rejected.");
	});

	it("keeps MessageBars quiet under SilentAnnouncements (Properties)", async () => {
		await show(
			createElement(
				SilentAnnouncements,
				null,
				createElement(
					MessageBar,
					{ intent: "warning" },
					createElement(MessageBarBody, null, "Dataverse ignores this."),
				),
			),
		);
		await flushed();
		expect(assertive()).toBe("");
	});
});

describe("useAnnouncer outside the provider", () => {
	it("does nothing", async () => {
		await render(createElement(Probe));
		expect(() => announce("Hello")).not.toThrow();
	});
});

describe("the regions while a dialog is open", () => {
	it("are top-level children of the body that a modal never hides", async () => {
		await show(
			createElement(Dialog, {
				open: true,
				modalType: "modal",
				children: createElement(
					DialogSurface,
					null,
					createElement(DialogBody, null, createElement(DialogTitle, null, "Save view")),
				),
			}),
		);
		await flushed();
		for (const region of document.querySelectorAll("[data-fxs-announcer]")) {
			expect(region.parentElement).toBe(document.body);
			expect(region.hasAttribute("data-tabster-never-hide")).toBe(true);
			expect(region.closest('[aria-hidden="true"]')).toBeNull();
		}
		await act(async () => announce("Saved"));
		await flushed();
		expect(polite()).toBe("Saved");
	});
});

describe("AnnouncementScope", () => {
	const heard = vi.fn();
	const bar = (text: string, withAction = false) =>
		createElement(
			MessageBar,
			{ intent: "error" },
			createElement(MessageBarBody, null, text),
			withAction
				? createElement(MessageBarActions, null, createElement(Button, null, "Try again"))
				: null,
		);
	/** A tab: its scope, and its bar while the tab shows */
	const tab = (shown: boolean, text: string | null, withAction = false) =>
		render(
			createElement(
				AnnounceProvider,
				{ value: { announce: heard } },
				createElement(AnnouncementScope, {
					shown,
					children: shown && text ? bar(text, withAction) : null,
				}),
			),
		);
	const readings = () => heard.mock.calls.map(([message]) => message as string);

	afterEach(() => heard.mockReset());

	it("reads a bar once, not again on each visit to its tab", async () => {
		await tab(true, "The query didn't run", true);
		await tab(false, null);
		await tab(true, "The query didn't run", true);
		await flush();
		expect(readings()).toEqual(["The query didn't run,Try again"]);
	});

	it("reads it again once it went while its tab showed (a new run cleared it)", async () => {
		await tab(true, "The query didn't run");
		await tab(true, null);
		await flush();
		await tab(true, "The query didn't run");
		await flush();
		expect(readings()).toEqual(["The query didn't run", "The query didn't run"]);
	});

	it("reads a changed bar", async () => {
		await tab(true, "Export failed. Disk full");
		await tab(true, null);
		await flush();
		await tab(true, "Export failed. Access denied");
		expect(readings()).toEqual(["Export failed. Disk full", "Export failed. Access denied"]);
	});
});

describe("AnnouncementScope in a tab that changes", () => {
	const heard = vi.fn();
	const readings = () => heard.mock.calls.map(([message]) => message as string);
	afterEach(() => heard.mockReset());

	let setShown: (shown: boolean) => void = () => {};
	let setExtra: (extra: number) => void = () => {};
	let slowHide = false;
	afterEach(() => {
		slowHide = false;
	});
	/** A tab with a bar that has an action, and something else in it that can change */
	function Tab() {
		const [shown, setShownState] = useState(true);
		const [extra, setExtraState] = useState(0);
		setShown = setShownState;
		setExtra = setExtraState;
		if (slowHide && !shown) {
			// Hiding takes a while (a big grid goes): the scheduler yields after the commit, before effects
			const end = Date.now() + 15;
			while (Date.now() < end) {
				/* busy */
			}
		}
		return createElement(AnnouncementScope, {
			shown,
			children: shown
				? [
						createElement(
							MessageBar,
							{ key: "bar", intent: "error" },
							createElement(MessageBarBody, null, "Export failed. Disk full"),
							createElement(
								MessageBarActions,
								null,
								createElement(Button, null, "Try again"),
							),
						),
						createElement("span", { key: "rows" }, `rows ${extra}`),
					]
				: null,
		});
	}
	const showTab = () =>
		render(createElement(AnnounceProvider, { value: { announce: heard } }, createElement(Tab)));

	it("a bar with an action stays heard while other things in its tab change", async () => {
		await showTab();
		await act(async () => setExtra(1));
		await flush();
		await act(async () => setShown(false));
		await act(async () => setShown(true));
		await flush();
		expect(readings()).toEqual(["Export failed. Disk full,Try again"]);
	});

	it("a tab hidden outside a click (a run moved to Results) doesn't forget its bars", async () => {
		await showTab();
		slowHide = true;
		const env = globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean };
		env.IS_REACT_ACT_ENVIRONMENT = false;
		try {
			// Not a discrete event, and a slow render: the observer fires (the bar went) before the
			// scope's effects learn the tab is hidden
			setShown(false);
			await new Promise((resolve) => setTimeout(resolve, 50));
			setShown(true);
			await new Promise((resolve) => setTimeout(resolve, 50));
		} finally {
			env.IS_REACT_ACT_ENVIRONMENT = true;
		}
		expect(readings()).toEqual(["Export failed. Disk full,Try again"]);
	});
});
