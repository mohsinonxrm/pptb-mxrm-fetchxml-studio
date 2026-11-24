/**
 * The app's announcer (#186 PR 3): a polite and an assertive live region; close messages read as
 * one, a batch replaces its waiting message, the same text is read again, the text clears later;
 * Fluent's MessageBar announces through it at its own politeness, except under SilentAnnouncements.
 */
import { act, createElement, type ReactNode } from "react";
import { MessageBar, MessageBarBody } from "@fluentui/react-components";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../testing/ui";
import { AnnouncerProvider, SilentAnnouncements } from "./Announcer";
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
