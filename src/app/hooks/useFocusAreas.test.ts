/**
 * Ctrl+F6 / Ctrl+Shift+F6 areas (#186): app bar → builder → output and round again, back to the last
 * focused element in each area, else its first stop; hidden areas are skipped; from outside every
 * area, next starts at the first and previous at the last; the rail beside a peek counts as the
 * builder.
 */
import { act, createElement, useRef } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { useFocusAreas, type FocusArea } from "./useFocusAreas";

let api: ReturnType<typeof useFocusAreas>;

function Probe({ hideBuilder = false, rail = false }: { hideBuilder?: boolean; rail?: boolean }) {
	const bar = useRef<HTMLElement>(null);
	const builder = useRef<HTMLElement>(null);
	const railEl = useRef<HTMLElement>(null);
	const output = useRef<HTMLElement>(null);
	const first = (el: HTMLElement) => el.querySelector<HTMLElement>("button, [tabindex]")?.focus();
	api = useFocusAreas((): FocusArea[] => [
		{ element: bar.current, focusFirst: first },
		{ element: builder.current, alsoContains: [railEl.current], focusFirst: first },
		{ element: output.current, focusFirst: first },
	]);
	return createElement(
		"div",
		null,
		createElement(
			"header",
			{ ref: bar },
			createElement("button", { id: "run" }, "Run"),
			createElement("button", { id: "save" }, "Save"),
		),
		rail &&
			createElement(
				"nav",
				{ ref: railEl },
				createElement("button", { id: "expand" }, "Expand"),
			),
		createElement(
			"aside",
			{ ref: builder, style: hideBuilder ? { display: "none" } : undefined },
			createElement("div", { id: "row1", tabIndex: 0 }, "Fetch query"),
			createElement("div", { id: "row2", tabIndex: -1 }, "account"),
		),
		createElement("button", { id: "splitter" }, "Resize"),
		createElement(
			"main",
			{ ref: output },
			createElement("button", { id: "tab" }, "Results"),
			createElement("button", { id: "copy" }, "Copy"),
		),
		createElement("button", { id: "outside" }, "Outside"),
	);
}

const focused = () => document.activeElement?.id;
const focus = (id: string) => act(async () => document.getElementById(id)?.focus());
const next = () => act(async () => api.next());
const previous = () => act(async () => api.previous());

afterEach(() => cleanup());

describe("useFocusAreas", () => {
	it("cycles app bar → builder → output → app bar, and back the other way", async () => {
		await render(createElement(Probe));
		await focus("run");
		await next();
		expect(focused()).toBe("row1");
		await next();
		expect(focused()).toBe("tab");
		await next();
		expect(focused()).toBe("run");
		await previous();
		expect(focused()).toBe("tab");
		await previous();
		expect(focused()).toBe("row1");
	});

	it("goes back to the last focused element in an area", async () => {
		await render(createElement(Probe));
		await focus("save");
		await focus("copy");
		await next();
		expect(focused()).toBe("save");
		await focus("row2");
		await next();
		expect(focused()).toBe("copy");
		await previous();
		expect(focused()).toBe("row2");
	});

	it("uses the first stop when the remembered element is gone", async () => {
		await render(createElement(Probe));
		await focus("copy");
		await act(async () => document.getElementById("copy")?.remove());
		await focus("row1");
		await next();
		expect(focused()).toBe("tab");
	});

	it("skips a hidden area", async () => {
		await render(createElement(Probe, { hideBuilder: true }));
		await focus("run");
		await next();
		expect(focused()).toBe("tab");
		await previous();
		expect(focused()).toBe("run");
	});

	it("starts at the first area going forward and the last going back, from outside them", async () => {
		await render(createElement(Probe));
		await focus("outside");
		await next();
		expect(focused()).toBe("run");
		await focus("outside");
		await previous();
		expect(focused()).toBe("tab");
	});

	it("from between two areas (the pane splitter), goes to the next one after it or the one before", async () => {
		await render(createElement(Probe));
		await focus("splitter");
		await next();
		expect(focused()).toBe("tab");
		await focus("splitter");
		await previous();
		expect(focused()).toBe("row1");
	});

	it("counts the rail beside a peek as the builder", async () => {
		await render(createElement(Probe, { rail: true }));
		await focus("expand");
		await next();
		expect(focused()).toBe("tab");
	});
});
