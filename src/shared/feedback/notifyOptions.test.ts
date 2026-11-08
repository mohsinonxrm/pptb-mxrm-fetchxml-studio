import { describe, expect, it } from "vitest";
import { dedupeKey, resolveOptions } from "./notifyOptions";

describe("resolveOptions", () => {
	it("times out success at 4 s, info at 5 s and warning at 8 s", () => {
		expect(resolveOptions("success", { title: "a" }).timeout).toBe(4000);
		expect(resolveOptions("info", { title: "a" }).timeout).toBe(5000);
		expect(resolveOptions("warning", { title: "a" }).timeout).toBe(8000);
	});

	it("keeps an error until it's dismissed", () => {
		expect(resolveOptions("error", { title: "a" })).toMatchObject({
			timeout: -1,
			sticky: true,
		});
	});

	it("raises the timeout to at least 8 s when there's an action", () => {
		const action = { label: "Undo", onClick: () => undefined };
		expect(resolveOptions("success", { title: "a", action }).timeout).toBe(8000);
		expect(resolveOptions("info", { title: "a", action }).timeout).toBe(8000);
		expect(resolveOptions("warning", { title: "a", action }).timeout).toBe(8000);
		expect(resolveOptions("success", { title: "a", action, timeout: 12000 }).timeout).toBe(
			12000,
		);
		expect(resolveOptions("error", { title: "a", action }).sticky).toBe(true);
		const secondaryAction = { label: "Copy", onClick: () => undefined };
		expect(resolveOptions("success", { title: "a", action, secondaryAction }).timeout).toBe(
			8000,
		);
	});

	it("honours an explicit timeout, including sticky", () => {
		expect(resolveOptions("success", { title: "a", timeout: 1000 }).timeout).toBe(1000);
		expect(resolveOptions("success", { title: "a", timeout: "sticky" })).toMatchObject({
			timeout: -1,
			sticky: true,
		});
		expect(resolveOptions("error", { title: "a", timeout: 3000 })).toMatchObject({
			timeout: 3000,
			sticky: false,
		});
	});

	it("is always polite", () => {
		for (const intent of ["success", "info", "warning", "error"] as const) {
			expect(resolveOptions(intent, { title: "a" }).politeness).toBe("polite");
		}
	});
});

describe("dedupeKey", () => {
	it("defaults to intent, title and body", () => {
		expect(dedupeKey("success", { title: "Copied 3 rows" })).toBe("success|Copied 3 rows|");
		expect(dedupeKey("error", { title: "T", body: "B" })).toBe("error|T|B");
		expect(dedupeKey("success", { title: "T" })).not.toBe(dedupeKey("info", { title: "T" }));
	});

	it("uses the key when there is one", () => {
		expect(dedupeKey("success", { title: "T", key: "copy" })).toBe("copy");
		expect(resolveOptions("info", { title: "T", key: "copy" }).key).toBe("copy");
	});
});
