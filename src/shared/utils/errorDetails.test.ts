import { describe, expect, it } from "vitest";
import { errorDetails, errorMessage } from "./errorDetails";

describe("error details for Copy details (#185)", () => {
	const time = new Date("2026-10-25T06:00:00.000Z");

	it("names the tool version, where it failed, the error and the time", () => {
		const text = errorDetails({ where: "Results panel", error: new TypeError("x is undefined"), version: "1.6.0", time });
		expect(text).toBe("FetchXML Studio 1.6.0 · Results panel\nTypeError: x is undefined\nTime: 2026-10-25T06:00:00.000Z");
	});

	it("keeps only the first lines of the component stack", () => {
		const componentStack = Array.from({ length: 20 }, (_, i) => `    at Component${i} (file.tsx:${i})`).join("\n");
		const text = errorDetails({ where: "Whole tool", error: new Error("boom"), componentStack, version: "1.6.0", time });
		expect(text).toContain("Component stack:\n  at Component0 (file.tsx:0)");
		expect(text).toContain("at Component7");
		expect(text).not.toContain("at Component8");
	});

	it("copes with things that aren't Errors", () => {
		expect(errorMessage("plain string")).toBe("plain string");
		expect(errorMessage({ weird: true })).toBe("Unknown error");
		expect(errorDetails({ where: "Code panel", error: 42, version: "1.6.0", time })).toContain("Error: Unknown error");
	});
});
