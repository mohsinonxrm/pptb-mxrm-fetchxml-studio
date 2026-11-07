import { describe, expect, it } from "vitest";
import { formatFinding, scan, titleCaseWords } from "./check-wording.mjs";

// Fixture tests only. There is deliberately NO test that scans the real src/ tree here: today's
// strings would fail it. The #231 wording sweep PR fixes them, fills wording-allowlist.json and
// adds that "no findings in src/" assertion.

const rules = (src, options) => scan(src, options).map((f) => f.rule);

describe("ascii-ellipsis", () => {
	it('flags "Loading..." in JSX text, an attribute and a template', () => {
		const src = [
			"export const A = () => <p>Loading...</p>;",
			'export const B = () => <Spinner label="Loading..." />;',
			"export const c = (n: number) => `Loading ${n} rows...`;",
		].join("\n");
		const findings = scan(src);
		expect(findings.map((f) => [f.line, f.rule])).toEqual([
			[1, "ascii-ellipsis"],
			[2, "ascii-ellipsis"],
			[3, "ascii-ellipsis"],
		]);
		expect(formatFinding(findings[0])).toBe('source.tsx:1: ascii-ellipsis: "Loading..."');
	});

	it("ignores log calls, spread, comments and the single ellipsis character", () => {
		const src = [
			'debugLog("Loading...");',
			'debugWarn(`retrying ${n}...`);',
			'console.error("failed...");',
			"// Loading... in a comment",
			"/* Loading... */",
			"export const A = (props: object) => <div {...props}>Loading…</div>;",
			"export const xs = [...[1, 2]];",
		].join("\n");
		expect(scan(src)).toEqual([]);
	});
});

describe("title-case", () => {
	it('flags "Export (Server)" and Title Case props or JSX text', () => {
		expect(rules('export const A = () => <Button label="Export (Server)" />;')).toEqual(["title-case"]);
		expect(rules('export const ops = [{ label: "Less Than" }];')).toEqual(["title-case"]);
		expect(rules("export const A = () => <b>Status Reason</b>;")).toEqual(["title-case"]);
	});

	it('passes "Open in Advanced Find", proper nouns, acronyms and sentence case', () => {
		expect(rules('export const A = () => <Button label="Open in Advanced Find" />;')).toEqual([]);
		expect(rules('export const A = () => <Tab title="Copy FetchXML to the Power Automate flow" />;')).toEqual([]);
		expect(rules('export const A = () => <Tab title="Export as CSV" />;')).toEqual([]);
		expect(rules('export const A = () => <Tab title="Not equals" />;')).toEqual([]);
		expect(titleCaseWords("Done. Close it")).toEqual([]);
	});

	it("treats Entity and Attribute as common nouns: capitalized only at the start", () => {
		expect(titleCaseWords("Entity filter")).toEqual([]);
		expect(titleCaseWords("Select Entity")).toEqual(["Entity"]);
		expect(titleCaseWords("All Attributes")).toEqual(["Attributes"]);
	});

	it("doesn't check case for a 7-word label", () => {
		expect(rules('export const A = () => <Tip content="One Two Three Four Five Six Seven" />;')).toEqual([]);
		expect(rules('export const A = () => <Tip content="One Two Three Four Five Six" />;')).toEqual(["title-case"]);
	});

	it("only checks label props and JSX text, not other strings", () => {
		expect(rules('export const x = { name: "Less Than" }; export const y = "Less Than";')).toEqual([]);
		expect(rules('export const A = () => <Foo className="Less Than" />;')).toEqual([]);
	});
});

describe("allowlist", () => {
	const src = 'export const A = () => <Spinner label="Loading..." />;';
	it("skips an exact file and text match, and nothing else", () => {
		const entry = { file: "src/A.tsx", text: "Loading...", reason: "quoted from another product" };
		expect(scan([{ file: "src/A.tsx", text: src }], { allowlist: [entry] })).toEqual([]);
		expect(scan([{ file: "src/B.tsx", text: src }], { allowlist: [entry] })).toHaveLength(1);
		expect(scan([{ file: "src/A.tsx", text: src }], { allowlist: [{ ...entry, text: "Loading…" }] })).toHaveLength(1);
		expect(scan(src, { file: "src/A.tsx", allowlist: [entry] })).toEqual([]);
	});
});
