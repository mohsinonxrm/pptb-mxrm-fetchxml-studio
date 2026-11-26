import { beforeAll, describe, expect, it } from "vitest";
import {
	formatFinding,
	loadAllowlist,
	scan,
	sourceFiles,
	titleCaseWords,
} from "./check-wording.mjs";

// The fixture tests below, and one on the real src/ tree: it blocks CI (#231 decision 3). A
// justified exception goes in wording-allowlist.json with its reason.

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
			"debugWarn(`retrying ${n}...`);",
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
		expect(rules('export const A = () => <Button label="Export (Server)" />;')).toEqual([
			"title-case",
		]);
		expect(rules('export const ops = [{ label: "Less Than" }];')).toEqual(["title-case"]);
		expect(rules("export const A = () => <b>Status Reason</b>;")).toEqual(["title-case"]);
	});

	it('passes "Open in Advanced Find", proper nouns, acronyms and sentence case', () => {
		expect(rules('export const A = () => <Button label="Open in Advanced Find" />;')).toEqual(
			[],
		);
		expect(
			rules(
				'export const A = () => <Tab title="Copy FetchXML to the Power Automate flow" />;',
			),
		).toEqual([]);
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
		expect(
			rules('export const A = () => <Tip content="One Two Three Four Five Six Seven" />;'),
		).toEqual([]);
		expect(
			rules('export const A = () => <Tip content="One Two Three Four Five Six" />;'),
		).toEqual(["title-case"]);
	});

	it("only checks label props and JSX text, not other strings", () => {
		expect(
			rules('export const x = { name: "Less Than" }; export const y = "Less Than";'),
		).toEqual([]);
		expect(rules('export const A = () => <Foo className="Less Than" />;')).toEqual([]);
	});
});

describe("allowlist", () => {
	const src = 'export const A = () => <Spinner label="Loading..." />;';
	it("skips an exact file and text match, and nothing else", () => {
		const entry = {
			file: "src/A.tsx",
			text: "Loading...",
			reason: "quoted from another product",
		};
		expect(scan([{ file: "src/A.tsx", text: src }], { allowlist: [entry] })).toEqual([]);
		expect(scan([{ file: "src/B.tsx", text: src }], { allowlist: [entry] })).toHaveLength(1);
		expect(
			scan([{ file: "src/A.tsx", text: src }], {
				allowlist: [{ ...entry, text: "Loading…" }],
			}),
		).toHaveLength(1);
		expect(scan(src, { file: "src/A.tsx", allowlist: [entry] })).toEqual([]);
	});
});

// Scanning all of src/ takes seconds on a busy runner: one scan, and room for it (#231 review)
const TREE_TIMEOUT = 60_000;

describe("the tool's own strings (#231)", () => {
	/** Every finding in src/, before the allowlist */
	let raw = [];
	let allowlist = [];
	beforeAll(() => {
		allowlist = loadAllowlist();
		raw = scan(sourceFiles());
	}, TREE_TIMEOUT);
	const allowed = (f) => allowlist.some((a) => a.file === f.file && a.text === f.text);

	it("follow the wording rules: no findings in src/", () => {
		expect(raw.filter((f) => !allowed(f)).map(formatFinding)).toEqual([]);
	});

	it("every allowlist entry still matches a string, and says why", () => {
		const unused = allowlist.filter(
			(entry) => !raw.some((f) => f.file === entry.file && f.text === entry.text),
		);
		expect(unused.map((e) => `${e.file}: ${e.text}`)).toEqual([]);
		expect(allowlist.filter((e) => !e.reason?.trim())).toEqual([]);
	});

	it("plural acronyms aren't Title Case", () => {
		expect(titleCaseWords("Raw (GUIDs, numbers, ISO dates)")).toEqual([]);
	});

	it(
		"skips only src/shared/testing (test infrastructure), not other folders named testing",
		() => {
			const files = sourceFiles();
			expect(files.some((file) => /[\\/]shared[\\/]testing[\\/]/.test(file))).toBe(false);
			expect(files.some((file) => /[\\/]features[\\/]/.test(file))).toBe(true);
		},
		TREE_TIMEOUT,
	);
});
