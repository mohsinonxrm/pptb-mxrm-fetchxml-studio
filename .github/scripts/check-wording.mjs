// UI wording check (#231): finds "..." and Title Case labels in the strings users see, so they
// can't come back after the wording sweep. The rules are the style table in CONTRIBUTING.md
// ("UI changes"), which links each rule to its Fluent 2 / Microsoft Style Guide source.
//
// Uses the `typescript` devDependency's AST (no new dependencies). It walks non-test
// src/**/*.{ts,tsx}, skips __fixtures__, comments and log calls (console.*, debugLog,
// debugWarn), and reports:
//   ascii-ellipsis  "..." in JSX text, JSX attribute strings, and string or template literals.
//                   Use the single character "…" (U+2026) instead.
//   title-case      a short label (at most 6 words) in a `label`, `title`, `aria-label`,
//                   `placeholder`, `content` or `hint` prop, or in JSX text, where a word after the
//                   first starts with a capital and isn't a proper noun. Use sentence case.
//
// Exceptions go in wording-allowlist.json: [{ file, text, reason }] (exact file and text).
// NOT enforced yet: it isn't run by CI or `npm test` until the #231 sweep PR fixes today's strings.
//
// CLI: `npm run check:wording` prints `file:line: rule: "text"` and exits 1 on any finding.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import ts from "typescript";

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, "..", "..");
export const ALLOWLIST_PATH = path.join(here, "wording-allowlist.json");

/** Props whose string value is a short UI label */
export const LABEL_PROPS = new Set([
	"label",
	"title",
	"aria-label",
	"placeholder",
	"content",
	"hint",
]);
/** Calls whose arguments are developer logs, not UI */
const LOG_CALL = /^(console\.\w+|debugLog|debugWarn)$/;
export const MAX_LABEL_WORDS = 6;

/**
 * Proper nouns that keep their capitals inside a sentence-case label. Multi-word names are
 * matched as phrases. "Entity" and "Attribute" are deliberately NOT here: they are common nouns
 * in sentence case ("All entities", "Show entities from"), capitalized only at the start of a
 * label or sentence, so "Select Entity" is a finding. All-caps acronyms (CSV, JSON, XML) are
 * always fine.
 */
export const PROPER_NOUNS = [
	"Power Platform ToolBox",
	"Power Automate",
	"Power Apps",
	"Power BI",
	"Power Pages",
	"Advanced Find",
	"Microsoft Learn",
	"Web API",
	"pac CLI",
	"FetchXML",
	"LayoutXML",
	"Dataverse",
	"Excel",
	"Microsoft",
	"Learn",
	"SQL",
	"TDS",
	"OData",
	"Xrm.WebApi",
	"C#",
	"QueryExpression",
	"FetchExpression",
	"PPTB",
	"ID",
	"GUID",
	"URL",
	"cURL",
	"PowerShell",
	"Windows",
	"GitHub",
	"Monaco",
	"Fluent",
];

const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const PROPER_PHRASES = [...PROPER_NOUNS]
	.sort((a, b) => b.length - a.length)
	.map((p) => new RegExp(`(?<![\\w])${escapeRegExp(p)}(?![\\w])`, "g"));

/** The Title Case words in a short label (after the first word, not proper nouns), or [] */
export function titleCaseWords(text) {
	let t = text.replace(/\s+/g, " ").trim();
	for (const re of PROPER_PHRASES) t = t.replace(re, "§");
	const tokens = t.split(" ").filter((w) => /[\p{L}§]/u.test(w));
	if (tokens.length === 0 || tokens.length > MAX_LABEL_WORDS) return [];
	const bad = [];
	let sentenceStart = true;
	for (const token of tokens) {
		const word = token.replace(/^[^\p{L}§]+/u, "");
		const startsSentence = sentenceStart;
		sentenceStart = /[.!?:]$/.test(token);
		if (startsSentence || !word || word.startsWith("§")) continue;
		if (!/^\p{Lu}/u.test(word)) continue;
		const letters = word.replace(/[^\p{L}]/gu, "");
		if (letters.length > 1 && letters === letters.toUpperCase()) continue; // acronym
		bad.push(word.replace(/[^\p{L}\p{N}]+$/u, ""));
	}
	return bad;
}

function isLogCall(node, sf) {
	for (let p = node.parent; p; p = p.parent) {
		if (ts.isCallExpression(p) && LOG_CALL.test(p.expression.getText(sf))) return true;
	}
	return false;
}

/** Is this string a short-label value: `label="x"`, `label={"x"}` or `{ label: "x" }`? */
function isLabelValue(node) {
	let parent = node.parent;
	if (parent && ts.isJsxExpression(parent)) parent = parent.parent;
	if (parent && ts.isJsxAttribute(parent)) return LABEL_PROPS.has(parent.name.getText());
	const p = node.parent;
	if (p && ts.isPropertyAssignment(p) && p.initializer === node) {
		const name =
			ts.isIdentifier(p.name) || ts.isStringLiteral(p.name) ? p.name.text : undefined;
		return name !== undefined && LABEL_PROPS.has(name);
	}
	return false;
}

function scanSource(file, text, allowlist) {
	const kind = /\.ts$/.test(file) ? ts.ScriptKind.TS : ts.ScriptKind.TSX;
	const sf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, kind);
	const findings = [];
	const allowed = (t) => allowlist.some((a) => a.file === file && a.text === t);
	const report = (pos, rule, t) => {
		if (allowed(t)) return;
		findings.push({
			file,
			line: sf.getLineAndCharacterOfPosition(pos).line + 1,
			rule,
			text: t,
		});
	};
	const check = (raw, pos, isLabel) => {
		const t = raw.replace(/\s+/g, " ").trim();
		if (!t) return;
		if (t.includes("...")) report(pos, "ascii-ellipsis", t);
		if (isLabel && titleCaseWords(t).length > 0) report(pos, "title-case", t);
	};
	const visit = (node) => {
		if (ts.isJsxText(node)) {
			const raw = node.getText(sf);
			check(raw, node.getStart(sf) + (raw.length - raw.trimStart().length), true);
		} else if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) {
			const p = node.parent;
			const skip =
				ts.isImportDeclaration(p) ||
				ts.isExportDeclaration(p) ||
				ts.isLiteralTypeNode(p) ||
				isLogCall(node, sf);
			if (!skip) check(node.text, node.getStart(sf), isLabelValue(node));
		} else if (ts.isTemplateExpression(node)) {
			if (!isLogCall(node, sf)) {
				// the literal parts only; `${...}` expressions are visited as their own nodes
				const parts = [node.head, ...node.templateSpans.map((s) => s.literal)];
				check(parts.map((x) => x.text).join(" "), node.getStart(sf), false);
			}
		}
		ts.forEachChild(node, visit);
	};
	visit(sf);
	return findings;
}

/** Non-test .ts and .tsx files under `root`/src, skipping __fixtures__ and .d.ts */
export function sourceFiles(root = repoRoot) {
	const out = [];
	const walk = (dir) => {
		for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
			const p = path.join(dir, e.name);
			if (e.isDirectory()) {
				if (e.name !== "__fixtures__") walk(p);
			} else if (
				/\.(ts|tsx)$/.test(e.name) &&
				!/\.test\.(ts|tsx)$/.test(e.name) &&
				!/\.d\.ts$/.test(e.name)
			) {
				out.push(p);
			}
		}
	};
	walk(path.join(root, "src"));
	return out;
}

export function loadAllowlist(file = ALLOWLIST_PATH) {
	return JSON.parse(fs.readFileSync(file, "utf8"));
}

/**
 * Findings `{ file, line, rule, text }` for:
 *  - a string: source text, scanned as `options.file` (default "source.tsx");
 *  - an array of strings: file paths, read from disk and reported relative to `options.root`;
 *  - an array of `{ file, text }`: in-memory files.
 * `options.allowlist`: `[{ file, text, reason }]` to skip (default none; the CLI loads the file).
 */
export function scan(input, options = {}) {
	const allowlist = options.allowlist ?? [];
	const root = options.root ?? repoRoot;
	if (typeof input === "string")
		return scanSource(options.file ?? "source.tsx", input, allowlist);
	return input.flatMap((item) => {
		if (typeof item === "string") {
			const rel = path.relative(root, item).replace(/\\/g, "/");
			return scanSource(rel, fs.readFileSync(item, "utf8"), allowlist);
		}
		return scanSource(item.file, item.text, allowlist);
	});
}

export const formatFinding = (f) => `${f.file}:${f.line}: ${f.rule}: "${f.text}"`;

function main() {
	const findings = scan(sourceFiles(), { allowlist: loadAllowlist() });
	for (const f of findings) console.log(formatFinding(f));
	if (findings.length > 0) {
		console.error(
			`\n${findings.length} wording finding(s). See "UI changes" in CONTRIBUTING.md; a justified exception goes in .github/scripts/wording-allowlist.json.`,
		);
		process.exit(1);
	}
	console.log("No wording findings.");
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href)
	main();
