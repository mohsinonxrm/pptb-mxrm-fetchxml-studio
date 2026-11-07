/**
 * Safe FetchXML / XML pretty-printer (#186, #287).
 *
 * A small tokenizer, not DOMParser + XMLSerializer, because the serializer rewrites entities and
 * quotes. Only whitespace between tags changes; attribute values, text, comments, CDATA, the XML
 * declaration, processing instructions, attribute order and quote characters are kept as written.
 * Not used by the editor yet.
 *
 * Parsing and rendering use explicit stacks, so nesting depth is not limited by the call stack.
 *
 * Whitespace means XML whitespace only (space, tab, CR, LF). NBSP and other Unicode spaces are text.
 *
 * Checked: tag matching, one root, quoted unique attributes, the five predefined entities and numeric
 * references, comment syntax, declaration position, control characters. Not checked (accepted as
 * written): undeclared namespace prefixes, duplicate expanded names, and entities declared in a DOCTYPE.
 */

export type FormatFetchXmlResult = { ok: true; text: string } | { ok: false; error: string };

export interface FormatFetchXmlOptions {
	/** One level of indentation: spaces and tabs only. Default: two spaces, as the FetchXML generator writes. */
	indent?: string;
}

class FormatError extends Error {
	readonly offset: number;
	constructor(message: string, offset: number) {
		super(message);
		this.offset = offset;
	}
}

type XmlNode =
	| { kind: "text"; raw: string }
	| { kind: "other"; raw: string } // comment, processing instruction, DOCTYPE
	| { kind: "cdata"; raw: string }
	| {
			kind: "element";
			open: string; // normalised start tag, e.g. <a b="c"> or <a b="c" />
			selfClosing: boolean;
			name: string;
			children: XmlNode[];
			rawContent: string; // exact text between the start and end tags
	  };

const NAME = /[\p{L}_:][\p{L}\p{M}\p{N}_:.·-]*/uy;
const ENTITY = /&(?:#[0-9]+|#x[0-9a-fA-F]+|amp|lt|gt|quot|apos);/y;
const WS = /[ \t\r\n]/;
const ALL_WS = /^[ \t\r\n]*$/;
// eslint-disable-next-line no-control-regex -- rejecting XML-forbidden control characters is the point
const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/;
const INDENT = /^[ \t]*$/;

function checkEntities(s: string, base: number): void {
	for (let i = s.indexOf("&"); i !== -1; i = s.indexOf("&", i + 1)) {
		ENTITY.lastIndex = i;
		if (!ENTITY.test(s)) throw new FormatError("bad entity", base + i);
	}
}

interface Frame {
	name: string;
	open: string;
	children: XmlNode[];
	contentStart: number;
}

function parse(src: string, start: number): XmlNode[] {
	const control = CONTROL.exec(src);
	if (control) throw new FormatError("control character", control.index);

	let pos = start;
	const rootNodes: XmlNode[] = [];
	const stack: Frame[] = [];
	let cur = rootNodes;

	const skipWs = (): void => {
		while (pos < src.length && WS.test(src[pos])) pos++;
	};

	const readName = (): string => {
		NAME.lastIndex = pos;
		const m = NAME.exec(src);
		if (!m) throw new FormatError("expected a name", pos);
		pos += m[0].length;
		return m[0];
	};

	const parseStartTag = (): { open: string; name: string; selfClosing: boolean } => {
		pos++; // <
		const name = readName();
		const parts: string[] = [`<${name}`];
		const seen = new Set<string>();
		for (;;) {
			const before = pos;
			skipWs();
			if (src[pos] === ">") {
				pos++;
				return { open: parts.join("") + ">", name, selfClosing: false };
			}
			if (src[pos] === "/" && src[pos + 1] === ">") {
				pos += 2;
				return { open: parts.join("") + " />", name, selfClosing: true };
			}
			if (pos >= src.length) throw new FormatError("unterminated tag", pos);
			if (pos === before) throw new FormatError("expected whitespace", pos);
			const attrName = readName();
			if (seen.has(attrName)) throw new FormatError("duplicate attribute", pos);
			seen.add(attrName);
			skipWs();
			if (src[pos] !== "=") throw new FormatError("expected =", pos);
			pos++;
			skipWs();
			const q = src[pos];
			if (q !== '"' && q !== "'") throw new FormatError("expected a quoted value", pos);
			const end = src.indexOf(q, pos + 1);
			if (end === -1) throw new FormatError("unterminated value", pos);
			const value = src.slice(pos + 1, end);
			if (value.includes("<")) throw new FormatError("< in attribute value", pos);
			checkEntities(value, pos + 1);
			parts.push(` ${attrName}=${q}${value}${q}`);
			pos = end + 1;
		}
	};

	while (pos < src.length) {
		const atRoot = stack.length === 0;
		if (src[pos] !== "<") {
			let end = src.indexOf("<", pos);
			if (end === -1) end = src.length;
			const raw = src.slice(pos, end);
			if (atRoot) {
				if (!ALL_WS.test(raw)) throw new FormatError("text outside the root", pos);
			} else {
				checkEntities(raw, pos);
				if (raw.includes("]]>")) throw new FormatError("]]> in text", pos);
			}
			cur.push({ kind: "text", raw });
			pos = end;
		} else if (src.startsWith("<!--", pos)) {
			const end = src.indexOf("-->", pos + 4);
			if (end === -1) throw new FormatError("unterminated comment", pos);
			const body = src.slice(pos + 4, end);
			if (body.includes("--") || body.endsWith("-"))
				throw new FormatError("bad comment", pos);
			cur.push({ kind: "other", raw: src.slice(pos, end + 3) });
			pos = end + 3;
		} else if (src.startsWith("<![CDATA[", pos)) {
			if (atRoot) throw new FormatError("CDATA outside the root", pos);
			const end = src.indexOf("]]>", pos + 9);
			if (end === -1) throw new FormatError("unterminated CDATA", pos);
			cur.push({ kind: "cdata", raw: src.slice(pos, end + 3) });
			pos = end + 3;
		} else if (src.startsWith("<?", pos)) {
			const piStart = pos;
			pos += 2;
			const target = readName();
			if (target.toLowerCase() === "xml" && piStart !== start)
				throw new FormatError("misplaced XML declaration", piStart);
			if (!src.startsWith("?>", pos) && !WS.test(src[pos] ?? ""))
				throw new FormatError("malformed processing instruction", pos);
			const end = src.indexOf("?>", pos);
			if (end === -1) throw new FormatError("unterminated processing instruction", piStart);
			cur.push({ kind: "other", raw: src.slice(piStart, end + 2) });
			pos = end + 2;
		} else if (src.startsWith("<!", pos)) {
			if (!atRoot || !src.startsWith("<!DOCTYPE", pos) || !WS.test(src[pos + 9] ?? "")) {
				throw new FormatError("unexpected <!", pos);
			}
			// DOCTYPE: copy up to the closing > (quoted strings and an internal subset in [ ... ] may hold >).
			let i = pos + 9;
			let depth = 0;
			let quote = "";
			for (; i < src.length; i++) {
				const c = src[i];
				if (quote) {
					if (c === quote) quote = "";
				} else if (c === '"' || c === "'") quote = c;
				else if (c === "[") depth++;
				else if (c === "]") depth--;
				else if (c === ">" && depth <= 0) break;
			}
			if (i >= src.length) throw new FormatError("unterminated declaration", pos);
			cur.push({ kind: "other", raw: src.slice(pos, i + 1) });
			pos = i + 1;
		} else if (src[pos + 1] === "/") {
			if (atRoot) throw new FormatError("unexpected end tag", pos);
			const tagStart = pos;
			pos += 2;
			const name = readName();
			skipWs();
			if (src[pos] !== ">") throw new FormatError("malformed end tag", pos);
			const frame = stack[stack.length - 1];
			if (name !== frame.name)
				throw new FormatError(`</${name}> doesn't match <${frame.name}>`, tagStart);
			pos++;
			stack.pop();
			cur = stack.length ? stack[stack.length - 1].children : rootNodes;
			cur.push({
				kind: "element",
				open: frame.open,
				selfClosing: false,
				name,
				children: frame.children,
				rawContent: src.slice(frame.contentStart, tagStart),
			});
		} else {
			const tag = parseStartTag();
			if (tag.selfClosing) {
				cur.push({ kind: "element", ...tag, children: [], rawContent: "" });
			} else {
				const frame: Frame = {
					name: tag.name,
					open: tag.open,
					children: [],
					contentStart: pos,
				};
				stack.push(frame);
				cur = frame.children;
			}
		}
	}

	if (stack.length) throw new FormatError(`unclosed <${stack[stack.length - 1].name}>`, pos);
	const roots = rootNodes.filter((n) => n.kind === "element").length;
	if (roots === 0) throw new FormatError("no root element", src.length);
	if (roots > 1) throw new FormatError("more than one root element", src.length);
	return rootNodes;
}

function render(nodes: XmlNode[], indent: string): string {
	const out: string[] = [];
	const stack: { nodes: XmlNode[]; i: number; closeLine: string | null }[] = [
		{ nodes, i: 0, closeLine: null },
	];
	while (stack.length) {
		const f = stack[stack.length - 1];
		if (f.i >= f.nodes.length) {
			stack.pop();
			if (f.closeLine !== null) out.push(f.closeLine);
			continue;
		}
		const n = f.nodes[f.i++];
		if (n.kind === "text") continue; // whitespace between tags only; text is kept via rawContent
		const pad = indent.repeat(stack.length - 1);
		if (n.kind === "other" || n.kind === "cdata") {
			out.push(pad + n.raw);
			continue;
		}
		if (n.selfClosing) {
			out.push(pad + n.open);
			continue;
		}
		const close = `</${n.name}>`;
		if (n.children.length === 0) {
			out.push(pad + n.open + close);
			continue;
		}
		if (n.children.every((c) => c.kind !== "element")) {
			// Text, CDATA, a lone whitespace value, comments around whitespace: all of it is content, kept as written.
			out.push(pad + n.open + n.rawContent + close);
			continue;
		}
		if (
			n.children.some((c) => c.kind === "cdata" || (c.kind === "text" && !ALL_WS.test(c.raw)))
		) {
			// Mixed content: keep as written.
			out.push(pad + n.open + n.rawContent + close);
			continue;
		}
		out.push(pad + n.open);
		stack.push({ nodes: n.children, i: 0, closeLine: pad + close });
	}
	return out.join("\n");
}

/** 1-based line of an offset; CRLF, LF and a bare CR each count as one line break. */
function lineOf(src: string, offset: number): number {
	let line = 1;
	const end = Math.min(offset, src.length);
	for (let i = 0; i < end; i++) {
		const c = src.charCodeAt(i);
		if (c === 10 || (c === 13 && src.charCodeAt(i + 1) !== 10)) line++;
	}
	return line;
}

export function formatFetchXml(
	text: string,
	options?: FormatFetchXmlOptions,
): FormatFetchXmlResult {
	const indent = options?.indent ?? "  ";
	if (!INDENT.test(indent))
		return { ok: false, error: "Can't format: the indent must be spaces or tabs" };
	try {
		const start = text.charCodeAt(0) === 0xfeff ? 1 : 0; // the BOM is dropped; offsets stay in the original text
		return { ok: true, text: render(parse(text, start), indent) };
	} catch (e) {
		if (e instanceof FormatError) {
			return {
				ok: false,
				error: `Can't format: the FetchXML isn't well-formed (line ${lineOf(text, e.offset)})`,
			};
		}
		return { ok: false, error: "Can't format: the FetchXML couldn't be read" };
	}
}
