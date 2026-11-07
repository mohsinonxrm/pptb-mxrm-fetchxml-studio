/**
 * Safe FetchXML / XML pretty-printer (#186, #287).
 *
 * A small tokenizer, not DOMParser + XMLSerializer, because the serializer rewrites entities and
 * quotes. Only whitespace between tags changes; attribute values, text, comments, CDATA, the XML
 * declaration, processing instructions, attribute order and quote characters are kept as written.
 * Not used by the editor yet.
 */

export type FormatFetchXmlResult = { ok: true; text: string } | { ok: false; error: string };

export interface FormatFetchXmlOptions {
	/** One level of indentation. Default: two spaces, as the FetchXML generator writes. */
	indent?: string;
}

class FormatError extends Error {
	constructor(
		message: string,
		readonly offset: number,
	) {
		super(message);
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

const NAME = /[\p{L}_:][\p{L}\p{N}_:.·-]*/uy;
const ENTITY = /&(?:#[0-9]+|#x[0-9a-fA-F]+|[\p{L}_:][\p{L}\p{N}_:.-]*);/uy;
const WS = /[ \t\r\n]/;

function checkEntities(s: string, base: number): void {
	for (let i = s.indexOf("&"); i !== -1; i = s.indexOf("&", i + 1)) {
		ENTITY.lastIndex = i;
		if (!ENTITY.test(s)) throw new FormatError("bad entity", base + i);
	}
}

function parse(src: string): XmlNode[] {
	let pos = 0;

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
		const attrs: string[] = [];
		const seen = new Set<string>();
		for (;;) {
			const before = pos;
			skipWs();
			const joined = attrs.length ? " " + attrs.join(" ") : "";
			if (src[pos] === ">") {
				pos++;
				return { open: `<${name}${joined}>`, name, selfClosing: false };
			}
			if (src[pos] === "/" && src[pos + 1] === ">") {
				pos += 2;
				return { open: `<${name}${joined} />`, name, selfClosing: true };
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
			attrs.push(`${attrName}=${q}${value}${q}`);
			pos = end + 1;
		}
	};

	/** Parses nodes until the matching end tag (consumed) or, at depth 0, the end of input. */
	const parseContent = (parent: string | null): { nodes: XmlNode[]; rawEnd: number } => {
		const nodes: XmlNode[] = [];
		for (;;) {
			if (pos >= src.length) {
				if (parent !== null) throw new FormatError(`unclosed <${parent}>`, pos);
				return { nodes, rawEnd: pos };
			}
			if (src[pos] !== "<") {
				let end = src.indexOf("<", pos);
				if (end === -1) end = src.length;
				const raw = src.slice(pos, end);
				checkEntities(raw, pos);
				if (raw.includes("]]>")) throw new FormatError("]]> in text", pos);
				if (parent === null && raw.trim() !== "")
					throw new FormatError("text outside the root", pos);
				nodes.push({ kind: "text", raw });
				pos = end;
				continue;
			}
			if (src.startsWith("<!--", pos)) {
				const end = src.indexOf("-->", pos + 4);
				if (end === -1) throw new FormatError("unterminated comment", pos);
				nodes.push({ kind: "other", raw: src.slice(pos, end + 3) });
				pos = end + 3;
			} else if (src.startsWith("<![CDATA[", pos)) {
				if (parent === null) throw new FormatError("CDATA outside the root", pos);
				const end = src.indexOf("]]>", pos + 9);
				if (end === -1) throw new FormatError("unterminated CDATA", pos);
				nodes.push({ kind: "cdata", raw: src.slice(pos, end + 3) });
				pos = end + 3;
			} else if (src.startsWith("<?", pos)) {
				const end = src.indexOf("?>", pos + 2);
				if (end === -1) throw new FormatError("unterminated processing instruction", pos);
				nodes.push({ kind: "other", raw: src.slice(pos, end + 2) });
				pos = end + 2;
			} else if (src.startsWith("<!", pos)) {
				if (parent !== null) throw new FormatError("unexpected <!", pos);
				// DOCTYPE: copy up to the closing > (allowing an internal subset in [ ... ]).
				let i = pos + 2;
				let depth = 0;
				for (; i < src.length; i++) {
					const c = src[i];
					if (c === "[") depth++;
					else if (c === "]") depth--;
					else if (c === ">" && depth <= 0) break;
				}
				if (i >= src.length) throw new FormatError("unterminated declaration", pos);
				nodes.push({ kind: "other", raw: src.slice(pos, i + 1) });
				pos = i + 1;
			} else if (src[pos + 1] === "/") {
				if (parent === null) throw new FormatError("unexpected end tag", pos);
				const tagStart = pos;
				pos += 2;
				const name = readName();
				skipWs();
				if (src[pos] !== ">") throw new FormatError("malformed end tag", pos);
				if (name !== parent)
					throw new FormatError(`</${name}> doesn't match <${parent}>`, tagStart);
				pos++;
				return { nodes, rawEnd: tagStart };
			} else {
				const tag = parseStartTag();
				if (tag.selfClosing) {
					nodes.push({ kind: "element", ...tag, children: [], rawContent: "" });
				} else {
					const contentStart = pos;
					const inner = parseContent(tag.name);
					nodes.push({
						kind: "element",
						...tag,
						children: inner.nodes,
						rawContent: src.slice(contentStart, inner.rawEnd),
					});
				}
			}
		}
	};

	const { nodes } = parseContent(null);
	const roots = nodes.filter((n) => n.kind === "element");
	if (roots.length === 0) throw new FormatError("no root element", src.length);
	if (roots.length > 1) throw new FormatError("more than one root element", src.length);
	return nodes;
}

function render(nodes: XmlNode[], depth: number, indent: string, out: string[]): void {
	const pad = indent.repeat(depth);
	for (const n of nodes) {
		if (n.kind === "text") continue; // whitespace between tags only; text is kept via rawContent
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
		const hasData = n.children.some(
			(c) => c.kind === "cdata" || (c.kind === "text" && c.raw.trim() !== ""),
		);
		const onlyWhitespace = n.children.every((c) => c.kind === "text");
		if (hasData || onlyWhitespace) {
			// Text (or a lone whitespace value) is data: keep it byte for byte, on one line of tags.
			out.push(pad + n.open + n.rawContent + close);
			continue;
		}
		out.push(pad + n.open);
		render(n.children, depth + 1, indent, out);
		out.push(pad + close);
	}
}

function lineOf(src: string, offset: number): number {
	let line = 1;
	const end = Math.min(offset, src.length);
	for (let i = 0; i < end; i++) if (src[i] === "\n") line++;
	return line;
}

export function formatFetchXml(
	text: string,
	options?: FormatFetchXmlOptions,
): FormatFetchXmlResult {
	const indent = options?.indent ?? "  ";
	try {
		const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
		const nodes = parse(src);
		const out: string[] = [];
		render(nodes, 0, indent, out);
		return { ok: true, text: out.join("\n") };
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
