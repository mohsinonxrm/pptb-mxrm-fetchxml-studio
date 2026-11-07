/* eslint-disable no-irregular-whitespace -- the tests use NBSP and other Unicode spaces on purpose */
import { describe, expect, it } from "vitest";
import { formatFetchXml } from "./fetchxmlFormatter";

function fmt(text: string, indent?: string): string {
	const r = formatFetchXml(text, { indent });
	if (!r.ok) throw new Error(r.error);
	return r.text;
}

describe("formatFetchXml", () => {
	it("indents nested fetch / entity / link-entity / filter / condition", () => {
		const src =
			'<fetch top="5"><entity name="account"><attribute name="name"/><link-entity name="contact" from="parentcustomerid" to="accountid" alias="c"><filter type="and"><condition attribute="statecode" operator="eq" value="0"/></filter></link-entity></entity></fetch>';
		expect(fmt(src)).toBe(
			[
				'<fetch top="5">',
				'  <entity name="account">',
				'    <attribute name="name" />',
				'    <link-entity name="contact" from="parentcustomerid" to="accountid" alias="c">',
				'      <filter type="and">',
				'        <condition attribute="statecode" operator="eq" value="0" />',
				"      </filter>",
				"    </link-entity>",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("honours a custom indent", () => {
		expect(fmt("<a><b/></a>", "\t")).toBe("<a>\n\t<b />\n</a>");
	});

	it("keeps <value> text with surrounding spaces and stays on one line", () => {
		const out = fmt(
			'<condition attribute="x" operator="in"><value> 0 </value><value>abc</value></condition>',
		);
		expect(out).toBe(
			[
				'<condition attribute="x" operator="in">',
				"  <value> 0 </value>",
				"  <value>abc</value>",
				"</condition>",
			].join("\n"),
		);
	});

	it("keeps a whitespace-only value", () => {
		expect(fmt("<a><value>  </value></a>")).toBe("<a>\n  <value>  </value>\n</a>");
	});

	it("keeps entities in attributes and text as written", () => {
		const src = `<condition attribute="n" value="a &amp; b &lt; &quot;c&quot; &apos;d&apos; &#10; &#x41;"><value>x &amp; y</value></condition>`;
		const out = fmt(src);
		expect(out).toContain(`value="a &amp; b &lt; &quot;c&quot; &apos;d&apos; &#10; &#x41;"`);
		expect(out).toContain("<value>x &amp; y</value>");
	});

	it("keeps single and double quote characters", () => {
		const out = fmt(`<a   x='1'   y="2"   z='say "hi"'  w="it's"><b q='&apos;'/></a>`);
		expect(out).toBe(
			[`<a x='1' y="2" z='say "hi"' w="it's">`, `  <b q='&apos;' />`, "</a>"].join("\n"),
		);
	});

	it("keeps attribute order and a > inside a value", () => {
		expect(fmt(`<a z="1" b="x>y" a="2"/>`)).toBe(`<a z="1" b="x>y" a="2" />`);
	});

	it("keeps comments on their own line at the right depth", () => {
		const out = fmt(
			"<!-- top --><fetch><!-- c1 --><entity name='a'><!--  c2  --><attribute name='b'/></entity></fetch>",
		);
		expect(out).toBe(
			[
				"<!-- top -->",
				"<fetch>",
				"  <!-- c1 -->",
				"  <entity name='a'>",
				"    <!--  c2  -->",
				"    <attribute name='b' />",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("keeps CDATA as is", () => {
		const out = fmt("<a><value><![CDATA[ <x> & y ]]></value></a>");
		expect(out).toBe("<a>\n  <value><![CDATA[ <x> & y ]]></value>\n</a>");
	});

	it("keeps the XML declaration and processing instructions", () => {
		const out = fmt(
			`<?xml version="1.0" encoding="utf-8"?><?pi  data ?><fetch><entity name="a"/></fetch>`,
		);
		expect(out).toBe(
			[
				`<?xml version="1.0" encoding="utf-8"?>`,
				"<?pi  data ?>",
				"<fetch>",
				'  <entity name="a" />',
				"</fetch>",
			].join("\n"),
		);
	});

	it("leaves already-formatted input unchanged and is idempotent", () => {
		const formatted = [
			'<fetch top="5">',
			'  <entity name="account">',
			'    <attribute name="name" />',
			"    <!-- note -->",
			'    <filter type="or">',
			'      <condition attribute="a" operator="eq">',
			"        <value> 1 </value>",
			"      </condition>",
			"    </filter>",
			"  </entity>",
			"</fetch>",
		].join("\n");
		expect(fmt(formatted)).toBe(formatted);
		const messy =
			'<fetch  top="5">\r\n\r\n<entity name="a" >\n\n<attribute name="b"/></entity>\n</fetch>\n';
		const once = fmt(messy);
		expect(fmt(once)).toBe(once);
	});

	it("handles mixed self-closing and explicit close tags", () => {
		const out = fmt(
			'<fetch><entity name="a"><attribute name="x"/><attribute name="y"></attribute><all-attributes/></entity></fetch>',
		);
		expect(out).toBe(
			[
				"<fetch>",
				'  <entity name="a">',
				'    <attribute name="x" />',
				'    <attribute name="y"></attribute>',
				"    <all-attributes />",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("keeps mixed content (text with child elements) untouched", () => {
		expect(fmt("<a><b>one <i>two</i>  three</b></a>")).toBe(
			"<a>\n  <b>one <i>two</i>  three</b>\n</a>",
		);
	});

	it.each([
		["unclosed", "<fetch><entity>"],
		["mismatched", "<fetch><entity></fetch></entity>"],
		["stray text", "hello <fetch/>"],
		["two roots", "<a/><b/>"],
		["empty", "   "],
		["unquoted value", "<a b=1/>"],
		["duplicate attribute", '<a b="1" b="2"/>'],
		["bare ampersand", "<a>x & y</a>"],
		["< in value", '<a b="<"/>'],
		["unterminated comment", "<a><!-- x </a>"],
		["end tag only", "</a>"],
	])("returns ok: false for malformed input (%s)", (_n, src) => {
		const r = formatFetchXml(src);
		expect(r.ok).toBe(false);
		if (!r.ok) expect(r.error).toMatch(/^Can't format: /);
	});

	it("reports the line of the problem", () => {
		const r = formatFetchXml("<fetch>\n<entity>\n</fetch>");
		expect(r).toEqual({
			ok: false,
			error: "Can't format: the FetchXML isn't well-formed (line 3)",
		});
	});

	it("formats 5,000 conditions and stays idempotent", () => {
		const conds = Array.from(
			{ length: 5000 },
			(_, i) => `<condition attribute="a${i}" operator="eq" value="v &amp; ${i}"/>`,
		).join("");
		const src = `<fetch><entity name="account"><filter type="or">${conds}</filter></entity></fetch>`;
		const out = fmt(src);
		const lines = out.split("\n");
		expect(lines).toHaveLength(5000 + 6);
		expect(lines[3]).toBe('      <condition attribute="a0" operator="eq" value="v &amp; 0" />');
		expect(fmt(out)).toBe(out);
	});

	it("formats 10,000 levels of nesting (no recursion limit)", () => {
		const r = formatFetchXml("<a>".repeat(10000) + "</a>".repeat(10000), { indent: "" });
		expect(r.ok).toBe(true);
		if (r.ok) expect(r.text.split("\n")).toHaveLength(19999);
		const d = formatFetchXml("<a>".repeat(3000) + "<b/>" + "</a>".repeat(3000));
		expect(d.ok).toBe(true);
		if (d.ok) expect(d.text.split("\n")[3000]).toBe(`${"  ".repeat(3000)}<b />`);
	});

	it("never throws on extreme nesting", () => {
		const r = formatFetchXml("<a>".repeat(200000) + "</a>".repeat(200000), { indent: "" });
		expect(typeof r.ok).toBe("boolean");
	});

	it("handles 8,000 attributes quickly", () => {
		const attrs = Array.from({ length: 8000 }, (_, i) => `a${i}="${i}"`).join(" ");
		const t0 = Date.now();
		const r = formatFetchXml(`<a ${attrs}/>`);
		expect(r.ok).toBe(true);
		expect(Date.now() - t0).toBeLessThan(2000);
	});
});

describe("formatFetchXml: canonical layout of messy input", () => {
	const canonical = [
		"<fetch>",
		'  <entity name="a">',
		'    <filter type="and">',
		'      <condition attribute="x" operator="eq" value="1" />',
		"    </filter>",
		"  </entity>",
		"</fetch>",
	].join("\n");
	it.each([
		[
			"everything on one line",
			'<fetch><entity name="a"><filter type="and"><condition attribute="x" operator="eq" value="1"/></filter></entity></fetch>',
		],
		[
			"odd indents",
			'<fetch>\n      <entity name="a">\n <filter type="and">\n\t\t<condition attribute="x" operator="eq" value="1" />\n</filter>\n   </entity>\n</fetch>',
		],
		[
			"tabs",
			'<fetch>\n\t<entity name="a">\n\t\t<filter type="and">\n\t\t\t<condition attribute="x" operator="eq" value="1"/>\n\t\t</filter>\n\t</entity>\n</fetch>',
		],
		[
			"blank lines and CRLF",
			'\r\n<fetch>\r\n\r\n\r\n<entity name="a">\r\n\r\n<filter type="and">\r\n<condition attribute="x" operator="eq" value="1"/>\r\n\r\n</filter></entity>\r\n</fetch>\r\n\r\n',
		],
		[
			"spaces inside tags",
			'<fetch  ><entity   name="a"  ><filter type = "and" ><condition attribute="x"  operator="eq" value="1" /></filter></entity></fetch >',
		],
	])("%s", (_n, src) => {
		expect(formatFetchXml(src)).toEqual({ ok: true, text: canonical });
	});

	it("re-indents whitespace-only text between tags, not the leaf data", () => {
		expect(fmt("<a>\n\n   <b>  x  </b>\n\n</a>")).toBe("<a>\n  <b>  x  </b>\n</a>");
	});
});

describe("formatFetchXml: whitespace is XML whitespace only", () => {
	it("keeps NBSP and other Unicode spaces as text", () => {
		const nbsp = " ";
		expect(fmt(`<a>${nbsp}<b/></a>`)).toBe(`<a>${nbsp}<b/></a>`);
		expect(fmt(`<a>　</a>`)).toBe(`<a>　</a>`);
		expect(fmt(`<a> <!--c--></a>`)).toBe(`<a> <!--c--></a>`);
	});

	it("rejects non-XML whitespace outside the root", () => {
		expect(formatFetchXml(" <a/>").ok).toBe(false);
		expect(formatFetchXml("<a/> ").ok).toBe(false);
		expect(formatFetchXml("<a/>﻿").ok).toBe(false);
	});

	it("keeps the inner content of a leaf with only whitespace and comments or PIs", () => {
		expect(fmt("<value>  <!--c-->  </value>")).toBe("<value>  <!--c-->  </value>");
		expect(fmt("<a><value>  <?p x?>  </value></a>")).toBe(
			"<a>\n  <value>  <?p x?>  </value>\n</a>",
		);
	});
});

describe("formatFetchXml: errors and options", () => {
	it("reports the line in the original text, after a BOM and with bare CR line breaks", () => {
		expect(formatFetchXml("﻿<a>\n<b>\n</a>")).toEqual({
			ok: false,
			error: "Can't format: the FetchXML isn't well-formed (line 3)",
		});
		expect(formatFetchXml("<a>\r<b>\r</a>")).toEqual({
			ok: false,
			error: "Can't format: the FetchXML isn't well-formed (line 3)",
		});
		expect(formatFetchXml("<a>\r\n<b>\r\n</a>")).toEqual({
			ok: false,
			error: "Can't format: the FetchXML isn't well-formed (line 3)",
		});
	});

	it("strips a BOM from the output", () => {
		expect(fmt("﻿<a/>")).toBe("<a />");
		expect(fmt("﻿<?xml version='1.0'?><a/>")).toBe("<?xml version='1.0'?>\n<a />");
	});

	it("rejects an indent that isn't spaces or tabs", () => {
		expect(formatFetchXml("<a><b/></a>", { indent: "x" }).ok).toBe(false);
		expect(formatFetchXml("<a><b/></a>", { indent: " \n" }).ok).toBe(false);
		expect(fmt("<a><b/></a>", "    ")).toBe("<a>\n    <b />\n</a>");
	});
});

describe("formatFetchXml: accepted and rejected syntax", () => {
	it("accepts a DOCTYPE with a quoted >", () => {
		expect(fmt(`<!DOCTYPE a SYSTEM "x>y"><a/>`)).toBe(`<!DOCTYPE a SYSTEM "x>y">\n<a />`);
	});

	it("accepts combining marks in names", () => {
		expect(fmt("<á b̀c='1'/>")).toBe("<á b̀c='1' />");
	});

	it.each([
		["comment with --", "<a><!-- x -- y --></a>"],
		["comment ending --->", "<a><!-- x ---></a>"],
		["declaration not first", "<a/><?xml version='1.0'?>"],
		["whitespace before declaration", " <?xml version='1.0'?><a/>"],
		["declaration inside the root", "<a><?xml version='1.0'?></a>"],
		["<!foo> at the root", "<!foo><a/>"],
		["PI with no target", "<? x?><a/>"],
		["control character", "<a>\u0001</a>"],
		["vertical tab in an attribute", "<a b='\u000B'/>"],
		["undefined entity in text", "<a>&foo;</a>"],
		["undefined entity in a value", "<a b='&nbsp;'/>"],
		["]]> in text", "<a>x ]]> y</a>"],
		["bad entity in an attribute value", "<a b='x & y'/>"],
		["no whitespace between attributes", `<a b="1"c="2"/>`],
		["CDATA at the root", "<![CDATA[x]]><a/>"],
		["DOCTYPE inside an element", "<a><!DOCTYPE a></a>"],
		["lone <", "<a>1 < 2</a>"],
	])("rejects: %s", (_n, src) => {
		expect(formatFetchXml(src).ok).toBe(false);
	});

	it("accepts the five predefined and numeric entities", () => {
		const s = "<a b='&amp;&lt;&gt;&quot;&apos;&#65;&#x41;'>&amp;&lt;&gt;&quot;&apos;</a>";
		expect(fmt(s)).toBe(s);
	});
});
