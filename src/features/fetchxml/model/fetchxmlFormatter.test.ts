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

	it("never throws on very deep nesting", () => {
		const r = formatFetchXml("<a>".repeat(200000) + "</a>".repeat(200000));
		expect(typeof r.ok).toBe("boolean");
	});
});
