import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import {
	buildViewLayoutXml,
	generateLayoutFromFetchXml,
	generateLayoutXml,
	parseLayoutXml,
} from "./layoutxml";
import type { FetchNode } from "./nodes";

function tree(xml: string): FetchNode {
	const result = parseFetchXml(xml);
	if (!result.fetchNode) throw new Error("parse failed");
	return result.fetchNode;
}

const ACCOUNT = { objectTypeCode: 1, primaryIdAttribute: "accountid", primaryNameAttribute: "name" };

/** A query whose first column is a lookup, then a link-entity column */
const QUERY = tree(`<fetch><entity name="account">
	<attribute name="primarycontactid" />
	<attribute name="name" />
	<link-entity name="contact" from="contactid" to="primarycontactid" alias="pc"><attribute name="emailaddress1" /></link-entity>
</entity></fetch>`);

function grid(xml: string) {
	const doc = new DOMParser().parseFromString(xml, "text/xml");
	return { grid: doc.querySelector("grid")!, row: doc.querySelector("row")! };
}

describe("buildViewLayoutXml (#93)", () => {
	it("uses the primary name attribute for jump, not the first column", () => {
		const xml = buildViewLayoutXml(generateLayoutFromFetchXml(QUERY, "accountid"), QUERY, ACCOUNT);
		const { grid: g, row } = grid(xml);
		expect(g.getAttribute("jump")).toBe("name");
		expect(g.getAttribute("object")).toBe("1");
		expect(row.getAttribute("id")).toBe("accountid");
		expect(g.getAttribute("select")).toBe("1");
		expect(g.getAttribute("icon")).toBe("1");
		expect(g.getAttribute("preview")).toBe("1");
	});

	it("keeps a loaded view's own jump and takes object / row id from metadata", () => {
		const loaded = parseLayoutXml(
			`<grid name="resultset" object="0" jump="accountnumber" select="1" icon="1" preview="1"><row name="result" id="id"><cell name="name" width="300" /></row></grid>`,
		);
		const { grid: g, row } = grid(buildViewLayoutXml(loaded, QUERY, ACCOUNT));
		expect(g.getAttribute("jump")).toBe("accountnumber");
		expect(g.getAttribute("object")).toBe("1");
		expect(row.getAttribute("id")).toBe("accountid");
	});

	it("generates a layout when there is none yet instead of saving an empty one", () => {
		const linkOnly = tree(
			`<fetch><entity name="account"><link-entity name="contact" from="contactid" to="primarycontactid" alias="pc"><attribute name="fullname" /></link-entity></entity></fetch>`,
		);
		const xml = buildViewLayoutXml(null, linkOnly, ACCOUNT);
		expect(xml).toContain('<cell name="pc.fullname"');
		expect(grid(xml).row.getAttribute("id")).toBe("accountid");
	});

	it("leaves jump out when the Entity has no primary name attribute", () => {
		const xml = buildViewLayoutXml(null, QUERY, { objectTypeCode: 1, primaryIdAttribute: "accountid" });
		expect(grid(xml).grid.hasAttribute("jump")).toBe(false);
	});
});

describe("generateLayoutXml (#93)", () => {
	it("never writes a made-up row id", () => {
		const xml = generateLayoutXml({ gridName: "resultset", columns: [{ name: "name", width: 100 }] });
		expect(grid(xml).row.hasAttribute("id")).toBe(false);
	});

	it("round-trips the image provider web resource and function name", () => {
		const source = `<grid name="resultset" object="3" jump="name" select="1" icon="1" preview="1"><row name="result" id="opportunityid"><cell name="opportunityratingcode" width="150" imageproviderwebresource="new_SampleWebResource" imageproviderfunctionname="displayIconTooltip" /></row></grid>`;
		const xml = generateLayoutXml(parseLayoutXml(source));
		expect(xml).toContain('imageproviderwebresource="new_SampleWebResource"');
		expect(xml).toContain('imageproviderfunctionname="displayIconTooltip"');
	});

	it("escapes attribute values", () => {
		const xml = generateLayoutXml({ gridName: "resultset", columns: [{ name: 'a"b', width: 100 }] });
		expect(xml).toContain('name="a&quot;b"');
	});
});
