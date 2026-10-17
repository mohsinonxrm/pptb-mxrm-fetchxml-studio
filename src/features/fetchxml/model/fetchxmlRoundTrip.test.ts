/**
 * Round-trip fidelity (#90): parse → generate must preserve every FetchXML attribute and element,
 * so Parse to Tree never changes what a query means.
 *
 * Compared semantically: documented defaults (e.g. link-type="inner", filter type="and") and
 * equivalent boolean spellings ("1"/"true") are normalised, and child elements are compared per
 * element type. Dataverse "always orders attributes specified by the link-entity after attributes
 * for the entity element" (order-rows), and filter conditions are commutative, so regrouping
 * children by type doesn't change results.
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { generateFetchXml } from "./fetchxml";
import { LEARN_FETCHXML_EXAMPLES } from "./__fixtures__/learnFetchXml";

/** Boolean attributes per element, with their documented default when omitted. */
const BOOLEAN_DEFAULTS: Record<string, Record<string, boolean>> = {
	fetch: {
		aggregate: false,
		distinct: false,
		returntotalrecordcount: false,
		"no-lock": false,
		latematerialize: false,
		useraworderby: false,
	},
	entity: { enableprefiltering: false },
	attribute: { groupby: false, distinct: false, usertimezone: true },
	order: { descending: false },
	"link-entity": { intersect: false },
	filter: {
		isquickfindfields: false,
		overridequickfindrecordlimitenabled: false,
		overridequickfindrecordlimitdisabled: false,
	},
};

/** Non-boolean documented defaults. */
const VALUE_DEFAULTS: Record<string, Record<string, string>> = {
	"link-entity": { "link-type": "inner" },
	filter: { type: "and" },
};

const CHILD_RANK: Record<string, string[]> = {
	entity: ["all-attributes", "attribute", "order", "filter", "link-entity"],
	"link-entity": ["all-attributes", "attribute", "order", "filter", "link-entity"],
	filter: ["condition", "filter", "link-entity"],
};

interface CanonicalElement {
	name: string;
	attributes: Record<string, string>;
	text?: string;
	children: CanonicalElement[];
}

function canonicalize(el: Element): CanonicalElement {
	const name = el.nodeName;
	const attributes: Record<string, string> = {};
	for (const attr of Array.from(el.attributes)) {
		let value = attr.value;
		const lower = attr.name.toLowerCase();
		const boolDefault = BOOLEAN_DEFAULTS[name]?.[lower];
		if (boolDefault !== undefined) {
			const v = value.trim().toLowerCase();
			const bool = v === "true" || v === "1";
			if (bool === boolDefault) continue; // documented default
			value = String(bool);
		} else if (VALUE_DEFAULTS[name]?.[lower] === value) {
			continue;
		} else if (name === "fetch" && lower === "options") {
			value = value
				.split(",")
				.map((h) => h.trim())
				.filter(Boolean)
				.join(",");
		}
		attributes[attr.name] = value;
	}
	const rank = CHILD_RANK[name];
	const children = Array.from(el.children)
		.map(canonicalize)
		.map((c, index) => ({ c, index }))
		.sort((a, b) => {
			if (!rank) return a.index - b.index;
			return rank.indexOf(a.c.name) - rank.indexOf(b.c.name) || a.index - b.index;
		})
		.map(({ c }) => c);
	const text = el.children.length === 0 ? el.textContent?.trim() || undefined : undefined;
	return { name, attributes, ...(text ? { text } : {}), children };
}

function canonical(xml: string): CanonicalElement {
	const doc = new DOMParser().parseFromString(xml, "application/xml");
	expect(doc.querySelector("parsererror"), "source must be well-formed XML").toBeNull();
	return canonicalize(doc.documentElement);
}

function expectLosslessRoundTrip(xml: string) {
	const parsed = parseFetchXml(xml);
	expect(parsed.errors, "must parse").toEqual([]);
	const regenerated = generateFetchXml(parsed.fetchNode!);
	expect(canonical(regenerated)).toEqual(canonical(xml));
	// Idempotent: a second pass produces identical output
	expect(generateFetchXml(parseFetchXml(regenerated).fetchNode!)).toBe(regenerated);
	return { parsed, regenerated };
}

describe("Microsoft Learn FetchXML examples round-trip losslessly", () => {
	it("has the full corpus", () => {
		expect(LEARN_FETCHXML_EXAMPLES.length).toBeGreaterThanOrEqual(60);
	});

	it.each(LEARN_FETCHXML_EXAMPLES.map((e, i) => [`#${i + 1} ${e.source}`, e.xml]))("%s", (_label, xml) => {
		expectLosslessRoundTrip(xml);
	});
});

describe("real-world FetchXML round-trips losslessly", () => {
	it.each([
		[
			"Advanced Find export (version/output-format/mapping/distinct=false)",
			`<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="account">
    <attribute name="name" />
    <order attribute="name" descending="false" />
    <filter type="and">
      <condition attribute="statecode" operator="eq" value="0" />
    </filter>
  </entity>
</fetch>`,
		],
		[
			"saved view with uiname/uitype on a lookup condition",
			`<fetch version="1.0" mapping="logical"><entity name="account"><attribute name="name" /><filter type="and"><condition attribute="ownerid" operator="eq" value="{00000000-0000-0000-0000-000000000001}" uiname="Jane" uitype="systemuser" /></filter></entity></fetch>`,
		],
		[
			"SQL 4 CDS output with a namespace declaration (#48)",
			`<fetch xmlns:generator="MarkMpn.SQL4CDS" top="5"><entity name="account"><attribute name="name" /></entity></fetch>`,
		],
		[
			"namespaced attributes on several elements",
			`<fetch xmlns:x="urn:x" top="5" x:origin="tool"><entity name="account" x:note="n"><attribute name="name" x:hint="h" /><filter type="and" x:f="1"><condition attribute="statecode" operator="eq" value="0" x:c="1" /></filter><link-entity name="contact" from="parentcustomerid" to="accountid" alias="c" x:l="1"><attribute name="fullname" /></link-entity></entity></fetch>`,
		],
		[
			"report prefiltering",
			`<fetch><entity name="account" enableprefiltering="true" prefilterparametername="CRM_FilteredAccount"><attribute name="name" /></entity></fetch>`,
		],
		[
			"every documented fetch attribute",
			`<fetch aggregate="true" aggregatelimit="1000" datasource="retained" latematerialize="true" useraworderby="1" options="HashJoin, DisableRowGoal" no-lock="true"><entity name="task"><attribute name="prioritycode" alias="p" groupby="true" /><attribute name="activityid" alias="n" aggregate="count" /><order alias="n" descending="true" /></entity></fetch>`,
		],
		[
			"distinct count and row aggregate",
			`<fetch aggregate="true"><entity name="task"><attribute name="prioritycode" alias="n" aggregate="countcolumn" distinct="true" /></entity></fetch>`,
		],
		[
			"row aggregate with original casing",
			`<fetch><entity name="account"><attribute name="name" /><attribute name="accountid" alias="kids" rowaggregate="countchildren" /></entity></fetch>`,
		],
		[
			"quick find filter with record-limit override",
			`<fetch><entity name="account"><attribute name="name" /><filter type="or" isquickfindfields="1" overridequickfindrecordlimitdisabled="1"><condition attribute="name" operator="like" value="{0}" /><condition attribute="accountnumber" operator="like" value="{0}" /></filter></entity></fetch>`,
		],
		[
			"link-type natural (undocumented, accepted by Dataverse)",
			`<fetch top="5"><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="accountid" to="parentcustomerid" link-type="natural" alias="a"><attribute name="name" /></link-entity></entity></fetch>`,
		],
		[
			"link-entity without from/to (system many-to-many)",
			`<fetch top="5"><entity name="systemuser"><attribute name="fullname" /><link-entity name="team" alias="t"><attribute name="name" /></link-entity></entity></fetch>`,
		],
		[
			"unknown attributes are kept",
			`<fetch top="5" foo="bar"><entity name="account"><attribute name="name" bogus="x" /></entity></fetch>`,
		],
	])("%s", (_label, xml) => {
		expectLosslessRoundTrip(xml);
	});

	it("user time zone false is kept", () => {
		expectLosslessRoundTrip(
			`<fetch aggregate="true"><entity name="account"><attribute name="createdon" alias="d" groupby="true" dategrouping="day" usertimezone="false" /><attribute name="accountid" alias="n" aggregate="count" /></entity></fetch>`,
		);
	});
});

describe("Parse to Tree no longer reports documented attributes as changes", () => {
	it("a query using every documented attribute has no Parse to Tree changes", async () => {
		const { getParseToTreeChanges } = await import("./fetchxmlParser");
		const parsed = parseFetchXml(
			`<fetch aggregate="true" aggregatelimit="1000" datasource="retained" latematerialize="true" useraworderby="true" options="HashJoin"><entity name="task"><attribute name="prioritycode" alias="p" groupby="true" /><attribute name="activityid" alias="n" aggregate="countcolumn" distinct="true" /><order alias="n" /><filter type="or" isquickfindfields="1"><condition attribute="subject" operator="like" value="a%" /></filter></entity></fetch>`,
		);
		expect(getParseToTreeChanges(parsed)).toEqual([]);
	});
});
