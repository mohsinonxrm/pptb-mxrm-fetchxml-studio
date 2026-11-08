import { describe, expect, it } from "vitest";
import { isViewDirty, normalizeFetchXml } from "./viewDirty";
import {
	generateLayoutFromFetchXml,
	generateLayoutXml,
	parseLayoutXml,
	reorderColumns,
	updateColumnWidth,
} from "./layoutxml";
import { parseFetchXml } from "./fetchxmlParser";

/** FetchXML as Dataverse stores the Active Accounts system view: legacy attributes and spelled-out defaults */
const STORED_FETCHXML = `<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="account">
    <attribute name="name" />
    <attribute name="primarycontactid" />
    <attribute name="telephone1" />
    <attribute name="accountid" />
    <order attribute="name" descending="false" />
    <filter type="and">
      <condition attribute="statecode" operator="eq" value="0" />
    </filter>
    <link-entity name="contact" from="contactid" to="primarycontactid" visible="false" link-type="outer" alias="accountprimarycontactidcontactcontactid">
      <attribute name="emailaddress1" />
    </link-entity>
  </entity>
</fetch>`;

/** The same query: single quotes, attributes in another order, other whitespace, distinct="false" and descending="false" left out */
const SAME_FETCHXML_REFORMATTED = `<fetch mapping='logical' version='1.0' output-format='xml-platform'><entity name='account'><attribute name='name'/><attribute name='primarycontactid'/>
<attribute name='telephone1'/><attribute name='accountid'/><order attribute='name'/>
<filter type='and'><condition value='0' operator='eq' attribute='statecode'/></filter>
<link-entity alias='accountprimarycontactidcontactcontactid' link-type='outer' to='primarycontactid' from='contactid' name='contact' visible='false'>
<attribute name='emailaddress1'/></link-entity></entity></fetch>`;

/** The query with a filter added: accounts in Seattle */
const FETCHXML_WITH_CITY_FILTER = STORED_FETCHXML.replace(
	'<condition attribute="statecode" operator="eq" value="0" />',
	'<condition attribute="statecode" operator="eq" value="0" />\n      <condition attribute="address1_city" operator="eq" value="Seattle" />',
);

/** The Active Accounts LayoutXML as stored */
const STORED_LAYOUTXML = `<grid name="resultset" object="1" jump="name" select="1" icon="1" preview="1">
  <row name="result" id="accountid">
    <cell name="name" width="300" />
    <cell name="primarycontactid" width="150" />
    <cell name="telephone1" width="100" />
    <cell name="accountprimarycontactidcontactcontactid.emailaddress1" width="150" disableSorting="1" />
  </row>
</grid>`;

/** The same layout: single quotes, <cell> and <grid> attributes in another order, no whitespace */
const SAME_LAYOUTXML_REFORMATTED =
	"<grid preview='1' icon='1' select='1' jump='name' object='1' name='resultset'><row id='accountid' name='result'>" +
	"<cell width='300' name='name'/><cell width='150' name='primarycontactid'/><cell width='100' name='telephone1'/>" +
	"<cell disableSorting='1' width='150' name='accountprimarycontactidcontactcontactid.emailaddress1'/></row></grid>";

const config = () => parseLayoutXml(STORED_LAYOUTXML);
const withLayout = (layoutXml: string) => ({
	baselineFetchXml: STORED_FETCHXML,
	currentFetchXml: STORED_FETCHXML,
	baselineLayoutXml: STORED_LAYOUTXML,
	currentLayoutXml: layoutXml,
});

describe("normalizeFetchXml", () => {
	it("gives the stored and the reformatted query the same text", () => {
		expect(normalizeFetchXml(SAME_FETCHXML_REFORMATTED)).toBe(
			normalizeFetchXml(STORED_FETCHXML),
		);
	});

	it("collapses whitespace in FetchXML it can't parse", () => {
		expect(normalizeFetchXml("<fetch>\n  <entity name='account'>\n")).toBe(
			"<fetch><entity name='account'>",
		);
	});

	it("is empty for an empty query", () => {
		expect(normalizeFetchXml("  \n ")).toBe("");
	});
});

describe("isViewDirty (#278, Q14)", () => {
	it("whitespace, quotes, attribute order and defaults only: clean", () => {
		expect(
			isViewDirty({
				baselineFetchXml: STORED_FETCHXML,
				currentFetchXml: SAME_FETCHXML_REFORMATTED,
				baselineLayoutXml: STORED_LAYOUTXML,
				currentLayoutXml: SAME_LAYOUTXML_REFORMATTED,
			}),
		).toEqual({ dirty: false, reasons: [] });
	});

	it("the layout as FetchXML Studio regenerates it: clean", () => {
		expect(isViewDirty(withLayout(generateLayoutXml(config())))).toEqual({
			dirty: false,
			reasons: [],
		});
	});

	it("a filter added: query", () => {
		expect(
			isViewDirty({
				baselineFetchXml: STORED_FETCHXML,
				currentFetchXml: FETCHXML_WITH_CITY_FILTER,
			}),
		).toEqual({
			dirty: true,
			reasons: ["query"],
		});
	});

	it("an attribute removed: query", () => {
		const current = STORED_FETCHXML.replace('<attribute name="telephone1" />', "");
		expect(
			isViewDirty({ baselineFetchXml: STORED_FETCHXML, currentFetchXml: current }).reasons,
		).toEqual(["query"]);
	});

	it("columns reordered: columns", () => {
		const reordered = generateLayoutXml(reorderColumns(config(), 0, 2));
		expect(isViewDirty(withLayout(reordered))).toEqual({ dirty: true, reasons: ["columns"] });
	});

	it("a column added: columns", () => {
		const added = {
			...config(),
			columns: [...config().columns, { name: "address1_city", width: 150 }],
		};
		expect(isViewDirty(withLayout(generateLayoutXml(added)))).toEqual({
			dirty: true,
			reasons: ["columns"],
		});
	});

	it("a column removed: columns", () => {
		const removed = {
			...config(),
			columns: config().columns.filter((column) => column.name !== "telephone1"),
		};
		expect(isViewDirty(withLayout(generateLayoutXml(removed)))).toEqual({
			dirty: true,
			reasons: ["columns"],
		});
	});

	it("a width changed: widths", () => {
		const wider = generateLayoutXml(updateColumnWidth(config(), "telephone1", 125));
		expect(isViewDirty(withLayout(wider))).toEqual({ dirty: true, reasons: ["widths"] });
	});

	it("reordered and a width changed: columns and widths", () => {
		const both = generateLayoutXml(
			updateColumnWidth(reorderColumns(config(), 3, 0), "name", 250),
		);
		expect(isViewDirty(withLayout(both))).toEqual({
			dirty: true,
			reasons: ["columns", "widths"],
		});
	});

	it("a filter added and a width changed: query and widths", () => {
		const wider = generateLayoutXml(updateColumnWidth(config(), "name", 250));
		expect(
			isViewDirty({
				baselineFetchXml: STORED_FETCHXML,
				currentFetchXml: FETCHXML_WITH_CITY_FILTER,
				baselineLayoutXml: STORED_LAYOUTXML,
				currentLayoutXml: wider,
			}),
		).toEqual({ dirty: true, reasons: ["query", "widths"] });
	});

	it("query, columns and widths together", () => {
		const layout = generateLayoutXml(
			updateColumnWidth(reorderColumns(config(), 0, 1), "telephone1", 200),
		);
		expect(
			isViewDirty({
				baselineFetchXml: STORED_FETCHXML,
				currentFetchXml: FETCHXML_WITH_CITY_FILTER,
				baselineLayoutXml: STORED_LAYOUTXML,
				currentLayoutXml: layout,
			}).reasons,
		).toEqual(["query", "columns", "widths"]);
	});

	it("no view loaded: a query is unsaved, an empty one isn't", () => {
		expect(isViewDirty({ baselineFetchXml: null, currentFetchXml: STORED_FETCHXML })).toEqual({
			dirty: true,
			reasons: ["query"],
		});
		expect(isViewDirty({ baselineFetchXml: null, currentFetchXml: "" })).toEqual({
			dirty: false,
			reasons: [],
		});
	});

	it("columns aren't compared without a current LayoutXML, or without a loaded view", () => {
		expect(
			isViewDirty({ ...withLayout(STORED_LAYOUTXML), currentLayoutXml: undefined }).dirty,
		).toBe(false);
		expect(isViewDirty({ ...withLayout(STORED_LAYOUTXML), currentLayoutXml: "" }).dirty).toBe(
			false,
		);
		expect(
			isViewDirty({
				baselineFetchXml: null,
				currentFetchXml: "",
				currentLayoutXml: STORED_LAYOUTXML,
			}).dirty,
		).toBe(false);
	});

	describe("a view without a stored LayoutXML (compared with the layout generated from its FetchXML)", () => {
		const generated = () => {
			const parsed = parseFetchXml(STORED_FETCHXML);
			return generateLayoutFromFetchXml(parsed.fetchNode!);
		};

		it.each([null, undefined, "", "  "])(
			"baseline %j, layout as generated: clean",
			(baselineLayoutXml) => {
				expect(
					isViewDirty({
						...withLayout(generateLayoutXml(generated())),
						baselineLayoutXml,
					}),
				).toEqual({ dirty: false, reasons: [] });
			},
		);

		it("columns reordered: columns", () => {
			const reordered = generateLayoutXml(reorderColumns(generated(), 0, 2));
			expect(
				isViewDirty({ ...withLayout(reordered), baselineLayoutXml: null }).reasons,
			).toEqual(["columns"]);
		});

		it("a width changed: widths", () => {
			const wider = generateLayoutXml(updateColumnWidth(generated(), "name", 320));
			expect(isViewDirty({ ...withLayout(wider), baselineLayoutXml: "" }).reasons).toEqual([
				"widths",
			]);
		});

		it("default widths come from the attribute types passed", () => {
			const types = new Map([
				["name", "String"],
				["telephone1", "String"],
			]);
			const typed = generateLayoutFromFetchXml(
				parseFetchXml(STORED_FETCHXML).fetchNode!,
				undefined,
				types,
			);
			const layout = generateLayoutXml(typed);
			expect(
				isViewDirty({
					...withLayout(layout),
					baselineLayoutXml: null,
					attributeTypeMap: types,
				}).dirty,
			).toBe(false);
			expect(isViewDirty({ ...withLayout(layout), baselineLayoutXml: null }).reasons).toEqual(
				["widths"],
			);
		});
	});

	describe("widths that aren't numbers", () => {
		const cells = (width: string) =>
			`<grid name="resultset"><row name="result" id="accountid"><cell name="name" width="${width}" /></row></grid>`;
		const compare = (before: string, after: string) =>
			isViewDirty({ ...withLayout(cells(after)), baselineLayoutXml: cells(before) });

		it("the same text: clean", () => {
			expect(compare("auto", "auto").dirty).toBe(false);
		});

		it("other text, or text against a number: widths", () => {
			expect(compare("auto", "*").reasons).toEqual(["widths"]);
			expect(compare("auto", "300").reasons).toEqual(["widths"]);
		});

		it("numbers compare as numbers", () => {
			expect(compare("300", "300.0").dirty).toBe(false);
			expect(compare("300", "301").reasons).toEqual(["widths"]);
		});

		it("a missing width is the default width", () => {
			const noWidth = `<grid name="resultset"><row name="result" id="accountid"><cell name="name" /></row></grid>`;
			expect(
				isViewDirty({ ...withLayout(cells("150")), baselineLayoutXml: noWidth }).dirty,
			).toBe(false);
		});
	});

	describe("paging attributes", () => {
		it("the same query on page 3 with a paging cookie: clean", () => {
			const paged = STORED_FETCHXML.replace(
				'<fetch version="1.0"',
				`<fetch page="3" paging-cookie="&lt;cookie page=&quot;2&quot;&gt;&lt;accountid last=&quot;{1A2B}&quot; /&gt;&lt;/cookie&gt;" version="1.0"`,
			);
			expect(
				isViewDirty({ baselineFetchXml: STORED_FETCHXML, currentFetchXml: paged }),
			).toEqual({ dirty: false, reasons: [] });
		});

		it("count and top still count", () => {
			const withCount = STORED_FETCHXML.replace(
				'<fetch version="1.0"',
				'<fetch count="50" version="1.0"',
			);
			const withTop = STORED_FETCHXML.replace(
				'<fetch version="1.0"',
				'<fetch top="10" version="1.0"',
			);
			expect(
				isViewDirty({ baselineFetchXml: STORED_FETCHXML, currentFetchXml: withCount })
					.reasons,
			).toEqual(["query"]);
			expect(
				isViewDirty({ baselineFetchXml: STORED_FETCHXML, currentFetchXml: withTop })
					.reasons,
			).toEqual(["query"]);
		});
	});

	describe("condition values", () => {
		const inQuery = (value: string) =>
			`<fetch><entity name="account"><filter><condition attribute="accountnumber" operator="in"><value>${value}</value><value>A2</value></condition></filter></entity></fetch>`;

		it("whitespace inside a <value> counts: query", () => {
			expect(
				isViewDirty({ baselineFetchXml: inQuery("0"), currentFetchXml: inQuery(" 0 ") })
					.reasons,
			).toEqual(["query"]);
		});

		it("whitespace between <value> elements doesn't", () => {
			const spaced = inQuery("0").replace("</value><value>", "</value>\n        <value>");
			expect(
				isViewDirty({ baselineFetchXml: inQuery("0"), currentFetchXml: spaced }).dirty,
			).toBe(false);
		});

		it("whitespace inside a value attribute counts: query", () => {
			const eq = (value: string) =>
				`<fetch><entity name="account"><filter><condition attribute="name" operator="eq" value="${value}" /></filter></entity></fetch>`;
			expect(
				isViewDirty({ baselineFetchXml: eq("Contoso"), currentFetchXml: eq("Contoso ") })
					.reasons,
			).toEqual(["query"]);
		});
	});

	it("removing version and output-format counts as a change (Save would write the text without them)", () => {
		const trimmed = STORED_FETCHXML.replace(' version="1.0" output-format="xml-platform"', "");
		expect(
			isViewDirty({ baselineFetchXml: STORED_FETCHXML, currentFetchXml: trimmed }).reasons,
		).toEqual(["query"]);
	});

	it("a LayoutXML that can't be read is compared as text", () => {
		expect(
			isViewDirty({ ...withLayout("<grid><row>"), baselineLayoutXml: "<grid><row>" }).dirty,
		).toBe(false);
		expect(isViewDirty(withLayout("<grid><row>")).reasons).toEqual(["columns"]);
	});
});
