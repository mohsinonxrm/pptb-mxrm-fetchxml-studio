/**
 * Paging helpers (#58, #60). Annotation shapes are the ones the Web API returned live (2026-10-16)
 * and the Microsoft Learn page-results sample.
 */
import { describe, expect, it } from "vitest";
import {
	advancePaging,
	hasNextPage,
	nextLinkToQueryPath,
	parseFetchXmlPagingCookie,
	readFetchPagingOptions,
	RETRIEVE_ALL_ROW_LIMIT,
	retrieveAllLimitReached,
	type PagingState,
} from "./paging";
import { addPagingToFetchXml } from "../model/fetchxml";

// From https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results (Web API tab)
const LEARN_ANNOTATION =
	'<cookie pagenumber="2" pagingcookie="%253ccookie%2520page%253d%25221%2522%253e%253cfullname%2520last%253d%2522Susanna%2520Stubberod%2520%2528sample%2529%2522%2520first%253d%2522Yvonne%2520McKay%2520%2528sample%2529%2522%2520%252f%253e%253ccontactid%2520last%253d%2522%257b70BF4D48-34CB-ED11-B596-0022481D68CD%257d%2522%2520first%253d%2522%257b49B0BE2E-D01C-ED11-B83E-000D3A572421%257d%2522%2520%252f%253e%253c%252fcookie%253e" istracking="False" />';

describe("parseFetchXmlPagingCookie", () => {
	it("decodes the pagingcookie attribute twice", () => {
		expect(parseFetchXmlPagingCookie(LEARN_ANNOTATION)).toEqual({
			cookie:
				'<cookie page="1"><fullname last="Susanna Stubberod (sample)" first="Yvonne McKay (sample)" /><contactid last="{70BF4D48-34CB-ED11-B596-0022481D68CD}" first="{49B0BE2E-D01C-ED11-B83E-000D3A572421}" /></cookie>',
		});
	});

	it("treats an empty pagingcookie as no cookie (simple paging)", () => {
		expect(parseFetchXmlPagingCookie('<cookie pagenumber="2" pagingcookie="" istracking="False" />')).toEqual({});
	});

	it.each([
		["missing (last page)", undefined],
		["empty string", ""],
		["not XML", "garbage"],
		["a number", 5],
	])("returns nothing for %s", (_label, annotation) => {
		expect(parseFetchXmlPagingCookie(annotation)).toEqual({});
	});

	it("ignores a cookie whose encoding is broken", () => {
		expect(parseFetchXmlPagingCookie('<cookie pagenumber="3" pagingcookie="%E0%A4%A" />')).toEqual({});
	});
});

describe("readFetchPagingOptions (the executed FetchXML, either quote style)", () => {
	it.each([
		['<fetch top="10"><entity name="a" /></fetch>', { top: 10 }],
		["<fetch count='50' page='3'><entity name='a' /></fetch>", { count: 50, page: 3 }],
		['<fetch aggregate="true" count="5"><entity name="a"><attribute name="b" alias="count" /></entity></fetch>', { count: 5 }],
		['<fetch><entity name="a" /></fetch>', {}],
		["no fetch", {}],
	])("%s", (xml, expected) => expect(readFetchPagingOptions(xml)).toEqual(expected));
});

describe("nextLinkToQueryPath", () => {
	it("keeps what follows /api/data/v9.x/ (the host adds the org URL and version)", () => {
		expect(
			nextLinkToQueryPath("https://org.crm.dynamics.com/api/data/v9.2/accounts?savedQuery=00000000-0000-0000-00aa-000010001002&$skiptoken=%3Ccookie%3E"),
		).toBe("accounts?savedQuery=00000000-0000-0000-00aa-000010001002&$skiptoken=%3Ccookie%3E");
	});
	it("rejects anything else", () => expect(nextLinkToQueryPath("accounts?$skiptoken=x")).toBeUndefined());
});

describe("advancePaging / hasNextPage (#60)", () => {
	const start: PagingState = {
		mode: "fetchxml",
		xml: '<fetch count="50" distinct="true"><entity name="systemuser" /></fetch>',
		pageSize: 50,
		currentPage: 1,
		nextPage: 2,
		moreRecords: true,
		isRetrieveAllInProgress: true,
		retrieveAll: true,
	};

	it("advances one page at a time without a cookie (distinct: Dataverse's pagenumber stays at 2)", () => {
		const after2 = advancePaging(start, { moreRecords: true });
		const after3 = advancePaging(after2, { moreRecords: true });
		expect([after2.currentPage, after2.nextPage, after3.currentPage, after3.nextPage]).toEqual([2, 3, 3, 4]);
		expect(after3.pagingCookie).toBeUndefined();
	});

	it("carries the newest cookie and keeps the query, page size and Retrieve All flag", () => {
		const next = advancePaging({ ...start, pagingCookie: '<cookie page="1" />' }, { moreRecords: true, pagingCookie: '<cookie page="2" />' });
		expect(next).toMatchObject({ pagingCookie: '<cookie page="2" />', xml: start.xml, pageSize: 50, retrieveAll: true });
	});

	it("drops the cookie when a page returns none, and stops on the last page", () => {
		const last = advancePaging({ ...start, pagingCookie: '<cookie page="1" />' }, {});
		expect(last).toMatchObject({ moreRecords: false, pagingCookie: undefined });
	});

	it("continues from a query that starts at a later page", () => {
		expect(advancePaging({ ...start, currentPage: 3, nextPage: 4 }, { moreRecords: true })).toMatchObject({ currentPage: 4, nextPage: 5 });
	});

	it("pages saved views only while there's a next link", () => {
		expect(hasNextPage("view", { moreRecords: true, nextLink: "https://org/api/data/v9.2/accounts?$skiptoken=x" })).toBe(true);
		expect(hasNextPage("view", { moreRecords: true })).toBe(false);
		expect(hasNextPage("fetchxml", { moreRecords: true })).toBe(true);
		expect(hasNextPage("fetchxml", { moreRecords: false, nextLink: "x" })).toBe(false);
	});
});

describe("addPagingToFetchXml (#58, #60)", () => {
	it("replaces single-quoted page / paging-cookie / count (no duplicate attributes)", () => {
		const xml = addPagingToFetchXml("<fetch count='3' page='1' paging-cookie='&lt;cookie /&gt;'><entity name='account' /></fetch>", 2, "<cookie page=\"1\" />", 3);
		// A duplicate attribute would make the XML unparseable
		const doc = new DOMParser().parseFromString(xml, "application/xml");
		expect(doc.querySelector("parsererror")).toBeNull();
		const fetch = doc.documentElement;
		expect(fetch.getAttribute("page")).toBe("2");
		expect(fetch.getAttribute("count")).toBe("3");
		expect(fetch.getAttribute("paging-cookie")).toBe('<cookie page="1" />');
	});

	it("keeps the query's own count when no page size is given", () => {
		expect(addPagingToFetchXml('<fetch count="50"><entity name="a" /></fetch>', 2)).toContain('count="50"');
	});

	it("leaves other attributes alone, including ones that contain 'page'", () => {
		const xml = addPagingToFetchXml('<fetch aggregate="true" returntotalrecordcount="true"><entity name="a" /></fetch>', 2, undefined, 10);
		expect(xml).toContain('aggregate="true"');
		expect(xml).toContain('returntotalrecordcount="true"');
		expect(xml).toContain('page="2"');
	});

	it("keeps $ sequences in cookie values literally (record values can contain them)", () => {
		const cookie = `<cookie page="1"><name last="A $& B $1 $$ $' $\`" first="x" /></cookie>`;
		const xml = addPagingToFetchXml('<fetch count="3"><entity name="account" /></fetch>', 2, cookie, 3);
		const fetch = new DOMParser().parseFromString(xml, "application/xml").documentElement;
		expect(fetch.getAttribute("paging-cookie")).toBe(cookie);
		expect(xml.endsWith('<entity name="account" /></fetch>')).toBe(true);
	});

	it("round-trips the decoded Learn cookie into an XML-encoded paging-cookie attribute", () => {
		const { cookie } = parseFetchXmlPagingCookie(LEARN_ANNOTATION);
		const xml = addPagingToFetchXml('<fetch count="3"><entity name="contact" /></fetch>', 2, cookie, 3);
		const fetch = new DOMParser().parseFromString(xml, "application/xml").documentElement;
		expect(fetch.getAttribute("paging-cookie")).toBe(cookie);
		expect(fetch.getAttribute("page")).toBe("2");
	});
});

describe("Retrieve All row limit (#183)", () => {
	it("stops at 100,000 rows", () => {
		expect(RETRIEVE_ALL_ROW_LIMIT).toBe(100_000);
		expect(retrieveAllLimitReached(99_999)).toBe(false);
		expect(retrieveAllLimitReached(100_000)).toBe(true);
		expect(retrieveAllLimitReached(105_000)).toBe(true);
		expect(retrieveAllLimitReached(10, 10)).toBe(true);
	});
});

describe("a raw > inside a <fetch> attribute value (#200)", () => {
	const xml = '<fetch mapping="a>b" count="50" page="3"><entity name="account" /></fetch>';

	it("reads the paging attributes after it", () => {
		expect(readFetchPagingOptions(xml)).toEqual({ count: 50, page: 3 });
	});

	it("sets the paging attributes without cutting the tag", () => {
		expect(addPagingToFetchXml(xml, 4, undefined, 25)).toBe('<fetch mapping="a>b" page="4" count="25"><entity name="account" /></fetch>');
	});
});
