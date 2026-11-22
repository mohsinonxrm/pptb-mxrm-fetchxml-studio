/**
 * The Advisories tab's list (#97): FetchXML Studio's findings with titles, locations and Learn
 * links, merged with what ValidateFetchXmlExpression returned live (sandbox, 2026-10-18).
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import {
	advisorEntityNames,
	collectAdvisories,
	countByLevel,
	issueAttribute,
	issueEntity,
	mergeDataverseCheck,
	nodeLocations,
	type DataverseIssue,
} from "./advisories";
import type { AdvisorMetadata } from "./queryAdvisor";
import type { FetchNode } from "./nodes";

const parse = (xml: string): FetchNode => {
	const r = parseFetchXml(xml);
	expect(r.success, r.errors.map((e) => e.message).join("; ")).toBe(true);
	return r.fetchNode!;
};

const SCENARIO = `<fetch count="50" page="1"><entity name="contact"><attribute name="fullname" /><attribute name="statuscode" /><order attribute="statuscode" /><filter><condition attribute="fullname" operator="like" value="%smith%" /><condition attribute="description" operator="like" value="%vip%" /></filter><link-entity name="account" from="accountid" to="parentcustomerid" alias="a"><attribute name="name" /><order attribute="name" /></link-entity></entity></fetch>`;

const METADATA: AdvisorMetadata = new Map([
	[
		"contact",
		{
			primaryIdAttribute: "contactid",
			attributes: new Map([
				["fullname", { typeName: "StringType", maxLength: 160 }],
				["statuscode", { typeName: "StatusType" }],
				["description", { typeName: "MemoType", maxLength: 2000 }],
			]),
		},
	],
	[
		"account",
		{
			primaryIdAttribute: "accountid",
			attributes: new Map([["name", { typeName: "StringType", maxLength: 160 }]]),
		},
	],
]);

const bag = (pairs: Record<string, string>) => ({
	Count: Object.keys(pairs).length,
	Keys: Object.keys(pairs),
	Values: Object.values(pairs),
});

/** ValidateFetchXmlExpression's messages for SCENARIO, as returned live (2026-10-18) */
const LIVE_SCENARIO_ISSUES: DataverseIssue[] = [
	{
		Severity: 1,
		TypeCode: 3,
		LocalizedMessageText:
			"Like Condition on Attribute description will not perform well because the attribute is too long to be indexed. Attribute Type: Memo",
		OptionalPropertyBag: bag({
			Attribute: "description",
			Reason: "Memo",
			ConditionOperator: "Like",
		}),
	},
	{
		Severity: 1,
		TypeCode: 2,
		LocalizedMessageText:
			"Leading wild card on like/notlike conditions can lead to slow performance. Attribute fullname",
		OptionalPropertyBag: bag({ Attribute: "fullname", ConditionOperator: "Like" }),
	},
	{
		Severity: 1,
		TypeCode: 2,
		LocalizedMessageText:
			"Leading wild card on like/notlike conditions can lead to slow performance. Attribute description",
		OptionalPropertyBag: bag({ Attribute: "description", ConditionOperator: "Like" }),
	},
	{
		Severity: 1,
		TypeCode: 5,
		LocalizedMessageText:
			"Ordering on columns from joined tables can lead to slow queries. Attribute name",
		OptionalPropertyBag: bag({ Attribute: "name" }),
	},
	{
		Severity: 1,
		TypeCode: 4,
		LocalizedMessageText:
			"Ordering on Enum variables can lead to slow performance. Attribute statuscode",
		OptionalPropertyBag: bag({ Attribute: "statuscode" }),
	},
];
const HELP = "https://go.microsoft.com/fwlink/?linkid=2161205";

describe("collectAdvisories", () => {
	it("without metadata: the query rules' findings (wildcards, link-entity order)", () => {
		const list = collectAdvisories(parse(SCENARIO));
		expect(list.map((a) => `${a.rule}:${a.target}`)).toEqual([
			"leading-wildcard:fullname",
			"leading-wildcard:description",
			"order-link:name",
		]);
		expect(list.every((a) => a.level === "advice" && a.source === "fxs")).toBe(true);
	});

	it("with metadata: all six findings of the scenario, no request needed", () => {
		const list = collectAdvisories(parse(SCENARIO), METADATA);
		expect(list.map((a) => `${a.rule}:${a.target}`).sort()).toEqual([
			"large-text:description",
			"leading-wildcard:description",
			"leading-wildcard:fullname",
			"order-choice:statuscode",
			"order-link:name",
			"paging-order:contactid",
		]);
	});

	it("titles, locations and Learn links", () => {
		const list = collectAdvisories(parse(SCENARIO), METADATA);
		const wildcard = list.find((a) => a.target === "fullname")!;
		expect(wildcard).toMatchObject({
			title: "Leading wildcard",
			location: "contact › filter › fullname like",
			doc: { label: "Query anti-patterns" },
		});
		expect(wildcard.doc.href).toContain(
			"query-antipatterns#avoid-leading-wild-cards-in-filter-conditions",
		);
		expect(list.find((a) => a.rule === "order-link")).toMatchObject({
			location: "contact › account (a) › order name",
			doc: { label: "Page results" },
		});
		expect(list.find((a) => a.rule === "order-choice")!.doc.href).toContain(
			"#avoid-ordering-by-choice-columns",
		);
		expect(list.find((a) => a.rule === "large-text")!.doc.href).toContain(
			"#avoid-using-conditions-on-large-text-columns",
		);
		expect(list.find((a) => a.rule === "paging-order")).toMatchObject({
			location: "contact",
			title: "Paging without a unique order",
		});
		expect(list.find((a) => a.rule === "paging-order")!.doc.href).toContain(
			"order-rows#ordering-and-paging",
		);
	});

	it("errors and warnings come first, with level titles and reference links for other rules", () => {
		const fetch = parse(
			`<fetch top="10" count="5"><entity name="account"><attribute name="name" /><filter><condition attribute="name" operator="like" value="%a" /></filter></entity></fetch>`,
		);
		const list = collectAdvisories(fetch);
		expect(list[0]).toMatchObject({
			level: "error",
			title: "Dataverse rejects this",
			location: "fetch",
		});
		expect(list[0].doc.href).toContain("fetchxml/reference/fetch");
		expect(list[list.length - 1].level).toBe("advice");
		expect(countByLevel(list)).toMatchObject({ error: 1 });
	});

	it("leaves out settings Dataverse ignores (shown in the editors, not advice)", () => {
		const fetch = parse(
			`<fetch><entity name="account"><attribute name="name" distinct="true" /></entity></fetch>`,
		);
		expect(collectAdvisories(fetch)).toEqual([]);
	});
});

describe("nodeLocations", () => {
	it("names nested filters once and qualifies entityname", () => {
		const fetch = parse(
			`<fetch><entity name="contact"><attribute name="fullname" /><filter><filter type="or"><condition entityname="a" attribute="name" operator="eq" value="x" /></filter></filter><link-entity name="account" from="accountid" to="parentcustomerid" alias="a" /><order entityname="a" attribute="name" /></entity></fetch>`,
		);
		const locations = [...nodeLocations(fetch).values()];
		expect(locations).toContain("contact › filter › a.name eq");
		expect(locations).toContain("contact › order a.name");
		expect(locations).toContain("contact › account (a)");
	});
});

describe("issueAttribute", () => {
	it("reads the property bag, else the message", () => {
		expect(issueAttribute(LIVE_SCENARIO_ISSUES[0])).toBe("description");
		expect(
			issueAttribute({
				Severity: 1,
				TypeCode: 4,
				LocalizedMessageText:
					"Ordering on Enum variables can lead to slow performance. Attribute statuscode",
			}),
		).toBe("statuscode");
		expect(
			issueAttribute({
				Severity: 1,
				TypeCode: 3,
				LocalizedMessageText:
					"Like Condition on Attribute traversedpath will not perform well because the attribute is too long to be indexed. Attribute Length: 2500",
			}),
		).toBe("traversedpath");
		expect(
			issueAttribute({
				Severity: 1,
				TypeCode: 6,
				LocalizedMessageText:
					"Saved queries with too many select attributes can cause performance problems due to large data size. Select less columns",
			}),
		).toBeUndefined();
		expect(
			issueAttribute({
				Severity: 1,
				TypeCode: 9,
				LocalizedMessageText: "Attribute Type: Memo",
			}),
		).toBeUndefined();
	});
});

describe("mergeDataverseCheck", () => {
	it("scenario with metadata: Dataverse agrees with all five of its findings, nothing new", () => {
		const fetch = parse(SCENARIO);
		const { advisories, added } = mergeDataverseCheck(
			fetch,
			collectAdvisories(fetch, METADATA),
			{ issues: LIVE_SCENARIO_ISSUES, helplink: HELP },
		);
		expect(added).toBe(0);
		expect(
			advisories
				.filter((a) => a.source === "both")
				.map((a) => `${a.rule}:${a.target}`)
				.sort(),
		).toEqual([
			"large-text:description",
			"leading-wildcard:description",
			"leading-wildcard:fullname",
			"order-choice:statuscode",
			"order-link:name",
		]);
		expect(advisories.find((a) => a.rule === "paging-order")!.source).toBe("fxs");
		expect(advisories.find((a) => a.rule === "order-choice")!.dataverseText).toContain(
			"Enum variables",
		);
	});

	it("scenario without metadata: the choice order and large text rows come from Dataverse, at their nodes", () => {
		const fetch = parse(SCENARIO);
		const { advisories, added } = mergeDataverseCheck(fetch, collectAdvisories(fetch), {
			issues: LIVE_SCENARIO_ISSUES,
			helplink: HELP,
		});
		expect(added).toBe(2);
		const fromDataverse = advisories.filter((a) => a.source === "dataverse");
		expect(fromDataverse.map((a) => a.title).sort()).toEqual([
			"Condition on a large text attribute",
			"Order on a choice attribute",
		]);
		const choice = fromDataverse.find((a) => a.rule === "order-choice")!;
		expect(choice).toMatchObject({
			location: "contact › order statuscode",
			nodeId: fetch.entity.orders[0].id,
			level: "advice",
		});
		expect(choice.dataverseText).toBe(LIVE_SCENARIO_ISSUES[4].LocalizedMessageText);
		expect(fromDataverse.find((a) => a.rule === "large-text")!.nodeId).toBe(
			fetch.entity.filters[0].conditions[1].id,
		);
	});

	it("too many select attributes (TypeCode 6) agrees with all-attributes", () => {
		const fetch = parse(`<fetch><entity name="contact"><all-attributes /></entity></fetch>`);
		const issue: DataverseIssue = {
			Severity: 1,
			TypeCode: 6,
			LocalizedMessageText:
				"Saved queries with too many select attributes can cause performance problems due to large data size. Select less columns",
			OptionalPropertyBag: bag({ AttributeCount: "234", AttributeLimit: "100" }),
		};
		const { advisories, added } = mergeDataverseCheck(fetch, collectAdvisories(fetch), {
			issues: [issue],
		});
		expect(added).toBe(0);
		expect(advisories.find((a) => a.rule === "all-attributes")!.source).toBe("both");
	});

	it("an unknown TypeCode, or no attribute: a From Dataverse row quoting the message, with the help link", () => {
		const fetch = parse(
			`<fetch><entity name="contact"><attribute name="fullname" /></entity></fetch>`,
		);
		const issues: DataverseIssue[] = [
			{ Severity: 1, TypeCode: 99, LocalizedMessageText: "Something new." },
			{
				Severity: 1,
				TypeCode: 4,
				LocalizedMessageText: "Ordering on Enum variables can lead to slow performance.",
			},
		];
		const { advisories, added } = mergeDataverseCheck(fetch, collectAdvisories(fetch), {
			issues,
			helplink: HELP,
		});
		expect(added).toBe(2);
		expect(
			advisories.every(
				(a) =>
					a.source === "dataverse" &&
					a.title === "Reported by Dataverse" &&
					a.doc.href === HELP,
			),
		).toBe(true);
		expect(advisories.map((a) => a.dataverseText)).toEqual([
			"Something new.",
			"Ordering on Enum variables can lead to slow performance.",
		]);
		expect(advisories[0].nodeId).toBe(fetch.entity.id);
	});

	describe("which entity the attribute belongs to", () => {
		// The root's name and a link-entity's name, both with leading wildcards
		const XML = `<fetch><entity name="account"><attribute name="name" /><link-entity name="account" from="accountid" to="parentaccountid" alias="p"><filter><condition attribute="name" operator="like" value="%p%" /></filter></link-entity><filter><condition attribute="name" operator="like" value="%a%" /></filter></entity></fetch>`;
		const NAME_WILDCARD: DataverseIssue = {
			Severity: 1,
			TypeCode: 2,
			LocalizedMessageText:
				"Leading wild card on like/notlike conditions can lead to slow performance. Attribute name",
			OptionalPropertyBag: bag({ Attribute: "name", ConditionOperator: "Like" }),
		};

		it("records the entity of each finding (root, link-entity, entityname)", () => {
			const fetch = parse(
				`<fetch><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="accountid" to="parentcustomerid" alias="a" /><filter><condition entityname="a" attribute="name" operator="like" value="%x" /><condition attribute="fullname" operator="like" value="%y" /></filter></entity></fetch>`,
			);
			const list = collectAdvisories(fetch);
			expect(list.find((a) => a.target === "name")!.entity).toEqual({
				name: "account",
				alias: "a",
				isRoot: false,
			});
			expect(list.find((a) => a.target === "fullname")!.entity).toEqual({
				name: "contact",
				alias: undefined,
				isRoot: true,
			});
		});

		it("without an entity in the message, it agrees with the root entity's finding, even when the link-entity's comes first", () => {
			const fetch = parse(XML);
			// The link-entity's finding first in the list, so list order alone would pick it
			const own = collectAdvisories(fetch).sort(
				(a, b) => Number(!!a.entity?.isRoot) - Number(!!b.entity?.isRoot),
			);
			expect(own.map((a) => a.entity?.isRoot)).toEqual([false, true]);
			const { advisories, added } = mergeDataverseCheck(fetch, own, {
				issues: [NAME_WILDCARD],
			});
			expect(added).toBe(0);
			const agreed = advisories.filter((a) => a.source === "both");
			expect(agreed).toHaveLength(1);
			expect(agreed[0].entity).toMatchObject({ isRoot: true });
			expect(agreed[0].nodeId).toBe(fetch.entity.filters[0].conditions[0].id);
		});

		it("with an entity or alias in the message, it agrees only with that entity's finding", () => {
			const fetch = parse(XML);
			const withAlias: DataverseIssue = {
				...NAME_WILDCARD,
				OptionalPropertyBag: bag({ Attribute: "name", EntityAlias: "p" }),
			};
			const { advisories } = mergeDataverseCheck(fetch, collectAdvisories(fetch), {
				issues: [withAlias],
			});
			expect(
				advisories.filter((a) => a.source === "both").map((a) => a.entity?.alias),
			).toEqual(["p"]);
			expect(issueEntity(withAlias)).toBe("p");
			expect(issueEntity(NAME_WILDCARD)).toBeUndefined();
		});

		it("several plausible candidates: the first in tree order", () => {
			const fetch = parse(
				`<fetch><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="accountid" to="parentcustomerid" alias="a"><filter><condition attribute="name" operator="like" value="%a" /></filter></link-entity><link-entity name="account" from="accountid" to="parentcustomerid" alias="b"><filter><condition attribute="name" operator="like" value="%b" /></filter></link-entity></entity></fetch>`,
			);
			const { advisories } = mergeDataverseCheck(fetch, collectAdvisories(fetch), {
				issues: [NAME_WILDCARD],
			});
			expect(
				advisories.filter((a) => a.source === "both").map((a) => a.entity?.alias),
			).toEqual(["a"]);
		});
	});

	it("one Dataverse message agrees with one finding: two conditions on the same attribute", () => {
		const fetch = parse(
			`<fetch><entity name="contact"><attribute name="fullname" /><filter type="or"><condition attribute="fullname" operator="like" value="%a" /><condition attribute="fullname" operator="like" value="%b" /></filter></entity></fetch>`,
		);
		const issue = LIVE_SCENARIO_ISSUES[1];
		const { advisories } = mergeDataverseCheck(fetch, collectAdvisories(fetch), {
			issues: [issue, issue],
		});
		expect(advisories.filter((a) => a.source === "both")).toHaveLength(2);
	});
});

describe("advisorEntityNames", () => {
	it("the root and every link-entity, including those inside filters", () => {
		const fetch = parse(
			`<fetch><entity name="contact"><attribute name="fullname" /><filter><link-entity name="task" from="regardingobjectid" to="contactid" link-type="any" /></filter><link-entity name="account" from="accountid" to="parentcustomerid"><link-entity name="systemuser" from="systemuserid" to="ownerid" /></link-entity></entity></fetch>`,
		);
		expect(advisorEntityNames(fetch)).toEqual(["contact", "task", "account", "systemuser"]);
	});
});
