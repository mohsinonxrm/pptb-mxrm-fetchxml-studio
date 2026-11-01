import { describe, expect, it } from "vitest";
import { getParseToTreeChanges, parseFetchXml, validateFetchXmlSyntax } from "./fetchxmlParser";
import { generateFetchXml } from "./fetchxml";

const ISSUE_54_FETCHXML = `<fetch aggregate="true">
  <entity name="msdyn_timeentry">
    <attribute name="bookableresourceid" />
    <attribute name="msdyn_duration" alias="duration" aggregate="sum" />
    <filter type="and">
      <condition attribute="msdyn_date" operator="ge" value="2026-01-01 14:06:15" />
      <condition attribute="msdyn_entrystatus" operator="in">
        <value>192350002</value>
        <value>192350003</value>
      </condition>
    </filter>
    <link-entity name="bookableresource" from="bookableresourceid" to="msdyn_bookableresource" alias="BR">
      <filter type="and">
        <condition attribute="wtc_employeeid" operator="eq" value="61540674-5234-ef11-8e4e-0022481069ec" />
      </filter>
    </link-entity>
  </entity>
</fetch>`;

/** Strip generated node ids so trees can be compared structurally. */
function stripIds(value: unknown): unknown {
	if (Array.isArray(value)) return value.map(stripIds);
	if (value && typeof value === "object") {
		return Object.fromEntries(
			Object.entries(value as Record<string, unknown>)
				.filter(([k]) => k !== "id")
				.map(([k, v]) => [k, stripIds(v)]),
		);
	}
	return value;
}

describe("parseFetchXml", () => {
	it("parses a query with aggregates, filters, multi-values and link-entities", () => {
		const result = parseFetchXml(ISSUE_54_FETCHXML);
		expect(result.success).toBe(true);
		expect(result.errors).toEqual([]);

		const root = result.fetchNode!.entity;
		expect(result.fetchNode!.options.aggregate).toBe(true);
		expect(root.name).toBe("msdyn_timeentry");
		expect(root.attributes.map((a) => a.name)).toEqual(["bookableresourceid", "msdyn_duration"]);
		expect(root.attributes[1]).toMatchObject({ alias: "duration", aggregate: "sum" });

		const [statusCondition] = root.filters[0].conditions.filter((c) => c.operator === "in");
		expect(statusCondition.value).toEqual([192350002, 192350003]);

		expect(root.links).toHaveLength(1);
		expect(root.links[0]).toMatchObject({
			name: "bookableresource",
			from: "bookableresourceid",
			to: "msdyn_bookableresource",
			alias: "BR",
			linkType: "inner",
		});
	});

	it("fails on malformed XML", () => {
		const result = parseFetchXml("<fetch><entity name='account'></fetch>");
		expect(result.success).toBe(false);
		expect(result.errors[0].message).toMatch(/XML parsing error/);
	});

	it("fails when the root element is not <fetch>", () => {
		const result = parseFetchXml("<entity name='account' />");
		expect(result.success).toBe(false);
		expect(result.errors[0].message).toMatch(/Root element must be <fetch>/);
	});

	it("fails when <entity> is missing", () => {
		expect(parseFetchXml("<fetch />").success).toBe(false);
	});

	it("fails when entity has no name", () => {
		expect(parseFetchXml("<fetch><entity /></fetch>").success).toBe(false);
	});

	it("fails on empty input", () => {
		expect(parseFetchXml("   ").success).toBe(false);
	});

	it("warns on unknown <fetch> attributes", () => {
		const result = parseFetchXml("<fetch bogus='1'><entity name='account' /></fetch>");
		expect(result.success).toBe(true);
		expect(result.warnings.some((w) => w.attribute === "bogus")).toBe(true);
	});

	it("round-trips through generateFetchXml without structural change", () => {
		const first = parseFetchXml(ISSUE_54_FETCHXML);
		const regenerated = generateFetchXml(first.fetchNode!);
		const second = parseFetchXml(regenerated);
		expect(second.success).toBe(true);
		expect(stripIds(second.fetchNode)).toEqual(stripIds(first.fetchNode));
	});
});

describe("validateFetchXmlSyntax", () => {
	it("accepts a minimal query", () => {
		expect(validateFetchXmlSyntax("<fetch><entity name='account' /></fetch>")).toEqual({ valid: true });
	});

	it("rejects a non-fetch root", () => {
		expect(validateFetchXmlSyntax("<entity name='account' />").valid).toBe(false);
	});
});

describe("XML comments (#200)", () => {
	it("are listed as removed, so the confirmation shows them instead of dropping them silently", () => {
		const result = parseFetchXml(`<!-- saved from FXB --><fetch><!-- columns --><entity name="account"><attribute name="name" /><!-- todo --></entity></fetch>`);
		expect(result.success).toBe(true);
		expect(getParseToTreeChanges(result)).toEqual([
			expect.objectContaining({ kind: "dropped", message: "XML comments aren't kept in the tree; Parse to tree removes 3 comments." }),
		]);
	});

	it("say nothing when there are none, and use the singular for one", () => {
		expect(getParseToTreeChanges(parseFetchXml(`<fetch><entity name="account" /></fetch>`))).toEqual([]);
		expect(getParseToTreeChanges(parseFetchXml(`<fetch><!-- x --><entity name="account" /></fetch>`))[0].message).toMatch(/removes 1 comment\.$/);
	});
});
