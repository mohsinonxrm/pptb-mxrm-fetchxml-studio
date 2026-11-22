import { describe, expect, it } from "vitest";
import { functionWithString, odataString } from "./odata";

describe("odataString", () => {
	it("leaves ordinary names alone", () => {
		expect(odataString("account")).toBe("account");
		expect(odataString("new_custom_entity")).toBe("new_custom_entity");
	});

	it("doubles single quotes so the value can't end the literal", () => {
		expect(odataString("o'brien")).toBe("o''brien");
		expect(odataString("x' or 1 eq 1 or 'y")).toBe("x'' or 1 eq 1 or ''y");
	});
});

describe("functionWithString (#176)", () => {
	// A condition value with every character the issue names, as the serializer writes it
	const value = "O'Brien &amp; #1 50% a+b /?=";
	const fetchXml = `<fetch><entity name="account"><filter><condition attribute="name" operator="eq" value="${value}"/></filter></entity></fetch>`;

	it("passes the value as a parameter alias", () => {
		const url = functionWithString("FetchXmlToQueryExpression", "FetchXml", fetchXml);
		expect(url.startsWith("FetchXmlToQueryExpression(FetchXml=@p1)?@p1='")).toBe(true);
		expect(url.endsWith("'")).toBe(true);
	});

	it("keeps & # % + = / ? from splitting, cutting or changing the query string", () => {
		const url = functionWithString("FetchXMLToSQL", "FetchXml", fetchXml);
		const query = url.slice(url.indexOf("?") + 1);
		// One query parameter, nothing after a #
		expect(query.split("&")).toHaveLength(1);
		expect(query).not.toMatch(/[#+ ]/);
		// Every % starts a valid escape
		expect(query).not.toMatch(/%(?![0-9A-F]{2})/);
	});

	it("round-trips to the OData literal with doubled quotes", () => {
		const url = functionWithString("ValidateFetchXmlExpression", "FetchXml", fetchXml);
		const literal = decodeURIComponent(url.slice(url.indexOf("@p1=") + 4));
		expect(literal).toBe(`'${fetchXml.replace(/'/g, "''")}'`);
	});
});
