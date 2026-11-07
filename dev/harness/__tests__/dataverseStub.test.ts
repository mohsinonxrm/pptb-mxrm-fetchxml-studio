import { describe, expect, it, vi } from "vitest";
import {
	buildExecuteRequest,
	createDataverseStub,
	formatFunctionParameter,
	getEntitySetName,
	HARNESS_HEADER,
	type FetchLike,
} from "../dataverseStub";

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}) {
	return new Response(body === undefined ? "" : JSON.stringify(body), {
		status,
		headers: { "content-type": "application/json", ...headers },
	});
}

function mockFetch(response: () => Response) {
	return vi.fn<FetchLike>(async () => response());
}

function headersOf(fetchImpl: ReturnType<typeof mockFetch>, call = 0): Record<string, string> {
	return fetchImpl.mock.calls[call][1]?.headers as Record<string, string>;
}

describe("getEntitySetName (mirrors PPTB host pluralization)", () => {
	it.each([
		["account", "accounts"],
		["savedquery", "savedqueries"],
		["userquery", "userqueries"],
		["systemuser", "systemusers"],
		["opportunity", "opportunities"],
		["activityparty", "activityparties"],
		["survey", "surveys"],
		["mailbox", "mailboxes"],
		["status", "statuses"],
		["branch", "branches"],
		["webresource", "webresourceset"],
		["principalobjectaccess", "principalobjectaccessset"],
		["Account", "accounts"],
	])("%s → %s", (logicalName, expected) => {
		expect(getEntitySetName(logicalName)).toBe(expected);
	});
});

describe("formatFunctionParameter", () => {
	it("quotes and encodes plain strings, doubling single quotes", () => {
		expect(formatFunctionParameter("<fetch/>")).toBe(encodeURIComponent("'<fetch/>'"));
		expect(formatFunctionParameter("O'Neil x")).toBe(encodeURIComponent("'O''Neil x'"));
	});

	it("passes enum literals unquoted", () => {
		const value = "Microsoft.Dynamics.CRM.EntityFilters'Entity,Attributes'";
		expect(decodeURIComponent(formatFunctionParameter(value))).toBe(value);
	});

	it("formats primitives, null and dates", () => {
		expect(formatFunctionParameter(1033)).toBe("1033");
		expect(formatFunctionParameter(true)).toBe("true");
		expect(formatFunctionParameter(null)).toBe("null");
		expect(formatFunctionParameter(undefined)).toBe("null");
		expect(formatFunctionParameter(new Date("2026-01-02T03:04:05Z"))).toBe(
			"2026-01-02T03:04:05.000Z",
		);
	});

	it("converts entity references to @odata.id", () => {
		const id = "00000000-0000-0000-0000-000000000001";
		expect(
			decodeURIComponent(formatFunctionParameter({ entityLogicalName: "account", id })),
		).toBe(JSON.stringify({ "@odata.id": `accounts(${id})` }));
		expect(() => formatFunctionParameter({ entityLogicalName: 1, id: "x" })).toThrow(
			/EntityReference/,
		);
	});
});

describe("buildExecuteRequest", () => {
	it("builds unbound functions with parameter aliases (FetchXMLToSQL)", () => {
		const req = buildExecuteRequest({
			operationName: "FetchXMLToSQL",
			operationType: "function",
			parameters: { FetchXml: "<fetch/>" },
		});
		expect(req.method).toBe("GET");
		expect(req.body).toBeUndefined();
		expect(req.path).toBe(
			`FetchXMLToSQL(FetchXml=@p0)?@p0=${encodeURIComponent("'<fetch/>'")}`,
		);
	});

	it("leaves functions without parameters bare and adds () for empty parameters", () => {
		expect(
			buildExecuteRequest({ operationName: "WhoAmI", operationType: "function" }).path,
		).toBe("WhoAmI");
		expect(
			buildExecuteRequest({
				operationName: "WhoAmI",
				operationType: "function",
				parameters: {},
			}).path,
		).toBe("WhoAmI()");
	});

	it("sends action parameters as the POST body", () => {
		expect(
			buildExecuteRequest({
				operationName: "PublishXml",
				operationType: "action",
				parameters: { ParameterXml: "<importexportxml />" },
			}),
		).toEqual({
			path: "PublishXml",
			method: "POST",
			body: { ParameterXml: "<importexportxml />" },
		});
	});

	it("builds bound operations on entity sets and records", () => {
		expect(
			buildExecuteRequest({
				entityName: "savedquery",
				operationName: "ValidateSavedQuery",
				operationType: "action",
			}).path,
		).toBe("savedqueries/Microsoft.Dynamics.CRM.ValidateSavedQuery");
		expect(
			buildExecuteRequest({
				entityName: "workflow",
				entityId: "abc",
				operationName: "ExecuteWorkflow",
				operationType: "action",
			}).path,
		).toBe("workflows(abc)/Microsoft.Dynamics.CRM.ExecuteWorkflow");
	});
});

describe("createDataverseStub", () => {
	it("queryData hits the proxy with PPTB's default headers", async () => {
		const fetchImpl = mockFetch(() => jsonResponse({ value: [{ name: "x" }] }));
		const api = createDataverseStub(fetchImpl);
		const result = await api.queryData("?accounts?$select=name");
		expect(result.value).toHaveLength(1);
		const [url, init] = fetchImpl.mock.calls[0];
		expect(url).toBe("/__dv/accounts?$select=name");
		expect(init?.method).toBe("GET");
		const headers = headersOf(fetchImpl);
		expect(headers.Prefer).toBe('return=representation,odata.include-annotations="*"');
		expect(headers[HARNESS_HEADER]).toBe("1");
		expect(headers["OData-Version"]).toBe("4.0");
		expect(headers.Authorization).toBeUndefined();
	});

	it("lets an additional Prefer header replace the default extras", async () => {
		const fetchImpl = mockFetch(() => jsonResponse({ value: [] }));
		const api = createDataverseStub(fetchImpl);
		await api.queryData("accounts", undefined, {
			Prefer: "odata.maxpagesize=10",
			"MSCRM.SolutionUniqueName": "fxs",
		});
		const headers = headersOf(fetchImpl);
		expect(headers.Prefer).toBe("return=representation,odata.maxpagesize=10");
		expect(headers["MSCRM.SolutionUniqueName"]).toBe("fxs");
	});

	it("fetchXmlQuery pluralizes the root entity and requests annotations", async () => {
		const fetchImpl = mockFetch(() => jsonResponse({ value: [] }));
		const api = createDataverseStub(fetchImpl);
		const fetchXml = `<fetch top="5"><entity name='opportunity'><attribute name="name"/></entity></fetch>`;
		await api.fetchXmlQuery(fetchXml);
		expect(fetchImpl.mock.calls[0][0]).toBe(
			`/__dv/opportunities?fetchXml=${encodeURIComponent(fetchXml)}`,
		);
		expect(headersOf(fetchImpl).Prefer).toBe(
			'return=representation,odata.include-annotations="*"',
		);
		await expect(api.fetchXmlQuery("<fetch/>")).rejects.toThrow(/entity name/);
	});

	it("getEntityMetadata uses the LogicalName key and no annotations", async () => {
		const fetchImpl = mockFetch(() => jsonResponse({ LogicalName: "account" }));
		const api = createDataverseStub(fetchImpl);
		await api.getEntityMetadata("account", true);
		await api.getEntityMetadata("0000-1", false, ["LogicalName", "EntitySetName"]);
		expect(fetchImpl.mock.calls[0][0]).toBe("/__dv/EntityDefinitions(LogicalName='account')");
		expect(headersOf(fetchImpl).Prefer).toBe("return=representation");
		expect(fetchImpl.mock.calls[1][0]).toBe(
			"/__dv/EntityDefinitions(0000-1)?$select=LogicalName,EntitySetName",
		);
		await expect(api.getEntityMetadata(" ", true)).rejects.toThrow(/cannot be empty/);
	});

	it("create returns the id from OData-EntityId", async () => {
		const id = "11111111-2222-3333-4444-555555555555";
		const fetchImpl = mockFetch(() =>
			jsonResponse({ name: "My view" }, 201, {
				"OData-EntityId": `https://org.crm.dynamics.com/api/data/v9.2/userqueries(${id})`,
			}),
		);
		const api = createDataverseStub(fetchImpl);
		const created = await api.create("userquery", { name: "My view" });
		expect(created.id).toBe(id);
		expect(created.name).toBe("My view");
		const [url, init] = fetchImpl.mock.calls[0];
		expect(url).toBe("/__dv/userqueries");
		expect(init?.method).toBe("POST");
		expect(JSON.parse(init?.body as string)).toEqual({ name: "My view" });
	});

	it("update PATCHes and delete DELETEs the pluralized record", async () => {
		const fetchImpl = mockFetch(() => new Response(null, { status: 204 }));
		const api = createDataverseStub(fetchImpl);
		await api.update("savedquery", "abc", { name: "x" });
		await api.delete("account", "def");
		expect(fetchImpl.mock.calls[0][0]).toBe("/__dv/savedqueries(abc)");
		expect(fetchImpl.mock.calls[0][1]?.method).toBe("PATCH");
		expect(fetchImpl.mock.calls[1][0]).toBe("/__dv/accounts(def)");
		expect(fetchImpl.mock.calls[1][1]?.method).toBe("DELETE");
		expect(fetchImpl.mock.calls[1][1]?.body).toBeUndefined();
	});

	it("execute WhoAmI issues a GET", async () => {
		const fetchImpl = mockFetch(() => jsonResponse({ UserId: "u" }));
		const api = createDataverseStub(fetchImpl);
		expect(await api.execute({ operationName: "WhoAmI", operationType: "function" })).toEqual({
			UserId: "u",
		});
		expect(fetchImpl.mock.calls[0][0]).toBe("/__dv/WhoAmI");
		expect(fetchImpl.mock.calls[0][1]?.method).toBe("GET");
	});

	it("rejects with '<code>: <message>' for Dataverse errors", async () => {
		const api = createDataverseStub(
			mockFetch(() =>
				jsonResponse({ error: { code: "0x80040217", message: "Not found" } }, 404),
			),
		);
		await expect(api.queryData("accounts(x)")).rejects.toThrow("0x80040217: Not found");
	});

	it("rejects with 'HTTP <status>: <body>' for non-JSON errors", async () => {
		const api = createDataverseStub(
			mockFetch(() => new Response("Bad gateway", { status: 502 })),
		);
		await expect(api.queryData("accounts")).rejects.toThrow("HTTP 502: Bad gateway");
		await expect(api.queryData(" ")).rejects.toThrow(/cannot be empty/);
	});
});
