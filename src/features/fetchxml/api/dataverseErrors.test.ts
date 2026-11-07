import { describe, expect, it } from "vitest";
import {
	formatDataverseError,
	isResponseTooLargeError,
	isServiceProtectionError,
	parseDataverseError,
} from "./dataverseErrors";

// #96: errors as the PPTB v1.2.7 host rejects them ("<code>: <message>", "HTTP <status>"), with and
// without Electron's IPC prefix. Messages are the documented ones (Microsoft Learn).

const IPC = "Error invoking remote method 'dataverse.fetchXmlQuery': Error: ";

describe("parseDataverseError (#96)", () => {
	it("reads the hex code and keeps Dataverse's message, without the IPC prefix", () => {
		const info = parseDataverseError(
			new Error(
				`${IPC}Dataverse fetchXmlQuery failed: 0x8004e023: AggregateQueryRecordLimit exceeded. Cannot perform this operation.`,
			),
		);
		expect(info.code).toBe("0x8004e023");
		expect(info.kind).toBe("queryLimit");
		expect(info.message).toBe(
			"Dataverse fetchXmlQuery failed: 0x8004e023: AggregateQueryRecordLimit exceeded. Cannot perform this operation.",
		);
		expect(info.guidance).toMatch(/50,000/);
		expect(info.docUrl).toContain("aggregate-data#aggregate-query-record-limit-error");
	});

	it("matches codes case-insensitively and keeps the code as sent", () => {
		const info = parseDataverseError(
			new Error("0x8004430C: Number of conditions in query exceeded maximum limit."),
		);
		expect(info.code).toBe("0x8004430C");
		expect(info.guidance).toMatch(/500 condition and link-entity/);
	});

	it.each([
		[
			"0x80072322",
			"Number of requests exceeded the limit of 6000 over time window of 300 seconds.",
		],
		[
			"0x80072321",
			"Combined execution time of incoming requests exceeded limit of 1,200,000 milliseconds over time window of 300 seconds.",
		],
		["0x80072326", "Number of concurrent requests exceeded the limit of 52."],
	])("treats %s as a service protection error", (code, message) => {
		const error = new Error(`${code}: ${message}`);
		expect(parseDataverseError(error).kind).toBe("serviceProtection");
		expect(isServiceProtectionError(error)).toBe(true);
	});

	it("treats HTTP 429 without an OData body as a service protection error", () => {
		const info = parseDataverseError(new Error("HTTP 429"));
		expect(info.status).toBe(429);
		expect(info.kind).toBe("serviceProtection");
	});

	it("reads OData error JSON and Retry-After when an error object carries them", () => {
		const info = parseDataverseError({
			error: { code: "0x80072322", message: "Number of requests exceeded the limit." },
			status: 429,
			headers: { "retry-after": "7" },
		});
		expect(info.code).toBe("0x80072322");
		expect(info.message).toBe("0x80072322: Number of requests exceeded the limit.");
		expect(info.status).toBe(429);
		expect(info.retryAfterSeconds).toBe(7);
	});

	it("maps the privilege error and keeps the privilege Dataverse names", () => {
		const message =
			"0x80040220: Principal user (Id=1) is missing prvReadAccount privilege on account entity(OTC=1).";
		const info = parseDataverseError(new Error(message));
		expect(info.kind).toBe("privilege");
		expect(formatDataverseError(new Error(message))).toContain("prvReadAccount");
	});

	it("leaves undocumented codes and plain errors as they are", () => {
		const unknown = parseDataverseError(new Error("0x80049999: Something else."));
		expect(unknown.kind).toBe("unknown");
		expect(unknown.guidance).toBeUndefined();
		expect(formatDataverseError(new Error("0x80049999: Something else."))).toBe(
			"0x80049999: Something else.",
		);
		expect(formatDataverseError(new Error("Request failed: socket hang up"))).toBe(
			"Request failed: socket hang up",
		);
	});

	it("uses the fallback for values without a message", () => {
		expect(formatDataverseError(undefined, "Failed to save view")).toBe("Failed to save view");
		expect(formatDataverseError({}, "Failed")).toBe("Failed");
		expect(formatDataverseError("plain text")).toBe("plain text");
	});
});

describe("formatDataverseError (#96)", () => {
	it("puts the guidance first and Dataverse's text after Details:", () => {
		const text = formatDataverseError(
			new Error("0x8004430D: Number of link entities in query exceeded maximum limit."),
		);
		expect(text).toBe(
			"The query has more than 15 link-entity elements. Remove some link-entities. Details: 0x8004430D: Number of link entities in query exceeded maximum limit.",
		);
	});
});

describe("isResponseTooLargeError (#68)", () => {
	it("detects the host's 16 MB cap messages", () => {
		expect(
			isResponseTooLargeError(new Error("Dataverse response exceeds the maximum size")),
		).toBe(true);
		expect(
			isResponseTooLargeError(
				new Error(`${IPC}Dataverse metadata response exceeds the maximum size`),
			),
		).toBe(true);
		expect(
			isResponseTooLargeError(new Error("0x8004e023: AggregateQueryRecordLimit exceeded.")),
		).toBe(false);
	});

	it("explains the cap and keeps the host's text", () => {
		expect(
			formatDataverseError(new Error("Dataverse response exceeds the maximum size")),
		).toMatch(
			/16 MB limit of Power Platform ToolBox.*Details: Dataverse response exceeds the maximum size$/,
		);
	});
});
