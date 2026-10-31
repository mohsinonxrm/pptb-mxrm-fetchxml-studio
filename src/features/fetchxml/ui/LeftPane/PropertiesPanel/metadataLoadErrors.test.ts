import { afterEach, describe, expect, it } from "vitest";
import { clearMetadataLoadErrors, metadataLoadError, recordMetadataLoad, retryMetadataLoads } from "./metadataLoadErrors";

describe("metadata load failures for the property editors (#201)", () => {
	afterEach(() => clearMetadataLoadErrors());

	it("records a failure with Dataverse's message, per entity", () => {
		recordMetadataLoad("columns", "account", new Error("Principal user is missing prvReadEntity privilege"));
		expect(metadataLoadError("account")).toEqual({
			kind: "columns",
			message: expect.stringContaining("prvReadEntity"),
		});
		expect(metadataLoadError("contact")).toBeUndefined();
	});

	it("clears the failure once the load works", () => {
		recordMetadataLoad("details", "account", new Error("timeout"));
		recordMetadataLoad("details", "account");
		expect(metadataLoadError("account")).toBeUndefined();
	});

	it("names columns before entity details when both failed", () => {
		recordMetadataLoad("details", "account", new Error("a"));
		recordMetadataLoad("columns", "account", new Error("b"));
		expect(metadataLoadError("account")?.kind).toBe("columns");
	});

	it("keeps failures across a retry until the loads report back, and drops them on a connection change", () => {
		recordMetadataLoad("columns", "account", new Error("x"));
		retryMetadataLoads();
		expect(metadataLoadError("account")).toBeDefined();
		clearMetadataLoadErrors();
		expect(metadataLoadError("account")).toBeUndefined();
	});
});
