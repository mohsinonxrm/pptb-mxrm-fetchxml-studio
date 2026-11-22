/**
 * The Save privilege checks (#278): a privilege lookup that fails is "unknown" (failed: true), not a
 * "no", so Save still offers the command and Dataverse decides (owner rule: block only known-invalid).
 * A privilege the environment doesn't have is a known "no", and so is any known "no" among the
 * privileges a system view needs.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, removeHost } from "../../../shared/testing/ui";
import { checkSavedQueryPrivileges, checkUserQueryPrivileges } from "./pptbClient";
import { resetDataverseSessionCache } from "./sessionCache";

const USER = "aaaaaaaa-0000-0000-0000-000000000001";

/** A host whose privilege lookups answer per privilege: true, false, or an error message */
function host(answers: Record<string, boolean | string>) {
	installHost({
		dataverseAPI: {
			queryData: async (query: string) => {
				const name = /PrivilegeName='(\w+)'/.exec(query)?.[1] ?? "";
				const answer = answers[name];
				if (typeof answer === "string") throw new Error(answer);
				return { RolePrivileges: answer ? [{ PrivilegeName: name }] : [] } as never;
			},
			execute: async () => ({ UserId: USER }),
		} as never,
	});
}

afterEach(() => {
	removeHost();
	resetDataverseSessionCache();
	vi.restoreAllMocks();
});

describe("checkUserQueryPrivileges", () => {
	it("is a plain yes or no when the lookup answers", async () => {
		host({ prvWriteUserQuery: true });
		expect(await checkUserQueryPrivileges()).toEqual({ canWrite: true });
		resetDataverseSessionCache();
		host({ prvWriteUserQuery: false });
		expect(await checkUserQueryPrivileges()).toEqual({ canWrite: false });
	});

	it("says the check failed when the lookup fails", async () => {
		vi.spyOn(console, "warn").mockImplementation(() => {});
		host({ prvWriteUserQuery: "Network error" });
		expect(await checkUserQueryPrivileges()).toEqual({ canWrite: false, failed: true });
	});

	it("is a known no when the environment doesn't have the privilege", async () => {
		host({ prvWriteUserQuery: "0x80040203: privilege does not exist" });
		expect(await checkUserQueryPrivileges()).toEqual({ canWrite: false });
	});
});

describe("checkSavedQueryPrivileges", () => {
	it("needs both prvWriteQuery and prvWriteCustomization", async () => {
		host({ prvWriteQuery: true, prvWriteCustomization: true, prvPublishCustomization: true });
		expect(await checkSavedQueryPrivileges()).toEqual({ canWrite: true, canPublish: true });
	});

	it("says the check failed when a lookup fails and nothing known says no", async () => {
		vi.spyOn(console, "warn").mockImplementation(() => {});
		host({
			prvWriteQuery: "Timeout",
			prvWriteCustomization: true,
			prvPublishCustomization: true,
		});
		expect(await checkSavedQueryPrivileges()).toMatchObject({ canWrite: false, failed: true });
	});

	it("is a known no when one of them is no, even if the other lookup failed", async () => {
		vi.spyOn(console, "warn").mockImplementation(() => {});
		host({
			prvWriteQuery: "Timeout",
			prvWriteCustomization: false,
			prvPublishCustomization: true,
		});
		const result = await checkSavedQueryPrivileges();
		expect(result.canWrite).toBe(false);
		expect(result.failed).toBeUndefined();
	});
});
