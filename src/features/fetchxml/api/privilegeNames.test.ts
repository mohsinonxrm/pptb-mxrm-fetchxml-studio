import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { resetDataverseSessionCache } from "./sessionCache";
import {
	checkDeletePrivilege,
	checkWritePrivilege,
	resolveEntityPrivilegeName,
} from "./pptbClient";

// #327: entity privilege names come from EntityMetadata.Privileges, not prv{Type}{LogicalName}.

/** Privileges (SecurityPrivilegeMetadata) as the Web API returns them, per entity */
const PRIVILEGES: Record<string, { Name: string; PrivilegeType: string }[]> = {
	savedquery: [
		{ Name: "prvCreateQuery", PrivilegeType: "Create" },
		{ Name: "prvReadQuery", PrivilegeType: "Read" },
		{ Name: "prvWriteQuery", PrivilegeType: "Write" },
		{ Name: "prvDeleteQuery", PrivilegeType: "Delete" },
	],
	userquery: [
		{ Name: "prvWriteUserQuery", PrivilegeType: "Write" },
		{ Name: "prvDeleteUserQuery", PrivilegeType: "Delete" },
		{ Name: "prvShareUserQuery", PrivilegeType: "Share" },
	],
	systemuser: [
		{ Name: "prvWriteUser", PrivilegeType: "Write" },
		{ Name: "prvDeleteUser", PrivilegeType: "Delete" },
	],
	account: [
		{ Name: "prvWriteAccount", PrivilegeType: "Write" },
		{ Name: "prvDeleteAccount", PrivilegeType: "Delete" },
		{ Name: "prvAppendToAccount", PrivilegeType: "AppendTo" },
	],
};

let queryData: ReturnType<typeof vi.fn>;
/** Privilege names RetrieveUserPrivilegeByPrivilegeName was asked about */
let checked: string[];

/** A Dataverse holding PRIVILEGES, where the user has every privilege */
async function respond(query: string): Promise<unknown> {
	const definition = /^EntityDefinitions\(LogicalName='([^']+)'\)\?\$select=Privileges$/.exec(
		query,
	);
	if (definition) {
		const privileges = PRIVILEGES[definition[1]];
		if (!privileges) throw new Error("metadata unavailable");
		return { Privileges: privileges };
	}
	const privilege = /PrivilegeName='([^']+)'/.exec(query);
	if (privilege) {
		checked.push(privilege[1]);
		return { RolePrivileges: [{ Depth: "Global" }] };
	}
	throw new Error(`unexpected query ${query}`);
}

beforeEach(() => {
	checked = [];
	queryData = vi.fn(respond);
	const execute = vi
		.fn()
		.mockResolvedValue({ UserId: "u1", BusinessUnitId: "b", OrganizationId: "o" });
	(window as unknown as { dataverseAPI: unknown }).dataverseAPI = { execute, queryData };
});

afterEach(() => {
	resetDataverseSessionCache();
	delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
	vi.restoreAllMocks();
});

const metadataReads = () =>
	queryData.mock.calls.filter(([q]) => String(q).startsWith("EntityDefinitions")).length;

describe("Delete privilege name from entity metadata (#327)", () => {
	it.each([
		["savedquery", "prvDeleteQuery"],
		["userquery", "prvDeleteUserQuery"],
		["systemuser", "prvDeleteUser"],
		["account", "prvDeleteAccount"],
	])("%s checks %s", async (entity, privilege) => {
		await expect(checkDeletePrivilege(entity)).resolves.toBe(true);
		expect(checked).toEqual([privilege]);
	});

	it("resolves the other privilege types too", async () => {
		await expect(resolveEntityPrivilegeName("userquery", "Share")).resolves.toBe(
			"prvShareUserQuery",
		);
		await expect(resolveEntityPrivilegeName("account", "AppendTo")).resolves.toBe(
			"prvAppendToAccount",
		);
		await expect(resolveEntityPrivilegeName("account", "Assign")).resolves.toBeUndefined();
	});

	it("falls back to prvDelete{LogicalName} when the metadata can't be read", async () => {
		const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
		await expect(checkDeletePrivilege("contact")).resolves.toBe(true);
		expect(checked).toEqual(["prvDeleteContact"]);
		expect(warn).toHaveBeenCalled();
	});

	it("denies, without the prvDelete{LogicalName} fallback, when the metadata lists no Delete", async () => {
		PRIVILEGES.noDeleteEntity = [{ Name: "prvReadNoDeleteEntity", PrivilegeType: "Read" }];
		try {
			await expect(checkDeletePrivilege("noDeleteEntity")).resolves.toBe(false);
			expect(checked).toEqual([]);
		} finally {
			delete PRIVILEGES.noDeleteEntity;
		}
	});

	it("doesn't keep a null or malformed metadata answer for the session", async () => {
		queryData
			.mockResolvedValueOnce(null)
			.mockResolvedValueOnce({ Privileges: "not a list" })
			.mockImplementation(respond);
		await expect(resolveEntityPrivilegeName("savedquery", "Delete")).rejects.toThrow();
		await expect(resolveEntityPrivilegeName("savedquery", "Delete")).rejects.toThrow();
		await expect(resolveEntityPrivilegeName("savedquery", "Delete")).resolves.toBe(
			"prvDeleteQuery",
		);
		expect(metadataReads()).toBe(3);
	});

	it("keeps denying, with a warning, when the fallback privilege doesn't exist", async () => {
		const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
		queryData.mockImplementation(async (query: string) => {
			if (query.includes("PrivilegeName='prvDeleteContact'")) {
				throw new Error("0x80040203: privilege does not exist");
			}
			return respond(query);
		});
		await expect(checkDeletePrivilege("contact")).resolves.toBe(false);
		expect(warn).toHaveBeenCalledWith(expect.stringContaining("treating as denied"));
	});

	it("reads Privileges once per entity for the session, shared with the Write check", async () => {
		await checkDeletePrivilege("savedquery");
		await checkDeletePrivilege("savedquery");
		await checkWritePrivilege("savedquery");
		expect(metadataReads()).toBe(1);
		expect(checked).toEqual(["prvDeleteQuery", "prvWriteQuery"]);

		// A connection change clears it
		resetDataverseSessionCache();
		await checkDeletePrivilege("savedquery");
		expect(metadataReads()).toBe(2);
	});
});
