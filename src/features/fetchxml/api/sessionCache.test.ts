import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { memoizeForSession, resetDataverseSessionCache } from "./sessionCache";
import { checkPrivilegeByName, whoAmI } from "./pptbClient";

// #71: WhoAmI and privilege checks are sent once per connection.

afterEach(() => {
	resetDataverseSessionCache();
});

describe("memoizeForSession (#71)", () => {
	it("shares one in-flight request between concurrent callers", async () => {
		const load = vi.fn().mockResolvedValue("user");
		const [a, b] = await Promise.all([
			memoizeForSession("k", load),
			memoizeForSession("k", load),
		]);
		expect(a).toBe("user");
		expect(b).toBe("user");
		expect(load).toHaveBeenCalledTimes(1);
	});

	it("loads again after a failure, or a value it shouldn't keep", async () => {
		const failing = vi
			.fn()
			.mockRejectedValueOnce(new Error("offline"))
			.mockResolvedValueOnce("ok");
		await expect(memoizeForSession("f", failing)).rejects.toThrow("offline");
		await expect(memoizeForSession("f", failing)).resolves.toBe("ok");

		const nullFirst = vi.fn().mockResolvedValueOnce(null).mockResolvedValueOnce("user");
		const keep = (v: string | null) => v !== null;
		await expect(memoizeForSession("n", nullFirst, keep)).resolves.toBeNull();
		await expect(memoizeForSession("n", nullFirst, keep)).resolves.toBe("user");
	});

	it("asks again after resetDataverseSessionCache (connection change)", async () => {
		const load = vi.fn().mockResolvedValueOnce("first").mockResolvedValueOnce("second");
		await expect(memoizeForSession("r", load)).resolves.toBe("first");
		resetDataverseSessionCache();
		await expect(memoizeForSession("r", load)).resolves.toBe("second");
	});

	it("doesn't keep a request that was in flight during a reset", async () => {
		let resolveOld!: (v: string) => void;
		const load = vi
			.fn()
			.mockReturnValueOnce(new Promise<string>((r) => (resolveOld = r)))
			.mockResolvedValueOnce("new");
		const old = memoizeForSession("i", load);
		resetDataverseSessionCache();
		await expect(memoizeForSession("i", load)).resolves.toBe("new");
		resolveOld("old");
		await expect(old).resolves.toBe("old");
		await expect(memoizeForSession("i", load)).resolves.toBe("new");
		expect(load).toHaveBeenCalledTimes(2);
	});
});

describe("whoAmI and checkPrivilegeByName (#71)", () => {
	let execute: ReturnType<typeof vi.fn>;
	let queryData: ReturnType<typeof vi.fn>;

	beforeEach(() => {
		execute = vi
			.fn()
			.mockResolvedValue({ UserId: "u1", BusinessUnitId: "b", OrganizationId: "o" });
		queryData = vi.fn().mockResolvedValue({ RolePrivileges: [{ Depth: "Global" }] });
		(window as unknown as { dataverseAPI: unknown }).dataverseAPI = { execute, queryData };
	});

	afterEach(() => {
		delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
	});

	it("sends WhoAmI once for concurrent startup callers", async () => {
		await Promise.all([whoAmI(), whoAmI(), whoAmI(), whoAmI()]);
		expect(execute).toHaveBeenCalledTimes(1);
	});

	it("sends each privilege check once per user", async () => {
		const results = await Promise.all([
			checkPrivilegeByName("u1", "prvReadSolution"),
			checkPrivilegeByName("u1", "prvReadSolution"),
			checkPrivilegeByName("u1", "prvReadPublisher"),
		]);
		expect(results).toEqual([true, true, true]);
		expect(queryData).toHaveBeenCalledTimes(2);
	});

	it("doesn't keep a failed privilege check", async () => {
		vi.spyOn(console, "error").mockImplementation(() => {});
		queryData.mockRejectedValueOnce(new Error("Request failed: socket hang up"));
		await expect(checkPrivilegeByName("u1", "prvBulkDelete")).resolves.toBe(false);
		await expect(checkPrivilegeByName("u1", "prvBulkDelete")).resolves.toBe(true);
		expect(queryData).toHaveBeenCalledTimes(2);
	});
});
