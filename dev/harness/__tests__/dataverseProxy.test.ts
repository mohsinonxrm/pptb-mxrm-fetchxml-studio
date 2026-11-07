// @vitest-environment node
import { describe, expect, it, vi } from "vitest";
import {
	buildUpstreamUrl,
	createTokenProvider,
	injectStub,
	isTrustedRequest,
	normalizeCallerObjectId,
	normalizeOrgUrl,
	resolveAzCommand,
} from "../vite-plugin-dataverse-proxy";

const ORG = "https://contoso.crm.dynamics.com";

describe("normalizeOrgUrl", () => {
	it("returns the origin", () => {
		expect(normalizeOrgUrl("https://contoso.crm.dynamics.com/")).toBe(ORG);
		expect(normalizeOrgUrl(" https://contoso.crm4.dynamics.com/main.aspx ")).toBe(
			"https://contoso.crm4.dynamics.com",
		);
	});

	it("rejects non-https or odd hosts", () => {
		expect(() => normalizeOrgUrl("http://contoso.crm.dynamics.com")).toThrow();
		expect(() => normalizeOrgUrl("contoso")).toThrow();
		expect(() => normalizeOrgUrl("https://localhost")).toThrow();
	});
});

describe("normalizeCallerObjectId", () => {
	it("accepts a GUID or nothing", () => {
		expect(normalizeCallerObjectId(undefined)).toBeUndefined();
		expect(normalizeCallerObjectId(" ")).toBeUndefined();
		expect(normalizeCallerObjectId(" 3f2504e0-4f89-11d3-9a0c-0305e82c3301 ")).toBe(
			"3f2504e0-4f89-11d3-9a0c-0305e82c3301",
		);
	});

	it("rejects anything else", () => {
		expect(() => normalizeCallerObjectId("not-a-guid")).toThrow(/GUID/);
	});
});

describe("buildUpstreamUrl", () => {
	it("maps proxy paths under /api/data/v9.2/", () => {
		expect(buildUpstreamUrl(ORG, "/WhoAmI").href).toBe(`${ORG}/api/data/v9.2/WhoAmI`);
		expect(buildUpstreamUrl(ORG, "/EntityDefinitions(LogicalName='account')").pathname).toBe(
			"/api/data/v9.2/EntityDefinitions(LogicalName='account')",
		);
		expect(buildUpstreamUrl(ORG, "/accounts?fetchXml=%3Cfetch%2F%3E").search).toBe(
			"?fetchXml=%3Cfetch%2F%3E",
		);
	});

	it("refuses paths that escape the Web API root", () => {
		expect(() => buildUpstreamUrl(ORG, "/../../main.aspx")).toThrow();
		expect(() => buildUpstreamUrl(ORG, "/%2e%2e/%2e%2e/x")).toThrow();
		expect(buildUpstreamUrl(ORG, "//evil.example.com/x").origin).toBe(ORG);
		expect(buildUpstreamUrl(ORG, "@evil.example.com/x").origin).toBe(ORG);
	});
});

describe("isTrustedRequest", () => {
	const base = {
		headers: { host: "localhost:5173", "x-fxs-harness": "1" } as Record<string, string>,
		socket: { remoteAddress: "127.0.0.1" },
	};

	it("accepts the harness page on loopback", () => {
		expect(isTrustedRequest(base)).toBe(true);
		expect(
			isTrustedRequest({
				...base,
				headers: { ...base.headers, origin: "http://localhost:5173" },
			}),
		).toBe(true);
		expect(
			isTrustedRequest({ ...base, headers: { ...base.headers, host: "[::1]:5173" } }),
		).toBe(true);
	});

	it("rejects requests without the harness header", () => {
		expect(isTrustedRequest({ ...base, headers: { host: "localhost:5173" } })).toBe(false);
	});

	it("rejects cross-site origins and rebinding hosts", () => {
		expect(
			isTrustedRequest({
				...base,
				headers: { ...base.headers, origin: "https://evil.example" },
			}),
		).toBe(false);
		expect(
			isTrustedRequest({ ...base, headers: { ...base.headers, host: "evil.example:5173" } }),
		).toBe(false);
	});

	it("rejects non-loopback clients", () => {
		expect(isTrustedRequest({ ...base, socket: { remoteAddress: "10.0.0.5" } })).toBe(false);
	});
});

describe("injectStub", () => {
	it("puts the stub before the app entry", () => {
		const html = `<body><div id="root"></div><script type="module" src="/src/main.tsx"></script></body>`;
		const out = injectStub(html);
		expect(out.indexOf("/dev/harness/pptbStub.ts")).toBeGreaterThan(0);
		expect(out.indexOf("/dev/harness/pptbStub.ts")).toBeLessThan(out.indexOf("/src/main.tsx"));
	});
});

describe("resolveAzCommand", () => {
	it("prefers AZURE_CLI_PATH, then the Windows default, then PATH", () => {
		expect(resolveAzCommand({ AZURE_CLI_PATH: "/opt/az/bin/az" }, "linux", () => false)).toBe(
			"/opt/az/bin/az",
		);
		expect(resolveAzCommand({}, "win32", () => true)).toMatch(/CLI2\\wbin\\az\.cmd$/);
		expect(resolveAzCommand({}, "win32", () => false)).toBe("az");
		expect(resolveAzCommand({}, "linux", () => true)).toBe("az");
	});

	it("rejects shell metacharacters in AZURE_CLI_PATH", () => {
		expect(() => resolveAzCommand({ AZURE_CLI_PATH: 'az" & calc' }, "win32")).toThrow();
	});
});

describe("createTokenProvider", () => {
	it("caches the token until 5 minutes before expiry", async () => {
		let now = 1_000_000;
		const expiresOn = Math.floor((now + 60 * 60 * 1000) / 1000);
		const run = vi.fn(async () =>
			JSON.stringify({ accessToken: `tok-${run.mock.calls.length}`, expires_on: expiresOn }),
		);
		const getToken = createTokenProvider(ORG, run, () => now);

		expect(await getToken()).toBe("tok-1");
		expect(await getToken()).toBe("tok-1");
		expect(run).toHaveBeenCalledTimes(1);
		expect(run.mock.calls[0]).toEqual([
			["account", "get-access-token", "--resource", ORG, "-o", "json"],
		]);

		now += 54 * 60 * 1000; // 6 min before expiry → still cached
		expect(await getToken()).toBe("tok-1");
		now += 2 * 60 * 1000; // 4 min before expiry → refresh
		expect(await getToken()).toBe("tok-2");
		expect(run).toHaveBeenCalledTimes(2);
	});

	it("falls back to expiresOn and de-duplicates concurrent refreshes", async () => {
		const run = vi.fn(async () =>
			JSON.stringify({ accessToken: "tok", expiresOn: "2999-01-01 00:00:00.000000" }),
		);
		const getToken = createTokenProvider(ORG, run);
		const [a, b] = await Promise.all([getToken(), getToken()]);
		expect(a).toBe("tok");
		expect(b).toBe("tok");
		expect(run).toHaveBeenCalledTimes(1);
	});

	it("fails clearly when az returns no token", async () => {
		const getToken = createTokenProvider(ORG, async () => JSON.stringify({}));
		await expect(getToken()).rejects.toThrow(/no access token/);
	});
});
