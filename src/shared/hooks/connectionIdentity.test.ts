import { describe, expect, it } from "vitest";
import { connectionKey, nextConnectionScope, type ConnectionScope } from "./connectionIdentity";

// #66: only a change of the tool's own connection resets connection-scoped state

describe("connectionKey", () => {
	it("is the id and the normalized environment URL", () => {
		expect(connectionKey({ id: "c1", url: "https://Org.crm.dynamics.com/" })).toBe("c1|https://org.crm.dynamics.com");
		expect(connectionKey({ id: "c1", url: "https://org.crm.dynamics.com" })).toBe(connectionKey({ id: "c1", url: "https://ORG.crm.dynamics.com//" }));
	});

	it("differs for another connection to the same environment, or another environment", () => {
		const base = connectionKey({ id: "c1", url: "https://a.crm.dynamics.com" });
		expect(connectionKey({ id: "c2", url: "https://a.crm.dynamics.com" })).not.toBe(base);
		expect(connectionKey({ id: "c1", url: "https://b.crm.dynamics.com" })).not.toBe(base);
	});

	it("is empty without a connection", () => {
		expect(connectionKey(null)).toBe("");
		expect(connectionKey(undefined)).toBe("");
		expect(connectionKey({})).toBe("");
	});
});

describe("nextConnectionScope", () => {
	const start: ConnectionScope = { key: undefined, session: 0 };

	it("records the first connection without starting a new session", () => {
		const { scope, changed } = nextConnectionScope(start, "c1|a");
		expect(changed).toBe(false);
		expect(scope).toEqual({ key: "c1|a", session: 0 });
	});

	it("keeps the same scope object while the host hasn't answered or nothing changed", () => {
		expect(nextConnectionScope(start, undefined).scope).toBe(start);
		const known: ConnectionScope = { key: "c1|a", session: 0 };
		expect(nextConnectionScope(known, "c1|a").scope).toBe(known);
	});

	it("starts a new session when the connection changes, including to and from none", () => {
		let scope: ConnectionScope = { key: "c1|a", session: 0 };
		let step = nextConnectionScope(scope, "c2|b");
		expect(step).toEqual({ scope: { key: "c2|b", session: 1 }, changed: true });
		scope = step.scope;
		step = nextConnectionScope(scope, "");
		expect(step).toEqual({ scope: { key: "", session: 2 }, changed: true });
		step = nextConnectionScope(step.scope, "c1|a");
		expect(step.scope.session).toBe(3);
	});
});
