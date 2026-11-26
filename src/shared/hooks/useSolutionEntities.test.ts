/**
 * The entities of the selected solutions (#358): they arrive even when the solution components come
 * back before the startup preload of every entity, an empty list isn't cached for the session, and a
 * list from before a connection change isn't cached for the new one.
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, deferred, flush, installHost, removeHost, render } from "../testing/ui";
import { metadataCache } from "../../features/fetchxml/state/cache";
import type { EntityMetadata } from "../../features/fetchxml/api/pptbClient";
import { useSolutionEntities } from "./useSolutionEntities";
import { useSolutionFilter } from "./useSolutionFilter";
import { usePublisherFilter } from "./usePublisherFilter";

const entity = (LogicalName: string) => ({ LogicalName }) as EntityMetadata;
const ALL = [entity("account"), entity("contact"), entity("cr1_thing")];
const SOLUTION = "11111111-1111-1111-1111-111111111111";

/** queryData that answers the solution components at once and holds every entity until released */
function hostWithSlowPreload() {
	const preload = deferred<{ value: unknown[] }>();
	const queryData = vi.fn(async (query: string) => {
		if (query.startsWith("EntityDefinitions")) return preload.promise;
		if (query.startsWith("msdyn_solutioncomponentsummaries"))
			return { value: [{ msdyn_name: "contact" }, { msdyn_name: "cr1_thing" }] };
		return { value: [] };
	});
	installHost({ dataverseAPI: { queryData } as never });
	return { preload, queryData };
}

const names = (list: EntityMetadata[]) => list.map((e) => e.LogicalName);

beforeEach(() => {
	metadataCache.clear();
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	removeHost();
	metadataCache.clear();
	vi.restoreAllMocks();
});

describe("useSolutionEntities", () => {
	let latest: ReturnType<typeof useSolutionEntities>;
	function Probe({ ids }: { ids: string[] }) {
		latest = useSolutionEntities(ids);
		return null;
	}

	it("waits for the entity preload when the solution components come back first (#358)", async () => {
		const { preload } = hostWithSlowPreload();
		await render(createElement(Probe, { ids: [SOLUTION] }));
		// The components are in, the preload isn't: still loading, not an empty answer
		expect(latest.entitiesLoading).toBe(true);
		expect(latest.entities).toEqual([]);
		expect(metadataCache.getFilteredEntities([SOLUTION])).toBeUndefined();

		await act(async () => preload.resolve({ value: ALL }));
		await flush();
		expect(names(latest.entities)).toEqual(["contact", "cr1_thing"]);
		expect(latest.entitiesLoading).toBe(false);
		expect(latest.entitiesError).toBeNull();
		expect(names(metadataCache.getFilteredEntities([SOLUTION]) ?? [])).toEqual([
			"contact",
			"cr1_thing",
		]);
	});

	it("shares the preload already in flight instead of asking again", async () => {
		const { preload, queryData } = hostWithSlowPreload();
		const { getAllEntitiesForScopes } = await import("../../features/fetchxml/api/pptbClient");
		const started = getAllEntitiesForScopes(); // what useAccessMode does at launch
		await render(createElement(Probe, { ids: [SOLUTION] }));
		await act(async () => preload.resolve({ value: ALL }));
		await started;
		await flush();
		expect(names(latest.entities)).toEqual(["contact", "cr1_thing"]);
		expect(
			queryData.mock.calls.filter(([q]) => q.startsWith("EntityDefinitions")),
		).toHaveLength(1);
	});

	it("says so when the preload fails, and doesn't cache an empty list", async () => {
		vi.spyOn(console, "error").mockImplementation(() => {});
		const { preload } = hostWithSlowPreload();
		await render(createElement(Probe, { ids: [SOLUTION] }));
		await act(async () => preload.reject(new Error("Metadata request failed")));
		await flush();
		expect(latest.entitiesError).toBe("Metadata request failed");
		expect(latest.entitiesLoading).toBe(false);
		expect(metadataCache.getFilteredEntities([SOLUTION])).toBeUndefined();
	});

	it("doesn't cache a list from before a connection change (#66)", async () => {
		const { preload } = hostWithSlowPreload();
		await render(createElement(Probe, { ids: [SOLUTION] }));
		metadataCache.clear(); // the connection changed while the entities were loading
		await act(async () => preload.resolve({ value: ALL }));
		await flush();
		expect(metadataCache.getFilteredEntities([SOLUTION])).toBeUndefined();
	});

	it("uses the cached list for solutions it already resolved", async () => {
		const { queryData } = hostWithSlowPreload();
		metadataCache.setFilteredEntities([SOLUTION], [entity("contact")]);
		await render(createElement(Probe, { ids: [SOLUTION] }));
		expect(names(latest.entities)).toEqual(["contact"]);
		expect(queryData).not.toHaveBeenCalled();
	});

	it("has no entities without a selected solution", async () => {
		hostWithSlowPreload();
		await render(createElement(Probe, { ids: [] }));
		expect(latest.entities).toEqual([]);
		expect(latest.entitiesLoading).toBe(false);
	});
});

describe("useSolutionEntities when the selection changes while it loads (#358 review)", () => {
	const S1 = "11111111-0000-0000-0000-000000000001";
	const S2 = "22222222-0000-0000-0000-000000000002";
	let latest: ReturnType<typeof useSolutionEntities>;
	function Probe({ ids }: { ids: string[] }) {
		latest = useSolutionEntities(ids);
		return null;
	}
	const show = (ids: string[]) => render(createElement(Probe, { ids }));

	/** Every entity is preloaded; each solution's components wait until the test releases them */
	function hostWithSlowComponents(options: { preload?: Promise<{ value: unknown[] }> } = {}) {
		const components = new Map<string, ReturnType<typeof deferred<{ value: unknown[] }>>>();
		const forSolution = (id: string) => {
			if (!components.has(id)) components.set(id, deferred());
			return components.get(id)!;
		};
		const queryData = vi.fn(async (query: string) => {
			if (query.startsWith("EntityDefinitions")) return options.preload ?? { value: ALL };
			const id = /msdyn_solutionid eq '([^']+)'/.exec(query)?.[1];
			return id ? forSolution(id).promise : { value: [] };
		});
		installHost({ dataverseAPI: { queryData } as never });
		return { forSolution };
	}
	const release = async (
		host: ReturnType<typeof hostWithSlowComponents>,
		id: string,
		entityNames: string[],
	) => {
		await act(async () =>
			host
				.forSolution(id)
				.resolve({ value: entityNames.map((msdyn_name) => ({ msdyn_name })) }),
		);
		await flush();
	};

	it("isn't left loading when the selection is cleared mid-way", async () => {
		hostWithSlowComponents();
		await show([S1]);
		expect(latest.entitiesLoading).toBe(true);
		await show([]);
		expect(latest.entitiesLoading).toBe(false);
		expect(latest.entities).toEqual([]);
	});

	it("isn't left loading when it switches back to a selection an earlier request cached", async () => {
		const host = hostWithSlowComponents();
		await show([S1]);
		await show([S2]);
		await release(host, S1, ["contact"]); // S1's earlier request finishes and caches
		await show([S1]);
		expect(names(latest.entities)).toEqual(["contact"]);
		expect(latest.entitiesLoading).toBe(false);
	});

	it("isn't left loading when it switches to a cached selection while another loads", async () => {
		hostWithSlowComponents();
		metadataCache.setFilteredEntities([S2], [entity("account")]);
		await show([S1]);
		expect(latest.entitiesLoading).toBe(true);
		await show([S2]);
		expect(names(latest.entities)).toEqual(["account"]);
		expect(latest.entitiesLoading).toBe(false);
	});

	it("drops the earlier selection's error when it switches to a cached one", async () => {
		vi.spyOn(console, "error").mockImplementation(() => {});
		const host = hostWithSlowComponents();
		metadataCache.setFilteredEntities([S2], [entity("account")]);
		await show([S1]);
		await act(async () => host.forSolution(S1).reject(new Error("boom")));
		await flush();
		expect(latest.entitiesError).toBe("boom");
		await show([S2]);
		expect(latest.entitiesError).toBeNull();
		expect(names(latest.entities)).toEqual(["account"]);
	});

	it("keeps the newer selection's entities when an older request finishes later", async () => {
		const host = hostWithSlowComponents();
		await show([S1]);
		await show([S2]);
		await release(host, S2, ["account"]);
		await release(host, S1, ["contact"]);
		expect(names(latest.entities)).toEqual(["account"]);
	});

	it("doesn't wait for the entity list when the solution has no entities", async () => {
		const host = hostWithSlowComponents({ preload: new Promise(() => {}) });
		await show([S1]);
		await release(host, S1, []);
		expect(latest.entitiesLoading).toBe(false);
		expect(latest.entities).toEqual([]);
		expect(latest.entitiesError).toBeNull();
	});
});

describe("the scope hooks get their entities through it (#358)", () => {
	it("useSolutionFilter", async () => {
		const { preload } = hostWithSlowPreload();
		let latest!: ReturnType<typeof useSolutionFilter>;
		function Probe() {
			latest = useSolutionFilter();
			return null;
		}
		await render(createElement(Probe));
		await act(async () => latest.updateSelectedSolutions([SOLUTION]));
		await flush();
		expect(latest.entitiesLoading).toBe(true);
		await act(async () => preload.resolve({ value: ALL }));
		await flush();
		expect(names(latest.entities)).toEqual(["contact", "cr1_thing"]);
	});

	it("usePublisherFilter", async () => {
		const { preload } = hostWithSlowPreload();
		let latest!: ReturnType<typeof usePublisherFilter>;
		function Probe() {
			latest = usePublisherFilter();
			return null;
		}
		await render(createElement(Probe));
		await act(async () => latest.updateSelectedSolutions([SOLUTION]));
		await flush();
		expect(latest.entitiesLoading).toBe(true);
		await act(async () => preload.resolve({ value: ALL }));
		await flush();
		expect(names(latest.entities)).toEqual(["contact", "cr1_thing"]);
	});
});
