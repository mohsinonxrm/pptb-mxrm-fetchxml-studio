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
