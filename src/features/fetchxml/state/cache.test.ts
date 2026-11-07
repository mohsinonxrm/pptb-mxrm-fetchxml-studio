import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { metadataCache } from "./cache";
import type { EntityMetadata, PublisherWithSolutions } from "../api/pptbClient";

// #66: everything in the metadata cache belongs to one connection

const entity = (name: string) => ({ LogicalName: name, EntitySetName: `${name}s` }) as EntityMetadata;

beforeEach(() => {
	vi.spyOn(console, "log").mockImplementation(() => {});
});

afterEach(() => {
	metadataCache.clear();
	vi.restoreAllMocks();
});

describe("metadataCache.clear", () => {
	it("clears every cached list, including the scope lists and solution components", () => {
		metadataCache.setEntityMetadata("account", entity("account"));
		metadataCache.setAllEntities([entity("account")]);
		metadataCache.setAllEntityMetadata([entity("account")]);
		metadataCache.setPublishers([]);
		metadataCache.setPublishersWithSolutions([{} as PublisherWithSolutions]);
		metadataCache.setSolutions(["p"], []);
		metadataCache.setFilteredEntities(["s"], [entity("account")]);
		metadataCache.setSolutionComponents("s", ["account"]);

		metadataCache.clear();

		expect(metadataCache.getEntityMetadata("account")).toBeUndefined();
		expect(metadataCache.getAllEntities()).toBeUndefined();
		expect(metadataCache.getAllEntityMetadata()).toBeUndefined();
		expect(metadataCache.getPublishers()).toBeUndefined();
		expect(metadataCache.getPublishersWithSolutions()).toBeUndefined();
		expect(metadataCache.getSolutions(["p"])).toBeUndefined();
		expect(metadataCache.getFilteredEntities(["s"])).toBeUndefined();
		expect(metadataCache.getSolutionComponents("s")).toBeUndefined();
	});

	it("drops in-flight request promises", () => {
		metadataCache.setAllEntityMetadataPromise(new Promise(() => {}));
		metadataCache.setAttributesPromise("account__all", new Promise(() => {}));
		metadataCache.clear();
		expect(metadataCache.getAllEntityMetadataPromise()).toBeUndefined();
		expect(metadataCache.getAttributesPromise("account__all")).toBeUndefined();
	});
});

describe("metadataCache.generationGuard", () => {
	it("tells a request started before a clear that its answer is stale", () => {
		const isCurrent = metadataCache.generationGuard();
		expect(isCurrent()).toBe(true);
		metadataCache.clear();
		expect(isCurrent()).toBe(false);
		expect(metadataCache.generationGuard()()).toBe(true);
	});
});
