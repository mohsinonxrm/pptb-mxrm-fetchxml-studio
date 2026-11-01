import { afterEach, describe, expect, it, vi } from "vitest";
import { searchCompetitors } from "./pptbClient";

// #200: contains() is a LIKE, so the competitor picker's term matches %, _ and [ literally

describe("searchCompetitors", () => {
	afterEach(() => {
		delete (window as { dataverseAPI?: unknown }).dataverseAPI;
	});

	it("escapes LIKE wildcards in the term, as record search does", async () => {
		const queryData = vi.fn(async (_query: string) => ({ value: [] }));
		(window as { dataverseAPI?: unknown }).dataverseAPI = { queryData };
		await searchCompetitors("50%_[x]");
		const query = decodeURIComponent(String(queryData.mock.calls[0][0]));
		expect(query).toContain("contains(name,'50[%][_][[]x]')");
	});
});
