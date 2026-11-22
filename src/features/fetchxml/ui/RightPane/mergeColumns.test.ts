import { describe, it, expect } from "vitest";
import { mergeColumns } from "./mergeColumns";

const col = (name: string, linkEntityAlias?: string) => ({ name, width: 100, linkEntityAlias });

describe("mergeColumns", () => {
	it("keeps an unapplied reorder when a column is added, appending the new one", () => {
		const local = [col("c"), col("a"), col("b")];
		const incoming = [col("a"), col("b"), col("c"), col("d")];
		expect(mergeColumns(local, incoming).map((c) => c.name)).toEqual(["c", "a", "b", "d"]);
	});

	it("drops removed columns and keeps the local order of the rest", () => {
		const local = [col("c"), col("a"), col("b")];
		expect(mergeColumns(local, [col("a"), col("c")]).map((c) => c.name)).toEqual(["c", "a"]);
	});

	it("distinguishes columns by link-entity alias", () => {
		const merged = mergeColumns([col("name", "x")], [col("name"), col("name", "x")]);
		expect(merged.map((c) => c.linkEntityAlias)).toEqual(["x", undefined]);
	});
});
