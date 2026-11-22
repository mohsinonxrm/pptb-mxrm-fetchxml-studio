import { describe, expect, it } from "vitest";
import { selectionLabel } from "./selectionLabel";

describe("the selection bar's label (#234)", () => {
	it("names each kind of node, then what names it", () => {
		expect(selectionLabel(null)).toBe("Nothing selected");
		expect(selectionLabel({ type: "fetch" })).toBe("Query options");
		expect(selectionLabel({ type: "entity", name: "account" })).toBe("Entity · account");
		expect(selectionLabel({ type: "all-attributes" })).toBe("All attributes");
		expect(selectionLabel({ type: "attribute", name: "name" })).toBe("Attribute · name");
		expect(selectionLabel({ type: "attribute", name: "name", alias: "n" })).toBe(
			"Attribute · n",
		);
		expect(selectionLabel({ type: "order", attribute: "createdon" })).toBe("Sort · createdon");
		expect(selectionLabel({ type: "filter", conjunction: "or" })).toBe("Filter (or)");
		expect(selectionLabel({ type: "filter", conjunction: "and" })).toBe("Filter (and)");
		expect(selectionLabel({ type: "condition", attribute: "revenue" })).toBe(
			"Condition · revenue",
		);
		expect(selectionLabel({ type: "condition", attribute: "" })).toBe("Condition");
		expect(selectionLabel({ type: "link-entity", name: "contact", alias: "pc" })).toBe(
			"Link-entity · contact (pc)",
		);
	});
});
