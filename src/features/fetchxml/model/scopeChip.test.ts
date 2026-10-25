import { describe, expect, it } from "vitest";
import { scopeChipText, solutionDetail, solutionMatchesSearch, solutionTagLabel } from "./scopeChip";

const contoso = { friendlyname: "Contoso" };
const fabrikam = { friendlyname: "Fabrikam" };
const core = { friendlyname: "Sales Core", uniquename: "cso_salescore", version: "1.2.0.0" };
const patch = { friendlyname: "Sales Core", uniquename: "cso_salescore_Patch_7f3e", version: "1.2.1.0" };

describe("scope chip text (#227)", () => {
	it("labels a solution by name and the first three version parts, so a patch differs from its parent", () => {
		expect(solutionTagLabel(core)).toBe("Sales Core v1.2.0");
		expect(solutionTagLabel(patch)).toBe("Sales Core v1.2.1");
		expect(solutionTagLabel({ friendlyname: "No version" })).toBe("No version");
		expect(solutionDetail(core)).toBe("cso_salescore · v1.2.0.0");
	});

	it("says All entities in that mode", () => {
		expect(scopeChipText("all", [], [])).toMatchObject({ configured: true, label: "All entities", ariaLabel: "Entity scope: All entities" });
	});

	it("says Choose scope until a solution is picked", () => {
		expect(scopeChipText("publisher-solution", [], [])).toMatchObject({ configured: false, label: "Choose scope", ariaLabel: "Choose scope" });
		expect(scopeChipText("publisher-solution", [contoso], [])).toMatchObject({ configured: false, label: "Choose scope" });
		expect(scopeChipText("solution-only", [], [])).toMatchObject({ configured: false, label: "Choose scope" });
	});

	it("shows publisher › solution, with +N for more than one", () => {
		expect(scopeChipText("publisher-solution", [contoso], [core]).label).toBe("Contoso › Sales Core v1.2.0");
		const both = scopeChipText("publisher-solution", [contoso, fabrikam], [core, patch]);
		expect(both.label).toBe("Contoso +1 › Sales Core v1.2.0 +1");
		expect(both.ariaLabel).toBe("Entity scope: Contoso +1 › Sales Core v1.2.0 +1");
		expect(both.detail).toBe(
			"Publishers: Contoso, Fabrikam\nSolutions: Sales Core — cso_salescore · v1.2.0.0; Sales Core — cso_salescore_Patch_7f3e · v1.2.1.0",
		);
	});

	it("shows only solutions in Solution only mode", () => {
		const text = scopeChipText("solution-only", [contoso], [core, patch]);
		expect(text.label).toBe("Sales Core v1.2.0 +1");
		expect(text.detail.startsWith("Solutions:")).toBe(true);
	});

	it("searches solutions by name, unique name and version", () => {
		expect(solutionMatchesSearch(core, "sales")).toBe(true);
		expect(solutionMatchesSearch(core, "CSO_SALES")).toBe(true);
		expect(solutionMatchesSearch(patch, "1.2.1")).toBe(true);
		expect(solutionMatchesSearch(core, "1.2.1")).toBe(false);
		expect(solutionMatchesSearch(core, "  ")).toBe(true);
	});
});
