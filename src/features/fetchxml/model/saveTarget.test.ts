import { describe, expect, it } from "vitest";
import { SAVE_AS_LABEL, saveOptions, type SaveAction } from "./saveTarget";

const ME = "11111111-1111-1111-1111-111111111111";
const COLLEAGUE = "22222222-2222-2222-2222-222222222222";
const CUSTOMIZER = { canWrite: true, canPublish: true };

const saveAs: SaveAction = { kind: "save-as", label: "Save as new view…", needsConfirm: false, publishes: false };

const myView = { type: "personal" as const, name: "My open opportunities", ownerId: ME };
const sharedView = { type: "personal" as const, name: "Team pipeline", ownerId: COLLEAGUE };
const activeAccounts = { type: "system" as const, name: "Active Accounts", isCustomizable: true };

describe("saveOptions (#278, Q13)", () => {
	it("nothing loaded: Save as new view… only", () => {
		expect(saveOptions({ view: null, systemViewPrivileges: CUSTOMIZER })).toEqual({ primary: saveAs, menu: [] });
	});

	it("uses U+2026 for the ellipsis", () => {
		expect(SAVE_AS_LABEL.endsWith("…")).toBe(true);
		expect(SAVE_AS_LABEL).not.toContain("...");
	});

	describe("personal views", () => {
		it("the user's own view: Save changes to it, Save as new view… in the menu", () => {
			expect(saveOptions({ view: myView, currentUserId: ME })).toEqual({
				primary: { kind: "overwrite-personal", label: 'Save changes to "My open opportunities"', needsConfirm: false, publishes: false },
				menu: [saveAs],
			});
		});

		it("owner IDs compare case-insensitively", () => {
			const result = saveOptions({ view: { ...myView, ownerId: ME.toUpperCase() }, currentUserId: ME, personalWriteAccess: false });
			expect(result.primary.kind).toBe("overwrite-personal");
		});

		it("shared with Write: overwrite, and it asks first", () => {
			expect(saveOptions({ view: sharedView, currentUserId: ME, personalWriteAccess: true })).toEqual({
				primary: { kind: "overwrite-personal", label: 'Save changes to "Team pipeline"', needsConfirm: true, publishes: false },
				menu: [saveAs],
			});
		});

		it("shared read-only (#142): Save as new view… only", () => {
			expect(saveOptions({ view: sharedView, currentUserId: ME, personalWriteAccess: false })).toEqual({ primary: saveAs, menu: [] });
		});

		it("shared, access unknown: overwrite is offered and Dataverse decides", () => {
			expect(saveOptions({ view: sharedView, currentUserId: ME }).primary.kind).toBe("overwrite-personal");
		});

		it("owner unknown: treated as the user's own view", () => {
			expect(saveOptions({ view: { type: "personal", name: "Mine" } }).primary).toEqual({
				kind: "overwrite-personal",
				label: 'Save changes to "Mine"',
				needsConfirm: false,
				publishes: false,
			});
		});

		it("system view privileges don't matter", () => {
			const result = saveOptions({ view: sharedView, currentUserId: ME, personalWriteAccess: false, systemViewPrivileges: CUSTOMIZER });
			expect(result).toEqual({ primary: saveAs, menu: [] });
		});
	});

	describe("system views", () => {
		it("a customizer: Save as new view… primary, overwrite-and-publish in the menu", () => {
			expect(saveOptions({ view: activeAccounts, systemViewPrivileges: CUSTOMIZER })).toEqual({
				primary: saveAs,
				menu: [{ kind: "overwrite-system", label: 'Save changes to system view "Active Accounts"…', needsConfirm: true, publishes: true }],
			});
		});

		it("a managed but customizable view is still offered to a customizer", () => {
			const managedView = { ...activeAccounts, isManaged: true };
			const result = saveOptions({ view: managedView, systemViewPrivileges: CUSTOMIZER });
			expect(result.menu.map((action) => action.kind)).toEqual(["overwrite-system"]);
		});

		it("IsCustomizable unknown: offered (Dataverse decides)", () => {
			const result = saveOptions({ view: { type: "system", name: "Active Accounts" }, systemViewPrivileges: CUSTOMIZER });
			expect(result.menu.map((action) => action.kind)).toEqual(["overwrite-system"]);
		});

		it("not customizable (IsCustomizable false): Save as new view… only, even for a customizer", () => {
			expect(saveOptions({ view: { ...activeAccounts, isCustomizable: false }, systemViewPrivileges: CUSTOMIZER })).toEqual({
				primary: saveAs,
				menu: [],
			});
		});

		it.each([
			["no privileges", { canWrite: false, canPublish: false }],
			["write without publish", { canWrite: true, canPublish: false }],
			["privileges not known yet", null],
			["privileges not passed", undefined],
		])("%s: Save as new view… only", (_label, privileges) => {
			expect(saveOptions({ view: activeAccounts, systemViewPrivileges: privileges })).toEqual({ primary: saveAs, menu: [] });
		});

		it("never makes the system view overwrite the primary action", () => {
			expect(saveOptions({ view: activeAccounts, systemViewPrivileges: CUSTOMIZER }).primary.kind).toBe("save-as");
		});
	});
});
