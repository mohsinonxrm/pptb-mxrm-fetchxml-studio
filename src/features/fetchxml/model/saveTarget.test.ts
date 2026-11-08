import { describe, expect, it } from "vitest";
import {
	PRIVILEGES_UNKNOWN_WARNING,
	SAVE_AS_LABEL,
	VIEW_ACCESS_UNKNOWN_WARNING,
	saveOptions,
	type SaveAction,
	type SaveOptionsInput,
} from "./saveTarget";

const ME = "11111111-1111-1111-1111-111111111111";
const COLLEAGUE = "22222222-2222-2222-2222-222222222222";

const saveAs: SaveAction = {
	kind: "save-as",
	label: "Save as new view…",
	needsConfirm: false,
	publishes: false,
};
const saveAsOnly = { primary: saveAs, menu: [] };

const myView = { type: "personal" as const, name: "My open opportunities", ownerId: ME };
const sharedView = { type: "personal" as const, name: "Team pipeline", ownerId: COLLEAGUE };
const activeAccounts = { type: "system" as const, name: "Active Accounts", isCustomizable: true };

/** A personal view's input with prvWriteUserQuery granted, overridable */
const personal = (overrides: Partial<SaveOptionsInput>): SaveOptionsInput => ({
	view: myView,
	currentUserId: ME,
	canWriteUserQueries: true,
	...overrides,
});
/** A system view's input with write granted, overridable */
const system = (overrides: Partial<SaveOptionsInput>): SaveOptionsInput => ({
	view: activeAccounts,
	canWriteSystemViews: true,
	...overrides,
});

describe("saveOptions (#278, Q13)", () => {
	it("nothing loaded: Save as new view… only", () => {
		expect(
			saveOptions({ view: null, canWriteSystemViews: true, canWriteUserQueries: true }),
		).toEqual(saveAsOnly);
	});

	it("uses U+2026 for the ellipsis", () => {
		expect(SAVE_AS_LABEL.endsWith("…")).toBe(true);
		expect(SAVE_AS_LABEL).not.toContain("...");
	});

	describe("personal views", () => {
		it("the user's own view: Save changes to it, Save as new view… in the menu", () => {
			expect(saveOptions(personal({}))).toEqual({
				primary: {
					kind: "overwrite-personal",
					label: 'Save changes to "My open opportunities"',
					needsConfirm: false,
					publishes: false,
				},
				menu: [saveAs],
			});
		});

		it("owner IDs compare case-insensitively", () => {
			const result = saveOptions(
				personal({ view: { ...myView, ownerId: ME.toUpperCase() } }),
			);
			expect(result.primary).toMatchObject({
				kind: "overwrite-personal",
				needsConfirm: false,
			});
		});

		it("shared with Write: overwrite, and it asks first", () => {
			expect(saveOptions(personal({ view: sharedView, personalWriteAccess: true }))).toEqual({
				primary: {
					kind: "overwrite-personal",
					label: 'Save changes to "Team pipeline"',
					needsConfirm: true,
					publishes: false,
				},
				menu: [saveAs],
			});
		});

		it("shared read-only (#142): Save as new view… only", () => {
			expect(saveOptions(personal({ view: sharedView, personalWriteAccess: false }))).toEqual(
				saveAsOnly,
			);
		});

		it("shared, write access still loading: Save as new view… only", () => {
			expect(saveOptions(personal({ view: sharedView }))).toEqual(saveAsOnly);
		});

		it("shared, write access check failed: offered with a warning, and it asks first", () => {
			expect(
				saveOptions(personal({ view: sharedView, personalWriteAccess: "unknown" })).primary,
			).toEqual({
				kind: "overwrite-personal",
				label: 'Save changes to "Team pipeline"',
				needsConfirm: true,
				publishes: false,
				warning: VIEW_ACCESS_UNKNOWN_WARNING,
			});
		});

		it("read-only is blocked even while WhoAmI hasn't answered", () => {
			expect(
				saveOptions(
					personal({
						view: sharedView,
						currentUserId: undefined,
						personalWriteAccess: false,
					}),
				),
			).toEqual(saveAsOnly);
		});

		it("read-only is blocked even when the owner isn't known", () => {
			expect(
				saveOptions(
					personal({
						view: { type: "personal", name: "Team pipeline" },
						personalWriteAccess: false,
					}),
				),
			).toEqual(saveAsOnly);
		});

		it("owner unknown (WhoAmI pending), write not known false: offered, and it asks first", () => {
			expect(saveOptions(personal({ currentUserId: undefined })).primary).toEqual({
				kind: "overwrite-personal",
				label: 'Save changes to "My open opportunities"',
				needsConfirm: true,
				publishes: false,
			});
		});

		it("owner unknown (no ownerid), write check failed: offered with a warning, and it asks first", () => {
			expect(
				saveOptions(
					personal({
						view: { type: "personal", name: "Mine" },
						personalWriteAccess: "unknown",
					}),
				).primary,
			).toMatchObject({
				kind: "overwrite-personal",
				needsConfirm: true,
				warning: VIEW_ACCESS_UNKNOWN_WARNING,
			});
		});

		it("own view: a failed view access check doesn't matter", () => {
			expect(
				saveOptions(personal({ personalWriteAccess: "unknown" })).primary,
			).not.toHaveProperty("warning");
		});

		it("without prvWriteUserQuery: Save as new view… only, even for the user's own view", () => {
			expect(saveOptions(personal({ canWriteUserQueries: false }))).toEqual(saveAsOnly);
		});

		it("prvWriteUserQuery still loading: Save as new view… only", () => {
			expect(saveOptions(personal({ canWriteUserQueries: undefined }))).toEqual(saveAsOnly);
		});

		it("prvWriteUserQuery check failed: offered with the privileges warning", () => {
			expect(saveOptions(personal({ canWriteUserQueries: "unknown" })).primary).toEqual({
				kind: "overwrite-personal",
				label: 'Save changes to "My open opportunities"',
				needsConfirm: false,
				publishes: false,
				warning: PRIVILEGES_UNKNOWN_WARNING,
			});
		});

		it("system view privileges don't matter", () => {
			expect(
				saveOptions(
					personal({
						view: sharedView,
						personalWriteAccess: false,
						canWriteSystemViews: true,
					}),
				),
			).toEqual(saveAsOnly);
		});
	});

	describe("system views", () => {
		it("can write system views: Save as new view… primary, overwrite-and-publish in the menu", () => {
			expect(saveOptions(system({}))).toEqual({
				primary: saveAs,
				menu: [
					{
						kind: "overwrite-system",
						label: 'Save changes to system view "Active Accounts"…',
						needsConfirm: true,
						publishes: true,
					},
				],
			});
		});

		it("the publish privilege isn't a gate (prvPublishCustomization is deprecated)", () => {
			// SystemViewPrivileges { canWrite: true, canPublish: false } maps to canWriteSystemViews: true
			expect(saveOptions(system({ canWriteSystemViews: true })).menu).toEqual([
				{
					kind: "overwrite-system",
					label: 'Save changes to system view "Active Accounts"…',
					needsConfirm: true,
					publishes: true,
				},
			]);
		});

		it("a managed but customizable view is still offered", () => {
			const managedView = { ...activeAccounts, isManaged: true };
			expect(
				saveOptions(system({ view: managedView })).menu.map((action) => action.kind),
			).toEqual(["overwrite-system"]);
		});

		it("IsCustomizable unknown: offered (Dataverse decides)", () => {
			expect(
				saveOptions(system({ view: { type: "system", name: "Active Accounts" } })).menu.map(
					(action) => action.kind,
				),
			).toEqual(["overwrite-system"]);
		});

		it("not customizable (IsCustomizable false): Save as new view… only, even with write", () => {
			expect(
				saveOptions(system({ view: { ...activeAccounts, isCustomizable: false } })),
			).toEqual(saveAsOnly);
		});

		it("can't write system views: Save as new view… only", () => {
			expect(saveOptions(system({ canWriteSystemViews: false }))).toEqual(saveAsOnly);
		});

		it("privileges still loading: Save as new view… only, so the item doesn't flash", () => {
			expect(saveOptions(system({ canWriteSystemViews: undefined }))).toEqual(saveAsOnly);
		});

		it("privilege check failed: offered with a warning", () => {
			expect(saveOptions(system({ canWriteSystemViews: "unknown" })).menu).toEqual([
				{
					kind: "overwrite-system",
					label: 'Save changes to system view "Active Accounts"…',
					needsConfirm: true,
					publishes: true,
					warning: PRIVILEGES_UNKNOWN_WARNING,
				},
			]);
		});

		it("privilege check failed on a non-customizable view: still Save as new view… only", () => {
			expect(
				saveOptions(
					system({
						view: { ...activeAccounts, isCustomizable: false },
						canWriteSystemViews: "unknown",
					}),
				),
			).toEqual(saveAsOnly);
		});

		it("personal view checks don't matter", () => {
			expect(
				saveOptions(system({ canWriteUserQueries: false, personalWriteAccess: false }))
					.menu,
			).toHaveLength(1);
		});
	});
});
