/**
 * Save follows the owner's rules (#278, Q3) on the app bar: the button says Save; for a personal
 * view you can write it saves changes to that view, otherwise it saves a new view; a system view's
 * overwrite (with publish) is in the menu for those who can write system views. A check still
 * running hides the overwrite; one that failed offers it with a warning. The checks are mocked.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, flush, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { OverflowToolbar } from "../../../../shared/components/commandBar";
import { useSaveViewCommand, type SaveViewCommandOptions } from "./useSaveViewCommand";
import type { SaveViewDialogLoadedView } from "../Dialogs/SaveViewDialog";

const checks = vi.hoisted(() => ({
	whoAmI: vi.fn(),
	checkUserQueryPrivileges: vi.fn(),
	checkSavedQueryPrivileges: vi.fn(),
	checkUserQueryWriteAccess: vi.fn(),
}));
vi.mock("../../api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/pptbClient")>()),
	...checks,
}));

const ME = "aaaaaaaa-0000-0000-0000-000000000001";
const SOMEONE = "bbbbbbbb-0000-0000-0000-000000000002";

const personal = (extra: Partial<SaveViewDialogLoadedView> = {}) =>
	({
		id: "p1",
		type: "personal",
		name: "Mine",
		ownerId: ME,
		...extra,
	}) as SaveViewDialogLoadedView;
const system = (extra: Partial<SaveViewDialogLoadedView> = {}) =>
	({ id: "s1", type: "system", name: "Sys", ...extra }) as SaveViewDialogLoadedView;

function SaveBar(props: SaveViewCommandOptions) {
	const save = useSaveViewCommand(props);
	return createElement(
		"div",
		null,
		createElement(OverflowToolbar, { ariaLabel: "Test", commands: [save.command] }),
		save.dialog,
	);
}

const renderSave = async (targetView: SaveViewDialogLoadedView | null) => {
	await render(
		createElement(SaveBar, {
			fetchXml: "<fetch><entity name='account'/></fetch>",
			layoutXml: "",
			entityLogicalName: "account",
			objectTypeCode: 1,
			primaryIdAttribute: "accountid",
			targetView,
		}),
	);
	await flush(10);
};

const saveButton = () =>
	[...document.querySelectorAll<HTMLElement>('[role="toolbar"] button')].find(
		(b) => b.textContent?.trim() === "Save",
	);
const hint = () => {
	const id = saveButton()?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};
async function menuLabels(): Promise<(string | undefined)[]> {
	const menuHalf = saveButton()
		?.closest(".fui-SplitButton")
		?.querySelector<HTMLElement>(".fui-SplitButton__menuButton");
	await click(menuHalf);
	// The label only: the primary action also shows Ctrl+S (#186), checked on its own
	return [...document.querySelectorAll<HTMLElement>('[role="menu"] [role="menuitem"]')].map((i) =>
		i.querySelector(".fui-MenuItem__content")?.textContent?.trim(),
	);
}
const dialogText = () => document.querySelector(".fui-DialogSurface")?.textContent ?? "";

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
	stubOverflowWidths({ container: 2000 });
	checks.whoAmI.mockResolvedValue({ UserId: ME });
	checks.checkUserQueryPrivileges.mockResolvedValue({ canWrite: true });
	checks.checkSavedQueryPrivileges.mockResolvedValue({ canWrite: true, canPublish: true });
	checks.checkUserQueryWriteAccess.mockResolvedValue(true);
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
	for (const fn of Object.values(checks)) fn.mockReset();
});

describe("Save's rules", () => {
	it("no view: Save saves a new view", async () => {
		await renderSave(null);
		expect(hint()).toBe("Save as new view… (Ctrl+S)");
		expect(await menuLabels()).toEqual(["Save as new view…"]);
	});

	it("your own personal view: Save saves the changes to it", async () => {
		await renderSave(personal());
		expect(hint()).toBe('Save changes to "Mine" (Ctrl+S)');
		// Ctrl+S is the primary action's key (#186): on the button and its menu item
		expect(saveButton()?.getAttribute("aria-keyshortcuts")).toBe("Control+S");
		expect(await menuLabels()).toEqual(['Save changes to "Mine"', "Save as new view…"]);
		expect(checks.checkUserQueryWriteAccess).not.toHaveBeenCalled();
	});

	it("Save on your own personal view opens the overwrite, not a new view", async () => {
		await renderSave(personal());
		await click(saveButton());
		await flush(5);
		expect(dialogText()).toContain("Save view");
		expect(dialogText()).not.toContain("Save as new view");
	});

	it("a personal view shared with you read-only: only a new view", async () => {
		checks.checkUserQueryWriteAccess.mockResolvedValue(false);
		await renderSave(personal({ ownerId: SOMEONE }));
		expect(hint()).toBe("Save as new view… (Ctrl+S)");
		expect(await menuLabels()).toEqual(["Save as new view…"]);
	});

	it("no prvWriteUserQuery: only a new view", async () => {
		checks.checkUserQueryPrivileges.mockResolvedValue({ canWrite: false });
		await renderSave(personal());
		expect(await menuLabels()).toEqual(["Save as new view…"]);
	});

	it("a system view you can write: a new view first, and the overwrite in the menu", async () => {
		await renderSave(system());
		expect(hint()).toBe("Save as new view… (Ctrl+S)");
		expect(await menuLabels()).toEqual([
			"Save as new view…",
			'Save changes to system view "Sys"…',
		]);
	});

	it("a system view you can't write, or that isn't customizable: only a new view", async () => {
		checks.checkSavedQueryPrivileges.mockResolvedValue({ canWrite: false, canPublish: false });
		await renderSave(system());
		expect(await menuLabels()).toEqual(["Save as new view…"]);
		cleanup();
		checks.checkSavedQueryPrivileges.mockResolvedValue({ canWrite: true, canPublish: true });
		await renderSave(system({ isCustomizable: false }));
		expect(await menuLabels()).toEqual(["Save as new view…"]);
	});

	it("a check that failed still offers the overwrite, with a warning", async () => {
		checks.checkSavedQueryPrivileges.mockResolvedValue({
			canWrite: false,
			canPublish: false,
			failed: true,
		});
		await renderSave(system());
		expect(await menuLabels()).toEqual([
			"Save as new view…",
			'Save changes to system view "Sys"…',
		]);
	});

	it("a personal view whose privilege check failed: the overwrite, and the dialog lets you save", async () => {
		checks.checkUserQueryPrivileges.mockResolvedValue({ canWrite: false, failed: true });
		await renderSave(personal());
		expect(hint()).toContain('Save changes to "Mine"');
		expect(hint()).toContain("Dataverse decides");
		await click(saveButton());
		await flush(5);
		const save = [
			...document.querySelectorAll<HTMLButtonElement>(".fui-DialogSurface button"),
		].find((b) => b.textContent?.trim() === "Save");
		expect(save?.disabled).toBe(false);
	});

	it("a shared personal view while its Write check runs: only a new view", async () => {
		checks.checkUserQueryWriteAccess.mockReturnValue(new Promise(() => {}));
		await renderSave(personal({ ownerId: SOMEONE }));
		expect(await menuLabels()).toEqual(["Save as new view…"]);
	});

	it("a check still running hides the overwrite", async () => {
		checks.checkUserQueryPrivileges.mockReturnValue(new Promise(() => {}));
		await renderSave(personal());
		expect(hint()).toBe("Save as new view… (Ctrl+S)");
		expect(await menuLabels()).toEqual(["Save as new view…"]);
	});
});
