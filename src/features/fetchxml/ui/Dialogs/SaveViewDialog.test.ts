/**
 * Save view (#189, rules from #69 and #142): overwriting names the environment and the view and
 * says there's no undo; a personal view shared read-only can't be overwritten and offers
 * "Save a copy…", which switches to a new personal view named after it.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	button,
	cleanup,
	click,
	dialogText,
	deferred,
	expectDisabledWithReason,
	typeInto,
	flush,
	installHost,
	isDisabled,
	removeHost,
	render,
	toasts,
} from "../../../../shared/testing/ui";

const ME = "11111111-1111-1111-1111-111111111111";
const OTHER = "22222222-2222-2222-2222-222222222222";
const VIEW = "33333333-3333-3333-3333-333333333333";

const writeAccess = vi.fn(async (_id: string) => false as boolean | undefined);
const userQueryUpdate = vi.fn(async (_id: string, _data: unknown) => undefined as void);

vi.mock("../../api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/pptbClient")>()),
	whoAmI: vi.fn(async () => ({ UserId: ME, BusinessUnitId: "bu", OrganizationId: "org" })),
	checkSavedQueryPrivileges: vi.fn(async () => ({ canWrite: true, canPublish: true })),
	checkUserQueryPrivileges: vi.fn(async () => ({ canWrite: true })),
	checkUserQueryWriteAccess: (id: string) => writeAccess(id),
	validateFetchXmlExpression: vi.fn(async () => ({ ValidationResults: { Messages: [] } })),
	getUnmanagedSolutions: vi.fn(async () => []),
	getDefaultPublicView: vi.fn(async () => null),
	getSolutionLabel: vi.fn(async () => null),
	validateSavedQuery: vi.fn(async () => undefined),
	updateSavedQuery: vi.fn(async () => undefined),
	updateUserQuery: (id: string, data: unknown) => userQueryUpdate(id, data),
	clearOtherDefaultPublicViews: vi.fn(async () => [
		{ id: "44444444-4444-4444-4444-444444444444", name: "Old default", error: "Access denied" },
	]),
}));

const { SaveViewDialog } = await import("./SaveViewDialog");
const { checkUserQueryPrivileges } = await import("../../api/pptbClient");

const FETCH = "<fetch><entity name='contact'><attribute name='fullname'/></entity></fetch>";
const LAYOUT =
	"<grid name='resultset' object='2' jump='fullname' select='1' icon='1' preview='1'><row name='result' id='contactid'><cell name='fullname' width='300'/></row></grid>";

const props = (loadedView: Parameters<typeof SaveViewDialog>[0]["loadedView"]) => ({
	open: true,
	onClose: () => {},
	onSaveComplete: () => {},
	mode: "save" as const,
	shouldPublish: false,
	fetchXml: FETCH,
	layoutXml: LAYOUT,
	entityLogicalName: "contact",
	entityDisplayName: "Contact",
	objectTypeCode: 2,
	primaryIdAttribute: "contactid",
	loadedView,
});

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
	writeAccess.mockReset();
	userQueryUpdate.mockReset();
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

describe("Save view: a personal view shared read-only (#142)", () => {
	it("can't be overwritten, says who shared it, and offers Save a copy…", async () => {
		installHost();
		writeAccess.mockResolvedValue(false);
		await render(
			createElement(
				SaveViewDialog,
				props({
					id: VIEW,
					type: "personal",
					name: "Amy's contacts",
					ownerId: OTHER,
					ownerName: "Amy Alberts",
				}),
			),
		);
		await flush();
		const text = dialogText();
		expect(text).toContain("Can't overwrite this view");
		expect(text).toContain(
			'Amy Alberts shared "Amy\'s contacts" with you without write access.',
		);
		expectDisabledWithReason(
			button("Save"),
			"This view can't be overwritten. Save a copy instead",
		);
		expect(writeAccess).toHaveBeenCalledWith(VIEW);

		await click(button("Save a copy…"));
		const copy = dialogText();
		expect(copy).not.toContain("Can't overwrite this view");
		expect(copy).toContain("A new personal view");
		expect(
			document.querySelector<HTMLInputElement>('[role="dialog"] input[type="text"]')?.value,
		).toBe("Amy's contacts (copy)");
		expect(isDisabled(button("Save"))).toBe(false);
	});

	it("a shared view with write access can be overwritten", async () => {
		installHost();
		writeAccess.mockResolvedValue(true);
		await render(
			createElement(
				SaveViewDialog,
				props({
					id: VIEW,
					type: "personal",
					name: "Team contacts",
					ownerId: OTHER,
					ownerName: "Amy Alberts",
				}),
			),
		);
		await flush();
		expect(dialogText()).not.toContain("Can't overwrite this view");
		expect(isDisabled(button("Save"))).toBe(false);
	});
});

describe("Save view: overwriting your own view (#69)", () => {
	it("names the environment and the view, and says there's no undo", async () => {
		installHost({ connection: { name: "CRM279420" } });
		await render(
			createElement(
				SaveViewDialog,
				props({ id: VIEW, type: "personal", name: "My contacts", ownerId: ME }),
			),
		);
		await flush();
		const text = dialogText();
		expect(text).toContain("CRM279420");
		expect(text).toContain('The personal view "My contacts" (overwritten)');
		expect(text).toContain("There's no undo: the view's current definition is replaced.");
		expect(isDisabled(button("Save"))).toBe(false);
	});
});

describe("Save view: a default view that couldn't be cleared (#307, #277)", () => {
	it("is a warning toast that stays, not a host notification; the view is reported unpublished", async () => {
		const host = installHost({ connection: { name: "CRM279420" } });
		const onSaveComplete = vi.fn();
		await render(
			createElement(SaveViewDialog, {
				...props({
					id: VIEW,
					type: "system",
					name: "Active contacts",
					queryType: 0,
					isCustomizable: true,
					isManaged: false,
					isDefault: true,
				}),
				onSaveComplete,
			}),
		);
		await flush();
		await click(button(/^Save/));
		await flush(10);
		expect(onSaveComplete).toHaveBeenCalledWith(
			VIEW,
			"system",
			"Active contacts",
			expect.objectContaining({ published: false }),
		);
		expect(toasts()).toEqual([
			expect.stringMatching(/^Another view is still a default view\. Dismiss.*Old default/),
		]);
		expect(host.notifications).toEqual([]);
	});
});

describe("Save view: the buttons say why, and keep focus while saving (#279)", () => {
	it("Save asks for a name", async () => {
		installHost();
		await render(
			createElement(
				SaveViewDialog,
				props({ id: VIEW, type: "personal", name: "My contacts", ownerId: ME }),
			),
		);
		await flush();
		await typeInto(
			document.querySelector<HTMLInputElement>('[role="dialog"] input[type="text"]'),
			"  ",
		);
		expectDisabledWithReason(button("Save"), "Enter a name");
	});

	it("while saving: Save reads Saving… and keeps focus; Cancel says to wait", async () => {
		installHost();
		const saving = deferred<void>();
		userQueryUpdate.mockImplementation(() => saving.promise);
		await render(
			createElement(
				SaveViewDialog,
				props({ id: VIEW, type: "personal", name: "My contacts", ownerId: ME }),
			),
		);
		await flush();
		const save = button("Save");
		save?.focus();
		await click(save);
		await flush();
		const busy = button("Saving…");
		expect(busy).toBe(save);
		expect(document.activeElement).toBe(busy);
		expectDisabledWithReason(busy, "Wait for the save to finish");
		expectDisabledWithReason(button("Cancel"), "Wait for the save to finish");
		expectDisabledWithReason(
			document.querySelector<HTMLElement>('[role="dialog"] button[aria-label="Close"]'),
			"Wait for the save to finish",
		);
		saving.resolve();
		await flush();
	});
});

describe("Save view while your permissions are checked (#279 review)", () => {
	it("Save says it's checking, not that you lack the permission", async () => {
		installHost();
		const check = deferred<{ canWrite: boolean }>();
		vi.mocked(checkUserQueryPrivileges).mockImplementationOnce(() => check.promise);
		await render(
			createElement(SaveViewDialog, {
				...props(null),
				mode: "saveAs" as const,
			}),
		);
		await flush();
		await typeInto(
			document.querySelector<HTMLInputElement>('[role="dialog"] input[type="text"]'),
			"FXS TEST view",
		);
		expectDisabledWithReason(button("Save"), "Checking your permissions…");
		expect(dialogText()).not.toContain("You don't have permission");
		check.resolve({ canWrite: true });
		await flush();
		expect(isDisabled(button("Save"))).toBe(false);
	});
});
