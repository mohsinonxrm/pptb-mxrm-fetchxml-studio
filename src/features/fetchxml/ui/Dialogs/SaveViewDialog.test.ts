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
	flush,
	installHost,
	isDisabled,
	removeHost,
	render,
} from "../../../../shared/testing/ui";

const ME = "11111111-1111-1111-1111-111111111111";
const OTHER = "22222222-2222-2222-2222-222222222222";
const VIEW = "33333333-3333-3333-3333-333333333333";

const writeAccess = vi.fn(async (_id: string) => false as boolean | undefined);

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
}));

const { SaveViewDialog } = await import("./SaveViewDialog");

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
		expect(isDisabled(button("Save"))).toBe(true);
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
