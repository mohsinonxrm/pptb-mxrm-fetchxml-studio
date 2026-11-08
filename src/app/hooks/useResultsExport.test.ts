/**
 * useResultsExport (#301): the prvExportToExcel check on mount, the local export through the lazily
 * loaded exceljs module (#187), the server export of a loaded view, the exporting flag, the error
 * messages, and the error clearing after 10 seconds or on dismiss.
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, deferred, render } from "../../shared/testing/ui";
import { formatDataverseError } from "../../features/fetchxml/api/dataverseErrors";
import { defaultDisplaySettings } from "../../features/fetchxml/model/displaySettings";
import type { AttributeMetadata, EntityMetadata } from "../../features/fetchxml/api/pptbClient";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import type { LoadedViewState } from "../../features/fetchxml/state/builderStore";
import { useResultsExport, type ResultsExport, type ResultsExportInputs } from "./useResultsExport";

const api = vi.hoisted(() => ({
	whoAmI: vi.fn(),
	isDataverseAvailable: vi.fn(),
	exportToExcel: vi.fn(),
	downloadBase64File: vi.fn(),
	checkPrivilegeByName: vi.fn(),
	loadFormattingMetadata: vi.fn(),
	loadUserTimeZone: vi.fn(),
}));
vi.mock("../../features/fetchxml/api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<object>()),
	...api,
}));

// The chunk the local export imports on first use (#187)
const excel = vi.hoisted(() => ({ exportToExcelLocal: vi.fn(), downloadExcelFile: vi.fn() }));
vi.mock("../../features/fetchxml/api/excelExport", () => excel);

let latest: ResultsExport;

function Probe(props: ResultsExportInputs) {
	latest = useResultsExport(props);
	return null;
}

const fetchQuery = {
	id: "f",
	type: "fetch",
	options: {},
	entity: {
		id: "e",
		type: "entity",
		name: "account",
		attributes: [
			{ id: "a1", type: "attribute", name: "name" },
			{ id: "a2", type: "attribute", name: "primarycontactid" },
		],
		orders: [],
		filters: [],
		links: [],
	},
} as FetchNode;

const attribute = (label: string, type: string) =>
	({
		AttributeType: type,
		DisplayName: { UserLocalizedLabel: { Label: label } },
	}) as unknown as AttributeMetadata;

const attributeMetadata = new Map([
	[
		"account",
		new Map([
			["name", attribute("Account Name", "String")],
			["primarycontactid", attribute("Primary Contact", "Lookup")],
		]),
	],
]);

const entityMetadata = {
	LogicalName: "account",
	DisplayName: { UserLocalizedLabel: { Label: "Account" } },
} as unknown as EntityMetadata;

const view = { id: "v1", type: "system", name: "Active Accounts" } as LoadedViewState;

const rows = [
	{ accountid: "1", name: "Contoso", _primarycontactid_value: "c1" },
	{ accountid: "2", name: "Fabrikam", _primarycontactid_value: "c2" },
];

const inputs = (overrides: Partial<ResultsExportInputs> = {}): ResultsExportInputs => ({
	loadedView: null,
	fetchQuery,
	columnConfig: null,
	fetchXml: "<fetch/>",
	layoutXml: "<grid/>",
	queryResult: { rows, query: fetchQuery },
	attributeMetadata,
	entityMetadata,
	displaySettings: defaultDisplaySettings,
	...overrides,
});

const show = (overrides: Partial<ResultsExportInputs> = {}) =>
	render(createElement(Probe, inputs(overrides)));

/** Lets the async export (the lazy import included) move on until the condition holds */
async function until(condition: () => boolean) {
	for (let i = 0; i < 100 && !condition(); i++) {
		await act(async () => {
			await new Promise((resolve) => setTimeout(resolve, 0));
		});
	}
	expect(condition()).toBe(true);
}

/** Runs an export to the end */
async function run(action: () => Promise<void>) {
	await act(async () => {
		await action();
	});
}

beforeEach(() => {
	api.isDataverseAvailable.mockReturnValue(true);
	api.whoAmI.mockResolvedValue({ UserId: "u1" });
	api.checkPrivilegeByName.mockResolvedValue(true);
	api.exportToExcel.mockResolvedValue({ excelFile: "QUJD", filename: "Active Accounts.xlsx" });
	api.downloadBase64File.mockResolvedValue(undefined);
	api.loadFormattingMetadata.mockResolvedValue(new Map());
	api.loadUserTimeZone.mockResolvedValue(undefined);
	excel.exportToExcelLocal.mockResolvedValue({
		buffer: new ArrayBuffer(4),
		fileName: "Account.xlsx",
	});
	excel.downloadExcelFile.mockResolvedValue(undefined);
	vi.spyOn(console, "error").mockImplementation(() => undefined);
});

afterEach(() => {
	cleanup();
	vi.useRealTimers();
	vi.resetAllMocks();
	vi.restoreAllMocks();
});

describe("useResultsExport: the privilege", () => {
	it("checks prvExportToExcel for the current user on mount", async () => {
		await show({ loadedView: view });
		expect(api.checkPrivilegeByName).toHaveBeenCalledWith("u1", "prvExportToExcel");
		expect(latest.exportStatus).toEqual({
			isExporting: false,
			hasPrivilege: true,
			privilegeChecked: true,
		});
		expect(latest.canExport).toBe(true);
		expect(latest.exportDisabledReason).toBeUndefined();
	});

	it("without a loaded view, the server export is off and says to save a view", async () => {
		await show();
		expect(latest.canExport).toBe(false);
		expect(latest.exportDisabledReason).toBe("Save as a view first to enable export");
	});

	it("without the privilege, the server export is off and says so", async () => {
		api.checkPrivilegeByName.mockResolvedValue(false);
		await show({ loadedView: view });
		expect(latest.exportStatus.privilegeChecked).toBe(true);
		expect(latest.canExport).toBe(false);
		expect(latest.exportDisabledReason).toBe("You don't have the prvExportToExcel privilege");
	});

	it("a failed check counts as no privilege", async () => {
		api.whoAmI.mockRejectedValue(new Error("offline"));
		await show({ loadedView: view });
		expect(latest.exportStatus).toMatchObject({ hasPrivilege: false, privilegeChecked: true });
	});

	it("doesn't check without the Dataverse API", async () => {
		api.isDataverseAvailable.mockReturnValue(false);
		await show({ loadedView: view });
		expect(api.whoAmI).not.toHaveBeenCalled();
		expect(latest.exportStatus.privilegeChecked).toBe(false);
	});
});

describe("useResultsExport: the local export", () => {
	it("exports the grid's columns with their display names through the lazily loaded exceljs", async () => {
		await show();
		await run(latest.exportLocal);

		expect(api.loadFormattingMetadata).toHaveBeenCalledWith("account");
		expect(excel.exportToExcelLocal).toHaveBeenCalledTimes(1);
		const options = excel.exportToExcelLocal.mock.calls[0][0];
		expect(options).toMatchObject({
			records: rows,
			columns: ["name", "_primarycontactid_value"],
			attributeMetadata,
			entityName: "account",
			fileName: "Account",
			displaySettings: defaultDisplaySettings,
			timeZone: undefined,
		});
		expect([...options.columnDisplayNames]).toEqual([
			["name", "Account Name"],
			["_primarycontactid_value", "Primary Contact"],
		]);
		expect(options.formattingMetadata).toEqual(new Map([["account", new Map()]]));
		expect(excel.downloadExcelFile).toHaveBeenCalledWith(
			expect.any(ArrayBuffer),
			"Account.xlsx",
		);
		expect(api.exportToExcel).not.toHaveBeenCalled();
		expect(latest.exportStatus.isExporting).toBe(false);
		expect(latest.exportStatus.error).toBeUndefined();
	});

	it("names the file after the loaded view when there is one", async () => {
		await show({ loadedView: view });
		await run(latest.exportLocal);
		expect(excel.exportToExcelLocal.mock.calls[0][0].fileName).toBe("Active Accounts");
	});

	it("sets the exporting flag while it runs and clears it at the end", async () => {
		const download = deferred<void>();
		excel.downloadExcelFile.mockReturnValue(download.promise);
		await show();

		let pending: Promise<void> = Promise.resolve();
		await act(async () => {
			pending = latest.exportLocal();
		});
		await until(() => excel.downloadExcelFile.mock.calls.length === 1);
		expect(latest.exportStatus.isExporting).toBe(true);

		await act(async () => {
			download.resolve();
			await pending;
		});
		expect(latest.exportStatus.isExporting).toBe(false);
	});

	it("with no rows, says to execute a query first and exports nothing", async () => {
		await show({ queryResult: { rows: [] } });
		await run(latest.exportLocal);
		expect(latest.exportStatus.error).toBe("No data to export. Execute a query first.");
		expect(latest.exportStatus.isExporting).toBe(false);
		expect(excel.exportToExcelLocal).not.toHaveBeenCalled();
	});

	it("a failure shows the error's message and clears the exporting flag", async () => {
		excel.exportToExcelLocal.mockRejectedValue(new Error("Disk full"));
		await show();
		await run(latest.exportLocal);
		expect(latest.exportStatus.error).toBe("Disk full");
		expect(latest.exportStatus.isExporting).toBe(false);
	});

	it("a failure that isn't an Error shows the generic message", async () => {
		excel.downloadExcelFile.mockRejectedValue("nope");
		await show();
		await run(latest.exportLocal);
		expect(latest.exportStatus.error).toBe("Local export failed. Please try again.");
	});
});

describe("useResultsExport: the server export of a loaded view", () => {
	it("exports through ExportToExcel with the view, the FetchXML and the LayoutXML, then downloads", async () => {
		await show({ loadedView: view });
		await run(latest.exportToExcel);

		expect(api.exportToExcel).toHaveBeenCalledWith(
			"v1",
			"system",
			"<fetch/>",
			"<grid/>",
			"Active Accounts",
		);
		expect(api.downloadBase64File).toHaveBeenCalledWith("QUJD", "Active Accounts.xlsx");
		expect(excel.exportToExcelLocal).not.toHaveBeenCalled();
		expect(latest.exportStatus.isExporting).toBe(false);
		expect(latest.exportStatus.error).toBeUndefined();
	});

	it("sets the exporting flag while it runs and clears it at the end", async () => {
		const request = deferred<{ excelFile: string; filename: string }>();
		api.exportToExcel.mockReturnValue(request.promise);
		await show({ loadedView: view });

		let pending: Promise<void> = Promise.resolve();
		await act(async () => {
			pending = latest.exportToExcel();
		});
		expect(latest.exportStatus.isExporting).toBe(true);

		await act(async () => {
			request.resolve({ excelFile: "QUJD", filename: "Active Accounts.xlsx" });
			await pending;
		});
		expect(latest.exportStatus.isExporting).toBe(false);
	});

	it("without the privilege, says to ask an administrator and sends nothing", async () => {
		api.checkPrivilegeByName.mockResolvedValue(false);
		await show({ loadedView: view });
		await run(latest.exportToExcel);
		expect(latest.exportStatus.error).toBe(
			"You don't have permission to export to Excel. Contact your administrator to request the 'Export to Excel' privilege.",
		);
		expect(api.exportToExcel).not.toHaveBeenCalled();
	});

	it("without a loaded view, says to save the query as a view first", async () => {
		await show();
		await run(latest.exportToExcel);
		expect(latest.exportStatus.error).toBe(
			"Export to Excel requires a saved view. Please save your query as a view first.",
		);
		expect(api.exportToExcel).not.toHaveBeenCalled();
	});

	it("without LayoutXML, says both are needed", async () => {
		await show({ loadedView: view, layoutXml: "" });
		await run(latest.exportToExcel);
		expect(latest.exportStatus.error).toBe("Export to Excel requires FetchXML and LayoutXML");
		expect(api.exportToExcel).not.toHaveBeenCalled();
	});

	it("a failure shows the Dataverse error (#96) and clears the exporting flag", async () => {
		const error = new Error("The view couldn't be exported.");
		api.exportToExcel.mockRejectedValue(error);
		await show({ loadedView: view });
		await run(latest.exportToExcel);
		expect(latest.exportStatus.error).toBe(
			formatDataverseError(error, "Export to Excel failed. Please try again."),
		);
		expect(latest.exportStatus.isExporting).toBe(false);
	});

	it("doesn't block a second export itself: the command bar disables Export while one runs", async () => {
		const request = deferred<{ excelFile: string; filename: string }>();
		api.exportToExcel.mockReturnValue(request.promise);
		await show({ loadedView: view });

		const pending: Promise<void>[] = [];
		await act(async () => {
			pending.push(latest.exportToExcel());
		});
		expect(latest.exportStatus.isExporting).toBe(true);
		await act(async () => {
			pending.push(latest.exportToExcel());
		});
		expect(api.exportToExcel).toHaveBeenCalledTimes(2);

		await act(async () => {
			request.resolve({ excelFile: "QUJD", filename: "Active Accounts.xlsx" });
			await Promise.all(pending);
		});
		expect(latest.exportStatus.isExporting).toBe(false);
	});
});

describe("useResultsExport: the error", () => {
	it("clears itself after 10 seconds", async () => {
		await show({ queryResult: null });
		vi.useFakeTimers();
		await run(latest.exportLocal);
		expect(latest.exportStatus.error).toBe("No data to export. Execute a query first.");

		await act(async () => {
			vi.advanceTimersByTime(9999);
		});
		expect(latest.exportStatus.error).toBeDefined();
		await act(async () => {
			vi.advanceTimersByTime(1);
		});
		expect(latest.exportStatus.error).toBeUndefined();
	});

	it("clears on dismiss", async () => {
		await show({ queryResult: null });
		await run(latest.exportLocal);
		expect(latest.exportStatus.error).toBeDefined();
		await act(async () => {
			latest.dismissExportError();
		});
		expect(latest.exportStatus.error).toBeUndefined();
	});

	it("a new export clears the previous error", async () => {
		await show({ loadedView: view, layoutXml: "" });
		await run(latest.exportToExcel);
		expect(latest.exportStatus.error).toBeDefined();
		await show({ loadedView: view });
		await run(latest.exportToExcel);
		expect(latest.exportStatus.error).toBeUndefined();
	});
});
