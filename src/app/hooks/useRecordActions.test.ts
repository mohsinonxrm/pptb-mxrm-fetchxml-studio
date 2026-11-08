/**
 * useRecordActions and useRecordActionAvailability (#301), against the in-memory Dataverse: the
 * privileges and why Activate / Deactivate / Close are off, dropped when the entity changes (#179);
 * Open and Copy URL; Delete for one or many records, Delete all and Bulk delete (#174, #177); Run
 * workflow; Activate / Deactivate (#143); Close (#149); and a delete that stops when the tool's
 * connection changes (#175).
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import {
	cleanup,
	deferred,
	flush,
	installHost,
	removeHost,
	render,
	type HostStub,
} from "../../shared/testing/ui";
import {
	createFakeDataverse,
	sampleTables,
	type FakeDataverseOptions,
	type FakeTable,
} from "../../shared/testing/fakeDataverse";
import { clearMetadataCache } from "../../features/fetchxml/api/dataverseMetadata";
import { resetDataverseSessionCache } from "../../features/fetchxml/api/sessionCache";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { findCloseAction } from "../../features/fetchxml/model/closeActions";
import { generateFetchXml } from "../../features/fetchxml/model/fetchxml";
import { deleteRecord, deleteRecordsBatch } from "../../features/fetchxml/api/pptbClient";
import type { PagingState } from "../../features/fetchxml/api/paging";
import type { EntityMetadata, WorkflowInfo } from "../../features/fetchxml/api/pptbClient";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import {
	useRecordActionAvailability,
	useRecordActions,
	type RecordActionAvailability,
	type RecordActions,
	type RecordActionsInputs,
} from "./useRecordActions";

// The real deletes, wrapped so a test can make one throw (the close hold must still be released)
vi.mock("../../features/fetchxml/api/pptbClient", async (importOriginal) => {
	const actual = await importOriginal<typeof import("../../features/fetchxml/api/pptbClient")>();
	return {
		...actual,
		deleteRecord: vi.fn(actual.deleteRecord),
		deleteRecordsBatch: vi.fn(actual.deleteRecordsBatch),
	};
});

let latest: RecordActions;
let availability: RecordActionAvailability;

type ProbeProps = Omit<RecordActionsInputs, "entityLogicalName" | "availability"> & {
	/** The selected entity; the query's root by default */
	entityLogicalName?: string;
};

function Probe({ entityLogicalName: entity, ...props }: ProbeProps) {
	const entityLogicalName = entity ?? props.fetchQuery?.entity.name;
	availability = useRecordActionAvailability(entityLogicalName);
	latest = useRecordActions({ ...props, entityLogicalName, availability });
	return null;
}

/** A query rooted at `entity`; options, attributes and link-entities as given */
const query = (
	entity: string,
	{
		options = {},
		links = [],
		attributes = ["name"],
	}: {
		options?: Record<string, unknown>;
		links?: Record<string, unknown>[];
		attributes?: string[];
	} = {},
) =>
	({
		id: "f",
		type: "fetch",
		options,
		entity: {
			id: "e",
			type: "entity",
			name: entity,
			attributes: attributes.map((name, i) => ({ id: `a${i}`, type: "attribute", name })),
			orders: [],
			filters: [],
			links,
		},
	}) as unknown as FetchNode;

const metadata = (
	logicalName: string,
	primaryId: string,
	primaryName: string,
	entitySet: string,
): EntityMetadata =>
	({
		LogicalName: logicalName,
		PrimaryIdAttribute: primaryId,
		PrimaryNameAttribute: primaryName,
		EntitySetName: entitySet,
	}) as EntityMetadata;

const ACCOUNT = metadata("account", "accountid", "name", "accounts");
const INCIDENT = metadata("incident", "incidentid", "title", "incidents");

const accountId = (n: number) => `aa000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const incidentId = (n: number) => `ii000000-0000-0000-0000-${String(n).padStart(12, "0")}`;

/** The sample tables plus a few cases, for Close */
function tables(accountCount = 5): FakeTable[] {
	return [
		...sampleTables(accountCount),
		{
			logicalName: "incident",
			entitySetName: "incidents",
			displayName: "Case",
			primaryId: "incidentid",
			primaryName: "title",
			objectTypeCode: 112,
			attributes: [
				{ name: "incidentid", type: "Uniqueidentifier" },
				{ name: "title", type: "String" },
				{ name: "statecode", type: "State" },
				{ name: "statuscode", type: "Status" },
			],
			rows: [1, 2].map((n) => ({
				incidentid: incidentId(n),
				title: `Case ${n}`,
				statecode: 0,
				statuscode: 1,
			})),
		},
	];
}

let dataverse: ReturnType<typeof createFakeDataverse>;
let host: HostStub;
/** The tool's connection; a test changes it in place (#66, #175) */
let connection: { id: string; name: string; url: string; environment: string } | null;
let closeHolds: { prevented: number; released: number };
let rerun: Mock<() => Promise<void>>;
let pagingRef: { current: PagingState | null };

/** An in-memory Dataverse behind a stubbed host, with a connection that can change */
function connect(
	options: FakeDataverseOptions & { tables?: FakeTable[]; connected?: boolean } = {},
) {
	resetDataverseSessionCache();
	clearMetadataCache();
	metadataCache.clear();
	dataverse = createFakeDataverse(options.tables ?? tables(), options);
	connection =
		options.connected === false
			? null
			: {
					id: "c1",
					name: "Test environment",
					url: "https://test.crm.dynamics.com/",
					environment: "Dev",
				};
	host = installHost({
		connection: null,
		dataverseAPI: dataverse.api as unknown as HostStub["dataverseAPI"],
	});
	const toolbox = (window as unknown as { toolboxAPI: Record<string, unknown> }).toolboxAPI;
	toolbox.connections = {
		getConnection: async () => connection,
		getActiveConnection: async () => connection,
	};
	closeHolds = { prevented: 0, released: 0 };
	toolbox.preventClose = async () => {
		closeHolds.prevented++;
	};
	toolbox.releasePreventClose = async () => {
		closeHolds.released++;
	};
}

/** The rows a query returned: the account table's, as the grid has them */
const accountRows = () =>
	dataverse.tables
		.find((t) => t.logicalName === "account")!
		.rows.map((row) => ({ accountid: row.accountid, name: row.name }));

const show = (props: Partial<ProbeProps> = {}) => {
	const fetchQuery = props.fetchQuery === undefined ? query("account") : props.fetchQuery;
	return render(
		createElement(Probe, {
			fetchQuery,
			entityMetadata: ACCOUNT,
			queryResult: { query: fetchQuery ?? undefined, rows: accountRows() },
			executeError: undefined,
			pagingRef,
			handleExecuteRef: { current: rerun },
			...props,
		}),
	);
};

/** Calls a hook function inside act, so its state updates render */
async function run<T>(fn: () => T | Promise<T>): Promise<T> {
	let result!: T;
	await act(async () => {
		result = await fn();
	});
	await flush();
	return result;
}

/** Waits until the predicate holds, letting requests and effects settle */
async function waitFor(predicate: () => boolean, ms = 3000) {
	const end = Date.now() + ms;
	while (!predicate()) {
		if (Date.now() > end) throw new Error("waitFor timed out");
		await act(async () => {
			await new Promise((resolve) => setTimeout(resolve, 5));
		});
	}
}

/** Answers a query from `answer` when it matches, and from the fake otherwise */
function answerQuery(match: (query: string) => boolean, answer: () => unknown) {
	const fake = host.dataverseAPI.queryData;
	host.dataverseAPI.queryData = async (q: unknown) =>
		match(String(q)) ? answer() : fake(q as string);
}

/** Records the requests passed to dataverseAPI.execute */
function captureExecute() {
	const requests: { operationName: string; parameters?: Record<string, unknown> }[] = [];
	const fake = host.dataverseAPI.execute;
	host.dataverseAPI.execute = async (request: unknown) => {
		requests.push(request as (typeof requests)[number]);
		return fake(request);
	};
	return requests;
}

const accounts = () => dataverse.tables.find((t) => t.logicalName === "account")!;

beforeEach(() => {
	connect();
	rerun = vi.fn<() => Promise<void>>(async () => undefined);
	pagingRef = { current: null };
	vi.spyOn(console, "error").mockImplementation(() => undefined);
	vi.spyOn(console, "warn").mockImplementation(() => undefined);
});

afterEach(() => {
	cleanup();
	removeHost();
	resetDataverseSessionCache();
	clearMetadataCache();
	metadataCache.clear();
	vi.restoreAllMocks();
});

describe("useRecordActionAvailability: privileges (#179)", () => {
	it("allows Delete, Bulk delete and Run workflow when the user has the privileges", async () => {
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		expect(availability.recordActionPrivileges).toEqual({
			canDelete: true,
			canBulkDelete: true,
			canRunWorkflow: true,
			privilegesChecked: true,
		});
		expect(latest.canDelete).toBe(true);
		expect(latest.canBulkDelete).toBe(true);
		expect(latest.canRunWorkflow).toBe(true);
	});

	it("turns each action off when its privilege is missing", async () => {
		cleanup();
		connect({ privileges: new Set(["prvReadWorkflow"]) });
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		expect(latest.canDelete).toBe(false);
		expect(latest.canBulkDelete).toBe(false);
		// Run workflow needs both prvReadWorkflow and prvWorkflowExecution
		expect(latest.canRunWorkflow).toBe(false);
	});

	it("drops the privileges when the entity changes, until the new entity's checks finish", async () => {
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		expect(latest.canDelete).toBe(true);

		const gate = deferred<void>();
		const fake = host.dataverseAPI.queryData;
		host.dataverseAPI.queryData = async (q: unknown) => {
			if (String(q).includes("prvDeleteContact")) await gate.promise;
			return fake(q as string);
		};
		await show({ fetchQuery: query("contact") });
		expect(availability.recordActionPrivileges).toEqual({
			canDelete: false,
			canBulkDelete: false,
			canRunWorkflow: false,
			privilegesChecked: false,
		});

		gate.resolve();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		expect(latest.canDelete).toBe(true);
	});

	it("drops a check for an entity that's no longer selected", async () => {
		cleanup();
		connect({ privileges: new Set(["prvDeleteAccount", "prvBulkDelete"]) });
		const gate = deferred<void>();
		const fake = host.dataverseAPI.queryData;
		host.dataverseAPI.queryData = async (q: unknown) => {
			if (String(q).includes("prvDeleteAccount")) await gate.promise;
			return fake(q as string);
		};
		await show();
		await show({ fetchQuery: query("contact") });
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		expect(latest.canDelete).toBe(false);

		// The account check answers last: it would allow Delete on contacts
		gate.resolve();
		await flush(10);
		expect(latest.canDelete).toBe(false);
		expect(latest.canBulkDelete).toBe(true);
	});
});

describe("useRecordActionAvailability: Activate / Deactivate / Close (#143, #149)", () => {
	it("leaves Activate and Deactivate on for an entity with a state model, and no Close menu", async () => {
		await show();
		await waitFor(() => dataverse.calls.some((c) => c.includes("EnforceStateTransitions")));
		await flush(10);
		expect(latest.activateDisabledReason).toBeUndefined();
		expect(latest.deactivateDisabledReason).toBeUndefined();
		expect(latest.closeActions).toBeUndefined();
	});

	it("says why when the user doesn't have the entity's Write privilege", async () => {
		cleanup();
		connect({ privileges: new Set(["prvDeleteAccount"]) });
		answerQuery(
			(q) => q === "EntityDefinitions(LogicalName='account')?$select=Privileges",
			() => ({ Privileges: [{ Name: "prvWriteAccount", PrivilegeType: "Write" }] }),
		);
		await show();
		const reason =
			"You don't have the Write privilege on this entity, so you can't change its records' status.";
		await waitFor(() => latest.activateDisabledReason === reason);
		expect(latest.deactivateDisabledReason).toBe(reason);
	});

	it("gives a close-message entity its Close menu, and says why Deactivate is off", async () => {
		await show({ fetchQuery: query("incident"), entityMetadata: INCIDENT });
		expect(latest.deactivateDisabledReason).toBe(
			"Cases are resolved with Resolve Case (CloseIncident). Use Close.",
		);
		expect(latest.closeActions?.map((a) => [a.key, a.label])).toEqual([
			["resolve", "Resolve case"],
			["cancel", "Cancel case"],
		]);
		expect(latest.closeActions?.every((a) => a.disabledReason === undefined)).toBe(true);
	});

	it("turns a close message off without the Create privilege on its activity", async () => {
		cleanup();
		connect({ privileges: new Set(["prvWriteIncident"]) });
		answerQuery(
			(q) => q === "EntityDefinitions(LogicalName='incidentresolution')?$select=Privileges",
			() => ({ Privileges: [{ Name: "prvCreateActivity", PrivilegeType: "Create" }] }),
		);
		await show({ fetchQuery: query("incident"), entityMetadata: INCIDENT });
		await waitFor(() => latest.closeActions?.[0].disabledReason !== undefined);
		expect(latest.closeActions?.map((a) => a.disabledReason)).toEqual([
			"You don't have the Create privilege on activities, and CloseIncident creates a incidentresolution activity.",
			undefined,
		]);
	});

	it("drops a slow answer for an entity that's no longer selected (#143, #149)", async () => {
		const gate = deferred<void>();
		const fake = host.dataverseAPI.queryData;
		host.dataverseAPI.queryData = async (q: unknown) => {
			if (String(q).startsWith("EntityDefinitions(LogicalName='incident')/Attributes"))
				await gate.promise;
			return fake(q as string);
		};
		await show({ fetchQuery: query("incident"), entityMetadata: INCIDENT });
		expect(latest.closeActions).toHaveLength(2);

		await show();
		await waitFor(() =>
			dataverse.calls.some(
				(c) =>
					c ===
					"queryData EntityDefinitions(LogicalName='account')?$select=EnforceStateTransitions",
			),
		);
		await flush(10);
		expect(latest.deactivateDisabledReason).toBeUndefined();
		expect(latest.closeActions).toBeUndefined();

		// The case answers last: it would bring back the Close menu and the case's Deactivate reason
		gate.resolve();
		await flush(10);
		await waitFor(() =>
			dataverse.calls.some((c) => c.includes("LogicalName='incident')?$select=Enforce")),
		);
		await flush(10);
		expect(latest.activateDisabledReason).toBeUndefined();
		expect(latest.deactivateDisabledReason).toBeUndefined();
		expect(latest.closeActions).toBeUndefined();
	});

	it("says rows repeat when the query has a one-to-many link-entity", async () => {
		await show({
			fetchQuery: query("account", {
				links: [
					{
						id: "l1",
						type: "link-entity",
						name: "contact",
						relationshipType: "1N",
						attributes: [],
						orders: [],
						filters: [],
						links: [],
					},
				],
			}),
		});
		await flush(10);
		expect(latest.hasOneToManyRelationship).toBe(true);
		expect(latest.activateDisabledReason).toBe(
			"The query has a one-to-many or many-to-many link-entity, so rows repeat. Remove it to activate or deactivate records.",
		);
		expect(latest.deactivateDisabledReason).toBe(latest.activateDisabledReason);
		expect(latest.closeDisabledReason).toBe(
			"The query has a one-to-many or many-to-many link-entity, so rows repeat. Remove it to close records.",
		);
	});
});

describe("useRecordActions: why record actions are off", () => {
	it("says why when the results have no record IDs: aggregate, distinct, or no primary key", async () => {
		const rows = [{ name: "Account 01" }];
		const aggregate = query("account", { options: { aggregate: true } });
		await show({ fetchQuery: aggregate, queryResult: { query: aggregate, rows } });
		expect(latest.isAggregateQuery).toBe(true);
		expect(latest.recordActionsDisabledReason).toBe(
			"Aggregate query results don't contain record IDs, so record actions aren't available.",
		);

		const distinct = query("account", { options: { distinct: true } });
		await show({ fetchQuery: distinct, queryResult: { query: distinct, rows } });
		expect(latest.recordActionsDisabledReason).toMatch(/^Distinct query results/);

		await show({ queryResult: { query: query("account"), rows } });
		expect(latest.recordActionsDisabledReason).toBe(
			"These results don't include the table's primary key, so record actions aren't available.",
		);

		await show();
		expect(latest.recordActionsDisabledReason).toBeUndefined();
	});

	it("turns Delete all off until the query on screen ran without an error (#174)", async () => {
		await show();
		expect(latest.deleteAllDisabledReason).toBeUndefined();

		await show({ executeError: "0x80040216: An unexpected error occurred." });
		expect(latest.deleteAllDisabledReason).toBe(
			"Run the query first: Delete all deletes the records the query on screen matches.",
		);

		const edited = query("account", { options: { top: 3 } });
		await show({ fetchQuery: edited, queryResult: { query: query("account"), rows: [] } });
		expect(latest.deleteAllDisabledReason).toMatch(/^The query has changed/);
	});
});

describe("useRecordActions: selection, Open and Copy URL", () => {
	it("keeps the grid's selection for the command bar", async () => {
		await show();
		expect(latest.getSelectedRecordIds()).toEqual([]);
		await run(() => latest.onSelectionChange([accountId(1), accountId(2)]));
		expect(latest.getSelectedRecordIds()).toEqual([accountId(1), accountId(2)]);
	});

	it("opens each record in the connection's browser", async () => {
		const opened: string[] = [];
		const utils = (window as unknown as { toolboxAPI: { utils: Record<string, unknown> } })
			.toolboxAPI.utils;
		utils.openInConnectionBrowser = async (url: string) => {
			opened.push(url);
		};
		await show();
		await run(() => latest.openRecord([accountId(1), accountId(2)]));
		expect(opened).toEqual([
			`https://test.crm.dynamics.com/main.aspx?etn=account&id=${accountId(1)}&pagetype=entityrecord`,
			`https://test.crm.dynamics.com/main.aspx?etn=account&id=${accountId(2)}&pagetype=entityrecord`,
		]);
	});

	it("opens nothing without an environment URL, and says so in the console", async () => {
		cleanup();
		connect({ connected: false });
		const open = vi.spyOn(window, "open").mockImplementation(() => null);
		await show();
		await run(() => latest.openRecord([accountId(1)]));
		expect(open).not.toHaveBeenCalled();
		expect(console.error).toHaveBeenCalledWith("Failed to get environment URL");

		await run(() => latest.copyRecordUrl([accountId(1)]));
		expect(console.error).toHaveBeenCalledTimes(2);
	});

	it("does nothing with no records", async () => {
		const open = vi.spyOn(window, "open").mockImplementation(() => null);
		await show();
		await run(() => latest.openRecord([]));
		expect(open).not.toHaveBeenCalled();
	});

	it("copies the record URLs, one per line", async () => {
		const copied: string[] = [];
		const utils = (window as unknown as { toolboxAPI: { utils: Record<string, unknown> } })
			.toolboxAPI.utils;
		utils.copyToClipboard = async (text: string) => {
			copied.push(text);
		};
		await show();
		await run(() => latest.copyRecordUrl([accountId(1), accountId(2)]));
		expect(copied).toEqual([
			[1, 2]
				.map(
					(n) =>
						`https://test.crm.dynamics.com/main.aspx?etn=account&id=${accountId(n)}&pagetype=entityrecord`,
				)
				.join("\n"),
		]);
		expect(console.error).not.toHaveBeenCalled();
	});

	it("logs when the clipboard can't be written", async () => {
		const utils = (window as unknown as { toolboxAPI: { utils: Record<string, unknown> } })
			.toolboxAPI.utils;
		utils.copyToClipboard = async () => {
			throw new Error("denied");
		};
		await show();
		await run(() => latest.copyRecordUrl([accountId(1)]));
		expect(console.error).toHaveBeenCalledWith("Failed to copy to clipboard");
	});
});

describe("useRecordActions: Delete", () => {
	it("opens the Delete dialog with the record's name for one record", async () => {
		await show();
		await run(() => latest.deleteRecords([accountId(1)]));
		expect(latest.deleteDialogState).toEqual({
			open: true,
			recordIds: [accountId(1)],
			recordName: "Account 01",
			isBatchDelete: false,
		});
		expect(dataverse.deleted).toEqual([]);
	});

	it("deletes one record, runs the query again and releases the close hold", async () => {
		await show();
		await run(() => latest.deleteRecords([accountId(1)]));
		const result = await run(() => latest.confirmDelete());
		expect(result).toBeUndefined();
		expect(dataverse.deleted).toEqual([`account:${accountId(1)}`]);
		expect(accounts().rows).toHaveLength(4);
		expect(rerun).toHaveBeenCalledTimes(1);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});

	it("deletes 4 to 100 records in $batch requests, with progress", async () => {
		await show();
		const ids = [1, 2, 3, 4].map(accountId);
		await run(() => latest.deleteRecords(ids));
		expect(latest.deleteDialogState.isBatchDelete).toBe(true);
		expect(latest.deleteDialogState.recordName).toBeUndefined();

		const result = await run(() => latest.confirmDelete());
		expect(result).toEqual({ succeeded: 4, failed: 0, errors: [] });
		expect(dataverse.deleted).toEqual(ids.map((id) => `account:${id}`));
		expect(latest.deleteProgress).toMatchObject({ completed: 4, total: 4, succeeded: 4 });
		expect(rerun).toHaveBeenCalledTimes(1);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });

		await run(() => latest.dismissDeleteDialog());
		expect(latest.deleteDialogState).toEqual({ open: false, recordIds: [] });
		expect(latest.deleteProgress).toBeNull();
	});

	it("releases the close hold when the batch delete throws (#70)", async () => {
		const failure = new Error("0x80040216: An unexpected error occurred.");
		vi.mocked(deleteRecordsBatch).mockImplementationOnce(async () => {
			throw failure;
		});
		await show();
		await run(() => latest.deleteRecords([1, 2, 3, 4].map(accountId)));
		await expect(run(() => latest.confirmDelete())).rejects.toBe(failure);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
		expect(dataverse.deleted).toEqual([]);
		expect(rerun).not.toHaveBeenCalled();
	});

	it("releases the close hold when a single delete throws, and returns the failure (#70)", async () => {
		vi.mocked(deleteRecord).mockImplementationOnce(async () => {
			throw new Error("0x80040216: An unexpected error occurred.");
		});
		await show();
		await run(() => latest.deleteRecords([accountId(1)]));
		const result = await run(() => latest.confirmDelete());
		expect(result).toMatchObject({ succeeded: 0, failed: 1 });
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
		expect(dataverse.deleted).toEqual([]);
	});

	it("runs the latest query runner when a delete finishes, not the one it was opened with", async () => {
		const first = vi.fn<() => Promise<void>>(async () => undefined);
		const second = vi.fn<() => Promise<void>>(async () => undefined);
		const handleExecuteRef = { current: first };
		await show({ handleExecuteRef });
		await run(() => latest.deleteRecords([accountId(1)]));
		// AppShell's layout effect updates the ref after every render; no render in between here
		handleExecuteRef.current = second;
		await run(() => latest.confirmDelete());
		expect(second).toHaveBeenCalledTimes(1);
		expect(first).not.toHaveBeenCalled();
	});

	it("returns the failures and doesn't run the query again when nothing was deleted", async () => {
		await show();
		const missing = accountId(99);
		await run(() => latest.deleteRecords([missing]));
		const result = await run(() => latest.confirmDelete());
		expect(result).toMatchObject({ succeeded: 0, failed: 1 });
		expect(result?.errors[0]).toMatch(new RegExp(`^Failed to delete ${missing}: `));
		expect(rerun).not.toHaveBeenCalled();
		expect(console.error).toHaveBeenCalledWith("Delete errors:", result?.errors);
	});

	it("stops deleting when the tool's connection changes, so the next delete can't run in the new environment (#175)", async () => {
		const fake = host.dataverseAPI.executeBatch;
		host.dataverseAPI.executeBatch = async (requests: unknown) => {
			const answer = await fake(requests);
			// Connected to another environment in place, after the first delete (PPTB ≥ 1.2.7, #66)
			connection = {
				id: "c2",
				name: "Production",
				url: "https://prod.crm.dynamics.com/",
				environment: "Prod",
			};
			return answer;
		};
		await show();
		const ids = [1, 2, 3].map(accountId);
		await run(() => latest.deleteRecords(ids));
		expect(latest.deleteDialogState.isBatchDelete).toBe(false);

		const result = await run(() => latest.confirmDelete());
		expect(dataverse.deleted).toEqual([`account:${accountId(1)}`]);
		expect(result).toMatchObject({ succeeded: 1, failed: 2 });
		expect(result?.errors).toEqual([
			`Failed to delete ${accountId(2)}: Not deleted: the tool's connection changed during the delete`,
			`Failed to delete ${accountId(3)}: Not deleted: the tool's connection changed during the delete`,
		]);
		// One record was deleted, so the results are refreshed
		expect(rerun).toHaveBeenCalledTimes(1);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});

	it("sends more than 100 records to a Bulk delete job, only with the privilege (#177)", async () => {
		const ids = Array.from({ length: 101 }, (_, i) => accountId(i + 1));
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.deleteRecords(ids));
		expect(latest.deleteDialogState.open).toBe(false);
		expect(latest.bulkDeleteDialogState).toEqual({
			open: true,
			recordIds: ids,
			isAllRecords: false,
		});

		const requests = captureExecute();
		const job = await run(() => latest.confirmBulkDelete("Delete 101 accounts"));
		expect(job).toEqual({ asyncOperationId: "", bulkDeleteOperationId: "", jobUrl: "" });
		const bulkDelete = requests.find((r) => r.operationName === "BulkDelete");
		expect(bulkDelete?.parameters?.JobName).toBe("Delete 101 accounts");
		expect(JSON.stringify(bulkDelete?.parameters?.QuerySet)).toContain(accountId(101));

		cleanup();
		connect({ privileges: new Set(["prvDeleteAccount"]) });
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.deleteRecords(ids));
		expect(latest.bulkDeleteDialogState.open).toBe(false);
		expect(latest.deleteDialogState.open).toBe(false);
	});

	it("Delete all opens the Bulk delete dialog for every record the query matches (#174)", async () => {
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.deleteRecords([]));
		expect(latest.bulkDeleteDialogState).toEqual({
			open: true,
			recordIds: [],
			isAllRecords: true,
			totalViewRecords: 5,
		});
		await run(() => latest.dismissBulkDeleteDialog());
		expect(latest.bulkDeleteDialogState).toEqual({ open: false, recordIds: [] });

		await run(() => latest.bulkDeleteRecords([]));
		expect(latest.bulkDeleteDialogState.isAllRecords).toBe(true);
		expect(latest.bulkDeleteDialogState.totalViewRecords).toBe(5);
	});

	it("Delete all submits the FetchXML that ran, with the quick find value (#142)", async () => {
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.deleteRecords([]));
		// Read when the job is submitted, not when the dialog opened
		pagingRef.current = {
			mode: "fetchxml",
			xml: "<fetch><entity name='account'><filter><condition attribute='name' operator='like' value='Contoso%'/></filter></entity></fetch>",
			currentPage: 1,
			nextPage: 2,
			moreRecords: false,
			isRetrieveAllInProgress: false,
			retrieveAll: false,
		};
		const requests = captureExecute();
		await run(() => latest.confirmBulkDelete("Delete all accounts"));
		const conversion = dataverse.calls.find((c) => c.includes("FetchXmlToQueryExpression"));
		expect(decodeURIComponent(conversion ?? "")).toContain("Contoso%");
		expect(requests.map((r) => r.operationName)).toEqual(["BulkDelete"]);
	});

	it("Delete all without a paging state submits the query that ran, not the tree (#174)", async () => {
		// A column added in the tree after the run: Delete all is still allowed, it matches the same records
		const executed = query("account", { attributes: ["name"] });
		const tree = query("account", { attributes: ["name", "telephone1"] });
		await show({ fetchQuery: tree, queryResult: { query: executed, rows: accountRows() } });
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		expect(latest.deleteAllDisabledReason).toBeUndefined();
		await run(() => latest.deleteRecords([]));
		await run(() => latest.confirmBulkDelete("Delete all accounts"));
		const conversion = dataverse.calls.find((c) => c.includes("FetchXmlToQueryExpression"));
		const submitted = decodeURIComponent(conversion ?? "");
		expect(generateFetchXml(executed)).not.toBe(generateFetchXml(tree));
		expect(submitted).toContain(generateFetchXml(executed).replace(/'/g, "''"));
		expect(submitted).not.toContain("telephone1");
	});

	it("keeps Delete all closed when the results are stale or there's no Bulk delete privilege", async () => {
		await show({ executeError: "0x80040216: An unexpected error occurred." });
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.deleteRecords([]));
		await run(() => latest.bulkDeleteRecords([]));
		expect(latest.bulkDeleteDialogState.open).toBe(false);

		cleanup();
		connect({ privileges: new Set(["prvDeleteAccount"]) });
		await show();
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.deleteRecords([]));
		await run(() => latest.bulkDeleteRecords([accountId(1)]));
		expect(latest.bulkDeleteDialogState.open).toBe(false);
	});

	it("Bulk delete fails with the exact messages when it can't run", async () => {
		const requests = captureExecute();
		await show({ entityMetadata: null });
		await waitFor(() => availability.recordActionPrivileges.privilegesChecked);
		await run(() => latest.bulkDeleteRecords([accountId(1)]));
		await expect(latest.confirmBulkDelete("job")).rejects.toThrow(
			"Missing entity information for bulk delete",
		);

		// Delete all after the tree changed which records match: Delete all's own reason (#174)
		await show();
		await run(() => latest.bulkDeleteRecords([]));
		expect(latest.bulkDeleteDialogState.isAllRecords).toBe(true);
		await show({
			fetchQuery: query("account", { options: { top: 3 } }),
			queryResult: { query: query("account"), rows: accountRows() },
		});
		await expect(latest.confirmBulkDelete("job")).rejects.toThrow(
			"The query has changed since these results were loaded. Run it again to see what Delete all would delete.",
		);

		// Delete all with no query that ran and no tree to compare: the fallback message
		await show({ fetchQuery: null, entityLogicalName: "account", queryResult: null });
		expect(latest.deleteAllDisabledReason).toBeUndefined();
		await run(() => latest.bulkDeleteRecords([]));
		expect(latest.bulkDeleteDialogState.isAllRecords).toBe(true);
		await expect(latest.confirmBulkDelete("job")).rejects.toThrow(
			"Run the query first: Delete all deletes the records the query on screen matches.",
		);
		expect(requests.map((r) => r.operationName)).not.toContain("BulkDelete");
		expect(dataverse.calls.some((c) => c.includes("FetchXmlToQueryExpression"))).toBe(false);
	});
});

describe("useRecordActions: Run workflow", () => {
	const workflow: WorkflowInfo = {
		workflowid: "ww000000-0000-0000-0000-000000000001",
		name: "Notify owner",
		primaryentity: "account",
		category: 0,
		type: 1,
		statuscode: 2,
		mode: 1,
	};

	it("opens the workflow dialog with the workflow chosen, for selected records only", async () => {
		await show();
		await run(() => latest.runSpecificWorkflow(workflow, []));
		expect(latest.workflowDialogState.open).toBe(false);
		await run(() => latest.runSpecificWorkflow(workflow, [accountId(1)]));
		expect(latest.workflowDialogState).toEqual({
			open: true,
			recordIds: [accountId(1)],
			preSelectedWorkflow: workflow,
		});
		await run(() => latest.dismissWorkflowDialog());
		expect(latest.workflowDialogState).toEqual({ open: false, recordIds: [] });
	});

	it("lists the on-demand workflows of the entity asked for, or of the selected entity", async () => {
		await show();
		await latest.fetchWorkflows("contact");
		await latest.fetchWorkflows();
		const listed = dataverse.calls.filter((c) => c.startsWith("queryData workflows?"));
		expect(listed).toHaveLength(2);
		expect(listed[0]).toContain("primaryentity eq 'contact'");
		expect(listed[1]).toContain("primaryentity eq 'account'");

		await show({ fetchQuery: null, queryResult: null });
		expect(await latest.fetchWorkflows()).toEqual([]);
	});

	it("runs the workflow on each record while holding the close warning", async () => {
		const requests = captureExecute();
		await show();
		const progress = vi.fn();
		const result = await run(() =>
			latest.executeWorkflow(workflow, [accountId(1), accountId(2)], progress),
		);
		expect(result).toEqual({ succeeded: 2, failed: 0, errors: [] });
		expect(
			requests.filter((r) => r.operationName === "ExecuteWorkflow").map((r) => r.parameters),
		).toEqual([{ EntityId: accountId(1) }, { EntityId: accountId(2) }]);
		expect(progress).toHaveBeenCalled();
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
		// Workflows don't change the rows on screen: the query doesn't run again
		expect(rerun).not.toHaveBeenCalled();
	});

	it("runs nothing without an entity", async () => {
		const requests = captureExecute();
		await show({ fetchQuery: null, queryResult: null });
		const result = await latest.executeWorkflow(workflow, [accountId(1)], vi.fn());
		expect(result).toEqual({ succeeded: 0, failed: 1, errors: ["Entity not selected"] });
		expect(requests).toEqual([]);
	});
});

describe("useRecordActions: Activate / Deactivate (#143)", () => {
	it("opens the confirmation for the action and the records", async () => {
		await show();
		await run(() => latest.activateRecords([]));
		expect(latest.setStateDialogState.open).toBe(false);
		await run(() => latest.deactivateRecords([accountId(1), accountId(2)]));
		expect(latest.setStateDialogState).toEqual({
			open: true,
			action: "deactivate",
			recordIds: [accountId(1), accountId(2)],
		});
		await run(() => latest.activateRecords([accountId(3)]));
		expect(latest.setStateDialogState).toEqual({
			open: true,
			action: "activate",
			recordIds: [accountId(3)],
		});
		// Closing keeps the action and records, so the dialog doesn't change while it animates out
		await run(() => latest.dismissSetStateDialog());
		expect(latest.setStateDialogState).toEqual({
			open: false,
			action: "activate",
			recordIds: [accountId(3)],
		});
	});

	it("sets statecode and statuscode on the records, then runs the query again", async () => {
		await show();
		await run(() => latest.deactivateRecords([accountId(1), accountId(2)]));
		const result = await run(() =>
			latest.executeStateChange({ statecode: 1, statuscode: 2 }, vi.fn()),
		);
		expect(result).toMatchObject({ succeeded: 2, failed: 0 });
		const batch = dataverse.calls.find((c) => c.startsWith("executeBatch"));
		expect(batch).toBe(
			`executeBatch PATCH accounts(${accountId(1)}), PATCH accounts(${accountId(2)})`,
		);
		expect(rerun).toHaveBeenCalledTimes(1);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});

	it("runs the latest query runner when the change finishes, not the one it was opened with", async () => {
		const first = vi.fn<() => Promise<void>>(async () => undefined);
		const second = vi.fn<() => Promise<void>>(async () => undefined);
		const handleExecuteRef = { current: first };
		await show({ handleExecuteRef });
		await run(() => latest.activateRecords([accountId(1)]));
		handleExecuteRef.current = second;
		await run(() => latest.executeStateChange({ statecode: 0, statuscode: 1 }, vi.fn()));
		expect(second).toHaveBeenCalledTimes(1);
		expect(first).not.toHaveBeenCalled();
	});

	it("doesn't run without an entity", async () => {
		await show({ fetchQuery: null, queryResult: null });
		await run(() => latest.activateRecords([accountId(1)]));
		expect(latest.setStateDialogState.open).toBe(false);
		await expect(
			latest.executeStateChange({ statecode: 0, statuscode: 1 }, vi.fn()),
		).rejects.toThrow("Entity not selected");
	});
});

describe("useRecordActions: Close (#149)", () => {
	const showCases = () =>
		show({
			fetchQuery: query("incident"),
			entityMetadata: INCIDENT,
			queryResult: { query: query("incident"), rows: [] },
		});

	it("a plain state change (Cancel case) opens the Deactivate confirmation limited to its state", async () => {
		await showCases();
		await run(() => latest.closeRecords("cancel", [incidentId(1)]));
		expect(latest.setStateDialogState).toEqual({
			open: true,
			action: "deactivate",
			recordIds: [incidentId(1)],
			closeAction: findCloseAction("incident", "cancel"),
		});
		expect(latest.closeDialogState.open).toBe(false);
	});

	it("a close message (Resolve case) opens its dialog, and runs on each record", async () => {
		await showCases();
		await run(() => latest.closeRecords("resolve", [incidentId(1), incidentId(2)]));
		expect(latest.closeDialogState).toEqual({
			open: true,
			action: findCloseAction("incident", "resolve"),
			recordIds: [incidentId(1), incidentId(2)],
		});

		const result = await run(() =>
			latest.executeClose({ statuscode: 5, resolution: "Fixed" }, vi.fn()),
		);
		expect(result).toMatchObject({ succeeded: 2, failed: 0 });
		const batch = dataverse.calls.find((c) => c.startsWith("executeBatch"));
		expect(batch).toBe("executeBatch POST CloseIncident, POST CloseIncident");
		expect(rerun).toHaveBeenCalledTimes(1);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });

		await run(() => latest.dismissCloseDialog());
		expect(latest.closeDialogState.open).toBe(false);
		expect(latest.closeDialogState.recordIds).toEqual([incidentId(1), incidentId(2)]);
	});

	it("runs the latest query runner when the close finishes, not the one it was opened with", async () => {
		const first = vi.fn<() => Promise<void>>(async () => undefined);
		const second = vi.fn<() => Promise<void>>(async () => undefined);
		const handleExecuteRef = { current: first };
		await show({
			fetchQuery: query("incident"),
			entityMetadata: INCIDENT,
			queryResult: { query: query("incident"), rows: [] },
			handleExecuteRef,
		});
		await run(() => latest.closeRecords("resolve", [incidentId(1)]));
		handleExecuteRef.current = second;
		await run(() => latest.executeClose({ statuscode: 5, resolution: "Fixed" }, vi.fn()));
		expect(second).toHaveBeenCalledTimes(1);
		expect(first).not.toHaveBeenCalled();
	});

	it("ignores an unknown action or no records", async () => {
		await showCases();
		await run(() => latest.closeRecords("win", [incidentId(1)]));
		await run(() => latest.closeRecords("resolve", []));
		expect(latest.closeDialogState.open).toBe(false);
		expect(latest.setStateDialogState.open).toBe(false);
		await expect(latest.executeClose({ statuscode: 5 }, vi.fn())).rejects.toThrow(
			"Entity not selected",
		);
	});
});
