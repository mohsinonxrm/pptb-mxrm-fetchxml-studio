/**
 * useQueryExecution (#301), against the in-memory Dataverse: the first page (as written, by a loaded
 * view's ID, smaller after the 16 MB limit, with a quick find value); paging with and without a
 * paging cookie, and Load more (#58, #60); Retrieve all to the end, with Stop, at the 100,000-row
 * limit, and while Dataverse is throttling (#60, #183); a failed page and Retry; the execute error
 * (#96); clearing the results on an entity change (#60, #102); responses dropped after a newer run or
 * a connection change (#60, #66); the close hold (#70); and handleExecuteRef (#65).
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
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
	type FakeView,
} from "../../shared/testing/fakeDataverse";
import { clearMetadataCache } from "../../features/fetchxml/api/dataverseMetadata";
import { resetDataverseSessionCache } from "../../features/fetchxml/api/sessionCache";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { RETRIEVE_ALL_LIMIT_MESSAGE } from "../../features/fetchxml/api/paging";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import type { LoadedViewState } from "../../features/fetchxml/state/builderStore";
import {
	SIMPLE_PAGING_MESSAGE,
	useQueryExecution,
	type QueryExecution,
	type QueryExecutionInputs,
} from "./useQueryExecution";

let latest: QueryExecution;

function Probe(props: QueryExecutionInputs) {
	latest = useQueryExecution(props);
	return null;
}

/** A tree rooted at `entity`; its entity node's id decides when the results are cleared */
const tree = (entity = "account", entityId = "e1") =>
	({
		id: "f",
		type: "fetch",
		options: {},
		entity: {
			id: entityId,
			type: "entity",
			name: entity,
			attributes: [],
			orders: [],
			filters: [],
			links: [],
		},
	}) as unknown as FetchNode;

/** The account query, with a count when given */
const accountXml = (count?: number) =>
	`<fetch${count === undefined ? "" : ` count='${count}'`}><entity name='account'><attribute name='name'/></entity></fetch>`;

const VIEW_ID = "00000000-0000-0000-0000-00000000a001";

let dataverse: ReturnType<typeof createFakeDataverse>;
let closeHolds: { prevented: number; released: number };

/** An in-memory Dataverse behind a stubbed host that counts close holds */
function connect(options: FakeDataverseOptions & { tables?: FakeTable[] } = {}) {
	resetDataverseSessionCache();
	clearMetadataCache();
	metadataCache.clear();
	dataverse = createFakeDataverse(options.tables ?? sampleTables(5), options);
	installHost({ dataverseAPI: dataverse.api as unknown as HostStub["dataverseAPI"] });
	const toolbox = (window as unknown as { toolboxAPI: Record<string, unknown> }).toolboxAPI;
	closeHolds = { prevented: 0, released: 0 };
	toolbox.preventClose = async () => {
		closeHolds.prevented++;
	};
	toolbox.releasePreventClose = async () => {
		closeHolds.released++;
	};
}

let props: QueryExecutionInputs;

/** Renders the hook; later calls change the given inputs and keep the rest */
const show = (changes: Partial<QueryExecutionInputs> = {}) => {
	props = { ...props, ...changes };
	return render(createElement(Probe, props));
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

/** Starts a hook function without waiting for it (a page is held) */
async function start(fn: () => unknown) {
	await act(async () => {
		void fn();
	});
	await flush();
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

/** The FetchXML of each page request sent for accounts */
const pageRequests = () =>
	dataverse.calls
		.filter((c) => c.startsWith("queryData accounts?fetchXml="))
		.map((c) => decodeURIComponent(c.slice("queryData accounts?fetchXml=".length)));

/** A fetch attribute of a page request */
const fetchAttribute = (xml: string, name: string) =>
	new DOMParser().parseFromString(xml, "application/xml").documentElement.getAttribute(name);

const rows = () => latest.queryResult?.rows.length;

/** Gates page `n`: it's answered once the returned gate is resolved (only the first time it's asked for) */
function holdPage(n: number) {
	const gate = deferred<void>();
	let held = false;
	return {
		gate,
		beforePage: (page: number) => {
			if (page !== n || held) return;
			held = true;
			return gate.promise;
		},
	};
}

beforeEach(() => {
	connect();
	props = {
		fetchQuery: tree(),
		fetchXml: accountXml(),
		quickFindSearchValue: "",
		loadedView: null,
		retrieveAllRecords: false,
	};
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

describe("useQueryExecution: the first page", () => {
	it("runs the FetchXML as written and shows its rows", async () => {
		await show();
		expect(latest.queryResult).toBeNull();
		await run(() => latest.execute());
		expect(pageRequests()).toEqual([accountXml()]);
		expect(latest.isExecuting).toBe(false);
		expect(latest.queryResult).toMatchObject({
			entityLogicalName: "account",
			moreRecords: false,
		});
		expect(rows()).toBe(5);
		expect(latest.queryResult?.rows[0]).toMatchObject({ name: "Account 01" });
		expect(latest.pagingNotice).toBeNull();
		expect(latest.executeError).toBeUndefined();
	});

	it("shows the spinner while the first page loads", async () => {
		const { gate, beforePage } = holdPage(1);
		cleanup();
		connect({ beforePage });
		await show();
		await start(() => latest.execute());
		expect(latest.isExecuting).toBe(true);
		gate.resolve();
		await waitFor(() => !latest.isExecuting);
		expect(rows()).toBe(5);
	});

	it("runs nothing without FetchXML", async () => {
		await show({ fetchQuery: null, fetchXml: "" });
		await run(() => latest.execute());
		expect(pageRequests()).toEqual([]);
		expect(latest.queryResult).toBeNull();
	});

	it("runs the synced editor query instead of the tree (#102)", async () => {
		await show();
		await run(() => latest.execute({ xml: accountXml(2), fetchQuery: tree() }));
		expect(pageRequests()).toEqual([accountXml(2)]);
		expect(rows()).toBe(2);
	});

	it("needs a quick find value for {0}, and puts it in the FetchXML that runs (#142)", async () => {
		const quickFind =
			"<fetch><entity name='account'><attribute name='name'/><filter><condition attribute='name' operator='like' value='{0}'/></filter></entity></fetch>";
		await show({ fetchXml: quickFind });
		await run(() => latest.execute());
		expect(pageRequests()).toEqual([]);

		await show({ quickFindSearchValue: "Acc%" });
		await run(() => latest.execute());
		expect(pageRequests()).toEqual([quickFind.replace("{0}", "Acc%")]);
	});

	it("runs an unmodified loaded view by its ID, and pages with its nextLink (#93)", async () => {
		const view: FakeView = {
			id: VIEW_ID,
			name: "Active accounts",
			entity: "account",
			fetchxml: accountXml(2),
			layoutxml: "",
			type: "system",
		};
		cleanup();
		connect({ views: [view] });
		const loadedView: LoadedViewState = {
			id: VIEW_ID,
			type: "system",
			originalFetchXml: accountXml(2),
			entitySetName: "accounts",
			name: "Active accounts",
		};
		await show({ fetchXml: accountXml(2), loadedView });
		await run(() => latest.execute());
		expect(dataverse.calls).toContain(`queryData accounts?savedQuery=${VIEW_ID}`);
		expect(pageRequests()).toEqual([]);
		expect(rows()).toBe(2);
		expect(latest.queryResult?.moreRecords).toBe(true);
		expect(latest.pagingRef.current?.mode).toBe("view");

		await run(() => latest.loadMore());
		expect(dataverse.calls).toContain(`queryData accounts?savedQuery=${VIEW_ID}&$skiptoken=2`);
		expect(rows()).toBe(4);
		// Views page with nextLink, so this isn't simple paging
		expect(latest.pagingNotice).toBeNull();

		// Modified, the view runs as FetchXML
		await show({ fetchXml: accountXml(3) });
		await run(() => latest.execute());
		expect(pageRequests()).toEqual([accountXml(3)]);
	});

	it("asks again with half the count when the first page is over the host's 16 MB limit (#68)", async () => {
		let failed = false;
		cleanup();
		connect({
			beforePage: () => {
				if (failed) return;
				failed = true;
				throw new Error("Dataverse response exceeds the maximum size of 16 MB");
			},
		});
		await show();
		await run(() => latest.execute());
		const requests = pageRequests();
		expect(requests).toHaveLength(2);
		expect(fetchAttribute(requests[1], "count")).toBe("2500");
		expect(fetchAttribute(requests[1], "page")).toBe("1");
		expect(rows()).toBe(5);
		expect(latest.pagingRef.current?.pageSize).toBe(2500);
		expect(latest.executeError).toBeUndefined();
	});
});

describe("useQueryExecution: paging and Load more (#58, #60)", () => {
	it("pages with the paging cookie and the query's count, until there are no more", async () => {
		await show({ fetchXml: accountXml(2) });
		await run(() => latest.execute());
		expect(rows()).toBe(2);
		expect(latest.queryResult?.moreRecords).toBe(true);
		expect(latest.queryResult?.pagingCookie).toContain("<cookie page=");

		await run(() => latest.loadMore());
		expect(rows()).toBe(4);
		const page2 = pageRequests()[1];
		expect(fetchAttribute(page2, "page")).toBe("2");
		expect(fetchAttribute(page2, "count")).toBe("2");
		expect(fetchAttribute(page2, "paging-cookie")).toContain("<cookie page=");
		expect(latest.pagingNotice).toBeNull();

		await run(() => latest.loadMore());
		expect(rows()).toBe(5);
		expect(latest.queryResult?.moreRecords).toBe(false);
		expect(latest.queryResult?.rows.map((r) => r.name)).toEqual([
			"Account 01",
			"Account 02",
			"Account 03",
			"Account 04",
			"Account 05",
		]);
		expect(dataverse.pagingErrors).toEqual([]);

		// No more pages: Load more sends nothing
		await run(() => latest.loadMore());
		expect(pageRequests()).toHaveLength(3);
		// Load more doesn't hold the close warning; only Retrieve all does
		expect(closeHolds.prevented).toBe(0);
	});

	it("without a paging cookie, pages by number and says so (simple paging)", async () => {
		cleanup();
		connect({ simplePaging: true });
		await show({ fetchXml: accountXml(2) });
		await run(() => latest.execute());
		expect(latest.queryResult?.moreRecords).toBe(true);
		expect(latest.queryResult?.pagingCookie).toBeUndefined();
		expect(latest.pagingNotice).toBeNull();

		await run(() => latest.loadMore());
		const page2 = pageRequests()[1];
		expect(fetchAttribute(page2, "page")).toBe("2");
		expect(fetchAttribute(page2, "paging-cookie")).toBeNull();
		expect(rows()).toBe(4);
		expect(latest.pagingNotice).toEqual({ kind: "simple", message: SIMPLE_PAGING_MESSAGE });

		// The next page is page 3, not Dataverse's pagenumber hint (which stays at 2)
		await run(() => latest.loadMore());
		expect(fetchAttribute(pageRequests()[2], "page")).toBe("3");
		expect(rows()).toBe(5);
		expect(latest.pagingNotice).toEqual({ kind: "simple", message: SIMPLE_PAGING_MESSAGE });

		await run(() => latest.dismissPagingNotice());
		expect(latest.pagingNotice).toBeNull();
	});

	it("loads one page at a time: Load more while a page loads sends nothing", async () => {
		const { gate, beforePage } = holdPage(2);
		cleanup();
		connect({ beforePage });
		await show({ fetchXml: accountXml(2) });
		await run(() => latest.execute());
		await start(() => latest.loadMore());
		expect(latest.isLoadingMore).toBe(true);
		await run(() => latest.loadMore());
		expect(pageRequests()).toHaveLength(2);

		gate.resolve();
		await waitFor(() => !latest.isLoadingMore);
		expect(rows()).toBe(4);
		expect(pageRequests()).toHaveLength(2);
	});
});

describe("useQueryExecution: Retrieve all (#60, #183)", () => {
	it("loads every page after the first, holding the close warning until it ends (#70)", async () => {
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await run(() => latest.execute());
		await waitFor(() => rows() === 5 && !latest.isLoadingMore);
		expect(pageRequests()).toHaveLength(3);
		expect(latest.queryResult?.moreRecords).toBe(false);
		expect(latest.isRetrieveAllInProgress).toBe(false);
		expect(latest.pagingNotice).toBeNull();
		expect(dataverse.pagingErrors).toEqual([]);
		await flush();
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});

	it("doesn't start when the first page is the only one", async () => {
		await show({ retrieveAllRecords: true });
		await run(() => latest.execute());
		expect(rows()).toBe(5);
		expect(latest.isRetrieveAllInProgress).toBe(false);
		expect(closeHolds.prevented).toBe(0);
	});

	it("Stop keeps the loaded rows, drops the page in flight and says so; pages then load one at a time", async () => {
		const { gate, beforePage } = holdPage(2);
		cleanup();
		connect({ beforePage });
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await start(() => latest.execute());
		await waitFor(() => latest.isRetrieveAllInProgress && latest.isLoadingMore);
		expect(closeHolds.prevented).toBe(1);

		await run(() => latest.stopRetrieveAll());
		expect(latest.isRetrieveAllInProgress).toBe(false);
		expect(latest.isLoadingMore).toBe(false);
		expect(latest.pagingNotice).toEqual({ kind: "stopped", message: "" });
		expect(latest.pagingRef.current?.retrieveAll).toBe(false);

		gate.resolve();
		await flush(10);
		expect(rows()).toBe(2);
		expect(pageRequests()).toHaveLength(2);
		expect(latest.pagingNotice).toEqual({ kind: "stopped", message: "" });
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });

		// Stop again does nothing; scrolling loads the next page and the notice goes
		await run(() => latest.stopRetrieveAll());
		await run(() => latest.loadMore());
		expect(rows()).toBe(4);
		expect(latest.pagingNotice).toBeNull();
		expect(pageRequests()).toHaveLength(3);
	});

	it("stops at 100,000 rows and says so; further pages load one at a time", async () => {
		cleanup();
		connect({ tables: sampleTables(100_050) });
		await show({ fetchXml: accountXml(5000), retrieveAllRecords: true });
		await run(() => latest.execute());
		await waitFor(() => latest.pagingNotice?.kind === "limit", 30000);
		expect(latest.pagingNotice).toEqual({
			kind: "limit",
			message: RETRIEVE_ALL_LIMIT_MESSAGE,
		});
		expect(rows()).toBe(100_000);
		expect(pageRequests()).toHaveLength(20);
		expect(latest.queryResult?.moreRecords).toBe(true);
		expect(latest.isRetrieveAllInProgress).toBe(false);
		expect(latest.isLoadingMore).toBe(false);
		expect(latest.pagingRef.current?.retrieveAll).toBe(false);
		await flush();
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });

		await run(() => latest.loadMore());
		expect(rows()).toBe(100_050);
		expect(latest.queryResult?.moreRecords).toBe(false);
		expect(latest.pagingNotice).toBeNull();
	}, 60000);

	it("shows the wait while a throttled page waits to retry, then carries on (#96, #183)", async () => {
		let throttled = false;
		cleanup();
		connect({
			beforePage: (page) => {
				if (page !== 2 || throttled) return;
				throttled = true;
				throw Object.assign(new Error("Number of requests exceeded the limit"), {
					status: 429,
					retryAfter: 1,
				});
			},
		});
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await start(() => latest.execute());
		await waitFor(() => latest.pagingWaitSeconds === 1);
		expect(latest.isRetrieveAllInProgress).toBe(true);
		await waitFor(() => rows() === 5 && !latest.isLoadingMore);
		expect(latest.pagingWaitSeconds).toBeUndefined();
		expect(latest.pagingNotice).toBeNull();
		await flush();
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});

	it("Stop during a throttle wait sends no retry", async () => {
		let throttled = false;
		cleanup();
		connect({
			beforePage: (page) => {
				if (page !== 2 || throttled) return;
				throttled = true;
				throw Object.assign(new Error("Number of requests exceeded the limit"), {
					status: 429,
					retryAfter: 1,
				});
			},
		});
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await start(() => latest.execute());
		await waitFor(() => latest.pagingWaitSeconds === 1);
		await run(() => latest.stopRetrieveAll());
		expect(latest.pagingWaitSeconds).toBeUndefined();
		await waitFor(() => closeHolds.released === 1, 3000);
		expect(pageRequests()).toHaveLength(2);
		expect(latest.pagingNotice).toEqual({ kind: "stopped", message: "" });
		expect(rows()).toBe(2);
	});
});

describe("useQueryExecution: a failed page and Retry (#60, #96)", () => {
	const failOnce = (n: number) => {
		let failed = false;
		return (page: number) => {
			if (page !== n || failed) return;
			failed = true;
			throw new Error("The service is temporarily unavailable.");
		};
	};

	it("Load more: says which page failed and why; Retry loads it", async () => {
		cleanup();
		connect({ beforePage: failOnce(2) });
		await show({ fetchXml: accountXml(2) });
		await run(() => latest.execute());
		await run(() => latest.loadMore());
		expect(latest.pagingNotice).toEqual({
			kind: "error",
			page: 2,
			message: "The service is temporarily unavailable.",
		});
		expect(rows()).toBe(2);
		expect(latest.isLoadingMore).toBe(false);
		expect(latest.executeError).toBeUndefined();

		await run(() => latest.retryPaging());
		await waitFor(() => rows() === 4);
		expect(latest.pagingNotice).toBeNull();
		expect(fetchAttribute(pageRequests()[2], "page")).toBe("2");
		expect(latest.pagingRef.current?.retrieveAll).toBe(false);
		expect(closeHolds.prevented).toBe(0);
	});

	it("Retrieve all: releases the close hold on failure; Retry continues Retrieve all", async () => {
		cleanup();
		connect({ beforePage: failOnce(3) });
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await run(() => latest.execute());
		await waitFor(() => latest.pagingNotice?.kind === "error");
		expect(latest.pagingNotice).toMatchObject({ kind: "error", page: 3 });
		expect(rows()).toBe(4);
		expect(latest.isRetrieveAllInProgress).toBe(false);
		expect(latest.isLoadingMore).toBe(false);
		await flush();
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });

		await run(() => latest.retryPaging());
		await waitFor(() => rows() === 5 && !latest.isLoadingMore);
		expect(latest.pagingNotice).toBeNull();
		expect(latest.queryResult?.moreRecords).toBe(false);
		await flush();
		expect(closeHolds).toEqual({ prevented: 2, released: 2 });
	});
});

describe("useQueryExecution: the execute error (#96)", () => {
	it("says why the run failed, keeps the query, and clears the error when the next run starts", async () => {
		let fail = true;
		const gate = deferred<void>();
		let hold = false;
		cleanup();
		connect({
			beforePage: () => {
				if (fail)
					throw new Error(
						"0x80041103: 'account' entity doesn't contain attribute with Name = 'nme'.",
					);
				if (hold) return gate.promise;
			},
		});
		await show();
		await run(() => latest.execute());
		expect(latest.executeError).toMatch(/doesn't contain attribute with Name = 'nme'/);
		expect(latest.isExecuting).toBe(false);
		expect(latest.queryResult).toEqual({ query: props.fetchQuery, rows: [] });

		fail = false;
		hold = true;
		await start(() => latest.execute());
		expect(latest.isExecuting).toBe(true);
		expect(latest.executeError).toBeUndefined();
		expect(latest.queryResult).toBeNull();
		gate.resolve();
		await waitFor(() => rows() === 5);
		expect(latest.executeError).toBeUndefined();
	});

	it("can be dismissed", async () => {
		cleanup();
		connect({
			beforePage: () => {
				throw new Error("Something went wrong.");
			},
		});
		await show();
		await run(() => latest.execute());
		expect(latest.executeError).toBe("Something went wrong.");
		await run(() => latest.dismissExecuteError());
		expect(latest.executeError).toBeUndefined();
	});
});

describe("useQueryExecution: clearing the results (#60, #102)", () => {
	it("clears the results, paging and the notice when the entity changes", async () => {
		cleanup();
		connect({ simplePaging: true });
		await show({ fetchXml: accountXml(2) });
		await run(() => latest.execute());
		await run(() => latest.loadMore());
		expect(latest.pagingNotice?.kind).toBe("simple");

		await show({ fetchQuery: tree("contact", "e2") });
		expect(latest.queryResult).toBeNull();
		expect(latest.pagingRef.current).toBeNull();
		expect(latest.pagingNotice).toBeNull();
		expect(latest.isLoadingMore).toBe(false);
		expect(latest.isExecuting).toBe(false);
	});

	it("keeps the results when the query changes but its entity node doesn't", async () => {
		await show();
		await run(() => latest.execute());
		await show({ fetchQuery: { ...tree(), options: { distinct: true } } as FetchNode });
		expect(rows()).toBe(5);
	});

	it("keeps results produced for the new tree: an editor-mode Execute (#102)", async () => {
		await show();
		const synced = tree("account", "e2");
		await run(() => latest.execute({ xml: accountXml(2), fetchQuery: synced }));
		expect(rows()).toBe(2);
		// The tree catches up on the next render
		await show({ fetchQuery: synced, fetchXml: accountXml(2) });
		expect(rows()).toBe(2);
	});

	it("drops a first page still loading for the old entity, and ends the spinner", async () => {
		const { gate, beforePage } = holdPage(1);
		cleanup();
		connect({ beforePage });
		await show();
		await start(() => latest.execute());
		expect(latest.isExecuting).toBe(true);

		await show({ fetchQuery: tree("contact", "e2") });
		expect(latest.isExecuting).toBe(false);
		gate.resolve();
		await flush(10);
		expect(latest.queryResult).toBeNull();
	});

	it("stops Retrieve all for the old entity's results; they don't come back", async () => {
		const { gate, beforePage } = holdPage(2);
		cleanup();
		connect({ beforePage });
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await start(() => latest.execute());
		await waitFor(() => latest.isRetrieveAllInProgress);

		await show({ fetchQuery: tree("contact", "e2") });
		expect(latest.queryResult).toBeNull();
		expect(latest.isLoadingMore).toBe(false);
		gate.resolve();
		await flush(10);
		expect(latest.queryResult).toBeNull();
		expect(pageRequests()).toHaveLength(2);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});
});

describe("useQueryExecution: stale responses (#60, #66)", () => {
	it("a newer run drops the first page of the older one", async () => {
		const { gate, beforePage } = holdPage(1);
		cleanup();
		connect({ beforePage });
		await show();
		await start(() => latest.execute());
		await show({ fetchXml: accountXml(3) });
		await run(() => latest.execute());
		expect(rows()).toBe(3);

		gate.resolve();
		await flush(10);
		expect(rows()).toBe(3);
		expect(latest.queryResult?.moreRecords).toBe(true);
		expect(latest.isExecuting).toBe(false);
	});

	it("a newer run stops the older run's Retrieve all", async () => {
		const { gate, beforePage } = holdPage(2);
		cleanup();
		connect({ beforePage });
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await start(() => latest.execute());
		await waitFor(() => latest.isRetrieveAllInProgress);

		await show({ fetchXml: accountXml(), retrieveAllRecords: false });
		await run(() => latest.execute());
		expect(rows()).toBe(5);
		gate.resolve();
		await flush(10);
		expect(rows()).toBe(5);
		expect(latest.pagingNotice).toBeNull();
		expect(pageRequests()).toHaveLength(3);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
	});

	it("unmounting when the connection changes stops Retrieve all and releases the close hold (#66, #70)", async () => {
		const { gate, beforePage } = holdPage(2);
		cleanup();
		connect({ beforePage });
		await show({ fetchXml: accountXml(2), retrieveAllRecords: true });
		await start(() => latest.execute());
		await waitFor(() => latest.isRetrieveAllInProgress);
		const errors = vi.mocked(console.error);
		errors.mockClear();

		// AppShell remounts AppContent for the new connection
		cleanup();
		gate.resolve();
		await flush(10);
		expect(pageRequests()).toHaveLength(2);
		expect(closeHolds).toEqual({ prevented: 1, released: 1 });
		expect(errors).not.toHaveBeenCalled();

		await show();
		expect(latest.queryResult).toBeNull();
	});
});

describe("useQueryExecution: handleExecuteRef (#65)", () => {
	it("is the same ref every render, and always runs the latest query", async () => {
		await show();
		const ref = latest.handleExecuteRef;
		expect(ref.current).toBe(latest.execute);

		await show({ fetchXml: accountXml(3) });
		expect(latest.handleExecuteRef).toBe(ref);
		expect(ref.current).toBe(latest.execute);
		await run(() => ref.current());
		expect(pageRequests()).toEqual([accountXml(3)]);
		expect(rows()).toBe(3);

		await show({ retrieveAllRecords: true, fetchXml: accountXml(2) });
		await run(() => ref.current());
		await waitFor(() => rows() === 5 && !latest.isLoadingMore);
	});
});
