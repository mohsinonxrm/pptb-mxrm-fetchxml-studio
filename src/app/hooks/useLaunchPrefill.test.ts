/**
 * useLaunchPrefill (#301), with the real builder and the in-memory Dataverse: the query carried over
 * a connection change (#66), the tool-to-tool prefill (entity, FetchXML, a view by viewRef, a view
 * that can't be opened, #197) applied once per tool instance, and loading a view as Load view does
 * (its layout, execution by ID and Save view's overwrite, the quick find reset and the load notice,
 * #93, #142).
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
	cleanup,
	deferred,
	flush,
	installHost,
	removeHost,
	render,
	toasts,
	type HostStub,
} from "../../shared/testing/ui";
import {
	createFakeDataverse,
	sampleTables,
	type FakeView,
} from "../../shared/testing/fakeDataverse";
import { clearMetadataCache } from "../../features/fetchxml/api/dataverseMetadata";
import { resetDataverseSessionCache } from "../../features/fetchxml/api/sessionCache";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { BuilderProvider } from "../../features/fetchxml/state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import { generateFetchXml } from "../../features/fetchxml/model/fetchxml";
import type { LoadedViewInfo } from "../../features/fetchxml/api/pptbClient";
import type { LaunchContextState } from "../../shared/hooks/useLaunchContext";
import {
	useLaunchPrefill,
	type ConnectionCarryOver,
	type LaunchPrefill,
	type LaunchPrefillInputs,
} from "./useLaunchPrefill";
import { useTreeNotice, type TreeNoticeInput } from "./useTreeNotice";

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

let builder: BuilderContextValue;
let latest: LaunchPrefill;

/** The MessageBar above the tree (#277): what it shows now */
let notice: TreeNoticeInput | null;

type ProbeProps = Omit<LaunchPrefillInputs, "builder" | "showTreeNotice"> & {
	/** Replaces the real tree notice's show */
	showTreeNotice?: LaunchPrefillInputs["showTreeNotice"];
};
function Probe({ showTreeNotice, ...props }: ProbeProps) {
	builder = useBuilder();
	const tree = useTreeNotice(builder.fetchQuery?.entity.id);
	notice = tree.notice;
	latest = useLaunchPrefill({
		...props,
		builder,
		showTreeNotice: showTreeNotice ?? tree.show,
	});
	return null;
}

const VIEW_ID = "00000000-0000-0000-0000-00000000a001";
const ACCOUNT_XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
const CONTACT_XML = "<fetch><entity name='contact'><attribute name='fullname'/></entity></fetch>";
const LAYOUT_XML =
	"<grid name='resultset' object='1' jump='name' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='name' width='321' /></row></grid>";

const activeAccounts: FakeView = {
	id: VIEW_ID,
	name: "Active accounts",
	entity: "account",
	fetchxml: ACCOUNT_XML,
	layoutxml: LAYOUT_XML,
	type: "system",
};

let host: HostStub;
let dataverse: ReturnType<typeof createFakeDataverse>;

/** The in-memory Dataverse behind a stubbed host */
function connect(views: FakeView[] = [activeAccounts]) {
	resetDataverseSessionCache();
	clearMetadataCache();
	metadataCache.clear();
	dataverse = createFakeDataverse(sampleTables(3), { views });
	host = installHost({ dataverseAPI: dataverse.api as unknown as HostStub["dataverseAPI"] });
}

const standalone: LaunchContextState = { loading: false, isCallee: false, context: null };
const launchedWith = (context: Record<string, unknown>): LaunchContextState => ({
	loading: false,
	isCallee: true,
	context,
});

let props: ProbeProps;
/** Renders the hook in a builder; later calls change the given inputs and keep the rest */
const show = (changes: Partial<ProbeProps> = {}) => {
	props = { ...props, ...changes };
	return render(createElement(BuilderProvider, null, createElement(Probe, props)));
};

const fresh = (changes: Partial<ProbeProps> = {}) => {
	props = {
		carryOver: { prefillConsumed: false },
		launch: standalone,
		onViewApplied: vi.fn(),
		...changes,
	};
	return show();
};

const rootEntity = () => builder.fetchQuery?.entity.name;

describe("the query carried over a connection change (#66)", () => {
	it("loads it once on mount and clears it", async () => {
		connect();
		const carryOver: ConnectionCarryOver = { prefillConsumed: true, fetchXml: CONTACT_XML };
		await fresh({ carryOver });
		expect(rootEntity()).toBe("contact");
		expect(carryOver.fetchXml).toBeUndefined();
	});

	it("logs a query it can't parse and leaves the builder empty", async () => {
		connect();
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		await fresh({ carryOver: { prefillConsumed: true, fetchXml: "<fetch><entity" } });
		expect(builder.fetchQuery).toBeNull();
		expect(error).toHaveBeenCalledWith(
			"Connection change: couldn't reload the query:",
			expect.anything(),
		);
	});

	it("does nothing without one", async () => {
		connect();
		await fresh({ carryOver: { prefillConsumed: true } });
		expect(builder.fetchQuery).toBeNull();
	});

	it("lets AppShell read the current query, and follows it as it changes", async () => {
		connect();
		const carryOver: ConnectionCarryOver = { prefillConsumed: true };
		await fresh({ carryOver });
		expect(carryOver.readFetchXml?.()).toBeUndefined();
		await act(async () => builder.loadFetchXml(ACCOUNT_XML));
		expect(carryOver.readFetchXml?.()).toBe(generateFetchXml(builder.fetchQuery!));
		await act(async () => builder.loadFetchXml(CONTACT_XML));
		expect(carryOver.readFetchXml?.()).toContain('name="contact"');
	});
});

describe("the tool-to-tool prefill (#197)", () => {
	it("waits while the launch context loads, then applies it once", async () => {
		connect();
		const carryOver: ConnectionCarryOver = { prefillConsumed: false };
		await fresh({ carryOver, launch: { loading: true, isCallee: false, context: null } });
		expect(carryOver.prefillConsumed).toBe(false);
		expect(builder.fetchQuery).toBeNull();

		await show({ launch: launchedWith({ entityLogicalName: "account" }) });
		expect(carryOver.prefillConsumed).toBe(true);
		expect(rootEntity()).toBe("account");

		// A new query, then a re-render with a new launch object: not applied again
		await act(async () => builder.loadFetchXml(CONTACT_XML));
		await show({ launch: launchedWith({ entityLogicalName: "account" }) });
		expect(rootEntity()).toBe("contact");
	});

	it("calls the current onViewApplied when a view arrives after later renders", async () => {
		connect();
		// Hold the view read open, so the prefill finishes after a render with a new callback
		const gate = deferred<void>();
		const queryData = host.dataverseAPI.queryData;
		host.dataverseAPI.queryData = async (query: unknown) => {
			if (String(query).startsWith("savedqueries")) await gate.promise;
			return queryData(query);
		};
		const first = vi.fn();
		const current = vi.fn();
		await fresh({
			onViewApplied: first,
			launch: launchedWith({ viewRef: { entityLogicalName: "savedquery", id: VIEW_ID } }),
		});
		await show({ onViewApplied: current });
		expect(builder.fetchQuery).toBeNull();
		gate.resolve();
		await flush(6);
		expect(rootEntity()).toBe("account");
		expect(first).not.toHaveBeenCalled();
		expect(current).toHaveBeenCalledTimes(1);
	});

	it("isn't applied again after a connection change (the carry-over outlives the remount)", async () => {
		connect();
		const carryOver: ConnectionCarryOver = { prefillConsumed: false };
		await fresh({ carryOver, launch: launchedWith({ entityLogicalName: "account" }) });
		expect(rootEntity()).toBe("account");
		// AppShell remounts the builder and AppContent for the new connection, with the same carry-over
		cleanup();
		await show();
		expect(builder.fetchQuery).toBeNull();
	});

	it("marks a standalone launch as consumed and does nothing", async () => {
		connect();
		const carryOver: ConnectionCarryOver = { prefillConsumed: false };
		await fresh({ carryOver });
		expect(carryOver.prefillConsumed).toBe(true);
		expect(builder.fetchQuery).toBeNull();
		expect(host.notifications).toEqual([]);
	});

	it("starts a query on the entity it's given", async () => {
		connect();
		await fresh({ launch: launchedWith({ entityLogicalName: " contact " }) });
		expect(rootEntity()).toBe("contact");
	});

	it("loads the FetchXML it's given, which wins over an entity hint", async () => {
		connect();
		await fresh({
			launch: launchedWith({ fetchXml: CONTACT_XML, entityLogicalName: "account" }),
		});
		expect(rootEntity()).toBe("contact");
		expect(builder.loadedView).toBeNull();
		expect(host.notifications).toEqual([]);
	});

	it("says when the FetchXML it's given can't be parsed: an error bar above the tree", async () => {
		connect();
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		await fresh({ launch: launchedWith({ fetchXml: "<fetch><entity" }) });
		expect(builder.fetchQuery).toBeNull();
		expect(error).toHaveBeenCalledWith(
			"T2T prefill: failed to parse incoming FetchXML:",
			expect.anything(),
		);
		expect(notice).toEqual({
			intent: "error",
			title: "Couldn't load the incoming query",
			body: "The tool that launched FetchXML Studio sent a query that couldn't be parsed.",
		});
		// In the app, not a host notification (#277, owner decision 2)
		expect(host.notifications).toEqual([]);

		// It stays until a query starts
		await act(async () => builder.setEntity("account"));
		expect(notice).toBeNull();
	});

	it("logs a prefill that fails, without throwing", async () => {
		connect();
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		await fresh({
			launch: launchedWith({ fetchXml: "<fetch><entity" }),
			showTreeNotice: () => {
				throw new Error("Notices unavailable");
			},
		});
		await flush();
		expect(error).toHaveBeenCalledWith(
			"T2T prefill: failed to load launch context:",
			expect.objectContaining({ message: "Notices unavailable" }),
		);
	});

	it("opens a view by viewRef as Load view does: entity set, layout, execution by ID, overwrite", async () => {
		connect();
		const onViewApplied = vi.fn();
		await fresh({
			onViewApplied,
			launch: launchedWith({ viewRef: { entityLogicalName: "savedquery", id: VIEW_ID } }),
		});
		await flush();
		expect(rootEntity()).toBe("account");
		expect(builder.loadedView).toMatchObject({
			id: VIEW_ID,
			type: "system",
			name: "Active accounts",
			entitySetName: "accounts",
			queryType: 0,
			isCustomizable: true,
			isManaged: false,
		});
		// The layout's width, not the default the FetchXML alone would give
		expect(builder.columnConfig?.columns.map((c) => [c.name, c.width])).toEqual([
			["name", 321],
		]);
		expect(onViewApplied).toHaveBeenCalledTimes(1);
		expect(host.notifications).toEqual([]);
	});

	it("says when the view doesn't exist", async () => {
		connect([]);
		const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
		await fresh({
			launch: launchedWith({ viewRef: { entityLogicalName: "userquery", id: VIEW_ID } }),
		});
		await flush();
		expect(builder.fetchQuery).toBeNull();
		expect(notice).toMatchObject({ intent: "error", title: "Couldn't open the view" });
		expect(notice?.body).toContain("personal view");
		expect(notice?.body).toContain(VIEW_ID);
		expect(host.notifications).toEqual([]);
		expect(warn).toHaveBeenCalledWith(
			"T2T prefill: the view couldn't be opened:",
			expect.objectContaining({ kind: "viewUnavailable", id: VIEW_ID }),
		);
	});

	it("a view whose FetchXML can't be parsed is a launch failure: the error bar, not a toast", async () => {
		connect([{ ...activeAccounts, fetchxml: "<fetch><entity name='account'>" }]);
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		await fresh({
			launch: launchedWith({ viewRef: { entityLogicalName: "savedquery", id: VIEW_ID } }),
		});
		await flush();
		expect(builder.fetchQuery).toBeNull();
		expect(notice).toMatchObject({ intent: "error", title: "Couldn't load the view" });
		expect(notice?.body).toContain('"Active accounts"');
		expect(toasts()).toEqual([]);
		expect(host.notifications).toEqual([]);
		expect(error).toHaveBeenCalledWith("Failed to parse view FetchXML:", expect.anything());
	});

	it("says when the view's entity has no entity set name", async () => {
		const tables = sampleTables(3).map((t) =>
			t.logicalName === "account" ? { ...t, entitySetName: "" } : t,
		);
		resetDataverseSessionCache();
		clearMetadataCache();
		metadataCache.clear();
		dataverse = createFakeDataverse(tables, { views: [activeAccounts] });
		host = installHost({ dataverseAPI: dataverse.api as unknown as HostStub["dataverseAPI"] });
		const onViewApplied = vi.fn();
		await fresh({
			onViewApplied,
			launch: launchedWith({ viewRef: { entityLogicalName: "savedquery", id: VIEW_ID } }),
		});
		await flush();
		expect(builder.fetchQuery).toBeNull();
		expect(onViewApplied).not.toHaveBeenCalled();
		expect(notice?.title).toBe("Couldn't open the view");
		expect(notice?.body).toContain("account has no entity set name in metadata.");
		expect(host.notifications).toEqual([]);
	});

	it("says when reading the view's entity fails", async () => {
		connect();
		host.dataverseAPI.getEntityMetadata = async () => {
			throw new Error("Metadata unavailable");
		};
		await fresh({
			launch: launchedWith({ viewRef: { entityLogicalName: "savedquery", id: VIEW_ID } }),
		});
		await flush();
		expect(builder.fetchQuery).toBeNull();
		expect(notice?.body).toContain("Metadata unavailable");
		expect(host.notifications).toEqual([]);
	});
});

describe("applyLoadedView (#93, #142, #197)", () => {
	const viewInfo = (changes: Partial<LoadedViewInfo> = {}): LoadedViewInfo => ({
		id: VIEW_ID,
		type: "personal",
		originalFetchXml: ACCOUNT_XML,
		entitySetName: "accounts",
		name: "My accounts",
		layoutxml: LAYOUT_XML,
		queryType: 0,
		isCustomizable: true,
		description: "Mine",
		isDefault: false,
		isManaged: false,
		solutionId: "sol-1",
		ownerId: "user-1",
		ownerName: "Test User",
		...changes,
	});

	it("loads the view with everything Save view needs, and resets the quick find value", async () => {
		connect();
		const onViewApplied = vi.fn();
		await fresh({ onViewApplied });
		await act(async () => latest.applyLoadedView(viewInfo()));
		expect(rootEntity()).toBe("account");
		expect(builder.loadedView).toMatchObject({
			id: VIEW_ID,
			type: "personal",
			entitySetName: "accounts",
			name: "My accounts",
			queryType: 0,
			isCustomizable: true,
			description: "Mine",
			isDefault: false,
			isManaged: false,
			solutionId: "sol-1",
			ownerId: "user-1",
			ownerName: "Test User",
		});
		// The layout's width, not the default the FetchXML alone would give
		expect(builder.columnConfig?.columns.map((c) => [c.name, c.width])).toEqual([
			["name", 321],
		]);
		expect(onViewApplied).toHaveBeenCalledTimes(1);
		expect(host.notifications).toEqual([]);
	});

	it("says when the tree changed the view's query: a warning bar above the tree, until the next load", async () => {
		connect();
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		await fresh();
		await act(async () =>
			latest.applyLoadedView(
				viewInfo({
					originalFetchXml: `<fetch><entity name="account"><attribute name="name" alias="1bad" /></entity></fetch>`,
				}),
			),
		);
		expect(notice).toMatchObject({ intent: "warning", title: "View loaded with changes" });
		expect(notice?.body).toContain('"My accounts"');
		expect(host.notifications).toEqual([]);
		expect(error).not.toHaveBeenCalled();

		// An edit keeps it; the next load clears it
		await act(async () => builder.addAttributeByName(builder.fetchQuery!.entity.id, "revenue"));
		expect(notice?.title).toBe("View loaded with changes");
		await act(async () => latest.applyLoadedView(viewInfo()));
		expect(notice).toBeNull();
	});

	it("says when the view's FetchXML can't be parsed (a sticky error toast), and logs it", async () => {
		connect();
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		const onViewApplied = vi.fn();
		await fresh({ onViewApplied });
		await act(async () =>
			latest.applyLoadedView(
				viewInfo({ originalFetchXml: "<fetch><entity name='account'>" }),
			),
		);
		expect(toasts()).toEqual([expect.stringMatching(/^Couldn't load the view\. Dismiss/)]);
		expect(notice).toBeNull();
		expect(host.notifications).toEqual([]);
		expect(error).toHaveBeenCalledWith("Failed to parse view FetchXML:", expect.anything());
		// The quick find value belonged to the previous view either way
		expect(onViewApplied).toHaveBeenCalledTimes(1);
	});
});
