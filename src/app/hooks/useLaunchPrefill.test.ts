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
	flush,
	installHost,
	removeHost,
	render,
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

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

let builder: BuilderContextValue;
let latest: LaunchPrefill;

type ProbeProps = Omit<LaunchPrefillInputs, "builder">;
function Probe(props: ProbeProps) {
	builder = useBuilder();
	latest = useLaunchPrefill({ ...props, builder });
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
		const first = vi.fn();
		const current = vi.fn();
		await fresh({
			onViewApplied: first,
			launch: { loading: true, isCallee: false, context: null },
		});
		await show({
			onViewApplied: current,
			launch: launchedWith({ viewRef: { entityLogicalName: "savedquery", id: VIEW_ID } }),
		});
		await flush();
		expect(rootEntity()).toBe("account");
		expect(first).not.toHaveBeenCalled();
		expect(current).toHaveBeenCalledTimes(1);
	});

	it("isn't applied again after a connection change (the carry-over outlives the remount)", async () => {
		connect();
		const carryOver: ConnectionCarryOver = { prefillConsumed: true };
		await fresh({ carryOver, launch: launchedWith({ entityLogicalName: "account" }) });
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

	it("says when the FetchXML it's given can't be parsed", async () => {
		connect();
		vi.spyOn(console, "error").mockImplementation(() => {});
		await fresh({ launch: launchedWith({ fetchXml: "<fetch><entity" }) });
		expect(builder.fetchQuery).toBeNull();
		expect(host.notifications).toEqual([
			{
				title: "Could not load incoming query",
				body: "The tool that launched FetchXML Studio sent a query that couldn't be parsed.",
				type: "error",
			},
		]);
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
		vi.spyOn(console, "warn").mockImplementation(() => {});
		await fresh({
			launch: launchedWith({ viewRef: { entityLogicalName: "userquery", id: VIEW_ID } }),
		});
		await flush();
		expect(builder.fetchQuery).toBeNull();
		expect(host.notifications).toHaveLength(1);
		expect(host.notifications[0]).toMatchObject({
			title: "Couldn't open the view",
			type: "error",
		});
		expect(host.notifications[0].body).toContain("personal view");
		expect(host.notifications[0].body).toContain(VIEW_ID);
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
		expect(host.notifications).toHaveLength(1);
		expect(host.notifications[0].title).toBe("Couldn't open the view");
		expect(host.notifications[0].body).toContain("account has no entity set name in metadata.");
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
		expect(host.notifications).toHaveLength(1);
		expect(host.notifications[0].body).toContain("Metadata unavailable");
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

	it("says when the tree changed the view's query", async () => {
		connect();
		await fresh();
		await act(async () =>
			latest.applyLoadedView(
				viewInfo({
					originalFetchXml: `<fetch><entity name="account"><attribute name="name" alias="1bad" /></entity></fetch>`,
				}),
			),
		);
		expect(host.notifications).toHaveLength(1);
		expect(host.notifications[0]).toMatchObject({
			title: "View loaded with changes",
			type: "warning",
		});
	});

	it("says when the view's FetchXML can't be parsed, and logs it", async () => {
		connect();
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		const onViewApplied = vi.fn();
		await fresh({ onViewApplied });
		await act(async () =>
			latest.applyLoadedView(
				viewInfo({ originalFetchXml: "<fetch><entity name='account'>" }),
			),
		);
		expect(host.notifications).toHaveLength(1);
		expect(host.notifications[0]).toMatchObject({
			title: "Could not load view",
			type: "error",
		});
		expect(error).toHaveBeenCalledWith("Failed to parse view FetchXML:", expect.anything());
		// The quick find value belonged to the previous view either way
		expect(onViewApplied).toHaveBeenCalledTimes(1);
	});
});
