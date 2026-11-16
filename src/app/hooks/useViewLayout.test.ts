/**
 * useViewLayout (#301), with the real builder: the column config follows the query (widths from the
 * root entity's attribute types, kept widths and order when the query changes, nothing until the
 * primary key is known), and the LayoutXML Save view and the export use (#93).
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { BuilderProvider } from "../../features/fetchxml/state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import type { AttributeMetadata, EntityMetadata } from "../../features/fetchxml/api/pptbClient";
import { useViewLayout, type ViewLayout, type ViewLayoutInputs } from "./useViewLayout";

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

let builder: BuilderContextValue;
let latest: ViewLayout;

type ProbeProps = Omit<ViewLayoutInputs, "builder"> & {
	/** Replaces the builder's syncLayoutWithFetchXml, to see how the hook calls it */
	sync?: BuilderContextValue["syncLayoutWithFetchXml"];
	/** Builder state the hook sees instead of the real one */
	state?: Partial<Pick<BuilderContextValue, "columnConfig" | "layoutNeedsSync">>;
};
function Probe({ sync, state, ...props }: ProbeProps) {
	builder = useBuilder();
	latest = useViewLayout({
		...props,
		builder: {
			...builder,
			...state,
			...(sync ? { syncLayoutWithFetchXml: sync } : {}),
		},
	});
	return null;
}

const attr = (LogicalName: string, AttributeType?: string) =>
	({ LogicalName, AttributeType }) as unknown as AttributeMetadata;
const attrs = (...list: AttributeMetadata[]) => new Map(list.map((a) => [a.LogicalName, a]));

const ACCOUNT = {
	LogicalName: "account",
	ObjectTypeCode: 1,
	PrimaryIdAttribute: "accountid",
	PrimaryNameAttribute: "name",
} as unknown as EntityMetadata;

const ACCOUNT_ATTRS = new Map([
	[
		"account",
		attrs(attr("name", "String"), attr("revenue", "Money"), attr("description", "Memo")),
	],
	["contact", attrs(attr("fullname", "String"))],
]);

const ACCOUNT_XML =
	"<fetch><entity name='account'><attribute name='name'/><attribute name='revenue'/></entity></fetch>";

let props: ProbeProps;
/** Renders the hook in a builder; later calls change the given inputs and keep the rest */
const show = (changes: Partial<ProbeProps> = {}) => {
	props = { ...props, ...changes };
	return render(createElement(BuilderProvider, null, createElement(Probe, props)));
};
const fresh = (changes: Partial<ProbeProps> = {}) => {
	props = { attributeMetadata: ACCOUNT_ATTRS, entityMetadata: ACCOUNT };
	return show(changes);
};
const act_ = (fn: () => void) => act(async () => fn());

/** A query built as the Entity picker and Add attribute do: the layout is synced, not parsed */
const build = async (...names: string[]) => {
	await act_(() => builder.setEntity("account"));
	for (const name of names)
		await act_(() => builder.addAttributeByName(builder.fetchQuery!.entity.id, name));
};

const widths = () =>
	Object.fromEntries((builder.columnConfig?.columns ?? []).map((c) => [c.name, c.width]));

describe("the column config follows the query", () => {
	it("is generated with widths from the root entity's attribute types", async () => {
		await fresh();
		await build("name", "revenue");
		expect(widths()).toEqual({ name: 200, revenue: 130 });
		expect(builder.columnConfig?.primaryIdAttribute).toBe("accountid");
		expect(builder.layoutNeedsSync).toBe(false);
	});

	it("waits for the primary key: nothing without the root entity's metadata, then generated", async () => {
		await fresh({ entityMetadata: null });
		await build("name", "revenue");
		expect(builder.columnConfig).toBeNull();
		expect(builder.layoutNeedsSync).toBe(true);

		await show({ entityMetadata: ACCOUNT });
		expect(widths()).toEqual({ name: 200, revenue: 130 });
		expect(builder.columnConfig?.primaryIdAttribute).toBe("accountid");
	});

	it("uses the default width for an attribute the root entity's metadata doesn't type", async () => {
		await fresh({ attributeMetadata: new Map() });
		await build("name", "revenue");
		expect(widths()).toEqual({ name: 150, revenue: 150 });
	});

	// The root-only type map itself is pinned under "when it syncs"; this is what the user sees
	it("gives a link-entity's column the default width", async () => {
		await fresh();
		await build("description");
		let link = "";
		await act_(() => {
			link = builder.addLinkEntityWithConfig(
				builder.fetchQuery!.entity.id,
				"contact",
				"contactid",
				"primarycontactid",
				"outer",
				"N1",
			);
		});
		await act_(() => builder.addAttributeByName(link, "fullname"));
		const columns = builder.columnConfig!.columns;
		expect(columns.map((c) => c.width)).toEqual([250, 150]);
		expect(columns[1].name).toMatch(/\.fullname$/);
	});

	it("keeps a view's widths and adds a new column at the end with its type's width", async () => {
		await fresh();
		await act_(() =>
			builder.loadView(
				ACCOUNT_XML,
				{ id: "v1", type: "system", entitySetName: "accounts", name: "Accounts" },
				"<grid name='resultset' object='1' jump='name' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='revenue' width='321' /><cell name='name' width='99' /></row></grid>",
			),
		);
		expect(widths()).toEqual({ revenue: 321, name: 99 });

		await act_(() => builder.addAttributeByName(builder.fetchQuery!.entity.id, "description"));
		expect(builder.columnConfig?.columns.map((c) => c.name)).toEqual([
			"revenue",
			"name",
			"description",
		]);
		expect(widths()).toEqual({ revenue: 321, name: 99, description: 250 });
		expect(builder.layoutNeedsSync).toBe(false);
	});

	it("leaves a loaded query's layout as the parser made it: nothing to sync", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		await act_(() => builder.loadFetchXml(ACCOUNT_XML));
		expect(sync).not.toHaveBeenCalled();
		expect(builder.layoutNeedsSync).toBe(false);
	});

	it("drops a removed attribute's column", async () => {
		await fresh();
		await build("name", "revenue");
		const revenue = builder.fetchQuery!.entity.attributes.find((a) => a.name === "revenue")!;
		await act_(() => builder.removeNode(revenue.id));
		expect(widths()).toEqual({ name: 200 });
	});
});

describe("when it syncs", () => {
	it("passes the root entity's attribute types (empty for an untyped one) and the primary key", async () => {
		const sync = vi.fn();
		const untyped = new Map([
			["account", attrs(attr("name", "String"), attr("x"))],
			["contact", attrs(attr("fullname", "String"), attr("name", "Memo"))],
		]);
		await fresh({ sync, attributeMetadata: untyped });
		await build("name");
		expect(sync).toHaveBeenCalled();
		const [map, primaryId] = sync.mock.calls.at(-1)!;
		expect([...map]).toEqual([
			["name", "String"],
			["x", ""],
		]);
		expect(primaryId).toBe("accountid");
	});

	it("passes no primary key without the root entity's metadata", async () => {
		const sync = vi.fn();
		await fresh({ sync, entityMetadata: null });
		await build("name");
		expect(sync.mock.calls.at(-1)![1]).toBeUndefined();
	});

	it("doesn't sync a query with no attributes and nothing to sync, or no query", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		expect(sync).not.toHaveBeenCalled();
		await act_(() => builder.setEntity("account"));
		expect(builder.fetchQuery?.entity.attributes).toEqual([]);
		expect(builder.layoutNeedsSync).toBe(false);
		expect(sync).not.toHaveBeenCalled();
	});

	it("syncs a query with attributes and no column config, even with nothing marked to sync", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		await act_(() => builder.loadFetchXml(ACCOUNT_XML));
		expect(sync).not.toHaveBeenCalled();
		await show({ state: { columnConfig: null, layoutNeedsSync: false } });
		expect(sync).toHaveBeenCalledTimes(1);
	});

	it("syncs when only the sync flag changes (an edit in the FetchXML editor)", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		await act_(() => builder.loadFetchXml(ACCOUNT_XML));
		await act_(() => builder.clearLoadedView());
		expect(builder.layoutNeedsSync).toBe(true);
		expect(sync).toHaveBeenCalledTimes(1);
	});

	it("syncs again when the query changes while a sync is still due", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		await build("name");
		const calls = sync.mock.calls.length;
		await act_(() => builder.addAttributeByName(builder.fetchQuery!.entity.id, "revenue"));
		expect(sync.mock.calls.length).toBe(calls + 1);
	});

	it("doesn't sync on a metadata change when no sync is due", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		await act_(() => builder.loadFetchXml(ACCOUNT_XML));
		await show({ entityMetadata: { ...ACCOUNT } as EntityMetadata });
		await show({ attributeMetadata: new Map(ACCOUNT_ATTRS) });
		expect(sync).not.toHaveBeenCalled();
	});

	it("syncs again when the root entity's metadata changes", async () => {
		const sync = vi.fn();
		await fresh({ sync });
		await build("name");
		const calls = sync.mock.calls.length;
		await show({ entityMetadata: { ...ACCOUNT } as EntityMetadata });
		expect(sync.mock.calls.length).toBe(calls + 1);
		await show({ attributeMetadata: new Map(ACCOUNT_ATTRS) });
		expect(sync.mock.calls.length).toBe(calls + 2);
	});
});

describe("layoutXml", () => {
	it("is empty without a query", async () => {
		await fresh();
		expect(latest.layoutXml).toBe("");
	});

	it("is empty without the root entity's metadata", async () => {
		await fresh({ entityMetadata: null });
		await act_(() => builder.loadFetchXml(ACCOUNT_XML));
		expect(latest.layoutXml).toBe("");
	});

	it("takes object, the row id and jump from the root entity's metadata", async () => {
		await fresh();
		await build("name", "revenue");
		expect(latest.layoutXml).toContain('object="1"');
		expect(latest.layoutXml).toContain('id="accountid"');
		expect(latest.layoutXml).toContain('jump="name"');
		expect(latest.layoutXml).toMatch(/<cell name="name" width="200"[^>]*\/>/);
		expect(latest.layoutXml).toMatch(/<cell name="revenue" width="130"[^>]*\/>/);
	});

	it("keeps a loaded view's jump and follows width changes", async () => {
		await fresh();
		await act_(() =>
			builder.loadView(
				ACCOUNT_XML,
				{ id: "v1", type: "system", entitySetName: "accounts", name: "Accounts" },
				"<grid name='resultset' object='1' jump='revenue' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='name' width='99' /><cell name='revenue' width='321' /></row></grid>",
			),
		);
		expect(latest.layoutXml).toContain('jump="revenue"');
		await act_(() => builder.updateColumnWidth("name", 222));
		expect(latest.layoutXml).toMatch(/<cell name="name" width="222"[^>]*\/>/);
	});

	it("follows the root entity's metadata when only it changes", async () => {
		await fresh();
		await build("name");
		expect(latest.layoutXml).toContain('object="1"');
		await show({ entityMetadata: { ...ACCOUNT, ObjectTypeCode: 2 } as EntityMetadata });
		expect(latest.layoutXml).toContain('object="2"');
	});

	it("uses the entity's object type code, primary key and name, not the query's", async () => {
		await fresh({
			entityMetadata: {
				...ACCOUNT,
				ObjectTypeCode: 10042,
				PrimaryIdAttribute: "new_thingid",
				PrimaryNameAttribute: "new_label",
			} as EntityMetadata,
		});
		await act_(() => builder.loadFetchXml(ACCOUNT_XML));
		expect(latest.layoutXml).toContain('object="10042"');
		expect(latest.layoutXml).toContain('id="new_thingid"');
		expect(latest.layoutXml).toContain('jump="new_label"');
	});
});
