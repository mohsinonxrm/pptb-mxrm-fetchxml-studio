/**
 * useResultsGridEdits (#301), with the real builder: column filters that run the query again once
 * the change commits (#23), reordering, removing and adding columns, related columns through
 * link-entities, and sorting by a column (#114).
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { BuilderProvider } from "../../features/fetchxml/state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../features/fetchxml/state/builderContext";
import type {
	AttributeMetadata,
	RelationshipMetadata,
} from "../../features/fetchxml/api/pptbClient";
import type { RelatedEntityColumn } from "../../features/fetchxml/ui/RightPane/AddColumnsPanel";
import { generateFetchXml } from "../../features/fetchxml/model/fetchxml";
import type { LinkEntityNode } from "../../features/fetchxml/model/nodes";
import {
	useResultsGridEdits,
	type ResultsGridEdits,
	type ResultsGridEditsInputs,
} from "./useResultsGridEdits";

let notify: Mock<(options: unknown) => Promise<void>>;
beforeEach(() => {
	notify = vi.fn(async () => {});
	(window as unknown as { toolboxAPI: unknown }).toolboxAPI = {
		utils: { showNotification: notify },
	};
});
afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
	delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
});

let builder: BuilderContextValue;
let edits: ResultsGridEdits;

type ProbeProps = Omit<ResultsGridEditsInputs, "builder"> & {
	/** Builder state and actions the hook sees instead of the real ones */
	override?: Partial<BuilderContextValue>;
};
function Probe({ override, ...props }: ProbeProps) {
	builder = useBuilder();
	edits = useResultsGridEdits({ ...props, builder: { ...builder, ...override } });
	return null;
}

const attr = (LogicalName: string, AttributeType: string, typeName?: string) =>
	({
		LogicalName,
		AttributeType,
		...(typeName ? { AttributeTypeName: { Value: typeName } } : {}),
	}) as unknown as AttributeMetadata;
const attrs = (...list: AttributeMetadata[]) => new Map(list.map((a) => [a.LogicalName, a]));

const METADATA = new Map([
	[
		"account",
		attrs(
			attr("name", "String"),
			attr("revenue", "Money"),
			attr("categories", "Virtual", "MultiSelectPicklistType"),
		),
	],
	[
		"contact",
		attrs(attr("fullname", "String"), attr("hobbies", "Virtual", "MultiSelectPicklistType")),
	],
]);

let run: Mock<() => Promise<void>>;
/** The query each run saw: running waits for the change to commit */
let ranWith: string[];
let props: ProbeProps;
/** Renders the hook in a builder; later calls change the given inputs and keep the rest */
const show = (changes: Partial<ProbeProps> = {}) => {
	props = { ...props, ...changes };
	return render(createElement(BuilderProvider, null, createElement(Probe, props)));
};
const fresh = (changes: Partial<ProbeProps> = {}) => {
	ranWith = [];
	run = vi.fn(async () => {
		ranWith.push(xml());
	});
	props = { attributeMetadata: METADATA, handleExecuteRef: { current: run } };
	return show(changes);
};
const act_ = (fn: () => void) => act(async () => fn());
const load = (fetchXml: string) => act_(() => builder.loadFetchXml(fetchXml));
const xml = () => (builder.fetchQuery ? generateFetchXml(builder.fetchQuery) : "");

const ACCOUNT_XML =
	"<fetch><entity name='account'><attribute name='name'/><attribute name='revenue'/></entity></fetch>";
const WITH_CONTACT =
	"<fetch><entity name='account'><attribute name='name'/>" +
	"<link-entity name='contact' from='contactid' to='primarycontactid' link-type='outer' alias='pc'>" +
	"<attribute name='fullname'/><attribute name='hobbies'/></link-entity></entity></fetch>";

const nameFilter = { attribute: "name", operator: "like", value: "%a%" } as const;
const revenueFilter = { attribute: "revenue", operator: "gt", value: 10 } as const;

describe("column filters", () => {
	it("lists none and doesn't run on its own", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		expect(edits.columnFilters).toEqual([]);
		expect(run).not.toHaveBeenCalled();
	});

	it("adds a condition and runs the query once, with the condition in it", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() => edits.addColumnFilter(nameFilter));
		expect(edits.columnFilters).toEqual([expect.objectContaining(nameFilter)]);
		expect(run).toHaveBeenCalledTimes(1);
		expect(run).toHaveBeenCalledWith();
		expect(ranWith[0]).toContain('attribute="name" operator="like" value="%a%"');
	});

	it("runs again for each change", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() => edits.addColumnFilter(nameFilter));
		await act_(() => edits.addColumnFilter(revenueFilter));
		expect(edits.columnFilters).toHaveLength(2);
		expect(run).toHaveBeenCalledTimes(2);
		expect(ranWith[1]).toContain('attribute="revenue" operator="gt" value="10"');
	});

	it("does nothing when asked to remove none", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() => edits.addColumnFilter(nameFilter));
		await act_(() => edits.removeColumnFilters([]));
		expect(edits.columnFilters).toHaveLength(1);
		expect(run).toHaveBeenCalledTimes(1);
	});

	it("removes some of them one by one, then runs the query", async () => {
		const clear = vi.fn();
		const remove = vi.fn();
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() => edits.addColumnFilter(nameFilter));
		await act_(() => edits.addColumnFilter(revenueFilter));
		const [first, second] = edits.columnFilters;

		await show({ override: { clearColumnFilters: clear, removeColumnFilter: remove } });
		await act_(() => edits.removeColumnFilters([second.id]));
		expect(remove.mock.calls).toEqual([[second.id]]);
		expect(clear).not.toHaveBeenCalled();
		expect(run).toHaveBeenCalledTimes(3);

		await show({ override: undefined });
		await act_(() => edits.removeColumnFilters([second.id]));
		expect(edits.columnFilters.map((c) => c.id)).toEqual([first.id]);
		expect(ranWith[3]).not.toContain('attribute="revenue"');
		expect(ranWith[3]).toContain('attribute="name"');
	});

	it("clears them all at once when all are removed, then runs the query", async () => {
		// Stable from the first render, so the hook keeps its own callbacks: they see the current count
		const clear = vi.fn(() => builder.clearColumnFilters());
		const remove = vi.fn((id: string) => builder.removeColumnFilter(id));
		await fresh({ override: { clearColumnFilters: clear, removeColumnFilter: remove } });
		await load(ACCOUNT_XML);
		await act_(() => edits.addColumnFilter(nameFilter));
		await act_(() => edits.addColumnFilter(revenueFilter));
		await act_(() => edits.removeColumnFilters(edits.columnFilters.map((c) => c.id)));
		expect(clear).toHaveBeenCalledTimes(1);
		expect(remove).not.toHaveBeenCalled();
		expect(edits.columnFilters).toEqual([]);
		expect(run).toHaveBeenCalledTimes(3);
		expect(ranWith[2]).not.toContain("<filter");
	});

	it("lists the conditions of the builder's column filter group", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() => edits.addColumnFilter(nameFilter));
		await show({ override: { columnFilterId: null } });
		expect(edits.columnFilters).toEqual([]);
	});
});

describe("columns", () => {
	it("reorders the column config's columns and keeps the rest of it", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		const before = builder.columnConfig!;
		await act_(() => edits.reorderColumns([...before.columns].reverse()));
		expect(builder.columnConfig!.columns.map((c) => c.name)).toEqual(["revenue", "name"]);
		expect(builder.columnConfig).toEqual({ ...before, columns: [...before.columns].reverse() });
	});

	it("doesn't reorder without a column config", async () => {
		const setColumnConfig = vi.fn();
		await fresh();
		await load(ACCOUNT_XML);
		await show({ override: { columnConfig: null, setColumnConfig } });
		await act_(() => edits.reorderColumns([]));
		expect(setColumnConfig).not.toHaveBeenCalled();
	});

	it("removes the columns' attributes, and an outer link-entity left empty", async () => {
		await fresh();
		await load(WITH_CONTACT);
		await act_(() => edits.removeColumns(["pc.fullname", "pc.hobbies", "nothere"]));
		expect(xml()).not.toContain("link-entity");
		expect(xml()).toContain('<attribute name="name"');
	});

	it("removes root columns", async () => {
		await fresh();
		await load(
			"<fetch><entity name='account'><attribute name='name'/><attribute name='revenue'/><attribute name='categories'/></entity></fetch>",
		);
		await act_(() => edits.removeColumns(["revenue", "name"]));
		expect(builder.fetchQuery!.entity.attributes.map((a) => a.name)).toEqual(["categories"]);
	});

	it("adds an attribute to the root entity", async () => {
		await fresh();
		await load(WITH_CONTACT);
		await act_(() => edits.addColumn("revenue"));
		expect(builder.fetchQuery!.entity.attributes.map((a) => a.name)).toEqual([
			"name",
			"revenue",
		]);
		expect(builder.fetchQuery!.entity.links[0].attributes).toHaveLength(2);
	});

	it("does nothing without a query", async () => {
		const removeNode = vi.fn();
		const addAttributeByName = vi.fn();
		const addLinkEntityWithConfig = vi.fn();
		await fresh({ override: { removeNode, addAttributeByName, addLinkEntityWithConfig } });
		await act_(() => {
			edits.removeColumns(["name"]);
			edits.addColumn("name");
			edits.addRelatedColumns([related(PRIMARY_CONTACT, "N1", "fullname")]);
		});
		expect(removeNode).not.toHaveBeenCalled();
		expect(addAttributeByName).not.toHaveBeenCalled();
		expect(addLinkEntityWithConfig).not.toHaveBeenCalled();
	});
});

const relationship = (r: Partial<RelationshipMetadata>) => r as RelationshipMetadata;
// account.primarycontactid → contact
const PRIMARY_CONTACT = relationship({
	SchemaName: "account_primary_contact",
	ReferencedEntity: "contact",
	ReferencedAttribute: "contactid",
	ReferencingEntity: "account",
	ReferencingAttribute: "primarycontactid",
});
// contact.parentcustomerid → account
const CONTACTS = relationship({
	SchemaName: "contact_customer_accounts",
	ReferencedEntity: "account",
	ReferencedAttribute: "accountid",
	ReferencingEntity: "contact",
	ReferencingAttribute: "parentcustomerid",
});
const related = (
	rel: RelationshipMetadata,
	relationshipType: "N1" | "1N",
	attribute: string,
): RelatedEntityColumn => ({
	relationship: rel,
	relationshipType,
	attribute: { LogicalName: attribute } as AttributeMetadata,
	displayName: attribute,
	lookupAttribute: rel.ReferencingAttribute,
});
const links = () =>
	builder.fetchQuery!.entity.links.map((l: LinkEntityNode) => ({
		name: l.name,
		from: l.from,
		to: l.to,
		linkType: l.linkType,
		relationshipType: l.relationshipType,
		attributes: l.attributes.map((a) => a.name),
	}));

describe("related columns", () => {
	it("joins a lookup's entity with an outer link-entity, from its key to the lookup", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() =>
			edits.addRelatedColumns([
				related(PRIMARY_CONTACT, "N1", "fullname"),
				related(PRIMARY_CONTACT, "N1", "emailaddress1"),
			]),
		);
		expect(links()).toEqual([
			{
				name: "contact",
				from: "contactid",
				to: "primarycontactid",
				linkType: "outer",
				relationshipType: "N1",
				attributes: ["fullname", "emailaddress1"],
			},
		]);
	});

	it("joins a one-to-many relationship's entity with an outer link-entity, from its lookup to the key", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() => edits.addRelatedColumns([related(CONTACTS, "1N", "fullname")]));
		expect(links()).toEqual([
			{
				name: "contact",
				from: "parentcustomerid",
				to: "accountid",
				linkType: "outer",
				relationshipType: "1N",
				attributes: ["fullname"],
			},
		]);
	});

	it("adds a link-entity for each relationship", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await act_(() =>
			edits.addRelatedColumns([
				related(PRIMARY_CONTACT, "N1", "fullname"),
				related(CONTACTS, "1N", "lastname"),
				related(PRIMARY_CONTACT, "N1", "telephone1"),
			]),
		);
		expect(links().map((l) => [l.from, l.attributes])).toEqual([
			["contactid", ["fullname", "telephone1"]],
			["parentcustomerid", ["lastname"]],
		]);
	});

	it("keeps an intersect entity's two joins apart (#57)", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		const side = (IntersectSide: 1 | 2, ReferencingAttribute: string) =>
			relationship({
				SchemaName: "accountleads_association",
				IntersectSide,
				ReferencedEntity: "account",
				ReferencedAttribute: "accountid",
				ReferencingEntity: "accountleads",
				ReferencingAttribute,
			});
		await act_(() =>
			edits.addRelatedColumns([
				related(side(1, "accountid"), "1N", "accountleadid"),
				related(side(2, "leadid"), "1N", "name"),
			]),
		);
		expect(links().map((l) => [l.from, l.attributes])).toEqual([
			["accountid", ["accountleadid"]],
			["leadid", ["name"]],
		]);
	});

	it("reuses the relationship's link-entity", async () => {
		await fresh();
		await load(WITH_CONTACT);
		await act_(() => edits.addRelatedColumns([related(PRIMARY_CONTACT, "N1", "telephone1")]));
		expect(links()).toEqual([
			{
				name: "contact",
				from: "contactid",
				to: "primarycontactid",
				linkType: "outer",
				attributes: ["fullname", "hobbies", "telephone1"],
			},
		]);
	});

	it("doesn't reuse a link-entity that returns no attributes", async () => {
		await fresh();
		await load(
			"<fetch><entity name='account'><attribute name='name'/>" +
				"<link-entity name='contact' from='contactid' to='primarycontactid' link-type='exists'>" +
				"<filter><condition attribute='statecode' operator='eq' value='0'/></filter></link-entity></entity></fetch>",
		);
		await act_(() => edits.addRelatedColumns([related(PRIMARY_CONTACT, "N1", "fullname")]));
		expect(links().map((l) => [l.linkType, l.attributes])).toEqual([
			["exists", []],
			["outer", ["fullname"]],
		]);
	});

	it("doesn't reuse a link-entity to another entity or through other attributes", async () => {
		await fresh();
		await load(
			"<fetch><entity name='account'><attribute name='name'/>" +
				"<link-entity name='contact' from='contactid' to='modifiedby' link-type='outer'/>" +
				"<link-entity name='systemuser' from='contactid' to='primarycontactid' link-type='outer'/>" +
				"<link-entity name='contact' from='parentcustomerid' to='primarycontactid' link-type='outer'/>" +
				"</entity></fetch>",
		);
		await act_(() => edits.addRelatedColumns([related(PRIMARY_CONTACT, "N1", "fullname")]));
		expect(links().map((l) => l.attributes)).toEqual([[], [], [], ["fullname"]]);
	});
});

describe("sorting by a column", () => {
	const sort = (columnId: string, direction: "ascending" | "descending", isMultiSort = false) =>
		act_(() => edits.sortByColumn({ columnId, direction, isMultiSort }));
	const orders = () =>
		builder.fetchQuery!.entity.orders.map((o) => ({
			attribute: o.attribute,
			entityname: o.entityname,
			descending: !!o.descending,
		}));

	it("sorts by a root column, ascending or descending", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await sort("revenue", "descending");
		expect(orders()).toEqual([
			{ attribute: "revenue", entityname: undefined, descending: true },
		]);
		await sort("revenue", "ascending");
		expect(orders()).toEqual([
			{ attribute: "revenue", entityname: undefined, descending: false },
		]);
	});

	it("adds to the sort with Shift, and replaces it without", async () => {
		const setSort = vi.fn();
		await fresh();
		await load(ACCOUNT_XML);
		await show({ override: { setSort } });
		await sort("name", "ascending", true);
		await sort("revenue", "descending", false);
		expect(setSort.mock.calls).toEqual([
			[{ spec: { attribute: "name" } }, false, true],
			[{ spec: { attribute: "revenue" } }, true, false],
		]);
	});

	it("sorts by a link-entity column through entityname", async () => {
		await fresh();
		await load(WITH_CONTACT);
		await sort("pc.fullname", "ascending");
		expect(orders()).toEqual([{ attribute: "fullname", entityname: "pc", descending: false }]);
		expect(notify).not.toHaveBeenCalled();
	});

	it("says why it can't sort by a root column Dataverse can't sort", async () => {
		await fresh();
		await load(
			"<fetch><entity name='account'><attribute name='name'/><attribute name='categories'/></entity></fetch>",
		);
		await sort("categories", "ascending");
		expect(orders()).toEqual([]);
		expect(notify).toHaveBeenCalledWith({
			title: "Can't sort by this column",
			body: expect.stringContaining("multi-select"),
			type: "warning",
		});
	});

	it("checks a link-entity column against its own entity's metadata", async () => {
		await fresh();
		await load(WITH_CONTACT);
		await sort("pc.hobbies", "ascending");
		expect(orders()).toEqual([]);
		expect(notify).toHaveBeenCalledWith(
			expect.objectContaining({ body: expect.stringContaining("multi-select") }),
		);
	});

	it("sorts a column with no metadata", async () => {
		await fresh({ attributeMetadata: new Map() });
		await load("<fetch><entity name='account'><attribute name='categories'/></entity></fetch>");
		await sort("categories", "ascending");
		expect(orders()).toEqual([
			{ attribute: "categories", entityname: undefined, descending: false },
		]);
	});

	it("checks against the metadata as it is now", async () => {
		await fresh({ attributeMetadata: new Map() });
		await load(
			"<fetch><entity name='account'><attribute name='name'/><attribute name='categories'/></entity></fetch>",
		);
		await show({ attributeMetadata: METADATA });
		await sort("categories", "ascending");
		expect(orders()).toEqual([]);
		expect(notify).toHaveBeenCalledTimes(1);
	});

	it("says why it can't sort by a column that isn't in the query", async () => {
		await fresh();
		await load(ACCOUNT_XML);
		await sort("nothere", "ascending");
		expect(orders()).toEqual([]);
		expect(notify).toHaveBeenCalledWith({
			title: "Can't sort by this column",
			body: "This column isn't an attribute of the query.",
			type: "warning",
		});
	});

	it("does nothing without a query", async () => {
		const setSort = vi.fn();
		await fresh({ override: { setSort } });
		await sort("name", "ascending");
		expect(setSort).not.toHaveBeenCalled();
		expect(notify).not.toHaveBeenCalled();
	});
});

describe("running the query", () => {
	it("runs only for column filters: the other edits leave running to the user", async () => {
		await fresh();
		await load(WITH_CONTACT);
		await act_(() => edits.reorderColumns([...builder.columnConfig!.columns].reverse()));
		await act_(() => edits.addColumn("revenue"));
		await act_(() => edits.addRelatedColumns([related(CONTACTS, "1N", "lastname")]));
		await act_(() =>
			edits.sortByColumn({ columnId: "name", direction: "ascending", isMultiSort: false }),
		);
		await act_(() => edits.removeColumns(["revenue"]));
		expect(xml()).toContain('<order attribute="name"');
		expect(run).not.toHaveBeenCalled();
	});
});
