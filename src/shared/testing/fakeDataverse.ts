/**
 * An in-memory Dataverse for app-level tests (#189, #301): answers the window.dataverseAPI calls
 * FetchXML Studio makes on start-up and when it runs a query, from a few small tables. Requests it
 * doesn't know get an empty answer and are recorded in `unknown`, so a test can see what's missing.
 *
 * Shapes follow what the Web API returns (and what the code reads); only what the app uses is here.
 * Only test files import this module.
 */

export interface FakeAttribute {
	name: string;
	/** AttributeType, e.g. "String", "Uniqueidentifier", "Money", "State", "Status", "Integer" */
	type: string;
	label?: string;
}

export interface FakeTable {
	logicalName: string;
	entitySetName: string;
	displayName: string;
	primaryId: string;
	primaryName: string;
	objectTypeCode: number;
	attributes: FakeAttribute[];
	rows: Record<string, unknown>[];
}

export interface FakeView {
	id: string;
	name: string;
	entity: string;
	fetchxml: string;
	layoutxml: string;
	type: "system" | "personal";
}

export interface FakeDataverseOptions {
	userId?: string;
	/** Privileges the user holds; "all" (the default) grants every RetrieveUserPrivilegeByPrivilegeName */
	privileges?: "all" | Set<string>;
	views?: FakeView[];
	/** Called before each FetchXML page is answered; a test can await a gate here to hold a page */
	beforePage?: (page: number, entity: string) => Promise<void> | void;
}

const label = (text: string) => ({ UserLocalizedLabel: { Label: text } });

const ATTRIBUTE_TYPE_NAMES: Record<string, string> = {
	String: "StringType",
	Uniqueidentifier: "UniqueidentifierType",
	Money: "MoneyType",
	State: "StateType",
	Status: "StatusType",
	Integer: "IntegerType",
	Lookup: "LookupType",
	DateTime: "DateTimeType",
};

/** A fetch element's attributes and its root entity, read with the same DOMParser the app uses */
function readFetch(xml: string): { entity: string; count: number; page: number; attributes: string[]; orders: { attribute: string; descending: boolean }[] } {
	const doc = new DOMParser().parseFromString(xml, "application/xml");
	const fetch = doc.documentElement;
	const entity = fetch.querySelector("entity");
	return {
		entity: entity?.getAttribute("name") ?? "",
		count: Number(fetch.getAttribute("count") ?? "5000"),
		page: Number(fetch.getAttribute("page") ?? "1"),
		attributes: [...(entity?.querySelectorAll(":scope > attribute") ?? [])].map((a) => a.getAttribute("name") ?? ""),
		orders: [...(entity?.querySelectorAll(":scope > order") ?? [])].map((o) => ({ attribute: o.getAttribute("attribute") ?? "", descending: o.getAttribute("descending") === "true" })),
	};
}

export function createFakeDataverse(tables: FakeTable[], options: FakeDataverseOptions = {}) {
	const userId = options.userId ?? "00000000-0000-0000-0000-0000000000a1";
	const calls: string[] = [];
	const unknown: string[] = [];
	const deleted: string[] = [];
	const byName = new Map(tables.map((t) => [t.logicalName, t]));
	const bySet = new Map(tables.map((t) => [t.entitySetName, t]));
	const views = options.views ?? [];

	const entityDefinition = (t: FakeTable) => ({
		LogicalName: t.logicalName,
		SchemaName: t.logicalName,
		DisplayName: label(t.displayName),
		EntitySetName: t.entitySetName,
		LogicalCollectionName: t.entitySetName,
		PrimaryIdAttribute: t.primaryId,
		PrimaryNameAttribute: t.primaryName,
		IsValidForAdvancedFind: true,
		IsCustomizable: { Value: true },
		ObjectTypeCode: t.objectTypeCode,
		MetadataId: `00000000-0000-0000-0000-${String(t.objectTypeCode).padStart(12, "0")}`,
	});

	const attributeDefinition = (t: FakeTable, a: FakeAttribute) => ({
		LogicalName: a.name,
		SchemaName: a.name,
		EntityLogicalName: t.logicalName,
		DisplayName: label(a.label ?? a.name),
		AttributeType: a.type,
		AttributeTypeName: { Value: ATTRIBUTE_TYPE_NAMES[a.type] ?? `${a.type}Type` },
		IsValidForAdvancedFind: { Value: true },
		IsValidForRead: true,
		IsPrimaryId: a.name === t.primaryId,
		IsPrimaryName: a.name === t.primaryName,
		MetadataId: `${t.logicalName}-${a.name}`,
	});

	const stateOptions = {
		LogicalName: "statecode",
		OptionSet: {
			Options: [
				{ Value: 0, Label: label("Active"), InvariantName: "Active", DefaultStatus: 1 },
				{ Value: 1, Label: label("Inactive"), InvariantName: "Inactive", DefaultStatus: 2 },
			],
		},
	};
	const statusOptions = {
		LogicalName: "statuscode",
		OptionSet: {
			Options: [
				{ Value: 1, Label: label("Active"), State: 0, TransitionData: null },
				{ Value: 2, Label: label("Inactive"), State: 1, TransitionData: null },
			],
		},
	};

	/** One page of a FetchXML query: page-number paging, as Dataverse does without a paging cookie */
	async function runFetch(set: string, xml: string) {
		const table = bySet.get(set);
		const fetch = readFetch(xml);
		if (!table || table.logicalName !== fetch.entity) throw new Error(`0x80060888: Resource not found for the segment '${set}'.`);
		await options.beforePage?.(fetch.page, fetch.entity);
		let rows = [...table.rows];
		for (const order of [...fetch.orders].reverse()) {
			rows.sort((a, b) => String(a[order.attribute] ?? "").localeCompare(String(b[order.attribute] ?? "")) * (order.descending ? -1 : 1));
		}
		const start = (fetch.page - 1) * fetch.count;
		const pageRows = rows.slice(start, start + fetch.count).map((row) => {
			const out: Record<string, unknown> = { "@odata.etag": 'W/"1"', [table.primaryId]: row[table.primaryId] };
			for (const name of fetch.attributes) out[name] = row[name] ?? null;
			return out;
		});
		const more = start + fetch.count < rows.length;
		return { value: pageRows, "@Microsoft.Dynamics.CRM.morerecords": more };
	}

	async function queryData(query: string): Promise<unknown> {
		calls.push(`queryData ${query}`);
		const fetchMatch = /^([a-z_]+)\?fetchXml=(.*)$/i.exec(query);
		if (fetchMatch) return runFetch(fetchMatch[1], decodeURIComponent(fetchMatch[2]));

		const privilege = /RetrieveUserPrivilegeByPrivilegeName\(PrivilegeName='([^']+)'\)/.exec(query);
		if (privilege) {
			const held = options.privileges === undefined || options.privileges === "all" || options.privileges.has(privilege[1]);
			return { RolePrivileges: held ? [{ PrivilegeName: privilege[1], Depth: "Global" }] : [] };
		}

		const entityScoped = /^EntityDefinitions\(LogicalName='([^']+)'\)(.*)$/.exec(query);
		if (entityScoped) {
			const table = byName.get(entityScoped[1]);
			const rest = entityScoped[2];
			if (!table) return { value: [] };
			if (rest.startsWith("/Attributes(LogicalName='statecode')")) return stateOptions;
			if (rest.startsWith("/Attributes(LogicalName='statuscode')")) return statusOptions;
			if (rest.startsWith("/Attributes")) return { value: table.attributes.map((a) => attributeDefinition(table, a)) };
			if (/Relationships/.test(rest)) return { value: [] };
			if (rest.includes("EnforceStateTransitions")) return { EnforceStateTransitions: false };
			if (rest.includes("Privileges")) return { Privileges: [] };
			return entityDefinition(table);
		}
		if (query.startsWith("EntityDefinitions")) return { value: tables.map(entityDefinition) };

		const viewMatch = /^(savedqueries|userqueries)(?:\(([0-9a-f-]+)\))?/.exec(query);
		if (viewMatch) {
			const type = viewMatch[1] === "savedqueries" ? "system" : "personal";
			const idKey = type === "system" ? "savedqueryid" : "userqueryid";
			const toRecord = (v: FakeView) => ({
				[idKey]: v.id,
				name: v.name,
				fetchxml: v.fetchxml,
				layoutxml: v.layoutxml,
				returnedtypecode: v.entity,
				querytype: 0,
				statecode: 0,
				iscustomizable: { Value: true },
				ismanaged: false,
				_ownerid_value: userId,
			});
			const matches = views.filter((v) => v.type === type && (!viewMatch[2] || v.id === viewMatch[2]));
			if (viewMatch[2]) {
				if (!matches[0]) throw new Error(`0x80040217: ${viewMatch[1]} With Id = ${viewMatch[2]} Does Not Exist`);
				return toRecord(matches[0]);
			}
			const filterId = /(savedqueryid|userqueryid) eq ([0-9a-f-]+)/.exec(query)?.[2];
			return { value: matches.filter((v) => !filterId || v.id === filterId).map(toRecord) };
		}

		if (/^(publishers|solutions|workflows|asyncoperations|systemusers)\b/.test(query)) return { value: [] };
		unknown.push(query);
		return { value: [] };
	}

	const api = {
		queryData,
		fetchXmlQuery: async (xml: string) => {
			calls.push(`fetchXmlQuery ${xml}`);
			const table = byName.get(readFetch(xml).entity);
			return runFetch(table?.entitySetName ?? "", xml);
		},
		getEntityMetadata: async (name: string) => {
			calls.push(`getEntityMetadata ${name}`);
			const table = byName.get(name);
			if (!table) throw new Error(`0x80060888: Could not find entity '${name}'`);
			return entityDefinition(table);
		},
		execute: async (request: { operationName: string }) => {
			calls.push(`execute ${request.operationName}`);
			if (request.operationName === "WhoAmI") return { UserId: userId, BusinessUnitId: "bu", OrganizationId: "org" };
			unknown.push(`execute ${request.operationName}`);
			return {};
		},
		executeBatch: async (requests: { method: string; url: string }[]) => {
			calls.push(`executeBatch ${requests.map((r) => `${r.method} ${r.url}`).join(", ")}`);
			return requests.map((r) => {
				const m = /^([a-z_]+)\(([0-9a-f-]+)\)$/i.exec(r.url);
				const table = m && bySet.get(m[1]);
				if (r.method === "DELETE" && table && m) {
					const before = table.rows.length;
					table.rows = table.rows.filter((row) => row[table.primaryId] !== m[2]);
					if (table.rows.length < before) {
						deleted.push(`${table.logicalName}:${m[2]}`);
						return { status: 204, headers: {} };
					}
					return { status: 404, headers: {}, body: { error: { code: "0x80040217", message: "Does Not Exist" } } };
				}
				return { status: 200, headers: {}, body: {} };
			});
		},
		delete: async (entity: string, id: string) => {
			calls.push(`delete ${entity} ${id}`);
			const table = byName.get(entity);
			if (table) table.rows = table.rows.filter((row) => row[table.primaryId] !== id);
			deleted.push(`${entity}:${id}`);
		},
		create: async (entity: string, record: Record<string, unknown>) => {
			calls.push(`create ${entity}`);
			const table = byName.get(entity);
			const id = `00000000-0000-0000-0000-${String((table?.rows.length ?? 0) + 900).padStart(12, "0")}`;
			if (table) table.rows.push({ ...record, [table.primaryId]: id });
			return { id };
		},
		update: async (entity: string, id: string) => {
			calls.push(`update ${entity} ${id}`);
		},
	};

	return { api, calls, unknown, deleted, tables };
}

/** Account and contact tables with a few rows each, enough for paging and record actions */
export function sampleTables(accountCount = 5): FakeTable[] {
	const id = (n: number, prefix: string) => `${prefix}000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
	return [
		{
			logicalName: "account",
			entitySetName: "accounts",
			displayName: "Account",
			primaryId: "accountid",
			primaryName: "name",
			objectTypeCode: 1,
			attributes: [
				{ name: "accountid", type: "Uniqueidentifier", label: "Account" },
				{ name: "name", type: "String", label: "Account Name" },
				{ name: "statecode", type: "State", label: "Status" },
				{ name: "statuscode", type: "Status", label: "Status Reason" },
			],
			rows: Array.from({ length: accountCount }, (_, i) => ({ accountid: id(i + 1, "aa"), name: `Account ${String(i + 1).padStart(2, "0")}`, statecode: 0, statuscode: 1 })),
		},
		{
			logicalName: "contact",
			entitySetName: "contacts",
			displayName: "Contact",
			primaryId: "contactid",
			primaryName: "fullname",
			objectTypeCode: 2,
			attributes: [
				{ name: "contactid", type: "Uniqueidentifier", label: "Contact" },
				{ name: "fullname", type: "String", label: "Full Name" },
				{ name: "statecode", type: "State", label: "Status" },
			],
			rows: [
				{ contactid: id(1, "cc"), fullname: "Avery Howard", statecode: 0 },
				{ contactid: id(2, "cc"), fullname: "Jordan Lee", statecode: 0 },
			],
		},
	];
}
