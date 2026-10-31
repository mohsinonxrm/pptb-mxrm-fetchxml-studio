/**
 * The QueryExpression sent in a BulkDelete job's QuerySet (#224).
 *
 * FetchXmlToQueryExpression returns non-string condition values as raw JSON numbers and booleans
 * (`{"Type":"System.Int32","Value":0}`). BulkDelete rejects those with 0x80048d19 "Error identified
 * in Payload provided by the user for Entity :''": the Web API's Object complex type, used by
 * ConditionExpression.Values, defines `Value` as Edm.String. So every query with a number, Yes/No,
 * choice, status, money or decimal condition failed, such as the common "Active …" views.
 * Verified live on 2026-10-30: the same values as strings are accepted, for every type.
 *
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/object
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/delete-data-bulk
 */

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };

/**
 * A copy of the query with every condition value (`Values[].Value`, in the root criteria, nested
 * filters and link-entity criteria) as a string; `Type` is kept. Null values stay null.
 */
export function withStringConditionValues<T>(query: T): T {
	const visit = (node: Json): Json => {
		if (Array.isArray(node)) return node.map(visit);
		if (node === null || typeof node !== "object") return node;
		const copy: { [key: string]: Json } = {};
		for (const [key, value] of Object.entries(node)) {
			if (key === "Values" && Array.isArray(value)) {
				copy[key] = value.map((item) =>
					item !== null && typeof item === "object" && !Array.isArray(item) && "Value" in item
						? { ...item, Value: item.Value === null || typeof item.Value === "string" ? item.Value : String(item.Value) }
						: visit(item),
				);
			} else {
				copy[key] = visit(value);
			}
		}
		return copy;
	};
	return visit(query as Json) as T;
}

/**
 * The query for "Bulk delete the selected records", built here rather than converted from FetchXML:
 * the conversion is a GET with the whole ID list in the URL, which failed above ~530 records (#177).
 */
export function selectedRecordsQuery(entityLogicalName: string, primaryIdAttribute: string, recordIds: readonly string[]) {
	return {
		"@odata.type": "Microsoft.Dynamics.CRM.QueryExpression",
		EntityName: entityLogicalName,
		ColumnSet: { AllColumns: false, Columns: [primaryIdAttribute] },
		Distinct: false,
		Criteria: {
			FilterOperator: "And",
			Conditions: [
				{
					AttributeName: primaryIdAttribute,
					Operator: "In",
					Values: recordIds.map((id) => ({ Type: "System.Guid", Value: id })),
				},
			],
		},
	};
}
