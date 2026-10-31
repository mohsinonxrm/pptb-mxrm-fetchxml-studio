import { describe, expect, it } from "vitest";
import { selectedRecordsQuery, withStringConditionValues } from "./bulkDeleteQuery";

// Shaped like FetchXmlToQueryExpression's answer for a query with typed conditions (captured live)
const converted = {
	EntityName: "account",
	ColumnSet: { AllColumns: false, Columns: ["name"], AttributeExpressions: [] },
	PageInfo: { PageNumber: 1, Count: 0, PagingCookie: null },
	Criteria: {
		FilterOperator: "And",
		FilterHint: null,
		Conditions: [
			{ EntityName: null, AttributeName: "name", Operator: "Equal", Values: [{ Type: "System.String", Value: "Contoso" }] },
			{ EntityName: null, AttributeName: "statecode", Operator: "Equal", Values: [{ Type: "System.Int32", Value: 0 }] },
			{ EntityName: null, AttributeName: "donotemail", Operator: "Equal", Values: [{ Type: "System.Boolean", Value: false }] },
			{ EntityName: null, AttributeName: "revenue", Operator: "GreaterThan", Values: [{ Type: "System.Decimal", Value: 100.5 }] },
			{ EntityName: null, AttributeName: "numberofemployees", Operator: "In", Values: [{ Type: "System.Int32", Value: 1 }, { Type: "System.Int32", Value: 2 }] },
			{ EntityName: null, AttributeName: "industrycode", Operator: "NotNull", Values: [] },
		],
		Filters: [
			{
				FilterOperator: "Or",
				Conditions: [{ AttributeName: "numberofemployees", Operator: "GreaterThan", Values: [{ Type: "System.Int32", Value: 50 }] }],
				Filters: [],
			},
		],
	},
	LinkEntities: [
		{
			LinkToEntityName: "contact",
			LinkCriteria: {
				FilterOperator: "And",
				Conditions: [{ AttributeName: "statecode", Operator: "Equal", Values: [{ Type: "System.Int32", Value: 0 }] }],
				Filters: [],
			},
			LinkEntities: [],
		},
	],
};

describe("BulkDelete query (#224)", () => {
	it("sends every condition value as a string, keeping its type", () => {
		const q = withStringConditionValues(converted);
		const values = q.Criteria.Conditions.map((c) => c.Values);
		expect(values[0]).toEqual([{ Type: "System.String", Value: "Contoso" }]);
		expect(values[1]).toEqual([{ Type: "System.Int32", Value: "0" }]);
		expect(values[2]).toEqual([{ Type: "System.Boolean", Value: "false" }]);
		expect(values[3]).toEqual([{ Type: "System.Decimal", Value: "100.5" }]);
		expect(values[4]).toEqual([
			{ Type: "System.Int32", Value: "1" },
			{ Type: "System.Int32", Value: "2" },
		]);
		expect(values[5]).toEqual([]);
	});

	it("also converts nested filters and link-entity criteria", () => {
		const q = withStringConditionValues(converted);
		expect(q.Criteria.Filters[0].Conditions[0].Values).toEqual([{ Type: "System.Int32", Value: "50" }]);
		expect(q.LinkEntities[0].LinkCriteria.Conditions[0].Values).toEqual([{ Type: "System.Int32", Value: "0" }]);
	});

	it("leaves everything else alone and doesn't change the input", () => {
		const before = JSON.stringify(converted);
		const q = withStringConditionValues(converted);
		expect(JSON.stringify(converted)).toBe(before);
		expect(q.PageInfo).toEqual(converted.PageInfo);
		expect(q.ColumnSet).toEqual(converted.ColumnSet);
		expect(q.Criteria.Conditions[0].EntityName).toBeNull();
	});

	it("keeps null values null", () => {
		const q = withStringConditionValues({ Criteria: { Conditions: [{ Values: [{ Type: "System.String", Value: null }] }] } });
		expect(q.Criteria.Conditions[0].Values[0].Value).toBeNull();
	});

	it("builds the selected-records query locally, with string GUIDs (no URL limit, #177)", () => {
		const ids = Array.from({ length: 1000 }, (_, i) => `00000000-0000-0000-0000-${String(i).padStart(12, "0")}`);
		const q = selectedRecordsQuery("account", "accountid", ids);
		expect(q["@odata.type"]).toBe("Microsoft.Dynamics.CRM.QueryExpression");
		expect(q.EntityName).toBe("account");
		expect(q.ColumnSet).toEqual({ AllColumns: false, Columns: ["accountid"] });
		const condition = q.Criteria.Conditions[0];
		expect(condition).toMatchObject({ AttributeName: "accountid", Operator: "In" });
		expect(condition.Values).toHaveLength(1000);
		expect(condition.Values[0]).toEqual({ Type: "System.Guid", Value: ids[0] });
	});
});
