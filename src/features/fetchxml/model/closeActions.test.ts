import { describe, expect, it } from "vitest";
import {
	CLOSE_ACTIONS,
	buildCloseParameters,
	closeActionsFor,
	closeDateValue,
	closeRequestUrl,
	closeStatusChoices,
	closedSummary,
	completedByStoppedBatch,
	findCloseAction,
	parseRevenue,
	planCloseStateChange,
	suggestedCloseStatus,
	validateCloseValues,
	type CloseMessageAction,
	type CloseStateAction,
	type CloseValues,
} from "./closeActions";
import { toEntityStateModel } from "./recordState";

// #149: close actions. State models as the sandbox returned them (2026-10-17).

const label = (text: string) => ({ UserLocalizedLabel: { Label: text } });
const model = (states: [number, string, number][], statuses: [number, string, number][]) =>
	toEntityStateModel(
		states.map(([Value, text, DefaultStatus]) => ({ Value, Label: label(text), DefaultStatus })),
		statuses.map(([Value, text, State]) => ({ Value, Label: label(text), State })),
		false,
	);

const opportunity = model(
	[[0, "Open", 1], [1, "Won", 3], [2, "Lost", 4]],
	[[1, "In Progress", 0], [2, "On Hold", 0], [3, "Won", 1], [4, "Canceled", 2], [5, "Out-Sold", 2]],
);
const incident = model(
	[[0, "Active", 1], [1, "Resolved", 5], [2, "Cancelled", 6]],
	[[5, "Problem Solved", 1], [1000, "Information Provided", 1], [6, "Cancelled", 2], [2000, "Merged", 2], [1, "In Progress", 0]],
);
const quote = model(
	[[0, "Draft", 1], [1, "Active", 2], [2, "Won", 4], [3, "Closed", 7]],
	[[1, "In Progress", 0], [2, "In Progress", 1], [3, "Open", 1], [4, "Won", 2], [5, "Lost", 3], [6, "Canceled", 3], [7, "Revised", 3]],
);
const invoice = model(
	[[0, "Active", 1], [1, "Closed (deprecated)", 100003], [2, "Paid", 100001], [3, "Canceled", 100003]],
	[[1, "New", 0], [3, "Canceled (deprecated)", 1], [7, "Paid in Full (deprecated)", 1], [100001, "Complete", 2], [100002, "Partial", 2], [100003, "Canceled", 3]],
);

const message = (entity: string, key: string) => findCloseAction(entity, key) as CloseMessageAction;
const state = (entity: string, key: string) => findCloseAction(entity, key) as CloseStateAction;

describe("CLOSE_ACTIONS", () => {
	it("defines the documented actions per entity, in menu order", () => {
		const summary = Object.fromEntries(
			Object.entries(CLOSE_ACTIONS).map(([entity, actions]) => [entity, actions.map((a) => (a.kind === "message" ? a.message : `PATCH ${a.targetState}`))]),
		);
		expect(summary).toEqual({
			opportunity: ["WinOpportunity", "LoseOpportunity"],
			incident: ["CloseIncident", "PATCH 2"],
			quote: ["WinQuote", "CloseQuote"],
			salesorder: ["FulfillSalesOrder", "CancelSalesOrder"],
			invoice: ["PATCH 2", "PATCH 3"],
		});
		expect(closeActionsFor("Opportunity")).toBe(CLOSE_ACTIONS.opportunity);
		expect(closeActionsFor("account")).toBeUndefined();
		expect(closeActionsFor("contract")).toBeUndefined();
	});

	it("titles and labels in sentence case, singular and plural", () => {
		expect(message("opportunity", "win").title(1)).toBe("Close opportunity as won");
		expect(message("opportunity", "lose").title(3)).toBe("Close 3 opportunities as lost");
		expect(message("incident", "resolve").title(2)).toBe("Resolve 2 cases");
		expect(state("invoice", "paid").title(1)).toBe("Mark invoice as paid");
		expect(closedSummary(message("salesorder", "fulfill"), 1)).toBe("1 order closed.");
		expect(closedSummary(message("quote", "win"), 4)).toBe("4 quotes closed.");
	});
});

describe("closeStatusChoices", () => {
	it("offers the target state's status reasons, default marked and preselected", () => {
		const lose = closeStatusChoices(opportunity, message("opportunity", "lose"));
		expect(lose.state?.label).toBe("Lost");
		expect(lose.choices).toEqual([
			{ statuscode: 4, label: "Canceled", isDefault: true },
			{ statuscode: 5, label: "Out-Sold", isDefault: false },
		]);
		expect(lose.preselected).toBe(4);
		expect(closeStatusChoices(opportunity, message("opportunity", "win")).choices.map((c) => c.statuscode)).toEqual([3]);
	});

	it("preselects Lost for Close quote, not the metadata default Revised", () => {
		const close = closeStatusChoices(quote, message("quote", "close"));
		expect(close.choices.find((c) => c.isDefault)?.label).toBe("Revised");
		expect(close.preselected).toBe(5);
	});

	it("falls back to the default when the preselected status isn't in metadata", () => {
		const custom = model([[3, "Closed", 7]], [[6, "Canceled", 3], [7, "Revised", 3]]);
		expect(closeStatusChoices(custom, message("quote", "close")).preselected).toBe(7);
	});

	it("suggests Out-Sold once a competitor is named", () => {
		const lose = message("opportunity", "lose");
		const choices = closeStatusChoices(opportunity, lose);
		expect(suggestedCloseStatus(lose, choices, undefined)).toBe(4);
		expect(suggestedCloseStatus(lose, choices, "c1")).toBe(5);
		expect(suggestedCloseStatus(message("opportunity", "win"), closeStatusChoices(opportunity, message("opportunity", "win")), "c1")).toBe(3);
	});
});

describe("planCloseStateChange", () => {
	it("Cancel case: Cancelled only, Merged left out, no choice needed", () => {
		expect(planCloseStateChange(incident, state("incident", "cancel"))).toEqual({
			kind: "simple",
			target: { statecode: 2, statuscode: 6 },
			stateLabel: "Cancelled",
			statusLabel: "Cancelled",
			// The way back is Activate (#69)
			undo: { available: true, picksStatusReason: false, transitionsEnforced: false },
		});
	});

	it("Mark as paid: Complete or Partial, the default preselected; deprecated states not offered", () => {
		const plan = planCloseStateChange(invoice, state("invoice", "paid"));
		expect(plan).toEqual({
			kind: "choose",
			groups: [
				{
					statecode: 2,
					label: "Paid",
					choices: [
						{ statuscode: 100001, label: "Complete", isDefault: true },
						{ statuscode: 100002, label: "Partial", isDefault: false },
					],
				},
			],
			defaultTarget: { statecode: 2, statuscode: 100001 },
			note: undefined,
			undo: { available: true, picksStatusReason: false, transitionsEnforced: false },
		});
		expect(planCloseStateChange(invoice, state("invoice", "cancel"))).toMatchObject({ kind: "simple", target: { statecode: 3, statuscode: 100003 } });
	});

	it("says so when the state is missing", () => {
		const noPaid = model([[0, "Active", 1]], [[1, "New", 0]]);
		expect(planCloseStateChange(noPaid, state("invoice", "paid")).kind).toBe("none");
	});
});

describe("validateCloseValues", () => {
	const base: CloseValues = { statuscode: 3, date: new Date(2026, 9, 17) };

	it("accepts empty optional fields", () => {
		expect(validateCloseValues(message("opportunity", "win"), base)).toEqual({});
	});

	it("checks the revenue amount (Money, 2 decimals)", () => {
		const win = message("opportunity", "win");
		expect(validateCloseValues(win, { ...base, revenue: "25000.50" })).toEqual({});
		expect(validateCloseValues(win, { ...base, revenue: "25,000" }).revenue).toBe("Enter an amount, like 25000 or 25000.50.");
		expect(validateCloseValues(win, { ...base, revenue: "1.234" }).revenue).toBeDefined();
		expect(validateCloseValues(win, { ...base, revenue: "-5" }).revenue).toBeDefined();
		expect(validateCloseValues(win, { ...base, revenue: "9999999999999999" }).revenue).toBeDefined();
	});

	it("requires the close date and, for Resolve case, the resolution", () => {
		expect(validateCloseValues(message("quote", "win"), { ...base, date: null }).date).toBe("Select a close date.");
		const resolve = message("incident", "resolve");
		expect(validateCloseValues(resolve, { statuscode: 5, resolution: "  " }).resolution).toBe("Enter a resolution.");
		expect(validateCloseValues(resolve, { statuscode: 5, resolution: "Fixed", billableMinutes: 30 })).toEqual({});
		expect(validateCloseValues(resolve, { statuscode: 5, resolution: "Fixed", billableMinutes: -1 }).billable).toBeDefined();
		expect(validateCloseValues(resolve, { statuscode: 5, resolution: "Fixed", billableMinutes: 1.5 }).billable).toBeDefined();
		expect(validateCloseValues(resolve, { statuscode: 5, resolution: "Fixed", billableMinutes: null })).toEqual({});
	});

	it("ignores fields the action doesn't have", () => {
		expect(validateCloseValues(message("salesorder", "cancel"), { ...base, revenue: "x", resolution: "" })).toEqual({});
	});
});

describe("parseRevenue / closeDateValue", () => {
	it("parses amounts; empty keeps the record's own value", () => {
		expect(parseRevenue("")).toBeUndefined();
		expect(parseRevenue(" 12.5 ")).toBe(12.5);
		expect(parseRevenue("abc")).toBe("invalid");
	});

	it("sends noon of the chosen local day (DateOnly / UserLocal)", () => {
		const value = closeDateValue(new Date(2026, 9, 17, 23, 59));
		expect(value).toBe(new Date(2026, 9, 17, 12).toISOString());
	});
});

describe("buildCloseParameters / closeRequestUrl", () => {
	const id = "{11111111-2222-3333-4444-555555555555}";
	const date = new Date(2026, 9, 17);

	it("WinOpportunity: typed close activity bound to the record, filled fields only, and Status", () => {
		expect(
			buildCloseParameters(message("opportunity", "win"), "opportunities", id, {
				statuscode: 3,
				revenue: "1234.5",
				date,
				description: "  Signed  ",
			}),
		).toEqual({
			OpportunityClose: {
				"@odata.type": "Microsoft.Dynamics.CRM.opportunityclose",
				"opportunityid@odata.bind": "/opportunities(11111111-2222-3333-4444-555555555555)",
				actualrevenue: 1234.5,
				actualend: closeDateValue(date),
				description: "Signed",
			},
			Status: 3,
		});
	});

	it("LoseOpportunity: competitor bound; empty revenue left out (keeps the opportunity's own)", () => {
		const params = buildCloseParameters(message("opportunity", "lose"), "opportunities", id, {
			statuscode: 5,
			revenue: "",
			date,
			competitorId: "c0000000-0000-0000-0000-000000000001",
		});
		expect(params.OpportunityClose).toEqual({
			"@odata.type": "Microsoft.Dynamics.CRM.opportunityclose",
			"opportunityid@odata.bind": "/opportunities(11111111-2222-3333-4444-555555555555)",
			actualend: closeDateValue(date),
			"competitorid@odata.bind": "/competitors(c0000000-0000-0000-0000-000000000001)",
		});
		expect(params.Status).toBe(5);
	});

	it("CloseIncident: the resolution is the subject; billable time is timespent", () => {
		expect(
			buildCloseParameters(message("incident", "resolve"), "incidents", id, {
				statuscode: 1000,
				resolution: " Replaced the roller ",
				billableMinutes: 45,
				description: "Remarks",
				revenue: "99",
				competitorId: "ignored",
			}),
		).toEqual({
			IncidentResolution: {
				"@odata.type": "Microsoft.Dynamics.CRM.incidentresolution",
				"incidentid@odata.bind": "/incidents(11111111-2222-3333-4444-555555555555)",
				subject: "Replaced the roller",
				timespent: 45,
				description: "Remarks",
			},
			Status: 1000,
		});
	});

	it("trims long values to the column lengths", () => {
		const params = buildCloseParameters(message("quote", "close"), "quotes", id, { statuscode: 5, date, description: "x".repeat(2500) });
		expect(((params.QuoteClose as Record<string, unknown>).description as string).length).toBe(2000);
		const resolve = buildCloseParameters(message("incident", "resolve"), "incidents", id, { statuscode: 5, resolution: "y".repeat(300) });
		expect(((resolve.IncidentResolution as Record<string, unknown>).subject as string).length).toBe(200);
	});

	it("FulfillSalesOrder is bound to the collection; the others are unbound", () => {
		expect(closeRequestUrl(message("salesorder", "fulfill"), "salesorders")).toBe("salesorders/Microsoft.Dynamics.CRM.FulfillSalesOrder");
		expect(closeRequestUrl(message("salesorder", "cancel"), "salesorders")).toBe("CancelSalesOrder");
		expect(closeRequestUrl(message("quote", "win"), "quotes")).toBe("WinQuote");
		expect(buildCloseParameters(message("salesorder", "fulfill"), "salesorders", id, { statuscode: 100001, date })).toMatchObject({
			OrderClose: { "@odata.type": "Microsoft.Dynamics.CRM.orderclose", "salesorderid@odata.bind": "/salesorders(11111111-2222-3333-4444-555555555555)" },
			Status: 100001,
		});
	});
});

describe("completedByStoppedBatch", () => {
	it("counts only records that moved into the target state during the run", () => {
		const before = new Map([
			["a", 0],
			["b", 0],
			["c", 1],
			["d", 0],
		]);
		const after = new Map([
			["a", 1],
			["b", 0],
			["c", 1],
			["e", 1],
		]);
		expect([...completedByStoppedBatch(before, after, 1, ["a", "b", "c", "d", "e"])]).toEqual(["a"]);
	});
});
