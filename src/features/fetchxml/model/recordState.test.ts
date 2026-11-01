import { describe, expect, it } from "vitest";
import {
	closeMessageReason,
	isSimpleStateModel,
	parseTransitionData,
	planStateChange,
	stateActionUnavailableReason,
	statePatchBody,
	toEntityStateModel,
	transitionFilter,
	type EntityStateModel,
} from "./recordState";

// #143: Activate / Deactivate from the entity's state model. Shapes and values as the sandbox
// returned them (StateAttributeMetadata / StatusAttributeMetadata options, 2026-10-17).

const label = (text: string) => ({ UserLocalizedLabel: { Label: text } });

const account = toEntityStateModel(
	[
		{ Value: 0, Label: label("Active"), InvariantName: "Active", DefaultStatus: 1 },
		{ Value: 1, Label: label("Inactive"), InvariantName: "Inactive", DefaultStatus: 2 },
	],
	[
		{ Value: 1, Label: label("Active"), State: 0, TransitionData: null },
		{ Value: 2, Label: label("Inactive"), State: 1, TransitionData: null },
	],
	false,
);

const lead = toEntityStateModel(
	[
		{ Value: 0, Label: label("Open"), DefaultStatus: 1 },
		{ Value: 1, Label: label("Qualified"), DefaultStatus: 3 },
		{ Value: 2, Label: label("Disqualified"), DefaultStatus: 4 },
	],
	[
		{ Value: 1, Label: label("New"), State: 0 },
		{ Value: 2, Label: label("Contacted"), State: 0 },
		{ Value: 3, Label: label("Qualified"), State: 1 },
		{ Value: 4, Label: label("Lost"), State: 2 },
		{ Value: 5, Label: label("Cannot Contact"), State: 2 },
		{ Value: 6, Label: label("No Longer Interested"), State: 2 },
		{ Value: 7, Label: label("Canceled"), State: 2 },
	],
	false,
);

const task = toEntityStateModel(
	[
		{ Value: 0, Label: label("Open"), DefaultStatus: 2 },
		{ Value: 1, Label: label("Completed"), DefaultStatus: 5 },
		{ Value: 2, Label: label("Canceled"), DefaultStatus: 6 },
	],
	[
		{ Value: 2, Label: label("Not Started"), State: 0 },
		{ Value: 3, Label: label("In Progress"), State: 0 },
		{ Value: 5, Label: label("Completed"), State: 1 },
		{ Value: 6, Label: label("Canceled"), State: 2 },
	],
	false,
);

/** A custom entity with enforced transitions: Active (1 New, 10 Review), Inactive (2 Done, 20 Dropped) */
const enforced = (transitions: Record<number, number[] | undefined>): EntityStateModel => ({
	states: [
		{ value: 0, label: "Active", defaultStatus: 1 },
		{ value: 1, label: "Inactive", defaultStatus: 2 },
	],
	statuses: [
		{ value: 1, label: "New", state: 0, allowedNext: transitions[1] },
		{ value: 10, label: "Review", state: 0, allowedNext: transitions[10] },
		{ value: 2, label: "Done", state: 1, allowedNext: transitions[2] },
		{ value: 20, label: "Dropped", state: 1, allowedNext: transitions[20] },
	],
	enforceTransitions: true,
});

describe("parseTransitionData", () => {
	it("reads the tostatusid values (Learn's incident example)", () => {
		const xml = `<allowedtransitions xmlns="https://schemas.microsoft.com/crm/2009/WebServices">
<allowedtransition sourcestatusid="1" tostatusid="6" />
<allowedtransition sourcestatusid="1" tostatusid="1000" />
<allowedtransition sourcestatusid="1" tostatusid="2000" />
<allowedtransition sourcestatusid="1" tostatusid="5" />
</allowedtransitions>`;
		expect(parseTransitionData(xml)).toEqual([6, 1000, 2000, 5]);
	});

	it("reads escaped XML and treats a missing document as no data", () => {
		expect(parseTransitionData('&lt;allowedtransitions&gt;&lt;allowedtransition sourcestatusid="1" tostatusid="2" /&gt;&lt;/allowedtransitions&gt;')).toEqual([2]);
		expect(parseTransitionData(null)).toBeUndefined();
		expect(parseTransitionData("  ")).toBeUndefined();
	});
});

describe("closeMessageReason", () => {
	it("names the documented action for close-message entities", () => {
		expect(closeMessageReason("opportunity")).toBe(
			"Opportunities close with Win or Lose (WinOpportunity / LoseOpportunity). Use the model-driven app for now.",
		);
		for (const entity of ["incident", "quote", "salesorder", "invoice", "contract"]) expect(closeMessageReason(entity)).toBeTruthy();
	});

	it("leaves other entities alone", () => {
		expect(closeMessageReason("account")).toBeUndefined();
		expect(closeMessageReason("lead")).toBeUndefined();
		expect(closeMessageReason(undefined)).toBeUndefined();
	});
});

describe("planStateChange", () => {
	it("gives simple entities their one target: the state's default status", () => {
		expect(isSimpleStateModel(account)).toBe(true);
		expect(planStateChange(account, "deactivate")).toEqual({
			kind: "simple",
			target: { statecode: 1, statuscode: 2 },
			stateLabel: "Inactive",
			statusLabel: "Inactive",
		});
		expect(planStateChange(account, "activate")).toMatchObject({ kind: "simple", target: { statecode: 0, statuscode: 1 } });
	});

	it("marks plans for entities that enforce transitions, whose way back may not be allowed (#69)", () => {
		const plan = planStateChange(enforced({ 1: [2, 20] }), "deactivate", [1]);
		expect(plan).toMatchObject({ kind: "choose", transitionsEnforced: true });
		expect(planStateChange(lead, "deactivate", undefined, "lead")).not.toHaveProperty("transitionsEnforced");
		expect(planStateChange(account, "deactivate")).not.toHaveProperty("transitionsEnforced");
	});

	it("offers status reasons grouped by state for multi-state entities", () => {
		expect(isSimpleStateModel(lead)).toBe(false);
		const activate = planStateChange(lead, "activate", undefined, "lead");
		expect(activate).toMatchObject({ kind: "choose", defaultTarget: { statecode: 0, statuscode: 1 } });
		const deactivate = planStateChange(task, "deactivate");
		if (deactivate.kind !== "choose") throw new Error("expected choices");
		expect(deactivate.groups.map((g) => [g.label, g.choices.map((c) => c.label)])).toEqual([
			["Completed", ["Completed"]],
			["Canceled", ["Canceled"]],
		]);
		expect(deactivate.defaultTarget).toEqual({ statecode: 1, statuscode: 5 });
	});

	it("preselects a state without a caveat: a lead's Deactivate defaults to Disqualified", () => {
		const plan = planStateChange(lead, "deactivate", undefined, "lead");
		if (plan.kind !== "choose") throw new Error("expected choices");
		expect(plan.groups.map((g) => g.label)).toEqual(["Qualified", "Disqualified"]);
		expect(plan.defaultTarget).toEqual({ statecode: 2, statuscode: 4 });
		expect(plan.groups[1].choices.find((c) => c.isDefault)?.label).toBe("Lost");
	});

	it("says why when there's no target state", () => {
		const single: EntityStateModel = { states: [{ value: 0, label: "Active", defaultStatus: 1 }], statuses: [{ value: 1, label: "Active", state: 0 }], enforceTransitions: false };
		expect(planStateChange(single, "deactivate")).toEqual({ kind: "none", reason: "This entity has no inactive state." });
		expect(stateActionUnavailableReason(single, "deactivate")).toBe("This entity has no inactive state.");
		expect(stateActionUnavailableReason(null, "activate")).toBe("This entity has no Status (statecode) column.");
		expect(stateActionUnavailableReason(account, "activate")).toBeUndefined();
	});
});

describe("status reason transitions", () => {
	it("doesn't filter when transitions aren't enforced", () => {
		expect(transitionFilter(lead, [1])).toEqual({});
	});

	it("offers only targets allowed from the current status", () => {
		const model = enforced({ 1: [10, 2], 10: [20] });
		const plan = planStateChange(model, "deactivate", [1]);
		if (plan.kind !== "choose") throw new Error("expected choices");
		expect(plan.groups[0].choices.map((c) => c.label)).toEqual(["Done"]);
		expect(plan.note).toMatch(/only targets allowed from the current status reason/);
	});

	it("intersects the allowed targets when the selected records differ, and says so", () => {
		const model = enforced({ 1: [2, 20], 10: [20] });
		const filter = transitionFilter(model, [1, 10, 10]);
		expect([...(filter.allowed ?? [])]).toEqual([20]);
		expect(filter.note).toMatch(/different status reasons/);
		const plan = planStateChange(model, "deactivate", [1, 10]);
		if (plan.kind !== "choose") throw new Error("expected choices");
		expect(plan.defaultTarget).toEqual({ statecode: 1, statuscode: 20 });
	});

	it("reports no target when the intersection is empty", () => {
		const model = enforced({ 1: [2], 10: [20] });
		expect(planStateChange(model, "deactivate", [1, 10])).toMatchObject({ kind: "none" });
	});

	it("doesn't restrict when the current statuses are unknown or have no TransitionData", () => {
		const model = enforced({ 1: [2] });
		expect(transitionFilter(model, undefined).allowed).toBeUndefined();
		expect(transitionFilter(model, undefined).note).toMatch(/couldn't be read/);
		expect(transitionFilter(model, [10]).allowed).toBeUndefined();
		// A status with data and one without: only the known one restricts
		expect([...(transitionFilter(model, [1, 10]).allowed ?? [])]).toEqual([2]);
	});
});

describe("statePatchBody", () => {
	it("always sets both statecode and statuscode", () => {
		expect(statePatchBody({ statecode: 1, statuscode: 2 })).toEqual({ statecode: 1, statuscode: 2 });
	});
});
