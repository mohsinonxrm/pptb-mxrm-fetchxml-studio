import { describe, expect, it } from "vitest";
import { runViewSaveSteps, viewSaveSteps, type ViewSaveStep } from "./viewSaveSteps";

describe("Save View follow-up steps (#180)", () => {
	it("adds to the solution before publishing, and only the steps asked for", () => {
		expect(viewSaveSteps({ addToSolution: true, publish: true })).toEqual(["solution", "publish"]);
		expect(viewSaveSteps({ addToSolution: false, publish: true })).toEqual(["publish"]);
		expect(viewSaveSteps({ addToSolution: true, publish: false })).toEqual(["solution"]);
		expect(viewSaveSteps({ addToSolution: false, publish: false })).toEqual([]);
	});

	it("runs every step and returns null when they all succeed", async () => {
		const ran: ViewSaveStep[] = [];
		const result = await runViewSaveSteps(["solution", "publish"], async (step) => {
			ran.push(step);
		});
		expect(result).toBeNull();
		expect(ran).toEqual(["solution", "publish"]);
	});

	it("stops at the first failure and returns it with the steps still to run", async () => {
		const ran: ViewSaveStep[] = [];
		const failure = new Error("not added");
		const result = await runViewSaveSteps(["solution", "publish"], async (step) => {
			ran.push(step);
			if (step === "solution") throw failure;
		});
		expect(ran).toEqual(["solution"]); // publish didn't run
		expect(result).toEqual({ remaining: ["solution", "publish"], error: failure });
	});

	it("a retry runs only what's left", async () => {
		const ran: ViewSaveStep[] = [];
		const first = await runViewSaveSteps(["solution", "publish"], async (step) => {
			ran.push(step);
			if (step === "publish") throw new Error("publish failed");
		});
		expect(first?.remaining).toEqual(["publish"]);
		const retry = await runViewSaveSteps(first!.remaining, async (step) => {
			ran.push(step);
		});
		expect(retry).toBeNull();
		expect(ran).toEqual(["solution", "publish", "publish"]); // the solution step ran once
	});
});
