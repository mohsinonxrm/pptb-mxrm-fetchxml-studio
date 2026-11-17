/**
 * The FetchXML editor's bars (#277, D3 (a)): one per condition, a check replaces what it checked,
 * most severe first.
 */
import { describe, expect, it } from "vitest";
import {
	applyCheck,
	FULL_CHECK,
	invalidProblem,
	mightRejectProblem,
	notesProblem,
	syntaxProblem,
	type EditorProblem,
} from "./editorProblems";

const treeChanged: EditorProblem = {
	key: "tree-changed",
	intent: "warning",
	title: "The tree changed while you were editing",
	body: "…",
};
const keys = (problems: EditorProblem[]) => problems.map((p) => p.key);

describe("applyCheck", () => {
	it("a second report of the same condition replaces the first: never two error bars", () => {
		const first = applyCheck([], ["error"], [syntaxProblem("Unexpected end")]);
		const second = applyCheck(first, ["error"], [invalidProblem([{ message: "No entity" }])]);
		expect(second).toEqual([invalidProblem([{ message: "No entity" }])]);
	});

	it("a clean check removes the bars for what it checked, and keeps the others", () => {
		const shown = applyCheck([treeChanged], FULL_CHECK, [
			invalidProblem([{ message: "x" }]),
			notesProblem([{ message: "n" }]),
		]);
		expect(keys(applyCheck(shown, ["error"], []))).toEqual(["tree-changed", "notes"]);
		expect(keys(applyCheck(shown, FULL_CHECK, []))).toEqual(["tree-changed"]);
	});

	it("a check that only looks at errors leaves a might-reject warning alone", () => {
		const shown = applyCheck([], ["warning"], [mightRejectProblem([{ message: "w" }])]);
		expect(keys(applyCheck(shown, ["error"], []))).toEqual(["warning"]);
	});

	it("a finding replaces its key even when the check didn't list it", () => {
		const shown = applyCheck([], ["warning"], [mightRejectProblem([{ message: "old" }])]);
		const next = applyCheck(shown, [], [mightRejectProblem([{ message: "new" }])]);
		expect(next).toEqual([mightRejectProblem([{ message: "new" }])]);
	});

	it("most severe first: error, might reject, tree changed, notes", () => {
		const shown = applyCheck([treeChanged], FULL_CHECK, [
			notesProblem([{ message: "n" }]),
			mightRejectProblem([{ message: "w" }]),
			syntaxProblem("e"),
		]);
		expect(keys(shown)).toEqual(["error", "warning", "tree-changed", "notes"]);
	});
});
