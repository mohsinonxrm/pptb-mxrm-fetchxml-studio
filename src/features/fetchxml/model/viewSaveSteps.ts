/**
 * The steps after a system view is created or updated (#180): add a new view to the chosen
 * solution, then publish. If one fails, the view already exists, so the Save View dialog retries
 * from the failed step instead of saving again, which created a second view with the same name.
 */
export type ViewSaveStep = "solution" | "publish";

/** The follow-up steps for a save, in the order they run */
export function viewSaveSteps(options: { addToSolution: boolean; publish: boolean }): ViewSaveStep[] {
	const steps: ViewSaveStep[] = [];
	if (options.addToSolution) steps.push("solution");
	if (options.publish) steps.push("publish");
	return steps;
}

/**
 * Runs the steps in order. The first failure stops it: the result has the error and the steps still
 * to run, starting with the one that failed. Null when every step succeeded.
 */
export async function runViewSaveSteps(
	steps: readonly ViewSaveStep[],
	run: (step: ViewSaveStep) => Promise<void>,
): Promise<{ remaining: ViewSaveStep[]; error: unknown } | null> {
	for (let i = 0; i < steps.length; i++) {
		try {
			await run(steps[i]);
		} catch (error) {
			return { remaining: steps.slice(i), error };
		}
	}
	return null;
}
