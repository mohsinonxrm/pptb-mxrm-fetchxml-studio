import { afterEach, describe, expect, it, vi } from "vitest";
import { executeWorkflowBatch } from "./pptbClient";
import { WORKFLOW_UNKNOWN_MESSAGE } from "./batchDelete";

// #264: background workflows in $batch requests; real-time ones (and unknown modes) one request each

const workflowid = "11111111-2222-3333-4444-555555555555";
const ids = ["00000000-0000-0000-0000-00000000000a", "00000000-0000-0000-0000-00000000000b"];

/** A host whose $batch stops early when `stops`, with what queryData answers for the workflow and jobs */
function host({ stops = false, workflow = { workflowid, _activeworkflowid_value: "act", asyncautodelete: false } as Record<string, unknown> | Error } = {}) {
	const execute = vi.fn(async () => ({}));
	const executeBatch = vi.fn(async (requests: unknown[]) => {
		if (stops && requests.length > 1) throw new Error("Batch stopped early");
		return requests.map(() => ({ status: 200, headers: {} }));
	});
	const queryData = vi.fn(async (query: string) => {
		if (query.startsWith("workflows?")) {
			if (workflow instanceof Error) throw workflow;
			return { value: [workflow] };
		}
		return { value: [] }; // no system jobs
	});
	(window as { dataverseAPI?: unknown }).dataverseAPI = { execute, executeBatch, queryData };
	return { execute, executeBatch, queryData };
}

afterEach(() => {
	delete (window as { dataverseAPI?: unknown }).dataverseAPI;
});

describe("executeWorkflowBatch (#264)", () => {
	it("sends a background workflow in $batch requests", async () => {
		const { execute, executeBatch, queryData } = host();
		const result = await executeWorkflowBatch({ workflowid, mode: 0 }, ids, "contact");
		expect(result).toEqual({ succeeded: 2, failed: 0, errors: [] });
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(execute).not.toHaveBeenCalled();
		// Its system jobs are matched by the definition and the activation
		expect(queryData.mock.calls.some(([q]) => String(q).includes(`_workflowactivationid_value eq ${workflowid} or _workflowactivationid_value eq act`))).toBe(true);
	});

	it("sends real-time workflows, and workflows of unknown mode, one request per record", async () => {
		for (const mode of [1, undefined]) {
			const { execute, executeBatch } = host();
			const result = await executeWorkflowBatch({ workflowid, mode }, ids, "contact");
			expect(result.succeeded).toBe(2);
			expect(execute).toHaveBeenCalledTimes(2);
			expect(executeBatch).not.toHaveBeenCalled();
		}
	});

	it("doesn't resend a stopped batch of a workflow that deletes its completed jobs", async () => {
		const { executeBatch } = host({ stops: true, workflow: { workflowid, _activeworkflowid_value: "act", asyncautodelete: true } });
		const result = await executeWorkflowBatch({ workflowid, mode: 0 }, ids, "contact");
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result.failed).toBe(2);
		expect(result.errors).toEqual(ids.map((id) => `Record ${id}: ${WORKFLOW_UNKNOWN_MESSAGE}`));
	});

	it("doesn't resend a stopped batch when the workflow couldn't be read", async () => {
		const { executeBatch } = host({ stops: true, workflow: new Error("Network error") });
		const result = await executeWorkflowBatch({ workflowid, mode: 0 }, ids, "contact");
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result.errors.every((e) => e.endsWith(WORKFLOW_UNKNOWN_MESSAGE))).toBe(true);
	});
});
