import { afterEach, describe, expect, it, vi } from "vitest";
import { getWorkflowJobs } from "./pptbClient";

// #264: a stopped $batch of background workflows is checked against the records' system jobs

describe("getWorkflowJobs", () => {
	afterEach(() => {
		delete (window as { dataverseAPI?: unknown }).dataverseAPI;
	});

	it("reads the system jobs of the workflow (definition or activation) for the records", async () => {
		const a = "00000000-0000-0000-0000-00000000000A";
		const b = "00000000-0000-0000-0000-00000000000B";
		const queryData = vi.fn(async () => ({
			value: [
				{ asyncoperationid: "job1", _regardingobjectid_value: a.toLowerCase() },
				{ asyncoperationid: "job2", _regardingobjectid_value: a.toLowerCase() },
				{ asyncoperationid: "job3", _regardingobjectid_value: "someone-else" },
			],
		}));
		(window as { dataverseAPI?: unknown }).dataverseAPI = { queryData };
		const jobs = await getWorkflowJobs(["def", "act"], [a, `{${b}}`]);
		expect(queryData).toHaveBeenCalledWith(
			`asyncoperations?$select=asyncoperationid,_regardingobjectid_value&$filter=(_workflowactivationid_value eq def or _workflowactivationid_value eq act) and (_regardingobjectid_value eq ${a} or _regardingobjectid_value eq ${b})&$orderby=createdon desc`,
		);
		expect(jobs).toEqual(new Map([[a, new Set(["job1", "job2"])]]));
	});

	it("asks for nothing without records", async () => {
		const queryData = vi.fn();
		(window as { dataverseAPI?: unknown }).dataverseAPI = { queryData };
		expect(await getWorkflowJobs(["def"], [])).toEqual(new Map());
		expect(queryData).not.toHaveBeenCalled();
	});
});
