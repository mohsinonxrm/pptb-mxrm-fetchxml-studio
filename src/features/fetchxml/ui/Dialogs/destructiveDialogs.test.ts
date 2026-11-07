/**
 * Destructive dialogs (#189, rules from #69): each names the environment and the scope before
 * anything runs, asks for the typed confirmation on Production, for more than 100 records or when
 * the total isn't known, and can't be submitted twice while it runs.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { BulkDeleteDialog } from "./BulkDeleteDialog";
import { DeleteConfirmDialog, type DeleteConfirmDialogProps } from "./DeleteConfirmDialog";
import { WorkflowPickerDialog, type WorkflowPickerDialogProps } from "./WorkflowPickerDialog";
import type { WorkflowInfo } from "../../api/pptbClient";
import { button, cleanup, click, deferred, dialogInput, dialogText, flush, installHost, isDisabled, removeHost, render, typeInto } from "../../../../shared/testing/ui";

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

const deleteProps = (over: Partial<DeleteConfirmDialogProps> = {}): DeleteConfirmDialogProps => ({
	open: true,
	recordName: "FXS TEST 1",
	entityDisplayName: "Contact",
	entityLogicalName: "contact",
	recordCount: 1,
	onClose: () => {},
	onConfirm: async () => undefined,
	...over,
});

describe("Delete record dialog", () => {
	it("names the environment, the entity and the scope, and says it can't be undone", async () => {
		installHost({ connection: { name: "CRM279420", url: "https://org382cc706.crm.dynamics.com/", environment: "Dev" } });
		await render(createElement(DeleteConfirmDialog, deleteProps()));
		const text = dialogText();
		expect(text).toContain("CRM279420");
		expect(text).toContain("https://org382cc706.crm.dynamics.com/");
		expect(text).toContain("Dev");
		expect(text).toContain("Contact");
		expect(text).toContain("1 record");
		expect(text).toContain("cannot be undone");
		expect(isDisabled(button("Delete"))).toBe(false);
	});

	it("deletes once, and can't be submitted or cancelled while it runs", async () => {
		installHost();
		const run = deferred<void>();
		const onConfirm = vi.fn(() => run.promise);
		const onClose = vi.fn();
		await render(createElement(DeleteConfirmDialog, deleteProps({ onConfirm, onClose })));

		await click(button("Delete"));
		expect(onConfirm).toHaveBeenCalledTimes(1);
		// While deleting, no Delete button, and the only Cancel is disabled
		expect(button("Delete")).toBeUndefined();
		expect(isDisabled(button("Cancel"))).toBe(true);

		run.resolve();
		await flush();
		expect(dialogText()).toMatch(/deleted successfully/i);
		expect(onConfirm).toHaveBeenCalledTimes(1);
	});

	it("shows Dataverse's error when the delete fails", async () => {
		installHost();
		const onConfirm = vi.fn(async () => {
			throw new Error("0x80040220: Principal user is missing prvDeleteContact privilege");
		});
		await render(createElement(DeleteConfirmDialog, deleteProps({ onConfirm })));
		await click(button("Delete"));
		expect(dialogText()).toMatch(/prvDeleteContact/);
	});

	it("asks for the entity name before deleting more than 100 records", async () => {
		installHost();
		await render(createElement(DeleteConfirmDialog, deleteProps({ recordCount: 101 })));
		expect(dialogText()).toContain("Type contact to confirm");
		expect(isDisabled(button("Delete"))).toBe(true);
		await typeInto(dialogInput(), "account");
		expect(isDisabled(button("Delete"))).toBe(true);
		await typeInto(dialogInput(), " Contact ");
		expect(isDisabled(button("Delete"))).toBe(false);
	});

	it("asks for it on a Production environment, even for one record", async () => {
		installHost({ connection: { name: "Contoso Prod", environment: "Production" } });
		await render(createElement(DeleteConfirmDialog, deleteProps()));
		expect(dialogText()).toContain("Production environment");
		expect(dialogText()).toContain("Required on a Production environment.");
		expect(isDisabled(button("Delete"))).toBe(true);
		await typeInto(dialogInput(), "contact");
		expect(isDisabled(button("Delete"))).toBe(false);
	});

	it("doesn't ask for 100 records or fewer on a non-production environment", async () => {
		installHost();
		await render(createElement(DeleteConfirmDialog, deleteProps({ recordCount: 100 })));
		expect(dialogInput()).toBeUndefined();
		expect(isDisabled(button("Delete"))).toBe(false);
	});
});

describe("Bulk delete dialog", () => {
	it("with no rows selected: the typed confirmation and the I understand box, then a job name", async () => {
		installHost();
		await render(
			createElement(BulkDeleteDialog, {
				open: true,
				recordCount: 0,
				entityDisplayName: "Contact",
				entityLogicalName: "contact",
				isAllRecords: true,
				totalViewRecords: 4,
				moreRecords: false,
				onClose: () => {},
				onConfirm: async () => ({ asyncOperationId: "job", jobUrl: "" }),
			}),
		);
		const text = dialogText();
		expect(text).toContain("Every record the query matches");
		expect(text).toContain("Type contact to confirm");
		const submit = () => button("Submit bulk delete");
		expect(isDisabled(submit())).toBe(true);

		await typeInto(dialogInput(), "contact");
		expect(isDisabled(submit())).toBe(true);
		const understand = [...document.querySelectorAll<HTMLInputElement>('[role="dialog"] input[type="checkbox"]')].find((c) =>
			/I understand/.test(c.closest("label, div")?.textContent ?? ""),
		);
		await click(understand);
		expect(isDisabled(submit())).toBe(false);
	});
});

const workflow = (id: string, name: string): WorkflowInfo =>
	({ workflowid: id, name, primaryentity: "contact", category: 0, type: 1, statuscode: 2, mode: 0 }) as WorkflowInfo;

const workflowProps = (over: Partial<WorkflowPickerDialogProps> = {}): WorkflowPickerDialogProps => ({
	open: true,
	recordCount: 1,
	entityDisplayName: "Contact",
	entityLogicalName: "contact",
	onClose: () => {},
	onFetchWorkflows: async () => [workflow("w1", "Reset Security Stamp"), workflow("w2", "Second flow")],
	onExecute: async () => ({ succeeded: 1, failed: 0, errors: [] }),
	...over,
});

describe("Run workflow dialog", () => {
	it("names the environment and scope, says there's no undo, and waits for a workflow to be chosen", async () => {
		installHost({ connection: { name: "CRM279420" } });
		await render(createElement(WorkflowPickerDialog, workflowProps()));
		const text = dialogText();
		expect(text).toContain("CRM279420");
		expect(text).toContain("1 record");
		expect(text).toContain("There's no undo");
		expect(isDisabled(button("Run workflow"))).toBe(true);
	});

	it("is ready to run with the workflow picked from the menu, and runs it once", async () => {
		installHost();
		const run = deferred<{ succeeded: number; failed: number; errors: string[] }>();
		const onExecute = vi.fn((_workflow: WorkflowInfo, _onProgress: unknown) => run.promise);
		await render(createElement(WorkflowPickerDialog, workflowProps({ preSelectedWorkflow: workflow("w1", "Reset Security Stamp"), onExecute })));
		expect(isDisabled(button("Run workflow"))).toBe(false);

		await click(button("Run workflow"));
		expect(onExecute).toHaveBeenCalledTimes(1);
		expect(onExecute.mock.calls[0][0]).toMatchObject({ workflowid: "w1" });
		expect(button("Run workflow")).toBeUndefined();

		run.resolve({ succeeded: 1, failed: 0, errors: [] });
		await flush();
		expect(dialogText()).toMatch(/1 record processed successfully/);
	});

	it("asks for the entity name before running on more than 100 records", async () => {
		installHost();
		await render(createElement(WorkflowPickerDialog, workflowProps({ recordCount: 150, preSelectedWorkflow: workflow("w1", "Reset Security Stamp") })));
		expect(isDisabled(button("Run workflow"))).toBe(true);
		await typeInto(dialogInput(), "contact");
		expect(isDisabled(button("Run workflow"))).toBe(false);
	});
});
