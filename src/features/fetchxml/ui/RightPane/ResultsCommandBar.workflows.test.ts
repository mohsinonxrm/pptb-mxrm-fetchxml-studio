/**
 * The Run workflow menu in the results command bar lists the results' entity's workflows (#182):
 * after the entity changes, opening the menu loads that entity's list, not the one kept before.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { WorkflowInfo } from "../../api/pptbClient";
import { ResultsCommandBar, type CommandBarProps } from "./ResultsCommandBar";

const wf = (workflowid: string, name: string) => ({ workflowid, name }) as WorkflowInfo;
const lists: Record<string, WorkflowInfo[]> = {
	account: [wf("a1", "Account flow")],
	contact: [wf("c1", "Contact flow")],
};

let root: Root;
let container: HTMLDivElement;
const fetchWorkflows = vi.fn(async (entity?: string) => lists[entity ?? ""] ?? []);

const render = (entityName: string | undefined) =>
	act(async () => {
		const props: CommandBarProps = {
			selectedCount: 1,
			canRunWorkflow: true,
			entityName,
			onFetchWorkflows: fetchWorkflows,
		};
		root.render(createElement(FluentProvider, { theme: webLightTheme }, createElement(ResultsCommandBar, props)));
		await Promise.resolve();
	});

const menuButton = () => document.querySelector<HTMLButtonElement>('button[aria-label="Run workflow on selected records"]');
const menuItems = () => [...document.querySelectorAll('[role="menuitem"]')].map((el) => el.textContent?.trim());

/** Click the menu button and wait for the list to load; returns the items shown */
async function openMenu() {
	await act(async () => {
		menuButton()?.click();
		await Promise.resolve();
		await Promise.resolve();
	});
	return menuItems();
}

async function closeMenu() {
	await act(async () => {
		document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
		menuButton()?.click();
		await Promise.resolve();
	});
	// Whichever way it closed, it must be closed before the next step
	if (menuItems().length > 0) {
		await act(async () => {
			menuButton()?.click();
			await Promise.resolve();
		});
	}
}

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	fetchWorkflows.mockClear();
	container = document.createElement("div");
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

describe("ResultsCommandBar: Run workflow menu (#182)", () => {
	it("lists the results' entity's workflows, and loads again after the entity changes", async () => {
		await render("account");
		expect(menuButton()).not.toBeNull();

		expect(await openMenu()).toContain("Account flow");
		expect(fetchWorkflows).toHaveBeenCalledWith("account");
		await closeMenu();

		// The entity changes while the command bar stays mounted (results cleared, then a new run)
		await render(undefined);
		await render("contact");

		const items = await openMenu();
		expect(fetchWorkflows).toHaveBeenLastCalledWith("contact");
		expect(items).toContain("Contact flow");
		expect(items).not.toContain("Account flow");
	});
});
