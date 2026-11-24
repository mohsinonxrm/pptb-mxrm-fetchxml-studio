/**
 * The builder pane's command beside the scope chip (#234): shown while access is being checked and
 * after, and the same element throughout, so focus on it (Collapse, Pin) survives the check.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, flush, installHost, removeHost, render } from "../../../../shared/testing/ui";
import { EntitySelector } from "./EntitySelector";

const access = vi.hoisted(() => ({
	current: { accessSummary: null as unknown, loading: true, noAccessMode: false, error: null },
}));
vi.mock("../../../../shared/hooks/useAccessMode", () => ({
	useAccessMode: () => access.current,
}));

afterEach(() => {
	cleanup();
	removeHost();
});

const selector = () =>
	createElement(EntitySelector, {
		selectedEntity: null,
		onEntityChange: () => {},
		entityScopeMode: "all",
		onEntityScopeModeChange: () => {},
		advancedFindOnly: false,
		onAdvancedFindOnlyChange: () => {},
		headerAction: createElement("button", { type: "button", id: "pane-action" }, "Pane action"),
	});

describe("the scope header's pane command", () => {
	it("is there while access is checked, and keeps focus once it has been", async () => {
		installHost();
		access.current = { accessSummary: null, loading: true, noAccessMode: false, error: null };
		await render(selector());
		await flush(2);
		expect(document.body.textContent).toContain("Checking access permissions");
		const action = document.getElementById("pane-action");
		expect(action).toBeTruthy();
		action?.focus();

		access.current = {
			accessSummary: {
				userId: "u1",
				canReadPublisher: true,
				canReadSolution: true,
				canReadCustomization: true,
				fullFilterMode: true,
				solutionsOnlyMode: false,
				publishersOnlyMode: false,
				metadataOnlyMode: false,
				noAccessMode: false,
			},
			loading: false,
			noAccessMode: false,
			error: null,
		};
		await render(selector());
		await flush(3);
		expect(document.body.textContent).not.toContain("Checking access permissions");
		// The same element, still focused
		expect(document.getElementById("pane-action")).toBe(action);
		expect(document.activeElement).toBe(action);
	});
});

describe("the Entity picker's hint (#279)", () => {
	it("says why the picker can't be used: no scope chosen yet", async () => {
		installHost();
		access.current = {
			accessSummary: {
				userId: "u1",
				canReadPublisher: true,
				canReadSolution: true,
				canReadCustomization: true,
				fullFilterMode: true,
				solutionsOnlyMode: false,
				publishersOnlyMode: false,
				metadataOnlyMode: false,
				noAccessMode: false,
			},
			loading: false,
			noAccessMode: false,
			error: null,
		};
		await render(
			createElement(EntitySelector, {
				selectedEntity: null,
				onEntityChange: () => {},
				entityScopeMode: "publisher-solution",
				onEntityScopeModeChange: () => {},
				advancedFindOnly: false,
				onAdvancedFindOnlyChange: () => {},
			}),
		);
		await flush(5);
		const combo = document.querySelector('[role="combobox"]');
		const hint = document.getElementById(combo?.getAttribute("aria-describedby") ?? "");
		expect(hint?.textContent).toBe("Choose a scope to list entities");
	});
});
