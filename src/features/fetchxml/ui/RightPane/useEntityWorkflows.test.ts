/**
 * The Run workflow menu's list follows the results' entity (#182): it's reloaded after the entity
 * changes, and a load that ends after the entity changed doesn't bring the old list back.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { WorkflowInfo } from "../../api/pptbClient";
import { useEntityWorkflows, type EntityWorkflows } from "./useEntityWorkflows";

const wf = (workflowid: string, name: string) => ({ workflowid, name }) as WorkflowInfo;

let latest: EntityWorkflows;
let root: Root;
let container: HTMLDivElement;

function Probe({ entityName, fetch }: { entityName?: string; fetch?: () => Promise<WorkflowInfo[]> }) {
	latest = useEntityWorkflows(entityName, fetch);
	return null;
}

const render = (entityName: string | undefined, fetch?: () => Promise<WorkflowInfo[]>) =>
	act(async () => root.render(createElement(Probe, { entityName, fetch })));
const load = () => act(async () => { await latest.load(); });

/** A fetch that resolves when the test says so */
function deferred() {
	let resolve!: (wfs: WorkflowInfo[]) => void;
	let reject!: (err: unknown) => void;
	const promise = new Promise<WorkflowInfo[]>((res, rej) => {
		resolve = res;
		reject = rej;
	});
	return { promise, resolve, reject };
}

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement("div");
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	vi.restoreAllMocks();
});

describe("useEntityWorkflows (#182)", () => {
	it("loads once per entity, and again after the entity changes", async () => {
		const fetchAccount = vi.fn(async () => [wf("a1", "Account flow")]);
		await render("account", fetchAccount);
		expect(latest).toMatchObject({ workflows: [], loaded: false, isLoading: false });

		await load();
		expect(latest.workflows.map((w) => w.name)).toEqual(["Account flow"]);
		expect(latest.loaded).toBe(true);

		// Opening the menu again for the same entity uses the list it has
		await load();
		expect(fetchAccount).toHaveBeenCalledTimes(1);

		// Another entity: the list is cleared, and the next open loads that entity's workflows
		const fetchContact = vi.fn(async () => [wf("c1", "Contact flow")]);
		await render("contact", fetchContact);
		expect(latest).toMatchObject({ workflows: [], loaded: false });
		await load();
		expect(fetchContact).toHaveBeenCalledTimes(1);
		expect(latest.workflows.map((w) => w.name)).toEqual(["Contact flow"]);
	});

	it("clears the list when the results go away (entity changed, results cleared)", async () => {
		await render("account", async () => [wf("a1", "Account flow")]);
		await load();
		await render(undefined, async () => []);
		expect(latest).toMatchObject({ workflows: [], loaded: false });
	});

	it("drops a load that ends after the entity changed", async () => {
		const slow = deferred();
		await render("account", () => slow.promise);
		let pending!: Promise<void>;
		act(() => {
			pending = latest.load();
		});
		expect(latest.isLoading).toBe(true);

		// The entity changes while the account list is still loading
		const fetchContact = vi.fn(async () => [wf("c1", "Contact flow")]);
		await render("contact", fetchContact);
		expect(latest.isLoading).toBe(false);

		await act(async () => {
			slow.resolve([wf("a1", "Account flow")]);
			await pending;
		});
		expect(latest).toMatchObject({ workflows: [], loaded: false, isLoading: false });

		// The new entity can load straight away
		await load();
		expect(latest.workflows.map((w) => w.name)).toEqual(["Contact flow"]);
	});

	it("drops a failure that ends after the entity changed, and logs a current one", async () => {
		const error = vi.spyOn(console, "error").mockImplementation(() => {});
		const slow = deferred();
		await render("account", () => slow.promise);
		let pending!: Promise<void>;
		act(() => {
			pending = latest.load();
		});
		await render("contact", async () => []);
		await act(async () => {
			slow.reject(new Error("old"));
			await pending;
		});
		expect(error).not.toHaveBeenCalled();
		expect(latest.isLoading).toBe(false);

		await render("contact", async () => {
			throw new Error("current");
		});
		// Same entity, new fetch function: the entity didn't change, so the state is kept
		await load();
		expect(error).toHaveBeenCalledTimes(1);
		expect(latest).toMatchObject({ loaded: false, isLoading: false });
	});

	it("does nothing without a fetch function", async () => {
		await render("account", undefined);
		await load();
		expect(latest).toMatchObject({ workflows: [], loaded: false, isLoading: false });
	});
});
