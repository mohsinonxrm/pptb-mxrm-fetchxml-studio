/**
 * Component-test helpers (#189): render into a FluentProvider with React's own `act`, a stubbed
 * PPTB host, and DOM queries that also find Fluent's portals (dialogs, menus, popovers render into
 * document.body). No Testing Library: React and the DOM only, like the hook tests.
 *
 * Only test files import this module.
 */
import { act, createElement, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

let root: Root | null = null;
let container: HTMLDivElement | null = null;

/** Lets pending promises and effects settle inside act */
export async function flush(times = 3): Promise<void> {
	await act(async () => {
		for (let i = 0; i < times; i++) await Promise.resolve();
	});
}

/** jsdom has no ResizeObserver; Fluent's MessageBar, Spinner and Overflow construct one */
class NoopResizeObserver {
	observe(): void {}
	unobserve(): void {}
	disconnect(): void {}
}

/** Renders the element inside a light FluentProvider; a second call re-renders the same root */
export async function render(ui: ReactElement): Promise<void> {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	const win = window as unknown as { ResizeObserver?: unknown };
	if (typeof win.ResizeObserver !== "function") win.ResizeObserver = NoopResizeObserver;
	if (!root) {
		container = document.createElement("div");
		document.body.appendChild(container);
		root = createRoot(container);
	}
	const target = root;
	await act(async () => {
		target.render(createElement(FluentProvider, { theme: webLightTheme }, ui));
	});
	await flush();
}

/** Unmounts and removes everything render added; call it in afterEach */
export function cleanup(): void {
	if (root) {
		const target = root;
		act(() => target.unmount());
	}
	container?.remove();
	root = null;
	container = null;
	document.body.innerHTML = "";
}

/** A real click on an element, inside act */
export async function click(element: Element | null | undefined): Promise<void> {
	if (!element) throw new Error("click: element not found");
	await act(async () => {
		(element as HTMLElement).click();
	});
	await flush();
}

/** Types into a React-controlled input: sets the value through the native setter, then fires input */
export async function typeInto(input: HTMLInputElement | null | undefined, text: string): Promise<void> {
	if (!input) throw new Error("typeInto: input not found");
	const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
	await act(async () => {
		setter?.call(input, text);
		input.dispatchEvent(new Event("input", { bubbles: true }));
	});
	await flush();
}

/** The open dialog (Fluent renders it in a portal), or undefined */
export function openDialog(): HTMLElement | undefined {
	return [...document.querySelectorAll<HTMLElement>('[role="dialog"], [role="alertdialog"]')].pop();
}

/** The open dialog's text, with whitespace collapsed */
export function dialogText(): string {
	return (openDialog()?.textContent ?? "").replace(/\s+/g, " ").trim();
}

/** A button by its visible text (exact, trimmed) or a pattern, searched in the open dialog first */
export function button(text: string | RegExp, scope: ParentNode | undefined = openDialog() ?? document): HTMLButtonElement | undefined {
	const matches = (value: string) => (typeof text === "string" ? value.trim() === text : text.test(value.trim()));
	return [...(scope ?? document).querySelectorAll<HTMLButtonElement>("button")].find((b) => matches(b.textContent ?? ""));
}

/** Whether a button can't be pressed: disabled, or Fluent's disabledFocusable */
export function isDisabled(element: HTMLElement | undefined): boolean {
	if (!element) return true;
	return (element as HTMLButtonElement).disabled || element.getAttribute("aria-disabled") === "true";
}

/** The first text input in the open dialog */
export function dialogInput(): HTMLInputElement | undefined {
	return openDialog()?.querySelector<HTMLInputElement>('input[type="text"], input:not([type])') ?? undefined;
}

export interface StubConnection {
	id: string;
	name: string;
	url: string;
	environment: string;
}

export interface HostStub {
	/** Notifications the tool showed through utils.showNotification */
	notifications: { title?: string; body?: string; type?: string }[];
	/** The stubbed dataverseAPI; replace its functions in a test as needed */
	dataverseAPI: Record<string, (...args: unknown[]) => unknown>;
}

/**
 * Installs window.toolboxAPI and window.dataverseAPI stubs: one connection (CRM279420-like, Dev by
 * default), notifications collected, close holding accepted, and no Dataverse calls unless a test
 * provides them.
 */
export function installHost(options: { connection?: Partial<StubConnection> | null; dataverseAPI?: HostStub["dataverseAPI"] } = {}): HostStub {
	const connection =
		options.connection === null
			? null
			: { id: "c1", name: "Test environment", url: "https://test.crm.dynamics.com/", environment: "Dev", ...options.connection };
	const notifications: HostStub["notifications"] = [];
	const dataverseAPI: HostStub["dataverseAPI"] = {
		queryData: async () => ({ value: [] }),
		...options.dataverseAPI,
	};
	(window as unknown as { toolboxAPI: unknown }).toolboxAPI = {
		connections: {
			getConnection: async () => connection,
			getActiveConnection: async () => connection,
		},
		utils: {
			showNotification: async (notification: HostStub["notifications"][number]) => {
				notifications.push(notification);
			},
			copyToClipboard: async () => undefined,
		},
		events: { on: () => undefined, off: () => undefined },
	};
	(window as unknown as { dataverseAPI: unknown }).dataverseAPI = dataverseAPI;
	return { notifications, dataverseAPI };
}

/** Removes the host stubs; call it in afterEach */
export function removeHost(): void {
	delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
	delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
}

/** A promise the test resolves or rejects when it chooses */
export function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void; reject: (error: unknown) => void } {
	let resolve!: (value: T) => void;
	let reject!: (error: unknown) => void;
	const promise = new Promise<T>((res, rej) => {
		resolve = res;
		reject = rej;
	});
	return { promise, resolve, reject };
}
