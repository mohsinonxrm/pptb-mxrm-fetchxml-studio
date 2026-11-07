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

/** The ResizeObservers alive now, so stubOverflowWidths can fire them */
const resizeObservers = new Set<RecordingResizeObserver>();

/**
 * jsdom has no ResizeObserver; Fluent's MessageBar, Spinner and Overflow construct one. This one does
 * nothing on its own, but records its callback and targets so a test can fire it (stubOverflowWidths).
 */
class RecordingResizeObserver {
	readonly callback: ResizeObserverCallback;
	readonly targets = new Set<Element>();
	constructor(callback: ResizeObserverCallback) {
		this.callback = callback;
		resizeObservers.add(this);
	}
	observe(target: Element): void {
		this.targets.add(target);
	}
	unobserve(target: Element): void {
		this.targets.delete(target);
	}
	disconnect(): void {
		this.targets.clear();
		resizeObservers.delete(this);
	}
	/** Calls back with one entry per observed element, as a resize would */
	fire(): void {
		if (!this.targets.size) return;
		const entries = [...this.targets].map(
			(target) => ({ target, contentRect: target.getBoundingClientRect(), borderBoxSize: [], contentBoxSize: [], devicePixelContentBoxSize: [] }) as unknown as ResizeObserverEntry,
		);
		this.callback(entries, this as unknown as ResizeObserver);
	}
}

function installResizeObserver(): void {
	const win = window as unknown as { ResizeObserver?: unknown };
	if (typeof win.ResizeObserver !== "function") win.ResizeObserver = RecordingResizeObserver;
}

/** Renders the element inside a light FluentProvider; a second call re-renders the same root */
export async function render(ui: ReactElement): Promise<void> {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	installResizeObserver();
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

/** Unmounts and removes everything render added (and any width stub); call it in afterEach */
export function cleanup(): void {
	if (root) {
		const target = root;
		act(() => target.unmount());
	}
	container?.remove();
	root = null;
	container = null;
	document.body.innerHTML = "";
	restoreWidths?.();
}

/** Widths, in px, that stubOverflowWidths reports. jsdom has no layout, so everything measures 0. */
export interface OverflowWidths {
	/** clientWidth of each Overflow container: the element Fluent's Overflow renders as (the Toolbar or TabList root) */
	container: number;
	/** offsetWidth of each overflow item; a function gets the element (the button, tab or SplitButton root). Default 80. */
	item?: number | ((element: HTMLElement) => number);
	/** offsetWidth of the More button. Default 40. */
	menu?: number;
	/** offsetWidth of each divider in the container. Default 24. */
	divider?: number;
}

let restoreWidths: (() => void) | null = null;

/**
 * Lays out Fluent's Overflow in jsdom. @fluentui/priority-overflow measures the container's
 * `clientWidth` and each visible item's, divider's and More button's `offsetWidth`, then hides the
 * lowest priority items until the rest fit in `clientWidth - padding`. It measures when it starts
 * observing (on mount, only if `clientWidth > 0`), when items or the More button register, and when
 * its ResizeObserver calls back.
 *
 * This stubs those two getters, by structure: the container is the element with Overflow's
 * `fui-Overflow` class; its direct children are the More button (`data-overflow-menu`, which Fluent
 * sets before registering it), dividers (role="separator") and items (everything else). The item
 * and divider data attributes aren't used because Fluent adds them only after it first measures.
 * Everything else keeps jsdom's 0.
 *
 * Call it before render so the first layout uses the widths. `resize` changes widths and fires every
 * recorded ResizeObserver inside act, as a real resize would. cleanup restores the getters.
 */
export function stubOverflowWidths(initial: OverflowWidths): { resize: (next: Partial<OverflowWidths>) => Promise<void>; restore: () => void } {
	restoreWidths?.();
	installResizeObserver();
	const widths: OverflowWidths = { ...initial };
	const proto = HTMLElement.prototype;
	const own = {
		offsetWidth: Object.getOwnPropertyDescriptor(proto, "offsetWidth"),
		clientWidth: Object.getOwnPropertyDescriptor(proto, "clientWidth"),
	};
	// jsdom's own getters (offsetWidth on HTMLElement, clientWidth on Element)
	const fallback = {
		offsetWidth: own.offsetWidth ?? Object.getOwnPropertyDescriptor(Element.prototype, "offsetWidth"),
		clientWidth: own.clientWidth ?? Object.getOwnPropertyDescriptor(Element.prototype, "clientWidth"),
	};
	const isContainer = (element: Element | null) => !!element?.classList.contains("fui-Overflow");
	const measure = (element: HTMLElement, which: "offsetWidth" | "clientWidth"): number => {
		if (isContainer(element)) return widths.container;
		if (which === "offsetWidth" && isContainer(element.parentElement)) {
			if (element.hasAttribute("data-overflow-menu")) return widths.menu ?? 40;
			if (element.getAttribute("role") === "separator") return widths.divider ?? 24;
			const item = widths.item ?? 80;
			return typeof item === "function" ? item(element) : item;
		}
		return (fallback[which]?.get?.call(element) as number | undefined) ?? 0;
	};
	for (const which of ["offsetWidth", "clientWidth"] as const) {
		Object.defineProperty(proto, which, {
			configurable: true,
			get(this: HTMLElement) {
				return measure(this, which);
			},
		});
	}
	const restore = () => {
		for (const which of ["offsetWidth", "clientWidth"] as const) {
			const descriptor = own[which];
			if (descriptor) Object.defineProperty(proto, which, descriptor);
			else delete (proto as unknown as Record<string, unknown>)[which];
		}
		if (restoreWidths === restore) restoreWidths = null;
	};
	restoreWidths = restore;
	return {
		restore,
		resize: async (next) => {
			Object.assign(widths, next);
			await act(async () => {
				for (const observer of [...resizeObservers]) observer.fire();
			});
			await flush();
		},
	};
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

/**
 * A row of the dialog's target block (TargetScope, #69): the value next to its "Environment",
 * "Entity" or "Scope" label, with each part separated by a space ("Contact (contact)"), or undefined
 */
export function scopeField(label: "Environment" | "Entity" | "Scope"): string | undefined {
	const dialog = openDialog();
	const labelElement = [...(dialog?.querySelectorAll("span, div") ?? [])].find((e) => e.children.length === 0 && e.textContent === label);
	const value = labelElement?.nextElementSibling;
	if (!value) return undefined;
	const parts = [...value.querySelectorAll("*")].filter((e) => e.children.length === 0).map((e) => e.textContent?.trim()).filter(Boolean);
	return (parts.length ? parts.join(" ") : (value.textContent ?? "")).trim();
}

/** A button by its visible text (exact, trimmed) or a pattern, searched in the open dialog first */
export function button(text: string | RegExp, scope: ParentNode | undefined = openDialog() ?? document): HTMLButtonElement | undefined {
	const matches = (value: string) => (typeof text === "string" ? value.trim() === text : text.test(value.trim()));
	return [...(scope ?? document).querySelectorAll<HTMLButtonElement>("button")].find((b) => matches(b.textContent ?? ""));
}

/** Whether a button can't be pressed: disabled, or Fluent's disabledFocusable. Throws when it's missing,
 * so a renamed or removed button fails the test instead of counting as disabled. */
export function isDisabled(element: HTMLElement | undefined): boolean {
	if (!element) throw new Error("isDisabled: element not found");
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
 * Installs window.toolboxAPI and window.dataverseAPI stubs: one connection (Dev by default),
 * notifications collected, and no Dataverse calls unless a test provides them. There's no
 * preventClose, so holdClose does nothing.
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
