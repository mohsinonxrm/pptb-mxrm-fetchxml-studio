/**
 * Failures of things the user started in the tool are in-app toasts, not host notifications
 * (#277): Return FetchXML (sticky, with Try again and Copy FetchXML, owner decision 4) and Send to
 * tool.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	button,
	cleanup,
	click,
	flush,
	installHost,
	removeHost,
	render,
	toasts,
	type HostStub,
} from "../testing/ui";
import { useReturnFetchXmlCommand } from "../../features/fetchxml/ui/Toolbar/useReturnFetchXmlCommand";
import {
	useSendToToolCommand,
	type SendToToolCommandOptions,
} from "../../features/fetchxml/ui/Toolbar/useSendToToolCommand";
import { OverflowToolbar } from "../components/commandBar";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string }) =>
		React.createElement("textarea", { readOnly: true, value: props.value ?? "" });
	return {
		default: Stub,
		Editor: Stub,
		DiffEditor: Stub,
		loader: { config() {}, init: async () => ({}) },
		useMonaco: () => null,
	};
});

const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";

/** The commands on a bar of their own, as the header bar renders them (#232) */
function ReturnBar({ getCurrentXml }: { getCurrentXml: () => string }) {
	return createElement(OverflowToolbar, {
		ariaLabel: "Test commands",
		commands: [useReturnFetchXmlCommand(getCurrentXml)],
	});
}
function SendBar(options: SendToToolCommandOptions) {
	const command = useSendToToolCommand(options);
	return createElement(OverflowToolbar, {
		ariaLabel: "Test commands",
		commands: command ? [command] : [],
	});
}

let host: HostStub;
let copied: string[];
let invocation: { returnData: ReturnType<typeof vi.fn>; launchTool: ReturnType<typeof vi.fn> };

beforeEach(() => {
	host = installHost();
	copied = [];
	const toolbox = (
		window as unknown as {
			toolboxAPI: { utils: Record<string, unknown>; invocation?: unknown };
		}
	).toolboxAPI;
	toolbox.utils.copyToClipboard = async (text: string) => {
		copied.push(text);
	};
	invocation = {
		returnData: vi.fn(async () => Promise.reject(new Error("The caller has closed"))),
		launchTool: vi.fn(async () => Promise.reject(new Error("Tool not installed"))),
	};
	toolbox.invocation = invocation;
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

const toastButton = (label: string) =>
	[...document.querySelectorAll<HTMLElement>(".fui-Toast button")].find(
		(b) => b.textContent?.trim() === label,
	);

describe("Return FetchXML", () => {
	it("a failure is a sticky error toast with Try again and Copy FetchXML", async () => {
		await render(createElement(ReturnBar, { getCurrentXml: () => XML }));
		await click(button("Return FetchXML"));
		expect(invocation.returnData).toHaveBeenCalledWith({ fetchXml: XML });
		expect(toasts()).toEqual([
			expect.stringMatching(/^Couldn't return the FetchXML\. Dismiss.*The caller has closed/),
		]);
		expect(host.notifications).toEqual([]);

		// Copy FetchXML keeps the query, and closes the toast
		await click(toastButton("Copy FetchXML"));
		expect(copied).toEqual([XML]);
		await flush(5);
		expect(toastButton("Copy FetchXML")).toBeUndefined();

		// Try again returns it again; failing again raises the toast again
		await click(button("Return FetchXML"));
		expect(invocation.returnData).toHaveBeenCalledTimes(2);
		invocation.returnData.mockImplementationOnce(async () =>
			Promise.reject(new Error("Still closed")),
		);
		await click(toastButton("Try again"));
		await flush(5);
		expect(invocation.returnData).toHaveBeenCalledTimes(3);
		expect(toasts().at(-1)).toMatch(/^Couldn't return the FetchXML\. Dismiss.*Still closed/);

		// Try again that works says nothing more
		invocation.returnData.mockImplementationOnce(async () => undefined);
		await click(toastButton("Try again"));
		await flush(5);
		expect(invocation.returnData).toHaveBeenCalledTimes(4);
		expect(toastButton("Try again")).toBeUndefined();
	});
});

describe("Send to tool", () => {
	it("a launch that fails is an error toast naming the tool", async () => {
		await render(
			createElement(SendBar, {
				fetchXml: XML,
				entityLogicalName: "account",
				tools: [{ id: "pptb-other-tool", name: "Other Tool" }],
			}),
		);
		await click(button("Send to tool"));
		await click(
			[...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].find((m) =>
				m.textContent?.includes("Other Tool"),
			),
		);
		await flush(5);
		expect(invocation.launchTool).toHaveBeenCalledWith(
			"pptb-other-tool",
			expect.objectContaining({ fetchXml: XML }),
			expect.anything(),
		);
		expect(toasts()).toEqual([
			expect.stringMatching(/^Couldn't open Other Tool\. Dismiss.*Tool not installed/),
		]);
		expect(host.notifications).toEqual([]);
	});

	it("a second launch the host rejects as already in progress says nothing", async () => {
		invocation.launchTool.mockImplementation(async () =>
			Promise.reject(new Error("A tool launch is already in progress")),
		);
		await render(
			createElement(SendBar, {
				fetchXml: XML,
				entityLogicalName: "account",
				tools: [{ id: "pptb-other-tool", name: "Other Tool" }],
			}),
		);
		await click(button("Send to tool"));
		await click(document.querySelector('[role="menuitem"]'));
		await flush(5);
		expect(toasts()).toEqual([]);
	});
});
