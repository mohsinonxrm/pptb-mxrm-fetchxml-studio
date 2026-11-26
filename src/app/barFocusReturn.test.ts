/**
 * A dialog a bar command opened returns focus to that command, though the command gained a reason
 * while the dialog was open (#393): after a delete the query runs again behind the dialog, and
 * Delete says "Wait for the query to finish". The command keeps its element (StableTooltip), so
 * the app's focus return (#186, useFocusReturn) finds it.
 */
import { act, createElement, useState } from "react";
import { Dialog, DialogBody, DialogSurface, DialogTitle } from "@fluentui/react-components";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../shared/testing/ui";
import { OverflowToolbar } from "../shared/components/commandBar/OverflowToolbar";
import type { BarCommand } from "../shared/components/commandBar/types";
import { useFocusReturn } from "./hooks/useFocusReturn";

afterEach(() => cleanup());

const wait = (ms: number) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));

describe("focus return to a bar command (#393)", () => {
	it("comes back to Delete after its dialog closes, though Delete gained a reason meanwhile", async () => {
		stubOverflowWidths({ container: 3000 });
		let setOpen: (open: boolean) => void = () => {};
		let setReason: (reason: string | undefined) => void = () => {};
		function Screen() {
			useFocusReturn();
			const [open, setOpenState] = useState(false);
			const [reason, setReasonState] = useState<string | undefined>();
			setOpen = setOpenState;
			setReason = setReasonState;
			const del = {
				id: "delete",
				label: "Delete",
				kind: "button",
				priority: 10,
				disabledReason: reason,
				onClick: () => setOpenState(true),
			} as BarCommand;
			return createElement(
				"div",
				null,
				createElement(OverflowToolbar, { ariaLabel: "Record actions", commands: [del] }),
				createElement(Dialog, {
					open,
					onOpenChange: (_e: unknown, data: { open: boolean }) => setOpenState(data.open),
					children: createElement(
						DialogSurface,
						null,
						createElement(
							DialogBody,
							null,
							createElement(DialogTitle, null, "Delete record"),
						),
					),
				}),
			);
		}
		await render(createElement(Screen));
		const del = [...document.querySelectorAll<HTMLElement>("button")].find(
			(b) => b.textContent?.trim() === "Delete",
		);
		del?.focus();
		await click(del);
		await wait(50);
		expect(document.querySelector('[role="dialog"]')).toBeTruthy();
		// The query runs again behind the dialog
		await act(async () => setReason("Wait for the query to finish"));
		await act(async () => setOpen(false));
		await wait(100);
		expect(document.activeElement).toBe(del);
	});
});
