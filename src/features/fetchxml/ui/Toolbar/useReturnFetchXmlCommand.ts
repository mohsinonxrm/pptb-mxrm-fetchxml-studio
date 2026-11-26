/**
 * Return FetchXML as a command for the shared bar (#232): returns the current FetchXML to the tool
 * that launched this window (T2T callee).
 *
 * This is distinct from the host-injected "Return to [Caller]" banner: that banner is a plain "go
 * back" affordance and resolves the caller's promise with null. Only an explicit returnData() call
 * carries the query back — which is what this does.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { ArrowReply20Filled, ArrowReply20Regular, bundleIcon } from "@fluentui/react-icons";
import { returnFetchXmlToInvokingTool } from "../../api/invocation";
import { useNotify } from "../../../../shared/feedback/useNotify";
import { useCopy } from "../../../../shared/feedback/useCopy";
import type { BarCommand } from "../../../../shared/components/commandBar";

const ReturnIcon = bundleIcon(ArrowReply20Filled, ArrowReply20Regular);

/**
 * How long FetchXML Studio waits, after returnData() resolves, for Power Platform ToolBox to close
 * it. PPTB closes the callee as soon as the caller has the data; when the caller was closed first,
 * returnData() resolves all the same and nothing closes (#363, verified in PPTB 1.2.9: the pending
 * invocation is dropped with the caller, and the return is ignored without an error).
 */
export const NOT_DELIVERED_AFTER_MS = 2000;

export function useReturnFetchXmlCommand(getCurrentXml: () => string): BarCommand {
	const [isReturning, setIsReturning] = useState(false);
	const { notify } = useNotify();
	const copy = useCopy();
	const notDeliveredTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
	useEffect(
		() => () => {
			if (notDeliveredTimer.current) clearTimeout(notDeliveredTimer.current);
		},
		[],
	);

	// A failure is a sticky error toast with Try again and Copy FetchXML, so the query isn't lost
	// (#277, owner decision 4)
	const handleReturn = useCallback(async () => {
		setIsReturning(true);
		const xml = getCurrentXml();
		try {
			await returnFetchXmlToInvokingTool(xml);
			// Delivered: PPTB closes this window. Still here after a moment: the caller had closed, so
			// the FetchXML went nowhere (owner decision on #363: detect it here). Trying again can't
			// help, so the toast offers only Copy FetchXML.
			await new Promise<void>((resolve) => {
				notDeliveredTimer.current = setTimeout(resolve, NOT_DELIVERED_AFTER_MS);
			});
			notDeliveredTimer.current = null;
			notify.error({
				title: "Couldn't return the FetchXML",
				body: "The tool that opened FetchXML Studio is closed, so the FetchXML can't go back to it. Copy it to use it elsewhere.",
				action: { label: "Copy FetchXML", onClick: () => void copy(xml, "FetchXML") },
			});
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			notify.error({
				title: "Couldn't return the FetchXML",
				body: message,
				action: { label: "Try again", onClick: () => void handleReturnRef.current() },
				secondaryAction: {
					label: "Copy FetchXML",
					onClick: () => void copy(xml, "FetchXML"),
				},
			});
		} finally {
			setIsReturning(false);
		}
	}, [getCurrentXml, notify, copy]);
	const handleReturnRef = useRef(handleReturn);
	useEffect(() => {
		handleReturnRef.current = handleReturn;
	}, [handleReturn]);

	return {
		id: "return",
		label: "Return FetchXML",
		icon: ReturnIcon,
		group: "return",
		priority: 10,
		kind: "button",
		busy: isReturning,
		hint: "Return the current FetchXML to the tool that launched this window",
		onClick: () => void handleReturn(),
	};
}
