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

export function useReturnFetchXmlCommand(getCurrentXml: () => string): BarCommand {
	const [isReturning, setIsReturning] = useState(false);
	const { notify } = useNotify();
	const copy = useCopy();

	// A failure is a sticky error toast with Try again and Copy FetchXML, so the query isn't lost
	// (#277, owner decision 4)
	const handleReturn = useCallback(async () => {
		setIsReturning(true);
		const xml = getCurrentXml();
		try {
			await returnFetchXmlToInvokingTool(xml);
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
		group: "share",
		priority: 10,
		kind: "button",
		busy: isReturning,
		hint: "Return the current FetchXML to the tool that launched this window",
		onClick: () => void handleReturn(),
	};
}
