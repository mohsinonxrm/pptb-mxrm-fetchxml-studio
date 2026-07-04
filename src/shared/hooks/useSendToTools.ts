/**
 * useSendToTools — discovers the PPTB tools that FetchXML Studio can send the current
 * query to, via the host capability registry (tools declaring the "fetchxml" capability).
 *
 * Discovery runs once on mount. Returns an empty list when the host doesn't expose
 * capability discovery (older hosts), so the "Send to Tool" button simply stays hidden.
 */

import { useEffect, useState } from "react";
import { discoverFetchXmlTools, type DiscoveredTool } from "../../features/fetchxml/api/invocation";

export interface SendToToolsState {
	loading: boolean;
	tools: DiscoveredTool[];
}

export function useSendToTools(): SendToToolsState {
	const [state, setState] = useState<SendToToolsState>({ loading: true, tools: [] });

	useEffect(() => {
		let isMounted = true;

		void (async () => {
			const tools = await discoverFetchXmlTools();
			if (isMounted) setState({ loading: false, tools });
		})();

		return () => {
			isMounted = false;
		};
	}, []);

	return state;
}
