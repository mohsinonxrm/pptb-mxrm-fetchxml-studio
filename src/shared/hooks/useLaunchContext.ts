/**
 * useLaunchContext — detects whether this tool was opened via PPTB Tool-to-Tool
 * (T2T) invocation and exposes the raw launch context once.
 *
 * The host's getLaunchContext() is called a single time on mount; both the
 * inbound prefill consumer (AppShell) and the "Return FetchXML" affordance read
 * from this shared result instead of calling the host API independently.
 */

import { useEffect, useState } from "react";
import { getLaunchContext, isT2TSupported } from "../../features/fetchxml/api/invocation";

export interface LaunchContextState {
	/** True until the initial getLaunchContext() resolves. */
	loading: boolean;
	/** True when this tool was launched by another tool (context is non-null). */
	isCallee: boolean;
	/** The raw prefill payload, or null for a standalone launch. */
	context: Record<string, unknown> | null;
}

const STANDALONE: LaunchContextState = { loading: false, isCallee: false, context: null };

export function useLaunchContext(): LaunchContextState {
	const [state, setState] = useState<LaunchContextState>({
		loading: true,
		isCallee: false,
		context: null,
	});

	useEffect(() => {
		let isMounted = true;

		if (!isT2TSupported()) {
			setState(STANDALONE);
			return;
		}

		void (async () => {
			try {
				const context = await getLaunchContext();
				if (isMounted) {
					setState({ loading: false, isCallee: context !== null, context });
				}
			} catch (error) {
				console.error("useLaunchContext: failed to read launch context", error);
				if (isMounted) setState(STANDALONE);
			}
		})();

		return () => {
			isMounted = false;
		};
	}, []);

	return state;
}
