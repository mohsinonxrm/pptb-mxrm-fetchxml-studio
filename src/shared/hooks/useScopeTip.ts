/**
 * Whether the scope chip's first-run tip (#227) was dismissed in this environment. Stored in
 * toolboxAPI.settings like the remembered scope (key "fetchxml-studio:scopeTipDismissed", a map
 * keyed by environment URL). Without toolboxAPI.settings it lasts for the session only.
 */
import { useCallback, useEffect, useState } from "react";
import { getPrimaryConnection } from "../../features/fetchxml/api/host";
import { environmentKey } from "../../features/fetchxml/model/scopeSelection";

export const SCOPE_TIP_DISMISSED_KEY = "fetchxml-studio:scopeTipDismissed";

const settingsApi = () => (typeof window === "undefined" ? undefined : window.toolboxAPI?.settings);

/** The stored map, or an empty one when the stored value isn't one */
const dismissedMap = (value: unknown): Record<string, true> =>
	value && typeof value === "object" && !Array.isArray(value)
		? Object.fromEntries(Object.entries(value as Record<string, unknown>).filter(([, v]) => v === true)) as Record<string, true>
		: {};

/** @returns `ready` once the stored value is read; `dismissed` for this environment; `dismiss()` */
export function useScopeTip(): { ready: boolean; dismissed: boolean; dismiss: () => void } {
	const [state, setState] = useState<{ ready: boolean; dismissed: boolean; envKey: string; stored: Record<string, true> }>({
		ready: false,
		dismissed: false,
		envKey: "",
		stored: {},
	});

	useEffect(() => {
		let mounted = true;
		const settings = settingsApi();
		if (!settings?.get) {
			setState((s) => ({ ...s, ready: true }));
			return;
		}
		Promise.all([settings.get(SCOPE_TIP_DISMISSED_KEY), getPrimaryConnection()])
			.then(([value, connection]) => {
				if (!mounted) return;
				const stored = dismissedMap(value);
				const envKey = environmentKey(connection?.url);
				setState({ ready: true, dismissed: !!envKey && stored[envKey] === true, envKey, stored });
			})
			.catch(() => mounted && setState((s) => ({ ...s, ready: true })));
		return () => {
			mounted = false;
		};
	}, []);

	const dismiss = useCallback(() => {
		setState((s) => {
			if (s.envKey) {
				const stored = { ...s.stored, [s.envKey]: true as const };
				settingsApi()
					?.set?.(SCOPE_TIP_DISMISSED_KEY, stored)
					.catch((err: unknown) => console.warn("Failed to remember the dismissed scope tip:", err));
				return { ...s, dismissed: true, stored };
			}
			return { ...s, dismissed: true };
		});
	}, []);

	return { ready: state.ready, dismissed: state.dismissed, dismiss };
}
