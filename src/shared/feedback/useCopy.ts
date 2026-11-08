import { useCallback } from "react";
import { copyText } from "../../features/fetchxml/api/host";
import { useNotify } from "./useNotify";

/**
 * Returns `copy(text, what)`: copies the text and says so, "Copied {what}" or "Couldn't copy to
 * the clipboard". Resolves to whether it worked.
 */
export function useCopy(): (text: string, what: string) => Promise<boolean> {
	const { notify } = useNotify();
	return useCallback(
		async (text, what) => {
			const ok = await copyText(text);
			if (ok) notify.success(`Copied ${what}`);
			else notify.error("Couldn't copy to the clipboard");
			return ok;
		},
		[notify],
	);
}
