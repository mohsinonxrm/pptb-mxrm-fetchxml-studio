/**
 * What a write confirmation's TargetScope says (#69, #279): whether the action may run, and why not
 * ("Reading the connection...", "Type account to confirm"), for the primary button's reason.
 */
import { useCallback, useState } from "react";

/** Until TargetScope has read the connection (also its first render) (#279) */
export const READING_CONNECTION = "Reading the connection...";

export interface TargetConfirmation {
	confirmed: boolean;
	/** Why the action can't run yet; undefined once confirmed */
	reason: string | undefined;
	/** TargetScope's onConfirmationChange */
	onConfirmationChange: (confirmed: boolean, reason?: string) => void;
}

export function useTargetConfirmation(): TargetConfirmation {
	// Not confirmed until TargetScope says so: its first report comes after the first render
	const [state, setState] = useState<{ confirmed: boolean; reason?: string }>({
		confirmed: false,
		reason: READING_CONNECTION,
	});
	const onConfirmationChange = useCallback((confirmed: boolean, reason?: string) => {
		setState({ confirmed, reason: confirmed ? undefined : (reason ?? READING_CONNECTION) });
	}, []);
	return { confirmed: state.confirmed, reason: state.reason, onConfirmationChange };
}
