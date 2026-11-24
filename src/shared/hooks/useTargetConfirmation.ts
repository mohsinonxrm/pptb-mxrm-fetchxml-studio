/**
 * What a write confirmation's TargetScope says (#69, #279): whether the action may run, and why not
 * ("Reading the connection...", "Type account to confirm"), for the primary button's reason.
 */
import { useCallback, useState } from "react";

export interface TargetConfirmation {
	confirmed: boolean;
	/** Why the action can't run yet; undefined once confirmed */
	reason: string | undefined;
	/** TargetScope's onConfirmationChange */
	onConfirmationChange: (confirmed: boolean, reason?: string) => void;
}

export function useTargetConfirmation(): TargetConfirmation {
	const [state, setState] = useState<{ confirmed: boolean; reason?: string }>({
		confirmed: false,
	});
	const onConfirmationChange = useCallback((confirmed: boolean, reason?: string) => {
		setState({ confirmed, reason: confirmed ? undefined : reason });
	}, []);
	return { confirmed: state.confirmed, reason: state.reason, onConfirmationChange };
}
