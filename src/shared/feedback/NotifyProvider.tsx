import { useCallback, useEffect, useMemo, useRef, type ReactNode } from "react";
import {
	Link,
	Toast,
	ToastBody,
	ToastFooter,
	ToastTitle,
	ToastTrigger,
	Toaster,
	useId,
	useToastController,
} from "@fluentui/react-components";
import { resolveOptions, type NotifyInput, type NotifyIntent } from "./notifyOptions";
import { NotifyContext, createNotify, type NotifyApi, type NotifyHandle } from "./useNotify";

export interface NotifyProviderProps {
	children: ReactNode;
	/** Clears every toast when it changes (a new session or a restart: their Undo and Try again are stale) */
	resetKey?: string;
}

/**
 * The app's one Toaster and the notify API (#277). Renders before its children so the Toaster's
 * effect, which subscribes to dispatches, runs before theirs.
 */
export function NotifyProvider({ children, resetKey }: NotifyProviderProps) {
	const toasterId = useId("app-toaster");
	const { dispatchToast, dismissToast, dismissAllToasts } = useToastController(toasterId);
	/** Visible toast id by dedupe key */
	const visible = useRef(new Map<string, string>());
	const counter = useRef(0);

	const dismissKey = useCallback(
		(key: string) => {
			const id = visible.current.get(key);
			if (id === undefined) return;
			visible.current.delete(key);
			dismissToast(id);
		},
		[dismissToast],
	);

	const dismissAll = useCallback(() => {
		visible.current.clear();
		dismissAllToasts();
	}, [dismissAllToasts]);

	const show = useCallback(
		(intent: NotifyIntent, input: NotifyInput): NotifyHandle => {
			const options = resolveOptions(intent, input);
			// The same key replaces the visible toast: a fresh id, because Fluent ignores a repeated one
			dismissKey(options.key);
			const toastId = `${toasterId}-${++counter.current}`;
			visible.current.set(options.key, toastId);
			const dismiss = () => {
				if (visible.current.get(options.key) === toastId) visible.current.delete(options.key);
				dismissToast(toastId);
			};
			const action = input.action;
			const actionLink = action ? (
				<Link
					as="button"
					onClick={() => {
						action.onClick();
						dismiss();
					}}
				>
					{action.label}
				</Link>
			) : undefined;
			dispatchToast(
				<Toast>
					<ToastTitle
						action={
							options.sticky ? (
								<ToastTrigger>
									<Link as="button">Dismiss</Link>
								</ToastTrigger>
							) : (
								// A timed toast keeps its action beside the title
								actionLink
							)
						}
					>
						{input.title}
					</ToastTitle>
					{input.body ? <ToastBody>{input.body}</ToastBody> : null}
					{options.sticky && actionLink ? <ToastFooter>{actionLink}</ToastFooter> : null}
				</Toast>,
				{
					toastId,
					intent,
					timeout: options.timeout,
					politeness: options.politeness,
					onStatusChange: (_event, data) => {
						if (data.status === "unmounted" && visible.current.get(options.key) === toastId) visible.current.delete(options.key);
					},
				},
			);
			return { dismiss };
		},
		[dismissKey, dismissToast, dispatchToast, toasterId],
	);

	const previousResetKey = useRef(resetKey);
	useEffect(() => {
		if (previousResetKey.current === resetKey) return;
		previousResetKey.current = resetKey;
		dismissAll();
	}, [resetKey, dismissAll]);

	const api = useMemo<NotifyApi>(() => ({ notify: createNotify(show), dismiss: dismissKey, dismissAll }), [show, dismissKey, dismissAll]);

	return (
		<NotifyContext.Provider value={api}>
			<Toaster toasterId={toasterId} position="bottom-end" offset={{ vertical: 48, horizontal: 16 }} limit={3} pauseOnHover pauseOnWindowBlur />
			{children}
		</NotifyContext.Provider>
	);
}
