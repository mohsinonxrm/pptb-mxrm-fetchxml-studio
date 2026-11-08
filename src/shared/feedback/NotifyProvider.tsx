import {
	Fragment,
	useCallback,
	useEffect,
	useLayoutEffect,
	useMemo,
	useRef,
	type ReactNode,
} from "react";
import {
	Link,
	Toast,
	ToastBody,
	ToastFooter,
	ToastTitle,
	ToastTrigger,
	Toaster,
	makeStyles,
	useId,
	useToastController,
} from "@fluentui/react-components";
import {
	resolveOptions,
	type NotifyInput,
	type NotifyIntent,
	type ResolvedOptions,
} from "./notifyOptions";
import { NotifyContext, createNotify, type NotifyApi, type NotifyHandle } from "./useNotify";

export interface NotifyProviderProps {
	children: ReactNode;
	/** Clears every toast when it changes (a new session or a restart: their Undo and Try again are stale) */
	resetKey?: string;
}

/** Toasts on screen at once; more wait in the provider's own queue */
const MAX_VISIBLE = 3;

/** Fluent announces a toast's textContent, so a hidden ". " keeps title, body and links apart */
const useStyles = makeStyles({
	separator: {
		position: "absolute",
		width: "1px",
		height: "1px",
		overflow: "hidden",
		clipPath: "inset(50%)",
		whiteSpace: "nowrap",
	},
});

interface Entry {
	toastId: string;
	intent: NotifyIntent;
	input: NotifyInput;
	options: ResolvedOptions;
}

/**
 * The app's one Toaster and the notify API (#277). Renders before its children so the Toaster's
 * effect, which subscribes to dispatches, runs before theirs.
 *
 * On mount, notify from `useEffect`, not `useLayoutEffect`: the Toaster subscribes in a passive
 * effect, so a toast raised earlier is lost.
 *
 * The queue is the provider's: Fluent's own limit queue can't drop a toast that hasn't shown yet,
 * so a duplicate, a dismiss or a reset would leave it to appear later with a stale action.
 */
export function NotifyProvider({ children, resetKey }: NotifyProviderProps) {
	const styles = useStyles();
	const toasterId = useId("app-toaster");
	const { dispatchToast, dismissToast, dismissAllToasts } = useToastController(toasterId);
	/** Ids on screen */
	const visible = useRef(new Set<string>());
	/** Waiting for room, oldest first */
	const pending = useRef<Entry[]>([]);
	/** Dedupe key to toast id, for visible and pending toasts */
	const byKey = useRef(new Map<string, string>());
	const counter = useRef(0);
	const pump = useRef<() => void>(() => undefined);

	const release = useCallback((entry: Entry) => {
		visible.current.delete(entry.toastId);
		if (byKey.current.get(entry.options.key) === entry.toastId)
			byKey.current.delete(entry.options.key);
		pump.current();
	}, []);

	const dismissEntry = useCallback(
		(entry: Entry) => {
			const queued = pending.current.indexOf(entry);
			if (queued >= 0) {
				pending.current.splice(queued, 1);
				if (byKey.current.get(entry.options.key) === entry.toastId)
					byKey.current.delete(entry.options.key);
				return;
			}
			if (!visible.current.has(entry.toastId)) return;
			dismissToast(entry.toastId);
			release(entry);
		},
		[dismissToast, release],
	);

	const dispatchEntry = useCallback(
		(entry: Entry) => {
			const { input, options, intent, toastId } = entry;
			visible.current.add(toastId);
			const dismiss = () => dismissEntry(entry);
			const separator = <span className={styles.separator}>. </span>;
			// A secondary action without a primary one is ignored (see NotifyInput)
			const actions = [input.action, input.action ? input.secondaryAction : undefined].filter(
				(a) => a !== undefined,
			);
			const actionLink = actions.length ? (
				<>
					{actions.map((action) => (
						<Fragment key={action.label}>
							<Link
								as="button"
								onClick={() => {
									try {
										action.onClick();
									} finally {
										dismiss();
									}
								}}
							>
								{action.label}
							</Link>
							{separator}
						</Fragment>
					))}
				</>
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
						{separator}
					</ToastTitle>
					{input.body ? (
						<ToastBody>
							{input.body}
							{separator}
						</ToastBody>
					) : null}
					{options.sticky && actionLink ? <ToastFooter>{actionLink}</ToastFooter> : null}
				</Toast>,
				{
					toastId,
					intent,
					timeout: options.timeout,
					politeness: options.politeness,
					onStatusChange: (_event, data) => {
						if (
							(data.status === "dismissed" || data.status === "unmounted") &&
							visible.current.has(toastId)
						)
							release(entry);
					},
				},
			);
		},
		[dismissEntry, dispatchToast, release, styles.separator],
	);

	pump.current = () => {
		while (visible.current.size < MAX_VISIBLE && pending.current.length > 0) {
			dispatchEntry(pending.current.shift()!);
		}
	};

	const dismissKey = useCallback(
		(key: string) => {
			const id = byKey.current.get(key);
			if (id === undefined) return;
			const entry = pending.current.find((e) => e.toastId === id);
			if (entry) return dismissEntry(entry);
			// On screen: dismiss by id (release needs the entry only for its key)
			visible.current.delete(id);
			byKey.current.delete(key);
			dismissToast(id);
			pump.current();
		},
		[dismissEntry, dismissToast],
	);

	const dismissAll = useCallback(() => {
		pending.current = [];
		byKey.current.clear();
		visible.current.clear();
		dismissAllToasts();
	}, [dismissAllToasts]);

	const show = useCallback(
		(intent: NotifyIntent, input: NotifyInput): NotifyHandle => {
			const options = resolveOptions(intent, input);
			// The same key replaces the toast, shown or waiting: a fresh id, because Fluent ignores a repeated one
			dismissKey(options.key);
			const entry: Entry = {
				toastId: `${toasterId}-${++counter.current}`,
				intent,
				input,
				options,
			};
			byKey.current.set(options.key, entry.toastId);
			if (visible.current.size < MAX_VISIBLE) dispatchEntry(entry);
			else pending.current.push(entry);
			return { dismiss: () => dismissEntry(entry) };
		},
		[dismissEntry, dismissKey, dispatchEntry, toasterId],
	);

	// A layout effect, so it runs before the children's passive effects: a toast the remounted
	// subtree raises on mount isn't wiped
	const previousResetKey = useRef(resetKey);
	useLayoutEffect(() => {
		if (previousResetKey.current === resetKey) return;
		previousResetKey.current = resetKey;
		dismissAll();
	}, [resetKey, dismissAll]);

	// Keep the queue honest if the provider unmounts
	useEffect(() => () => void (pending.current = []), []);

	const api = useMemo<NotifyApi>(
		() => ({ notify: createNotify(show), dismiss: dismissKey, dismissAll }),
		[show, dismissKey, dismissAll],
	);

	return (
		<NotifyContext.Provider value={api}>
			<Toaster
				toasterId={toasterId}
				position="bottom-end"
				offset={{ vertical: 48, horizontal: 16 }}
				pauseOnHover
				pauseOnWindowBlur
			/>
			{children}
		</NotifyContext.Provider>
	);
}
