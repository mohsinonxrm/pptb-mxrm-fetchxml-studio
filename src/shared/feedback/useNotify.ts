import { createContext, useContext } from "react";
import type { NotifyInput, NotifyIntent } from "./notifyOptions";

export interface NotifyHandle {
	dismiss: () => void;
}

type Shorthand = (input: NotifyInput | string) => NotifyHandle;

export interface Notify {
	(intent: NotifyIntent, input: NotifyInput | string): NotifyHandle;
	success: Shorthand;
	info: Shorthand;
	warning: Shorthand;
	error: Shorthand;
}

export interface NotifyApi {
	notify: Notify;
	/** Dismisses the visible toast with this key */
	dismiss: (key: string) => void;
	dismissAll: () => void;
}

const noopHandle: NotifyHandle = { dismiss: () => undefined };

export function createNotify(
	show: (intent: NotifyIntent, input: NotifyInput) => NotifyHandle,
): Notify {
	const normalize = (input: NotifyInput | string): NotifyInput =>
		typeof input === "string" ? { title: input } : input;
	const notify = ((intent: NotifyIntent, input: NotifyInput | string) =>
		show(intent, normalize(input))) as Notify;
	notify.success = (input) => show("success", normalize(input));
	notify.info = (input) => show("info", normalize(input));
	notify.warning = (input) => show("warning", normalize(input));
	notify.error = (input) => show("error", normalize(input));
	return notify;
}

/** Outside a provider (a component rendered on its own) every call does nothing */
const noopApi: NotifyApi = {
	notify: createNotify(() => noopHandle),
	dismiss: () => undefined,
	dismissAll: () => undefined,
};

export const NotifyContext = createContext<NotifyApi>(noopApi);

export function useNotify(): NotifyApi {
	return useContext(NotifyContext);
}
