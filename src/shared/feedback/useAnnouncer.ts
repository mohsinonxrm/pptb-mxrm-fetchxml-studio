/**
 * The app's announcer (#186 PR 3): `useAnnouncer()` gives `announce(message, options)`, read by
 * screen readers through AnnouncerProvider's live regions (Announcer.tsx). Outside the provider it
 * does nothing.
 */
import { createContext, useContext } from "react";

export type Politeness = "polite" | "assertive";

export interface AnnounceOptions {
	/** Default "polite": read when the user is idle. "assertive" interrupts (errors). */
	politeness?: Politeness;
	/** Messages with the same id replace each other while they wait */
	batchId?: string;
}

export type Announce = (message: string, options?: AnnounceOptions) => void;

/** Messages this close together are read as one */
export const FLUSH_MS = 150;
/** The text is cleared this long after it's read */
export const CLEAR_AFTER_MS = 7000;

export const AnnouncerContext = createContext<Announce>(() => {});

/** The app's announcer; outside the provider, a no-op */
export function useAnnouncer(): Announce {
	return useContext(AnnouncerContext);
}
