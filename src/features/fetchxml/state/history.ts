/**
 * Undo history core for the query builder (#192). Pure and generic: it stores immutable snapshots
 * of `S` by reference (the builder state is structurally shared, so a snapshot costs only its
 * references) and never reads the clock; callers pass `now`.
 *
 * Not wired yet. The #192 wiring PR (after #301) wraps `builderReducer` with it, roughly:
 *
 *   S = BuilderState (the whole state), with `createHistory(state, { restore })` (see Restore).
 *   Each action applies the reducer, then one of:
 *
 *   - **record** (one undo step, labelled in sentence case, with a `target` node): SET_ENTITY
 *     "Change entity", ADD_* "Add condition" / "Add filter" / … (target: the new node's parent),
 *     REMOVE_NODE "Delete filter group" / … (target: the deleted node), UPDATE_ATTRIBUTES "Select
 *     attributes", ADD_ALL_ATTRIBUTES "Add all attributes", LOAD_FETCHXML "Parse to tree", SET_SORT
 *     "Sort by column", ADD/REMOVE/CLEAR_COLUMN_FILTER(S) "Filter column" / …, SET_COLUMN_CONFIG
 *     "Add columns" / "Remove columns", REORDER_COLUMNS "Reorder columns".
 *     The label is what the toolbar shows: "Undo: Delete filter group (Ctrl+Z)".
 *   - **merge**: UPDATE_NODE (target: the edited node) is recorded with
 *     `mergeKey = nodeId + ":" + sorted update keys`, so typing into one field is one step while the
 *     edits are at most 1 s apart. Another field, another node, any other step, a skip that ends the
 *     merge (the default), undo or redo starts a new step.
 *   - **transaction**: multi-dispatch edits (alias rename, aggregate-off plan, many-to-many join,
 *     remove columns, add related columns, editor-mode sync before Execute / Save / Send) run
 *     between `beginTransaction` and `commitTransaction` (in a `finally`), so they are one step.
 *   - **skip, keeping redo, not ending a merge** (`{ keepRedo: true, endsMerge: false }`):
 *     SYNC_LAYOUT_WITH_FETCHXML. Every UPDATE_NODE sets `layoutNeedsSync`, and an AppShell effect
 *     dispatches the sync after the render, so ending the merge would make each keystroke its own
 *     step, and clearing redo would kill redo after every undo. Keeping redo is safe because the
 *     sync only derives `columnConfig` from the query: redo replaces the whole `columnConfig` with
 *     one that matches its own query.
 *   - **skip, keeping redo**: SELECT_NODE (selection only; it ends the merge, as a user action),
 *     UPDATE_COLUMN_WIDTH and SET_RETRIEVE_ALL_RECORDS (owner decisions: not recorded; safe because
 *     `restore` carries the present widths and `retrieveAllRecords` into every restored state).
 *     SET_LOADED_VIEW (after Save) also keeps redo, and patches the stored snapshots with
 *     `mapSnapshots` so an older query never sits next to the saved view's `originalFetchXml`.
 *   - **skip, clearing redo** (the default): CLEAR_LOADED_VIEW, and any non-recorded change that
 *     `restore` doesn't carry forward.
 *   - **clear** (not undoable; empties undo and redo): NEW_QUERY and LOAD_VIEW (Open view), plus the
 *     baseline loads (connection carry-over, T2T prefill, render-error restart), so their first
 *     Ctrl+Z doesn't blank the tree.
 *
 * Selection: `restore` gets the step's `target`. The store selects it if it exists in the restored
 * snapshot; otherwise its nearest ancestor that does (`nearestExisting`, with the target's path
 * taken from whichever tree holds it: the snapshot, or for a redone delete the present); otherwise
 * the snapshot's own `selectedNodeId`. So undoing a delete from the context menu selects the
 * restored node, even when another node was selected at the time, and redoing it selects its parent.
 *
 * Undo and redo while a transaction is open return the history unchanged rather than throw (a
 * throw inside `useReducer` would crash the provider). `commitTransaction` with no open
 * transaction does throw: that is a programming error in the store, not a user action.
 */

/** The owner's limits (#192 decision 6). */
export const HISTORY_LIMIT = 100;
export const MERGE_WINDOW_MS = 1000;

/** One undoable step. In `past`, `snapshot` is the state before the step; in `future`, after it. */
export interface HistoryEntry<S> {
	/** Stable across undo, redo and merges, so a toast can check that its step is still the last one */
	readonly id: number;
	/** Sentence case, for "Undo: {label}" */
	readonly label: string;
	readonly snapshot: S;
	/** The node to select after undoing or redoing this step, if it still exists */
	readonly target?: string;
}

/** What `restore` is told about the step being undone or redone. */
export interface RestoreContext {
	readonly direction: "undo" | "redo";
	readonly target: string | undefined;
}

/**
 * Builds the state to show from a stored snapshot and the current present: carries values the
 * history doesn't own (column widths, Retrieve all records) and picks the selection from `target`.
 * Given once to `createHistory`, so every undo and redo applies it.
 */
export type Restore<S> = (snapshot: S, present: S, context: RestoreContext) => S;

interface OpenTransaction<S> {
	readonly label: string;
	readonly mergeKey: string | undefined;
	readonly target: string | undefined;
	readonly now: number;
	/** `present` when the outermost transaction began */
	readonly base: S;
	/** Nesting depth; only the outermost commit records */
	readonly depth: number;
	/** Set once a `record` inside the transaction changed `present` */
	readonly recorded: boolean;
}

export interface History<S> {
	readonly past: readonly HistoryEntry<S>[];
	readonly present: S;
	readonly future: readonly HistoryEntry<S>[];
	/** The step the next same-key edit may merge into; null once anything else happens */
	readonly openMerge: { readonly key: string; readonly at: number } | null;
	readonly transaction: OpenTransaction<S> | null;
	readonly nextId: number;
	readonly limit: number;
	readonly mergeWindowMs: number;
	readonly restore: Restore<S> | null;
}

export interface HistoryOptions<S> {
	/** Most undo steps kept; the oldest drop off. Default 100. */
	limit?: number;
	/** Same-key edits at most this far apart merge. Default 1000 ms. */
	mergeWindowMs?: number;
	/** Applied to every undo and redo; without it the stored snapshot is shown as is */
	restore?: Restore<S>;
}

export interface RecordOptions {
	label: string;
	/** Edits with the same key (for example `nodeId:field`) merge within the merge window */
	mergeKey?: string;
	/** The node to select after undoing or redoing this step */
	target?: string;
	now: number;
}

export interface SkipOptions {
	/**
	 * Keep redo. Pass it only when redoing can't bring back what this change replaced: the change is
	 * selection-only, derived from the rest of the state, or carried forward by `restore`.
	 */
	keepRedo?: boolean;
	/**
	 * End any merge in progress. Default true, for user actions such as selecting another node.
	 * Pass false for derived, effect-driven changes (the layout sync after each edit).
	 */
	endsMerge?: boolean;
}

export function createHistory<S>(initial: S, options: HistoryOptions<S> = {}): History<S> {
	const limit = options.limit ?? HISTORY_LIMIT;
	const mergeWindowMs = options.mergeWindowMs ?? MERGE_WINDOW_MS;
	if (!Number.isInteger(limit) || limit < 1)
		throw new RangeError(`History limit must be a positive integer, got ${limit}`);
	if (!(mergeWindowMs >= 0))
		throw new RangeError(`Merge window must be 0 or more, got ${mergeWindowMs}`);
	return {
		past: [],
		present: initial,
		future: [],
		openMerge: null,
		transaction: null,
		nextId: 1,
		limit,
		mergeWindowMs,
		restore: options.restore ?? null,
	};
}

/**
 * Pushes a step from `before` to `h.present`, or merges it into the top step when the key matches
 * and the last merge was at most the merge window ago (the window slides with each merged edit).
 * `openMerge` is only ever set by a push, and cleared before the top entry can change otherwise,
 * so a match always refers to the top entry.
 */
function pushStep<S>(
	h: History<S>,
	before: S,
	{ label, mergeKey, target, now }: RecordOptions,
): History<S> {
	if (
		mergeKey !== undefined &&
		h.openMerge?.key === mergeKey &&
		now - h.openMerge.at <= h.mergeWindowMs
	) {
		return { ...h, future: [], openMerge: { key: mergeKey, at: now } };
	}
	const entry: HistoryEntry<S> =
		target === undefined
			? { id: h.nextId, label, snapshot: before }
			: { id: h.nextId, label, snapshot: before, target };
	const past =
		h.past.length >= h.limit
			? [...h.past.slice(h.past.length - h.limit + 1), entry]
			: [...h.past, entry];
	return {
		...h,
		past,
		future: [],
		openMerge: mergeKey === undefined ? null : { key: mergeKey, at: now },
		nextId: h.nextId + 1,
	};
}

/** Records `next` as one undoable step. A no-op (`next === present`) records nothing. Clears redo. */
export function record<S>(h: History<S>, next: S, options: RecordOptions): History<S> {
	if (Object.is(next, h.present)) return h;
	if (h.transaction) {
		return { ...h, present: next, transaction: { ...h.transaction, recorded: true } };
	}
	return pushStep({ ...h, present: next }, h.present, options);
}

/**
 * A change that isn't an undo step (selection, layout sync, column width, Retrieve all records).
 * It replaces `present`. By default it clears redo, because every redo snapshot predates this
 * change and redoing would silently revert it, and ends any merge; see `SkipOptions`.
 */
export function skip<S>(h: History<S>, next: S, options: SkipOptions = {}): History<S> {
	if (Object.is(next, h.present)) return h;
	const future = options.keepRedo ? h.future : [];
	const openMerge = options.endsMerge === false ? h.openMerge : null;
	return { ...h, present: next, future, openMerge };
}

/**
 * Starts grouping changes into one step. Nesting is supported: inner begins only deepen the
 * transaction, and the outermost label, merge key and target are used.
 */
export function beginTransaction<S>(h: History<S>, options: RecordOptions): History<S> {
	if (h.transaction) {
		return { ...h, transaction: { ...h.transaction, depth: h.transaction.depth + 1 } };
	}
	return {
		...h,
		transaction: {
			label: options.label,
			mergeKey: options.mergeKey,
			target: options.target,
			now: options.now,
			base: h.present,
			depth: 1,
			recorded: false,
		},
	};
}

/**
 * Ends the innermost transaction. The outermost commit records one step if a `record` inside it
 * changed `present` and `present` differs from where it began (merging by its key, unless a skip
 * inside it ended the merge); otherwise it records nothing and ends any merge.
 * Throws if no transaction is open: that is a bug in the caller.
 */
export function commitTransaction<S>(h: History<S>): History<S> {
	const tx = h.transaction;
	if (!tx) throw new Error("commitTransaction: no transaction is open");
	if (tx.depth > 1) return { ...h, transaction: { ...tx, depth: tx.depth - 1 } };
	const closed: History<S> = { ...h, transaction: null };
	if (!tx.recorded || Object.is(h.present, tx.base)) return { ...closed, openMerge: null };
	return pushStep(closed, tx.base, tx);
}

/** Goes back one step, through `restore`. Unchanged at the start or while a transaction is open. */
export function undo<S>(h: History<S>): History<S> {
	const entry = h.past[h.past.length - 1];
	if (!entry || h.transaction) return h;
	return {
		...h,
		past: h.past.slice(0, -1),
		present: h.restore
			? h.restore(entry.snapshot, h.present, { direction: "undo", target: entry.target })
			: entry.snapshot,
		future: [...h.future, { ...entry, snapshot: h.present }],
		openMerge: null,
	};
}

/** Goes forward one undone step, through `restore`. Unchanged at the end or while a transaction is open. */
export function redo<S>(h: History<S>): History<S> {
	const entry = h.future[h.future.length - 1];
	if (!entry || h.transaction) return h;
	return {
		...h,
		past: [...h.past, { ...entry, snapshot: h.present }],
		present: h.restore
			? h.restore(entry.snapshot, h.present, { direction: "redo", target: entry.target })
			: entry.snapshot,
		future: h.future.slice(0, -1),
		openMerge: null,
	};
}

/**
 * Sets `present` and empties undo and redo (New, Open view, baseline loads). Inside a transaction
 * the transaction stays open but restarts from `present`, so it records only later changes.
 */
export function clear<S>(h: History<S>, present: S): History<S> {
	const transaction = h.transaction ? { ...h.transaction, base: present, recorded: false } : null;
	return { ...h, past: [], present, future: [], openMerge: null, transaction };
}

/** Rewrites every stored snapshot (not `present`), for example to patch `loadedView` after Save. */
export function mapSnapshots<S>(h: History<S>, fn: (snapshot: S) => S): History<S> {
	const map = (e: HistoryEntry<S>): HistoryEntry<S> => {
		const snapshot = fn(e.snapshot);
		return Object.is(snapshot, e.snapshot) ? e : { ...e, snapshot };
	};
	const transaction = h.transaction ? { ...h.transaction, base: fn(h.transaction.base) } : null;
	return { ...h, past: h.past.map(map), future: h.future.map(map), transaction };
}

/**
 * The first id on `path` (a target followed by its ancestors, nearest first) that `exists`, or
 * null. The store builds `path` from whichever tree holds the target and checks existence in the
 * restored snapshot.
 */
export function nearestExisting(
	path: readonly string[],
	exists: (id: string) => boolean,
): string | null {
	return path.find(exists) ?? null;
}

export const canUndo = <S>(h: History<S>): boolean => h.past.length > 0 && !h.transaction;
export const canRedo = <S>(h: History<S>): boolean => h.future.length > 0 && !h.transaction;
export const undoLabel = <S>(h: History<S>): string | null =>
	h.past[h.past.length - 1]?.label ?? null;
export const redoLabel = <S>(h: History<S>): string | null =>
	h.future[h.future.length - 1]?.label ?? null;
/** The id of the step Undo would undo, or null */
export const lastEntryId = <S>(h: History<S>): number | null =>
	h.past[h.past.length - 1]?.id ?? null;
