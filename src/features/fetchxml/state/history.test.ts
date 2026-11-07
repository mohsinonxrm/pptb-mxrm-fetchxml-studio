import { describe, expect, it } from "vitest";
import {
	HISTORY_LIMIT,
	MERGE_WINDOW_MS,
	beginTransaction,
	canRedo,
	canUndo,
	clear,
	commitTransaction,
	createHistory,
	lastEntryId,
	mapSnapshots,
	nearestExisting,
	record,
	redo,
	redoLabel,
	skip,
	undo,
	undoLabel,
	type History,
	type RestoreContext,
} from "./history";

// #192: the undo history core, before it is wired into builderStore

interface Snap {
	text: string;
	width: number;
	selected: string | null;
}

const snap = (text: string, extra: Partial<Snap> = {}): Snap => ({ text, width: 100, selected: null, ...extra });

const at = (h: History<Snap>, text: string, now = 0, mergeKey?: string, label = `Edit ${text}`) =>
	record(h, { ...h.present, text }, { label, mergeKey, now });

describe("record, undo and redo", () => {
	it("starts empty", () => {
		const h = createHistory(snap("a"));
		expect(canUndo(h)).toBe(false);
		expect(canRedo(h)).toBe(false);
		expect(undoLabel(h)).toBeNull();
		expect(redoLabel(h)).toBeNull();
		expect(lastEntryId(h)).toBeNull();
	});

	it("restores the same snapshot references both ways", () => {
		const s0 = snap("a");
		const s1 = snap("b");
		let h = createHistory(s0);
		h = record(h, s1, { label: "Add condition", now: 0 });
		h = undo(h);
		expect(h.present).toBe(s0);
		h = redo(h);
		expect(h.present).toBe(s1);
	});

	it("records nothing for a no-op", () => {
		const h = createHistory(snap("a"));
		expect(record(h, h.present, { label: "Edit", now: 0 })).toBe(h);
	});

	it("does nothing at the boundaries", () => {
		let h = createHistory(snap("a"));
		expect(undo(h)).toBe(h);
		expect(redo(h)).toBe(h);
		h = at(h, "b");
		h = undo(h);
		expect(undo(h)).toBe(h);
		h = redo(h);
		expect(redo(h)).toBe(h);
		expect(h.present.text).toBe("b");
	});

	it("walks several steps back and forward", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = at(h, "c");
		h = at(h, "d");
		h = undo(undo(undo(h)));
		expect(h.present.text).toBe("a");
		expect(canUndo(h)).toBe(false);
		h = redo(redo(h));
		expect(h.present.text).toBe("c");
		expect(canUndo(h)).toBe(true);
		expect(canRedo(h)).toBe(true);
	});

	it("drops redo when a change is recorded after undo", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = at(h, "c");
		h = undo(h);
		expect(canRedo(h)).toBe(true);
		h = at(h, "x");
		expect(canRedo(h)).toBe(false);
		expect(redo(h)).toBe(h);
		h = undo(h);
		expect(h.present.text).toBe("b");
	});

	it("applies the restore given to createHistory on every undo and redo", () => {
		const calls: RestoreContext[] = [];
		const restore = (s: Snap, present: Snap, context: RestoreContext): Snap => {
			calls.push(context);
			return { ...s, width: present.width };
		};
		let h = createHistory(snap("a"), { restore });
		h = at(h, "b");
		h = skip(h, { ...h.present, width: 250 }, { keepRedo: true });
		h = undo(h);
		expect(h.present).toEqual(snap("a", { width: 250 }));
		h = redo(h);
		expect(h.present).toEqual(snap("b", { width: 250 }));
		expect(calls.map((c) => c.direction)).toEqual(["undo", "redo"]);
	});
});

describe("selection target", () => {
	it("stores the target with the step and hands it to restore both ways", () => {
		const seen: RestoreContext[] = [];
		const restore = (s: Snap, _present: Snap, context: RestoreContext): Snap => {
			seen.push(context);
			return s;
		};
		let h = createHistory(snap("a", { selected: "other" }), { restore });
		// A context-menu delete of n2 while "other" stays selected
		h = record(h, snap("b", { selected: "other" }), { label: "Delete filter", target: "n2", now: 0 });
		expect(h.past[0].target).toBe("n2");
		h = redo(undo(h));
		expect(seen).toEqual([
			{ direction: "undo", target: "n2" },
			{ direction: "redo", target: "n2" },
		]);
	});

	it("gives an undefined target when the step has none", () => {
		let seen: RestoreContext | undefined;
		let h = createHistory(snap("a"), {
			restore: (s, _p, c) => {
				seen = c;
				return s;
			},
		});
		h = undo(at(h, "b"));
		expect(seen).toEqual({ direction: "undo", target: undefined });
		expect("target" in h.future[0]).toBe(false);
	});

	it("uses the outermost transaction's target", () => {
		let h = createHistory(snap("a"));
		h = beginTransaction(h, { label: "Join", target: "link-1", now: 0 });
		h = beginTransaction(h, { label: "Inner", target: "inner", now: 0 });
		h = at(h, "b");
		h = commitTransaction(commitTransaction(h));
		expect(h.past[0].target).toBe("link-1");
	});

	it("nearestExisting picks the target, else its nearest existing ancestor, else null", () => {
		const inSnapshot = new Set(["root", "filter-1"]);
		const exists = (id: string) => inSnapshot.has(id);
		expect(nearestExisting(["filter-1", "root"], exists)).toBe("filter-1");
		expect(nearestExisting(["cond-9", "filter-2", "filter-1", "root"], exists)).toBe("filter-1");
		expect(nearestExisting(["gone", "also-gone"], exists)).toBeNull();
		expect(nearestExisting([], exists)).toBeNull();
	});
});

describe("depth", () => {
	it(`keeps the last ${HISTORY_LIMIT} steps and drops the oldest`, () => {
		let h = createHistory(snap("0"));
		for (let i = 1; i <= HISTORY_LIMIT + 5; i++) h = at(h, String(i), i * 10_000);
		expect(h.past).toHaveLength(HISTORY_LIMIT);
		expect(undoLabel(h)).toBe(`Edit ${HISTORY_LIMIT + 5}`);
		for (let i = 0; i < HISTORY_LIMIT; i++) h = undo(h);
		expect(canUndo(h)).toBe(false);
		// Steps 1-5 dropped off, so the oldest state left is after step 5
		expect(h.present.text).toBe("5");
	});

	it("keeps exactly the limit without dropping", () => {
		let h = createHistory(snap("0"));
		for (let i = 1; i <= HISTORY_LIMIT; i++) h = at(h, String(i), i * 10_000);
		expect(h.past).toHaveLength(HISTORY_LIMIT);
		expect(h.past[0].snapshot.text).toBe("0");
	});

	it("takes a custom limit and rejects a bad one", () => {
		let h = createHistory(snap("0"), { limit: 2 });
		h = at(h, "1");
		h = at(h, "2");
		h = at(h, "3");
		expect(h.past.map((e) => e.snapshot.text)).toEqual(["1", "2"]);
		expect(() => createHistory(snap("0"), { limit: 0 })).toThrow(RangeError);
	});
});

describe("merging", () => {
	const key = "node-1:value";

	it("merges same-key edits within the window into one step", () => {
		let h = createHistory(snap(""));
		h = at(h, "h", 0, key, "Edit condition");
		h = at(h, "he", 400, key, "Edit condition");
		h = at(h, "hel", 900, key, "Edit condition");
		expect(h.past).toHaveLength(1);
		expect(h.present.text).toBe("hel");
		h = undo(h);
		expect(h.present.text).toBe("");
		h = redo(h);
		expect(h.present.text).toBe("hel");
	});

	it("merges at exactly 1 s", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = at(h, "ab", MERGE_WINDOW_MS, key);
		expect(h.past).toHaveLength(1);
	});

	it("starts a new step after 1 s", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = at(h, "ab", MERGE_WINDOW_MS + 1, key);
		expect(h.past).toHaveLength(2);
		h = undo(h);
		expect(h.present.text).toBe("a");
	});

	it("slides the window with each merged edit", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = at(h, "ab", 800, key);
		h = at(h, "abc", 1600, key);
		expect(h.past).toHaveLength(1);
	});

	it("starts a new step for a different field or node", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, "node-1:value");
		h = at(h, "ab", 100, "node-1:operator");
		h = at(h, "abc", 200, "node-2:value");
		expect(h.past).toHaveLength(3);
	});

	it("starts a new step after an edit without a key", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = at(h, "ab", 100);
		h = at(h, "abc", 200, key);
		expect(h.past).toHaveLength(3);
	});

	it("starts a new step after a skipped change, an undo or a redo", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = skip(h, { ...h.present, selected: "node-2" }, { keepRedo: true });
		h = at(h, "ab", 100, key);
		expect(h.past).toHaveLength(2);

		h = undo(h);
		h = at(h, "ax", 200, key);
		expect(h.past).toHaveLength(2);

		h = undo(h);
		h = redo(h);
		h = at(h, "axy", 300, key);
		expect(h.past).toHaveLength(3);
	});

	it("keeps typing one step when the layout sync runs after each keystroke", () => {
		// UPDATE_NODE sets layoutNeedsSync; an AppShell effect then dispatches SYNC_LAYOUT_WITH_FETCHXML
		let h = createHistory(snap(""));
		let text = "";
		for (let i = 0; i < 5; i++) {
			text += "x";
			h = at(h, text, i * 150, key, "Edit condition");
			h = skip(h, { ...h.present, width: h.present.width + 1 }, { keepRedo: true, endsMerge: false });
		}
		expect(h.past).toHaveLength(1);
		expect(h.present.text).toBe("xxxxx");
		expect(undo(h).present.text).toBe("");
	});

	it("keeps the first edit's id and label", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key, "Edit condition");
		const id = lastEntryId(h);
		h = at(h, "ab", 100, key, "Edit something else");
		expect(lastEntryId(h)).toBe(id);
		expect(undoLabel(h)).toBe("Edit condition");
	});
});

describe("transactions", () => {
	it("groups several changes into one step", () => {
		let h = createHistory(snap("a"));
		h = beginTransaction(h, { label: "Parse to tree", now: 0 });
		expect(canUndo(h)).toBe(false);
		h = at(h, "b");
		h = at(h, "c");
		h = skip(h, { ...h.present, selected: "root" }, { keepRedo: true });
		h = commitTransaction(h);
		expect(h.past).toHaveLength(1);
		expect(undoLabel(h)).toBe("Parse to tree");
		h = undo(h);
		expect(h.present).toEqual(snap("a"));
		h = redo(h);
		expect(h.present).toEqual(snap("c", { selected: "root" }));
	});

	it("records nothing for an empty transaction", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = undo(h);
		const before = h;
		h = commitTransaction(beginTransaction(h, { label: "Add all attributes", now: 0 }));
		expect(h.past).toEqual(before.past);
		expect(h.future).toEqual(before.future);
		expect(canRedo(h)).toBe(true);
		expect(h.transaction).toBeNull();
	});

	it("records nothing when the changes cancel out", () => {
		const s0 = snap("a");
		let h = createHistory(s0);
		h = beginTransaction(h, { label: "Rename alias", now: 0 });
		h = at(h, "b");
		h = record(h, s0, { label: "x", now: 0 });
		h = commitTransaction(h);
		expect(canUndo(h)).toBe(false);
	});

	it("records nothing when only skipped changes happened", () => {
		let h = createHistory(snap("a"));
		h = beginTransaction(h, { label: "Sync layout", now: 0 });
		h = skip(h, { ...h.present, width: 300 }, { keepRedo: true });
		h = commitTransaction(h);
		expect(canUndo(h)).toBe(false);
		expect(h.present.width).toBe(300);
	});

	it("nests: only the outermost commit records, with the outermost label", () => {
		let h = createHistory(snap("a"));
		h = beginTransaction(h, { label: "Add related columns", now: 0 });
		h = at(h, "b");
		h = beginTransaction(h, { label: "Inner", now: 0 });
		h = at(h, "c");
		h = commitTransaction(h);
		expect(h.past).toHaveLength(0);
		expect(h.transaction).not.toBeNull();
		h = at(h, "d");
		h = commitTransaction(h);
		expect(h.past).toHaveLength(1);
		expect(undoLabel(h)).toBe("Add related columns");
		expect(undo(h).present.text).toBe("a");
	});

	it("clears redo when it records", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = undo(h);
		h = beginTransaction(h, { label: "Join", now: 0 });
		h = at(h, "x");
		expect(canRedo(h)).toBe(false);
		h = commitTransaction(h);
		expect(redo(h)).toBe(h);
	});

	it("merges a keyed transaction like a keyed record", () => {
		let h = createHistory(snap(""));
		for (const [text, now] of [
			["a", 0],
			["ab", 500],
		] as const) {
			h = beginTransaction(h, { label: "Rename alias", mergeKey: "attr-1:alias", now });
			h = at(h, text, now);
			h = commitTransaction(h);
		}
		expect(h.past).toHaveLength(1);
		expect(undo(h).present.text).toBe("");
	});

	it("throws clearly on a commit without a begin, and on one commit too many", () => {
		let h = createHistory(snap("a"));
		expect(() => commitTransaction(h)).toThrow(/no transaction is open/);
		h = commitTransaction(beginTransaction(h, { label: "T", now: 0 }));
		expect(() => commitTransaction(h)).toThrow(/no transaction is open/);
	});

	it("leaves undo and redo unchanged while a transaction is open, and says so", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = at(h, "c");
		h = undo(h);
		h = beginTransaction(h, { label: "T", now: 0 });
		expect(canUndo(h)).toBe(false);
		expect(canRedo(h)).toBe(false);
		expect(undo(h)).toBe(h);
		expect(redo(h)).toBe(h);
		h = commitTransaction(h);
		expect(canUndo(h)).toBe(true);
		expect(canRedo(h)).toBe(true);
	});

	it("doesn't merge across a selection change inside a transaction", () => {
		const key = "n1:value";
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = beginTransaction(h, { label: "Select", now: 100 });
		h = skip(h, { ...h.present, selected: "n2" }, { keepRedo: true });
		h = skip(h, { ...h.present, selected: "n1" }, { keepRedo: true });
		h = commitTransaction(h);
		h = at(h, "ab", 200, key);
		expect(h.past).toHaveLength(2);
	});

	it("ends a merge on an outermost commit that records nothing", () => {
		const key = "n1:value";
		let h = createHistory(snap(""));
		h = at(h, "a", 0, key);
		h = commitTransaction(beginTransaction(h, { label: "Empty", now: 100 }));
		expect(h.openMerge).toBeNull();
		h = at(h, "ab", 200, key);
		expect(h.past).toHaveLength(2);
	});

	it("restarts from a clear inside it", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = beginTransaction(h, { label: "T", now: 0 });
		h = at(h, "c");
		h = clear(h, snap("fresh"));
		h = commitTransaction(h);
		expect(canUndo(h)).toBe(false);
		expect(h.present.text).toBe("fresh");

		h = beginTransaction(h, { label: "T2", now: 0 });
		h = clear(h, snap("again"));
		h = at(h, "after");
		h = commitTransaction(h);
		expect(undoLabel(h)).toBe("T2");
		expect(undo(h).present.text).toBe("again");
	});
});

describe("skip and redo", () => {
	it("changes present without adding a step", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = skip(h, { ...h.present, width: 300 });
		expect(h.past).toHaveLength(1);
		expect(h.present.width).toBe(300);
	});

	it("ignores a no-op", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = undo(h);
		expect(skip(h, h.present)).toBe(h);
		expect(canRedo(h)).toBe(true);
	});

	it("clears redo by default, so redo can't bring back what it replaced", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = undo(h);
		h = skip(h, { ...h.present, width: 300 });
		expect(canRedo(h)).toBe(false);
		expect(redo(h).present.width).toBe(300);
	});

	it("keeps redo when asked, for selection-only changes", () => {
		let h = createHistory(snap("a", { selected: "x" }));
		h = record(h, snap("b", { selected: "y" }), { label: "Add condition", now: 0 });
		h = undo(h);
		h = skip(h, { ...h.present, selected: "z" }, { keepRedo: true });
		expect(redoLabel(h)).toBe("Add condition");
		h = redo(h);
		expect(h.present).toEqual(snap("b", { selected: "y" }));
		h = undo(h);
		// Undo returns to the state before the step, with the selection it had
		expect(h.present).toEqual(snap("a", { selected: "z" }));
	});

	it("with keepRedo and the history's restore carrying the change, redo keeps it", () => {
		const restore = (s: Snap, present: Snap): Snap => ({ ...s, width: present.width });
		let h = createHistory(snap("a"), { restore });
		h = at(h, "b");
		h = undo(h);
		h = skip(h, { ...h.present, width: 300 }, { keepRedo: true });
		h = redo(h);
		expect(h.present).toEqual(snap("b", { width: 300 }));
	});

	it("ends a merge by default, but not with endsMerge: false", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, "k");
		expect(skip(h, { ...h.present, width: 1 }).openMerge).toBeNull();
		expect(skip(h, { ...h.present, width: 1 }, { endsMerge: false }).openMerge).toEqual({ key: "k", at: 0 });
	});
});

describe("clear", () => {
	it("empties undo and redo and sets present", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = at(h, "c");
		h = undo(h);
		const opened = snap("view");
		h = clear(h, opened);
		expect(h.present).toBe(opened);
		expect(canUndo(h)).toBe(false);
		expect(canRedo(h)).toBe(false);
		expect(undoLabel(h)).toBeNull();
		expect(redoLabel(h)).toBeNull();
	});

	it("ends a merge, so the next edit is its own step", () => {
		let h = createHistory(snap(""));
		h = at(h, "a", 0, "k");
		h = clear(h, snap("new"));
		expect(h.openMerge).toBeNull();
		h = at(h, "newx", 100, "k");
		expect(h.past).toHaveLength(1);
		expect(undo(h).present.text).toBe("new");
	});

	it("keeps entry ids increasing, so a stale toast never matches", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		const oldId = lastEntryId(h);
		h = clear(h, snap("new"));
		h = at(h, "c");
		expect(lastEntryId(h)).not.toBe(oldId);
	});
});

describe("labels", () => {
	it("names the step undo and redo would take", () => {
		let h = createHistory(snap("a"));
		h = record(h, snap("b"), { label: "Add filter", now: 0 });
		h = record(h, snap("c"), { label: "Delete filter group", now: 0 });
		expect(undoLabel(h)).toBe("Delete filter group");
		h = undo(h);
		expect(undoLabel(h)).toBe("Add filter");
		expect(redoLabel(h)).toBe("Delete filter group");
		h = undo(h);
		expect(undoLabel(h)).toBeNull();
		expect(redoLabel(h)).toBe("Add filter");
		h = redo(h);
		expect(undoLabel(h)).toBe("Add filter");
		expect(redoLabel(h)).toBe("Delete filter group");
		h = redo(h);
		expect(undoLabel(h)).toBe("Delete filter group");
		expect(redoLabel(h)).toBeNull();
	});

	it("keeps a step's id across undo and redo", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		const id = lastEntryId(h);
		h = redo(undo(h));
		expect(lastEntryId(h)).toBe(id);
	});
});

describe("mapSnapshots", () => {
	it("rewrites stored snapshots but not present, reusing unchanged entries", () => {
		let h = createHistory(snap("a"));
		h = at(h, "b");
		h = at(h, "c");
		h = undo(h);
		const untouched = h.past[0];
		h = mapSnapshots(h, (s) => (s.text === "c" ? { ...s, selected: "patched" } : s));
		expect(h.future[0].snapshot.selected).toBe("patched");
		expect(h.past[0]).toBe(untouched);
		expect(h.present.selected).toBeNull();
	});

	it("patches an open transaction's base, so its step undoes to the patched state", () => {
		let h = createHistory(snap("a"));
		h = beginTransaction(h, { label: "T", now: 0 });
		h = at(h, "b");
		h = mapSnapshots(h, (s) => (s.text === "a" ? { ...s, selected: "patched" } : s));
		h = commitTransaction(h);
		expect(undo(h).present).toEqual(snap("a", { selected: "patched" }));
	});
});
