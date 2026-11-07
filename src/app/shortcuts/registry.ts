/**
 * The keyboard shortcut table (#186): one list of ids, labels and keys, plus the matcher. Nothing
 * here listens to the page; the global handler and the shortcuts dialog read this table later.
 *
 * `scope: "global"` entries are handled by the one document-level handler. `"documented"` entries
 * are handled locally (the tree, the grid, a popup, the editor) and are listed for the shortcuts
 * dialog and the README only.
 *
 * Owner decision on #186 Q5 (model-driven app keys): Ctrl+F6 and Ctrl+Shift+F6 move between areas,
 * F2 goes to the newest notification. They type no text, so they work in fields too. F2 also works
 * inside Monaco: Monaco binds it to "rename symbol", but FetchXML (XML) has no rename provider, so
 * nothing is lost. The handler must run before Monaco's own key handling for this to hold.
 * Not included on purpose: Ctrl+B (decided no), F5 (decided no) and anything PPTB 1.2.7 owns
 * (PPTB_RESERVED).
 */

export type ShortcutScope = "global" | "documented";
export type ShortcutPlatform = "mac" | "other";

export interface ShortcutEntry {
	id: string;
	/** Shown in tooltips, menus, the dialog and the README. Windows spelling, e.g. "Ctrl+Enter". */
	keys: string[];
	/** The `aria-keyshortcuts` value: one space-separated combination per key, e.g. "Control+Enter". */
	aria: string;
	label: string;
	scope: ShortcutScope;
	/** May the shortcut fire while focus is in a text field, contenteditable or Monaco? */
	inEditable: boolean;
}

export const SHORTCUTS: readonly ShortcutEntry[] = [
	{ id: "run", label: "Run query", keys: ["Ctrl+Enter"], aria: "Control+Enter", scope: "global", inEditable: true },
	{ id: "save", label: "Save view", keys: ["Ctrl+S"], aria: "Control+S", scope: "global", inEditable: true },
	{ id: "shortcuts", label: "Keyboard shortcuts", keys: ["Ctrl+/"], aria: "Control+/", scope: "global", inEditable: false },
	// Text fields and Monaco keep their own undo (owner decision on #192).
	{ id: "undo", label: "Undo", keys: ["Ctrl+Z"], aria: "Control+Z", scope: "global", inEditable: false },
	{ id: "redo", label: "Redo", keys: ["Ctrl+Y", "Ctrl+Shift+Z"], aria: "Control+Y Control+Shift+Z", scope: "global", inEditable: false },
	// Model-driven app keys (#186 Q5). Function keys, so they are safe in text fields and Monaco.
	{ id: "nextArea", label: "Go to the next area", keys: ["Ctrl+F6"], aria: "Control+F6", scope: "global", inEditable: true },
	{ id: "previousArea", label: "Go to the previous area", keys: ["Ctrl+Shift+F6"], aria: "Control+Shift+F6", scope: "global", inEditable: true },
	{ id: "toasts", label: "Go to the newest notification", keys: ["F2"], aria: "F2", scope: "global", inEditable: true },
	// Documented only: handled where the focus is.
	{ id: "removeNode", label: "Remove the node", keys: ["Delete"], aria: "Delete", scope: "documented", inEditable: false },
	{ id: "openMenu", label: "Open the row or cell menu", keys: ["Shift+F10", "Menu"], aria: "Shift+F10 ContextMenu", scope: "documented", inEditable: false },
	{ id: "closePopup", label: "Close the top-most popup or drawer", keys: ["Escape"], aria: "Escape", scope: "documented", inEditable: true },
	{ id: "formatXml", label: "Format FetchXML", keys: ["Shift+Alt+F"], aria: "Shift+Alt+F", scope: "documented", inEditable: true },
	{ id: "copyGrid", label: "Copy", keys: ["Ctrl+C"], aria: "Control+C", scope: "documented", inEditable: false },
];

/** Combinations the PPTB 1.2.7 host menu owns; no global shortcut may use one. */
export const PPTB_RESERVED: readonly string[] = [
	"Ctrl+R",
	"Ctrl+H",
	"Ctrl+,",
	"Ctrl+0",
	"Ctrl+Plus",
	"Ctrl+=",
	"Ctrl+-",
	"F11",
	"Ctrl+Shift+F",
	"Ctrl+Shift+T",
	"Ctrl+Shift+I",
	// Edit menu
	"Ctrl+X",
	"Ctrl+C",
	"Ctrl+V",
	"Ctrl+A",
];

export interface KeyEventLike {
	key: string;
	ctrlKey: boolean;
	metaKey: boolean;
	shiftKey: boolean;
	altKey: boolean;
}

interface Combo {
	key: string;
	ctrl: boolean;
	shift: boolean;
	alt: boolean;
}

const MODIFIERS: Record<string, "ctrl" | "shift" | "alt"> = {
	ctrl: "ctrl",
	control: "ctrl",
	cmd: "ctrl",
	meta: "ctrl",
	shift: "shift",
	alt: "alt",
	option: "alt",
};

const KEY_ALIASES: Record<string, string> = { plus: "+", esc: "escape", menu: "contextmenu", del: "delete" };

/** Parses "Ctrl+Shift+Z" or "Control+Shift+Z" (the last segment is the key; write "Plus", not "+"). */
function parseCombo(text: string): Combo {
	const parts = text.split("+");
	const last = parts.pop() as string;
	const combo: Combo = { key: "", ctrl: false, shift: false, alt: false };
	for (const part of parts) {
		const mod = MODIFIERS[part.toLowerCase()];
		if (!mod) throw new Error(`Unknown modifier "${part}" in "${text}"`);
		combo[mod] = true;
	}
	const lower = last.toLowerCase();
	combo.key = KEY_ALIASES[lower] ?? lower;
	return combo;
}

/** A canonical spelling for comparing combinations: "Control+shift+z" and "Ctrl+Shift+Z" are the same. */
export function normalizeCombo(text: string): string {
	const c = parseCombo(text);
	return [c.ctrl ? "ctrl" : "", c.alt ? "alt" : "", c.shift ? "shift" : "", c.key].filter(Boolean).join("+");
}

export function getEntry(id: string): ShortcutEntry {
	const entry = SHORTCUTS.find((s) => s.id === id);
	if (!entry) throw new Error(`Unknown shortcut "${id}"`);
	return entry;
}

export function detectPlatform(): ShortcutPlatform {
	if (typeof navigator === "undefined") return "other";
	const nav = navigator as Navigator & { userAgentData?: { platform?: string } };
	const platform = nav.userAgentData?.platform || nav.platform || "";
	return /mac|iphone|ipad|ipod/i.test(platform) ? "mac" : "other";
}

/**
 * Does the key event press this shortcut? The primary modifier is Ctrl, or Cmd on a Mac (the other
 * one must be up). Shift and Alt must match exactly, so Ctrl+Shift+Z is not Ctrl+Z. `event.key` is
 * compared without case (Shift turns "z" into "Z"). "/" is its own key: Shift+/ ("?") doesn't match.
 */
export function matches(event: KeyEventLike, id: string, platform: ShortcutPlatform = detectPlatform()): boolean {
	const entry = getEntry(id);
	const primary = platform === "mac" ? event.metaKey && !event.ctrlKey : event.ctrlKey && !event.metaKey;
	const noPrimary = !event.metaKey && !event.ctrlKey;
	const key = event.key.toLowerCase();
	return entry.aria.split(" ").some((text) => {
		const c = parseCombo(text);
		return c.key === key && c.shift === event.shiftKey && c.alt === event.altKey && (c.ctrl ? primary : noPrimary);
	});
}

/** Display keys for the platform: "Ctrl" becomes "⌘" on a Mac and "Alt" becomes "Option". */
export function formatKeys(entry: ShortcutEntry, platform: ShortcutPlatform = detectPlatform()): string[] {
	if (platform !== "mac") return [...entry.keys];
	return entry.keys.map((k) =>
		k
			.split("+")
			.map((part) => (part === "Ctrl" ? "⌘" : part === "Alt" ? "Option" : part))
			.join("+"),
	);
}

const NON_TEXT_INPUTS = new Set(["button", "checkbox", "radio", "submit", "reset", "range", "color", "file", "image"]);

/** Is focus somewhere the user types: a text input, a textarea, contenteditable, or Monaco? */
export function isEditableTarget(target: EventTarget | null): boolean {
	const el = target as Element | null;
	if (!el || typeof el.tagName !== "string") return false;
	const tag = el.tagName.toLowerCase();
	if (tag === "textarea") return true;
	if (tag === "input") return !NON_TEXT_INPUTS.has(((el as HTMLInputElement).type || "text").toLowerCase());
	if ((el as HTMLElement).isContentEditable) return true;
	return !!el.closest?.('[contenteditable]:not([contenteditable="false"]), .monaco-editor');
}
