/**
 * A keyboard- and pointer-operable pane splitter (#234), following the WAI-ARIA window splitter
 * pattern: https://www.w3.org/WAI/ARIA/apg/patterns/windowsplitter/
 *
 * The component is controlled: it reports every size change through `onChange` and reports the
 * settled size through `onCommit`, which is the moment to persist it: pointer release (only if the
 * drag changed the size), a double-click reset (only if it changed), Enter (at once), or 250 ms
 * after the last arrow, Home or End key (also flushed if the splitter unmounts first). Keys are
 * ignored while a drag is active, and only the primary pointer drags. The caller owns the size and
 * the layout it drives.
 */
import { useEffect, useRef, useState, type KeyboardEvent, type PointerEvent } from "react";
import {
	createFocusOutlineStyle,
	makeStyles,
	mergeClasses,
	Tooltip,
	tokens,
} from "@fluentui/react-components";

/** Arrow key step in pixels; Shift uses the large step */
export const SPLITTER_STEP = 16;
export const SPLITTER_LARGE_STEP = 64;
/** How long after the last key press the size counts as settled */
export const SPLITTER_COMMIT_DELAY_MS = 250;

export interface PaneSplitterProps {
	/** "vertical": a vertical line that resizes widths (Left/Right keys); "horizontal": a horizontal line that resizes heights (Up/Down keys) */
	orientation: "vertical" | "horizontal";
	/** The controlled pane's size in pixels */
	value: number;
	min: number;
	max: number;
	/** The size Enter and double-click reset to */
	defaultValue: number;
	/** Called with the clamped size on every change */
	onChange: (value: number) => void;
	/** Called once the change has settled: pointer release or lost capture after a change, a reset, Enter, or 250 ms after the last arrow, Home or End key */
	onCommit: (value: number) => void;
	/** The accessible name, for example "Resize the builder pane (size is remembered)" */
	label: string;
	/** Shown as a tooltip and exposed as the description, for example "Drag or use arrow keys; double-click resets." */
	description?: string;
	/** The id of the pane this splitter resizes */
	controls: string;
}

const useStyles = makeStyles({
	root: {
		position: "relative",
		flexShrink: 0,
		// The 8 px hit area; the visible line is drawn in the middle
		touchAction: "none",
		outlineStyle: "none",
		"::after": {
			content: '""',
			position: "absolute",
			backgroundColor: tokens.colorNeutralStroke2,
		},
		":hover::after": { backgroundColor: tokens.colorBrandStroke1 },
		":focus-visible::after": { backgroundColor: tokens.colorBrandStroke1 },
		"@media (forced-colors: active)": {
			"::after": { backgroundColor: "ButtonBorder", forcedColorAdjust: "none" },
			":hover::after": { backgroundColor: "Highlight" },
			":focus-visible::after": { backgroundColor: "Highlight" },
		},
		...createFocusOutlineStyle({ style: { outlineOffset: "-2px" } }),
	},
	vertical: {
		width: "8px",
		height: "100%",
		cursor: "col-resize",
		"::after": { top: 0, bottom: 0, left: "calc(50% - 0.5px)", width: "1px" },
		":hover::after": { left: "calc(50% - 1px)", width: "2px" },
		":focus-visible::after": { left: "calc(50% - 1px)", width: "2px" },
	},
	horizontal: {
		height: "8px",
		width: "100%",
		cursor: "row-resize",
		"::after": { left: 0, right: 0, top: "calc(50% - 0.5px)", height: "1px" },
		":hover::after": { top: "calc(50% - 1px)", height: "2px" },
		":focus-visible::after": { top: "calc(50% - 1px)", height: "2px" },
	},
	// Drag looks like hover (focus-visible has its own rules above)
	verticalActive: {
		"::after": {
			left: "calc(50% - 1px)",
			width: "2px",
			backgroundColor: tokens.colorBrandStroke1,
		},
		"@media (forced-colors: active)": { "::after": { backgroundColor: "Highlight" } },
	},
	horizontalActive: {
		"::after": {
			top: "calc(50% - 1px)",
			height: "2px",
			backgroundColor: tokens.colorBrandStroke1,
		},
		"@media (forced-colors: active)": { "::after": { backgroundColor: "Highlight" } },
	},
});

const clamp = (n: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, n));

interface DragState {
	pointerId: number;
	start: number;
	startValue: number;
}

export function PaneSplitter({
	orientation,
	value,
	min,
	max,
	defaultValue,
	onChange,
	onCommit,
	label,
	description,
	controls,
}: PaneSplitterProps) {
	const styles = useStyles();
	const vertical = orientation === "vertical";
	// Tolerate swapped or crossed limits (a window too small for the pane)
	const lo = Math.min(min, max);
	const hi = Math.max(min, max);
	const shown = clamp(value, lo, hi);
	const [dragging, setDragging] = useState(false);
	const [tipVisible, setTipVisible] = useState(false);

	// Latest size and callbacks, so rapid keys and the debounce never use a stale render
	const latest = useRef(shown);
	latest.current = shown;
	const callbacks = useRef({ onChange, onCommit });
	callbacks.current = { onChange, onCommit };

	const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
	const drag = useRef<DragState | null>(null);
	const previousUserSelect = useRef<string | null>(null);

	const change = (next: number) => {
		const clamped = clamp(next, lo, hi);
		latest.current = clamped;
		callbacks.current.onChange(clamped);
		return clamped;
	};

	const clearTimer = () => {
		if (timer.current !== null) {
			clearTimeout(timer.current);
			timer.current = null;
		}
	};

	const commitNow = () => {
		clearTimer();
		callbacks.current.onCommit(latest.current);
	};

	const restoreSelection = () => {
		if (previousUserSelect.current !== null) {
			document.body.style.userSelect = previousUserSelect.current;
			previousUserSelect.current = null;
		}
	};

	useEffect(
		() => () => {
			// A key change still waiting for its debounce is saved, not dropped
			if (timer.current !== null) commitNow();
			restoreSelection();
		},
		// eslint-disable-next-line react-hooks/exhaustive-deps -- only refs and helpers that read refs
		[],
	);

	const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
		const step = event.shiftKey ? SPLITTER_LARGE_STEP : SPLITTER_STEP;
		const decrease = vertical ? "ArrowLeft" : "ArrowUp";
		const increase = vertical ? "ArrowRight" : "ArrowDown";
		let next: number;
		if (event.key === decrease) next = latest.current - step;
		else if (event.key === increase) next = latest.current + step;
		else if (event.key === "Home") next = lo;
		else if (event.key === "End") next = hi;
		else if (event.key === "Enter") next = defaultValue;
		else return;
		event.preventDefault();
		if (drag.current) return;
		change(next);
		if (event.key === "Enter") {
			commitNow();
			return;
		}
		clearTimer();
		timer.current = setTimeout(commitNow, SPLITTER_COMMIT_DELAY_MS);
	};

	const endDrag = (event: PointerEvent<HTMLDivElement>) => {
		const current = drag.current;
		if (!current || event.pointerId !== current.pointerId) return;
		drag.current = null;
		setDragging(false);
		restoreSelection();
		// A plain click, or a drag that ended where it began, has nothing to save
		if (latest.current !== current.startValue) commitNow();
	};

	const onPointerDown = (event: PointerEvent<HTMLDivElement>) => {
		if (event.button !== 0 || !event.isPrimary || drag.current) return;
		event.preventDefault();
		event.currentTarget.setPointerCapture?.(event.pointerId);
		clearTimer();
		drag.current = {
			pointerId: event.pointerId,
			start: vertical ? event.clientX : event.clientY,
			startValue: latest.current,
		};
		previousUserSelect.current = document.body.style.userSelect;
		document.body.style.userSelect = "none";
		setDragging(true);
	};

	const onPointerMove = (event: PointerEvent<HTMLDivElement>) => {
		const current = drag.current;
		if (!current || event.pointerId !== current.pointerId) return;
		const position = vertical ? event.clientX : event.clientY;
		change(current.startValue + position - current.start);
	};

	const onPointerUp = (event: PointerEvent<HTMLDivElement>) => {
		const current = drag.current;
		if (!current || event.pointerId !== current.pointerId) return;
		try {
			event.currentTarget.releasePointerCapture?.(event.pointerId);
		} catch {
			// Already released
		}
		endDrag(event);
	};

	const onDoubleClick = () => {
		if (drag.current) return;
		const before = latest.current;
		change(defaultValue);
		if (latest.current !== before) commitNow();
	};

	const splitter = (
		<div
			role="separator"
			aria-orientation={orientation}
			aria-label={label}
			aria-valuenow={shown}
			aria-valuemin={lo}
			aria-valuemax={hi}
			aria-valuetext={`${shown} pixels`}
			aria-controls={controls}
			tabIndex={0}
			className={mergeClasses(
				styles.root,
				vertical ? styles.vertical : styles.horizontal,
				dragging && (vertical ? styles.verticalActive : styles.horizontalActive),
			)}
			onKeyDown={onKeyDown}
			onPointerDown={onPointerDown}
			onPointerMove={onPointerMove}
			onPointerUp={onPointerUp}
			onLostPointerCapture={endDrag}
			onDoubleClick={onDoubleClick}
		/>
	);

	return description ? (
		<Tooltip
			relationship="description"
			content={description}
			visible={tipVisible && !dragging}
			onVisibleChange={(_, data) => setTipVisible(data.visible)}
		>
			{splitter}
		</Tooltip>
	) : (
		splitter
	);
}
