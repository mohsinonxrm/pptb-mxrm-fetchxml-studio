/**
 * A keyboard- and pointer-operable pane splitter (#234), following the WAI-ARIA window splitter
 * pattern: https://www.w3.org/WAI/ARIA/apg/patterns/windowsplitter/
 *
 * The component is controlled: it reports every size change through `onChange` and reports the
 * settled size through `onCommit` (pointer release, or 250 ms after the last key), which is the
 * moment to persist it. The caller owns the size and the layout it drives.
 */
import { useCallback, useEffect, useRef, useState, type KeyboardEvent, type PointerEvent } from "react";
import { createFocusOutlineStyle, makeStyles, mergeClasses, Tooltip, tokens } from "@fluentui/react-components";

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
	/** Called once the change has settled: pointer release or lost capture, a reset, or 250 ms after the last key */
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
		"@media (forced-colors: active)": {
			"::after": { backgroundColor: "ButtonBorder", forcedColorAdjust: "none" },
			":hover::after": { backgroundColor: "Highlight" },
		},
		...createFocusOutlineStyle({ style: { outlineOffset: "-2px" } }),
	},
	vertical: {
		width: "8px",
		height: "100%",
		cursor: "col-resize",
		"::after": { top: 0, bottom: 0, left: "calc(50% - 0.5px)", width: "1px" },
		":hover::after": { left: "calc(50% - 1px)", width: "2px" },
	},
	horizontal: {
		height: "8px",
		width: "100%",
		cursor: "row-resize",
		"::after": { left: 0, right: 0, top: "calc(50% - 0.5px)", height: "1px" },
		":hover::after": { top: "calc(50% - 1px)", height: "2px" },
	},
	// Focus-visible and drag look like hover. Griffel can't share one rule, so these repeat it.
	verticalActive: {
		"::after": { left: "calc(50% - 1px)", width: "2px", backgroundColor: tokens.colorBrandStroke1 },
		"@media (forced-colors: active)": { "::after": { backgroundColor: "Highlight" } },
	},
	horizontalActive: {
		"::after": { top: "calc(50% - 1px)", height: "2px", backgroundColor: tokens.colorBrandStroke1 },
		"@media (forced-colors: active)": { "::after": { backgroundColor: "Highlight" } },
	},
});

const clamp = (n: number, min: number, max: number) => Math.min(max, Math.max(min, n));

export function PaneSplitter({ orientation, value, min, max, defaultValue, onChange, onCommit, label, description, controls }: PaneSplitterProps) {
	const styles = useStyles();
	const vertical = orientation === "vertical";
	const [dragging, setDragging] = useState(false);
	const [focusVisible, setFocusVisible] = useState(false);

	// Latest size and callbacks, so rapid keys and the debounce never use a stale render
	const latest = useRef(value);
	latest.current = value;
	const callbacks = useRef({ onChange, onCommit });
	callbacks.current = { onChange, onCommit };

	const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
	const drag = useRef<{ start: number; startValue: number } | null>(null);
	const previousUserSelect = useRef<string | null>(null);

	const change = useCallback(
		(next: number) => {
			const clamped = clamp(next, min, max);
			latest.current = clamped;
			callbacks.current.onChange(clamped);
			return clamped;
		},
		[min, max],
	);

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
			clearTimer();
			restoreSelection();
		},
		[],
	);

	const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
		const step = event.shiftKey ? SPLITTER_LARGE_STEP : SPLITTER_STEP;
		const decrease = vertical ? "ArrowLeft" : "ArrowUp";
		const increase = vertical ? "ArrowRight" : "ArrowDown";
		let next: number;
		if (event.key === decrease) next = latest.current - step;
		else if (event.key === increase) next = latest.current + step;
		else if (event.key === "Home") next = min;
		else if (event.key === "End") next = max;
		else if (event.key === "Enter") next = defaultValue;
		else return;
		event.preventDefault();
		change(next);
		clearTimer();
		timer.current = setTimeout(commitNow, SPLITTER_COMMIT_DELAY_MS);
	};

	const endDrag = () => {
		if (!drag.current) return;
		drag.current = null;
		setDragging(false);
		restoreSelection();
		commitNow();
	};

	const onPointerDown = (event: PointerEvent<HTMLDivElement>) => {
		if (event.button !== 0) return;
		event.preventDefault();
		event.currentTarget.setPointerCapture?.(event.pointerId);
		drag.current = { start: vertical ? event.clientX : event.clientY, startValue: latest.current };
		previousUserSelect.current = document.body.style.userSelect;
		document.body.style.userSelect = "none";
		setDragging(true);
	};

	const onPointerMove = (event: PointerEvent<HTMLDivElement>) => {
		if (!drag.current) return;
		const position = vertical ? event.clientX : event.clientY;
		change(drag.current.startValue + position - drag.current.start);
	};

	const onPointerUp = (event: PointerEvent<HTMLDivElement>) => {
		if (!drag.current) return;
		try {
			event.currentTarget.releasePointerCapture?.(event.pointerId);
		} catch {
			// Already released
		}
		endDrag();
	};

	const onDoubleClick = () => {
		change(defaultValue);
		commitNow();
	};

	const active = dragging || focusVisible;
	const splitter = (
		<div
			role="separator"
			aria-orientation={orientation}
			aria-label={label}
			aria-valuenow={value}
			aria-valuemin={min}
			aria-valuemax={max}
			aria-valuetext={`${value} pixels`}
			aria-controls={controls}
			tabIndex={0}
			className={mergeClasses(
				styles.root,
				vertical ? styles.vertical : styles.horizontal,
				active && (vertical ? styles.verticalActive : styles.horizontalActive),
			)}
			onKeyDown={onKeyDown}
			onPointerDown={onPointerDown}
			onPointerMove={onPointerMove}
			onPointerUp={onPointerUp}
			onLostPointerCapture={endDrag}
			onDoubleClick={onDoubleClick}
			onFocus={(event) => setFocusVisible(event.currentTarget.matches(":focus-visible"))}
			onBlur={() => setFocusVisible(false)}
		/>
	);

	return description ? (
		<Tooltip relationship="description" content={description}>
			{splitter}
		</Tooltip>
	) : (
		splitter
	);
}
