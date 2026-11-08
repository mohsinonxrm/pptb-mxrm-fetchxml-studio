/**
 * The builder pane's two splitters (#301): the tree/properties split and the builder width. A drag
 * starts on the handle's mousedown; while it lasts, document mousemove sizes against the element
 * marked `data-left-pane` (tree height, %) or `data-app-root` (builder width, px), and mouseup ends it.
 */
import { useCallback, useEffect, useState } from "react";

export interface Splitters {
	/** Percentage of the left pane's height for the tree; the properties get the rest */
	topHeight: number;
	/** Width of the left (builder) pane in pixels */
	leftPaneWidth: number;
	/** Whether the tree/properties split is being dragged */
	isDraggingHorizontal: boolean;
	/** Whether the left/right pane split is being dragged */
	isDraggingVertical: boolean;
	/** The tree/properties handle's mousedown */
	startHorizontalDrag: () => void;
	/** The left/right pane handle's mousedown */
	startVerticalDrag: () => void;
}

export function useSplitters(): Splitters {
	// State for resizable split
	const [topHeight, setTopHeight] = useState(58); // Percentage of left pane height for tree/properties
	const [leftPaneWidth, setLeftPaneWidth] = useState(480); // Width of left pane in pixels
	const [isDraggingHorizontal, setIsDraggingHorizontal] = useState(false); // Tree/Properties resize
	const [isDraggingVertical, setIsDraggingVertical] = useState(false); // Left/Right pane resize

	// Handle horizontal resize (tree/properties split)
	useEffect(() => {
		if (!isDraggingHorizontal) return;

		const handleMouseMove = (e: MouseEvent) => {
			const leftPane = document.querySelector("[data-left-pane]") as HTMLElement;
			if (!leftPane) return;

			const rect = leftPane.getBoundingClientRect();
			const newTopHeight = ((e.clientY - rect.top) / rect.height) * 100;

			// Constrain between 30% and 80%
			if (newTopHeight >= 30 && newTopHeight <= 80) {
				setTopHeight(newTopHeight);
			}
		};

		const handleMouseUp = () => {
			setIsDraggingHorizontal(false);
		};

		document.addEventListener("mousemove", handleMouseMove);
		document.addEventListener("mouseup", handleMouseUp);

		return () => {
			document.removeEventListener("mousemove", handleMouseMove);
			document.removeEventListener("mouseup", handleMouseUp);
		};
	}, [isDraggingHorizontal]);

	// Handle vertical resize (left/right pane split)
	useEffect(() => {
		if (!isDraggingVertical) return;

		const handleMouseMove = (e: MouseEvent) => {
			const appRoot = document.querySelector("[data-app-root]") as HTMLElement;
			if (!appRoot) return;

			const rect = appRoot.getBoundingClientRect();
			const newLeftPaneWidth = e.clientX - rect.left;

			// Constrain between 300px and 800px
			if (newLeftPaneWidth >= 300 && newLeftPaneWidth <= 800) {
				setLeftPaneWidth(newLeftPaneWidth);
			}
		};

		const handleMouseUp = () => {
			setIsDraggingVertical(false);
		};

		document.addEventListener("mousemove", handleMouseMove);
		document.addEventListener("mouseup", handleMouseUp);

		return () => {
			document.removeEventListener("mousemove", handleMouseMove);
			document.removeEventListener("mouseup", handleMouseUp);
		};
	}, [isDraggingVertical]);

	const startHorizontalDrag = useCallback(() => setIsDraggingHorizontal(true), []);
	const startVerticalDrag = useCallback(() => setIsDraggingVertical(true), []);

	return {
		topHeight,
		leftPaneWidth,
		isDraggingHorizontal,
		isDraggingVertical,
		startHorizontalDrag,
		startVerticalDrag,
	};
}
