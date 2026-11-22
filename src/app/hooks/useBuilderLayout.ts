/**
 * The builder pane's layout (#234; was useSplitters, #301): its width and the query tree's height,
 * where the Properties editor lives, and the inspector drawer.
 *
 * Sizes come from the layout rules (src/app/layout/layoutRules.ts) and are remembered: the
 * splitters report a settled size (pointer release, Enter, a reset, or a pause after arrow keys) and
 * only then is it saved, under its own settings key. A saved width is clamped to the window when
 * used, never written back clamped. The pane always starts docked; that isn't saved (owner decision
 * 2026-11-07). Properties move into the drawer when the window is under 1000 px or the docked pane
 * is under 320 px.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import {
	builderWidthLimits,
	clampBuilderWidth,
	clampTreeHeight,
	normalizePersistedLayout,
	propertiesHost as hostFor,
	TREE_HEIGHT_DEFAULT,
	TREE_HEIGHT_MAX,
	TREE_HEIGHT_MIN,
	type PersistedLayout,
	type PropertiesHost,
	type SizeLimits,
} from "../layout/layoutRules";

/** Namespaced, like the display settings ("<tool-id>:<setting-name>") */
export const LAYOUT_KEY = "fetchxml-studio:layout";

export interface BuilderLayout {
	/** The builder pane's width, clamped to the window */
	builderWidth: number;
	builderWidthLimits: SizeLimits;
	/** The query tree's height when Properties are docked below it */
	treeHeight: number;
	treeHeightLimits: SizeLimits;
	/** While dragging or pressing keys: the size follows */
	setBuilderWidth: (width: number) => void;
	setTreeHeight: (height: number) => void;
	/** The size has settled: remember it */
	commitBuilderWidth: (width: number) => void;
	commitTreeHeight: (height: number) => void;
	/** Where Properties are: docked below the tree, or in the drawer */
	propertiesHost: PropertiesHost;
	inspectorOpen: boolean;
	setInspectorOpen: (open: boolean) => void;
}

const TREE_LIMITS: SizeLimits = {
	min: TREE_HEIGHT_MIN,
	max: TREE_HEIGHT_MAX,
	defaultValue: TREE_HEIGHT_DEFAULT,
};

function useWindowWidth(): number {
	const [width, setWidth] = useState(() =>
		typeof window === "undefined" ? 1280 : window.innerWidth,
	);
	useEffect(() => {
		const onResize = () => setWidth(window.innerWidth);
		window.addEventListener("resize", onResize);
		return () => window.removeEventListener("resize", onResize);
	}, []);
	return width;
}

export function useBuilderLayout(): BuilderLayout {
	const windowWidth = useWindowWidth();
	// What the user chose (or what was saved), raw; undefined means the default for the window
	const [width, setWidth] = useState<number | undefined>(undefined);
	const [height, setHeight] = useState<number | undefined>(undefined);
	const [inspectorOpen, setInspectorOpen] = useState(false);
	// The saved layout is read once; nothing is saved before that, and a size the user set while it
	// was being read wins over the saved one
	const loaded = useRef(false);
	const touched = useRef({ width: false, height: false });
	const saved = useRef<PersistedLayout>({});

	useEffect(() => {
		let mounted = true;
		void (async () => {
			try {
				const raw = await window.toolboxAPI?.settings?.get?.(LAYOUT_KEY);
				if (!mounted) return;
				const layout = normalizePersistedLayout(raw);
				// A size committed while this was being read stays
				saved.current = { ...layout, ...saved.current };
				if (!touched.current.width && layout.builderWidth !== undefined)
					setWidth(layout.builderWidth);
				if (!touched.current.height && layout.treeHeight !== undefined)
					setHeight(layout.treeHeight);
			} catch (error) {
				console.warn("Couldn't read the saved pane sizes:", error);
			} finally {
				if (mounted) loaded.current = true;
			}
		})();
		return () => {
			mounted = false;
		};
	}, []);

	const save = useCallback((next: PersistedLayout) => {
		saved.current = { ...saved.current, ...next };
		if (!loaded.current) return;
		window.toolboxAPI?.settings
			?.set?.(LAYOUT_KEY, saved.current)
			?.catch((error: unknown) => console.warn("Couldn't save the pane sizes:", error));
	}, []);

	const widthLimits = builderWidthLimits(windowWidth);
	const builderWidth = clampBuilderWidth(width ?? widthLimits.defaultValue, windowWidth);
	const treeHeight = clampTreeHeight(height ?? TREE_HEIGHT_DEFAULT);
	const host = hostFor({ pane: "docked", windowWidth, builderWidth });

	return {
		builderWidth,
		builderWidthLimits: widthLimits,
		treeHeight,
		treeHeightLimits: TREE_LIMITS,
		setBuilderWidth: useCallback((next: number) => {
			touched.current.width = true;
			setWidth(next);
		}, []),
		setTreeHeight: useCallback((next: number) => {
			touched.current.height = true;
			setHeight(next);
		}, []),
		commitBuilderWidth: useCallback(
			(next: number) => {
				touched.current.width = true;
				setWidth(next);
				save({ builderWidth: next });
			},
			[save],
		),
		commitTreeHeight: useCallback(
			(next: number) => {
				touched.current.height = true;
				setHeight(next);
				save({ treeHeight: next });
			},
			[save],
		),
		propertiesHost: host,
		// The drawer exists only while Properties live in it
		inspectorOpen: host === "drawer" && inspectorOpen,
		setInspectorOpen,
	};
}
