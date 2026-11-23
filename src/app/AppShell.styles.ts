/**
 * AppShell's styles (#301): the panes and toolbars around the builder and the results. The splitters
 * are PaneSplitter (#234).
 */
import { makeStyles, tokens } from "@fluentui/react-components";

// ⚠️ IMPORTANT: makeStyles must be called OUTSIDE the component
// but tokens will automatically update when FluentProvider theme changes
// because Fluent UI v9 uses CSS variables under the hood
export const useStyles = makeStyles({
	// The app bar above the two panes (#278)
	shell: {
		// The theme background covers the whole tool; without it, surfaces with no background of
		// their own showed the page's white canvas in dark mode (#67). (Not on the FluentProvider:
		// Fluent copies its className onto portal containers.)
		backgroundColor: tokens.colorNeutralBackground1,
		color: tokens.colorNeutralForeground1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		position: "absolute",
		top: 0,
		left: 0,
		right: 0,
		bottom: 0,
	},
	// The app bar's banner landmark (#186): it keeps its height, as the bar did
	appBar: {
		flexShrink: 0,
		display: "flex",
		flexDirection: "column",
		minWidth: 0,
	},
	// The two panes, under the app bar
	root: {
		display: "flex",
		flexDirection: "row",
		flex: "1 1 0",
		minHeight: 0,
		width: "100%",
		overflow: "hidden",
		position: "relative",
	},
	// The builder pane: its width comes from useBuilderLayout, within the layout rules' limits (#234)
	leftPane: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		flexShrink: 0,
		height: "100%",
	},
	// The peek (#234): one part of the collapsed pane, over the output, beside the rail
	peek: {
		position: "absolute",
		top: 0,
		bottom: 0,
		left: "44px",
		zIndex: 30,
		backgroundColor: tokens.colorNeutralBackground1,
		boxShadow: tokens.shadow16,
		borderRight: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
	},
	peekBar: {
		paddingBlock: "6px",
		paddingRight: "8px",
	},
	peekBarEnd: {
		marginLeft: "auto",
	},
	leftPaneTop: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minHeight: 0, // Important for flex child with overflow
	},
	// Properties take the height the tree leaves (the tree has its own height while docked)
	leftPaneBottom: {
		flex: "1 1 0",
		display: "flex",
		flexDirection: "column",
		minHeight: "100px",
		overflow: "hidden",
		minWidth: 0, // Important for flex child with overflow
	},
	propertiesContent: {
		flex: 1,
		overflow: "hidden",
	},
	// Its own stacking context: what's raised inside (Monaco's widgets, the drawer) stays under the peek
	rightPane: {
		isolation: "isolate",
		flex: 1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minWidth: 0,
	},
	toolbar: {
		padding: "12px 16px",
		borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		display: "flex",
		gap: "8px",
		alignItems: "center",
		flexShrink: 0,
	},
	placeholder: {
		padding: "16px",
		color: tokens.colorNeutralForeground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
	},
});
