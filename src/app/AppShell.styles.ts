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
	leftPaneTop: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minHeight: 0, // Important for flex child with overflow
	},
	leftPaneBottom: {
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
	rightPane: {
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
