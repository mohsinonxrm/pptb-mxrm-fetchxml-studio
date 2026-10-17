/** Shared layout styles for the property editors (#90b). */
import { makeStyles, tokens } from "@fluentui/react-components";

export const useEditorStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalL,
		padding: tokens.spacingVerticalM,
	},
	section: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
	},
	mono: { fontFamily: tokens.fontFamilyMonospace },
	readOnly: { color: tokens.colorNeutralForeground1, paddingTop: tokens.spacingVerticalXXS },
	secondary: { color: tokens.colorNeutralForeground3 },
	optionText: { display: "flex", flexDirection: "column" },
	optionSecondary: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
	leftover: {
		display: "flex",
		alignItems: "center",
		justifyContent: "space-between",
		gap: tokens.spacingHorizontalS,
	},
	accordionHeader: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
	accordionPanel: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
		paddingBottom: tokens.spacingVerticalM,
	},
	checkboxGroup: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXXS },
});
