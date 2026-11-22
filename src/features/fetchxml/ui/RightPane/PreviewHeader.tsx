/**
 * The right pane's tab row (#232, #278, #233): FetchXML, LayoutXML, Results, Advisories and Code on
 * the shared overflow tab list, so tabs that don't fit move into "More" and the open tab always
 * stays. Every tab has an icon and text (#233). Results shows its row count and Advisories its total,
 * coloured by the worst finding; each count is in the tab's accessible name, and a tab in More shows
 * the same badge. The query's commands are in the app bar (#278).
 */
import { Badge, makeStyles, tokens } from "@fluentui/react-components";
import {
	Code20Filled,
	Code20Regular,
	DocumentData20Filled,
	DocumentData20Regular,
	DocumentTable20Filled,
	DocumentTable20Regular,
	ErrorCircle12Filled,
	Lightbulb20Filled,
	Lightbulb20Regular,
	Table20Filled,
	Table20Regular,
	Warning12Filled,
	bundleIcon,
} from "@fluentui/react-icons";
import { OverflowTabList, type BarTab } from "../../../../shared/components/commandBar";
import type { Advisory } from "../../model/advisories";
import type { PreviewTab } from "./previewHeaderRules";
import { advisoriesTabLabel, resultsTabLabel } from "./tabLabels";

export type { PreviewTab } from "./previewHeaderRules";

const FetchXmlIcon = bundleIcon(DocumentData20Filled, DocumentData20Regular);
const LayoutXmlIcon = bundleIcon(DocumentTable20Filled, DocumentTable20Regular);
const ResultsIcon = bundleIcon(Table20Filled, Table20Regular);
const AdvisoriesIcon = bundleIcon(Lightbulb20Filled, Lightbulb20Regular);
const CodeIcon = bundleIcon(Code20Filled, Code20Regular);

const useStyles = makeStyles({
	header: {
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		paddingRight: tokens.spacingHorizontalS,
	},
	badge: { marginInlineStart: tokens.spacingHorizontalXS },
});

export interface PreviewHeaderProps {
	selectedTab: PreviewTab;
	onTabSelect: (tab: PreviewTab) => void;
	/** The advisories the Advisories tab lists: its count and colour */
	advisories: readonly Advisory[];
	/** Rows loaded for the results; undefined before the first run (no count) */
	rowCount?: number;
}

export function PreviewHeader({
	selectedTab,
	onTabSelect,
	advisories,
	rowCount,
}: PreviewHeaderProps) {
	const styles = useStyles();
	const results = resultsTabLabel(rowCount);
	const advisoriesLabel = advisoriesTabLabel(advisories);
	// Colour is never the only cue: errors and warnings also have their own icon
	const advisoryIcon =
		advisoriesLabel.badgeColor === "danger" ? (
			<ErrorCircle12Filled />
		) : advisoriesLabel.badgeColor === "warning" ? (
			<Warning12Filled />
		) : undefined;

	const tabs: BarTab[] = [
		{ value: "xml", label: "FetchXML", icon: FetchXmlIcon },
		{ value: "layout", label: "LayoutXML", icon: LayoutXmlIcon },
		{
			value: "results",
			label: "Results",
			icon: ResultsIcon,
			ariaLabel: results.ariaLabel,
			content: results.badge ? (
				<Badge className={styles.badge} size="small" appearance="ghost" color="informative">
					{results.badge}
				</Badge>
			) : undefined,
		},
		{
			value: "advisories",
			label: "Advisories",
			icon: AdvisoriesIcon,
			ariaLabel: advisoriesLabel.ariaLabel,
			content: advisoriesLabel.badge ? (
				<Badge
					className={styles.badge}
					size="small"
					appearance={advisoriesLabel.badgeColor === "danger" ? "filled" : "tint"}
					color={advisoriesLabel.badgeColor}
					data-color={advisoriesLabel.badgeColor}
					icon={advisoryIcon}
				>
					{advisoriesLabel.badge}
				</Badge>
			) : undefined,
		},
		{ value: "code", label: "Code", icon: CodeIcon },
	];

	return (
		<div className={styles.header}>
			<OverflowTabList
				ariaLabel="Query views"
				tabs={tabs}
				selectedValue={selectedTab}
				onTabSelect={(value) => onTabSelect(value as PreviewTab)}
			/>
		</div>
	);
}
