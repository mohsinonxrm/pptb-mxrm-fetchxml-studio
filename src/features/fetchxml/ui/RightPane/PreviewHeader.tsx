/**
 * The right pane's tab row (#232, #278): FetchXML, LayoutXML, Results, Advisories and Code on the
 * shared overflow tab list, so tabs that don't fit move into "More" and the open tab always stays.
 * The query's commands (Run, Save, Send to, Return FetchXML, Settings) are in the app bar (#278).
 */
import { CounterBadge, makeStyles, tokens } from "@fluentui/react-components";
import { Code20Filled, Code20Regular, bundleIcon } from "@fluentui/react-icons";
import { OverflowTabList, type BarTab } from "../../../../shared/components/commandBar";
import type { PreviewTab } from "./previewHeaderRules";

export type { PreviewTab } from "./previewHeaderRules";

const CodeIcon = bundleIcon(Code20Filled, Code20Regular);

const useStyles = makeStyles({
	header: {
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		paddingRight: tokens.spacingHorizontalS,
	},
});

export interface PreviewHeaderProps {
	selectedTab: PreviewTab;
	onTabSelect: (tab: PreviewTab) => void;
	advisoryCount: number;
}

export function PreviewHeader({ selectedTab, onTabSelect, advisoryCount }: PreviewHeaderProps) {
	const styles = useStyles();
	const tabs: BarTab[] = [
		{ value: "xml", label: "FetchXML" },
		{ value: "layout", label: "LayoutXML" },
		{ value: "results", label: "Results" },
		{
			value: "advisories",
			label: "Advisories",
			ariaLabel: `Advisories (${advisoryCount})`,
			content: (
				<>
					{" "}
					<CounterBadge
						count={advisoryCount}
						size="small"
						appearance="ghost"
						color="brand"
					/>
				</>
			),
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
