/**
 * The Advisories tab (#97): every finding about the query in one list, from FetchXML Studio (no
 * request) and, on demand, from Dataverse's ValidateFetchXmlExpression. Advice never blocks
 * Execute; a location selects its node in the tree, and the paging advice has a quick fix.
 */
import { useState, type ReactNode } from "react";
import {
	Badge,
	Button,
	Link,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
	MessageBarTitle,
	Spinner,
	Text,
	ToggleButton,
	Toolbar,
	ToolbarDivider,
	Tooltip,
	makeStyles,
	mergeClasses,
	tokens,
} from "@fluentui/react-components";
import {
	CloudCheckmark20Regular,
	Dismiss16Regular,
	ErrorCircle20Regular,
	Lightbulb20Regular,
	Open16Regular,
	Warning20Regular,
	Wrench16Regular,
} from "@fluentui/react-icons";
import { LEVEL_LABELS, countByLevel, type Advisory, type AdvisoryLevel } from "../../model/advisories";
import type { NodeId } from "../../model/nodes";

const useStyles = makeStyles({
	card: {
		borderRadius: tokens.borderRadiusMedium,
		boxShadow: tokens.shadow16,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		height: "100%",
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
	},
	toolbar: {
		display: "flex",
		alignItems: "center",
		flexWrap: "wrap",
		gap: tokens.spacingHorizontalS,
		padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
		borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
	},
	status: { color: tokens.colorNeutralForeground3, minWidth: 0, flex: "1 1 240px" },
	filters: { display: "flex", alignItems: "center", flexWrap: "nowrap", gap: tokens.spacingHorizontalXXS, marginLeft: "auto" },
	messageBar: { padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS} 0` },
	scroll: { flex: 1, overflowY: "auto", minHeight: 0 },
	empty: {
		padding: tokens.spacingHorizontalL,
		color: tokens.colorNeutralForeground3,
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
	},
	footer: {
		padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
		color: tokens.colorNeutralForeground3,
		borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
	},
	row: {
		display: "grid",
		gridTemplateColumns: "20px minmax(0, 1fr) auto",
		columnGap: tokens.spacingHorizontalM,
		padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
		borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
	},
	selected: { backgroundColor: tokens.colorNeutralBackground1Selected },
	error: { color: tokens.colorPaletteRedForeground1 },
	warning: { color: tokens.colorPaletteDarkOrangeForeground1 },
	advice: { color: tokens.colorBrandForeground1 },
	body: { display: "flex", flexDirection: "column", gap: "2px", minWidth: 0 },
	secondary: { color: tokens.colorNeutralForeground3 },
	quote: { color: tokens.colorNeutralForeground3, fontStyle: "italic" },
	meta: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: tokens.spacingHorizontalS },
	side: { display: "flex", flexDirection: "column", alignItems: "flex-end", gap: tokens.spacingVerticalXS },
	linkIcon: { verticalAlign: "-3px" },
});

const LEVEL_ICON: Record<AdvisoryLevel, ReactNode> = {
	error: <ErrorCircle20Regular />,
	warning: <Warning20Regular />,
	advice: <Lightbulb20Regular />,
};

type Filter = AdvisoryLevel | "all";

export interface AdvisoriesPanelProps {
	advisories: Advisory[];
	/** No query yet (no entity) */
	noQuery: boolean;
	/** In editor mode, the editor's FetchXML doesn't parse */
	editorInvalid: boolean;
	/** The list is about the tree's query, so locations and fixes can act on the tree */
	inTree: boolean;
	selectedNodeId?: NodeId | null;
	onSelectNode?: (nodeId: NodeId) => void;
	onFix?: (advisory: Advisory) => void;
	// Check with Dataverse
	checking: boolean;
	checkedAt?: Date;
	/** The FetchXML changed since the check */
	stale: boolean;
	/** Rows only Dataverse reported */
	added: number;
	checkError?: string;
	onDismissCheckError: () => void;
	onCheck: () => void;
	/** Why Check with Dataverse is off (no connection, no Dataverse API) */
	checkDisabledReason?: string;
}

function statusText(p: Pick<AdvisoriesPanelProps, "checking" | "checkedAt" | "stale" | "added">): string {
	if (p.checking) return "Checking with Dataverse…";
	if (!p.checkedAt) return "From FetchXML Studio. Dataverse can also check choice and large text attributes.";
	const time = p.checkedAt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
	// A stale check's findings aren't shown (#170): they could describe a query Dataverse never saw
	return p.stale
		? `Checked with Dataverse at ${time}. The query has changed since: check again to see Dataverse's findings.`
		: `Checked with Dataverse at ${time}. ${p.added} new.`;
}

const NOT_IN_TREE = "The editor's FetchXML isn't in the tree yet. Use Parse to tree first.";

function AdvisoryRow({ a, props }: { a: Advisory; props: AdvisoriesPanelProps }) {
	const s = useStyles();
	const { inTree, selectedNodeId, onSelectNode, onFix } = props;
	const location =
		inTree && a.nodeId && onSelectNode ? (
			<Link as="button" onClick={() => onSelectNode(a.nodeId!)}>
				{a.location}
			</Link>
		) : (
			<Tooltip content={NOT_IN_TREE} relationship="description">
				<Text size={200} className={s.secondary}>
					{a.location}
				</Text>
			</Tooltip>
		);
	const fixButton = a.fix && (
		<Button size="small" icon={<Wrench16Regular />} disabled={!inTree || !onFix} onClick={() => onFix?.(a)}>
			Add {a.fix.attribute} as the last order
		</Button>
	);
	return (
		<div className={mergeClasses(s.row, inTree && a.nodeId && a.nodeId === selectedNodeId && s.selected)}>
			<span className={s[a.level]} aria-label={LEVEL_LABELS[a.level]} role="img">
				{LEVEL_ICON[a.level]}
			</span>
			<div className={s.body}>
				<Text weight="semibold">{a.title}</Text>
				<Text size={200} className={s.secondary}>
					{a.detail}
				</Text>
				{a.source === "dataverse" && a.dataverseText && (
					<Text size={200} className={s.quote}>
						Dataverse: “{a.dataverseText}”
					</Text>
				)}
				<div className={s.meta}>
					{location}
					<Text size={200} className={s.secondary} aria-hidden>
						·
					</Text>
					<Link href={a.doc.href} target="_blank" rel="noreferrer">
						{a.doc.label} <Open16Regular className={s.linkIcon} />
					</Link>
				</div>
			</div>
			<div className={s.side}>
				{a.source !== "fxs" && (
					<Tooltip
						content={a.source === "both" ? `Dataverse reported this too: “${a.dataverseText ?? ""}”` : "Only Dataverse reported this"}
						relationship="description"
					>
						<Badge appearance="tint" color="brand" size="small">
							{a.source === "both" ? "Dataverse agrees" : "From Dataverse"}
						</Badge>
					</Tooltip>
				)}
				{fixButton &&
					(inTree ? (
						fixButton
					) : (
						<Tooltip content={NOT_IN_TREE} relationship="description">
							<span>{fixButton}</span>
						</Tooltip>
					))}
			</div>
		</div>
	);
}

export function AdvisoriesPanel(props: AdvisoriesPanelProps) {
	const s = useStyles();
	const { advisories, checking, checkedAt, onCheck, checkDisabledReason } = props;
	const [filter, setFilter] = useState<Filter>("all");
	const counts = countByLevel(advisories);
	const list = advisories.filter((a) => filter === "all" || a.level === filter);

	const checkButton = (
		<Button
			icon={checking ? <Spinner size="extra-tiny" /> : <CloudCheckmark20Regular />}
			onClick={onCheck}
			disabled={checking || !!checkDisabledReason || props.noQuery}
		>
			{checkedAt ? "Check again" : "Check with Dataverse"}
		</Button>
	);

	const toggle = (value: Filter, label: string, count: number) => (
		<ToggleButton size="small" appearance="subtle" checked={filter === value} onClick={() => setFilter(value)}>
			{label} ({count})
		</ToggleButton>
	);

	let empty: ReactNode = null;
	if (props.noQuery) {
		empty = <Text>Choose an entity to start a query.</Text>;
	} else if (props.editorInvalid) {
		empty = <Text>The editor's FetchXML isn't valid FetchXML. Validate shows why.</Text>;
	} else if (advisories.length === 0) {
		empty = (
			<>
				<Text>Nothing to report for this query.</Text>
				<Text size={200}>FetchXML Studio checks documented anti-patterns and limits as you build. Check with Dataverse for Dataverse's own advice.</Text>
			</>
		);
	} else if (list.length === 0) {
		empty = <Text>Nothing to report here.</Text>;
	}

	return (
		<div className={s.card}>
			<Toolbar size="small" className={s.toolbar} aria-label="Advisories">
				{checkDisabledReason ? (
					// Disabled buttons don't get pointer events: the tooltip goes on a wrapper
					<Tooltip content={checkDisabledReason} relationship="description">
						<span>{checkButton}</span>
					</Tooltip>
				) : (
					checkButton
				)}
				<Text size={200} className={s.status} aria-live="polite">
					{statusText(props)}
				</Text>
				{/* The filters wrap together, so they never split across lines */}
				<div className={s.filters} role="group" aria-label="Show">
					{toggle("all", "All", advisories.length)}
					<ToolbarDivider />
					{toggle("error", LEVEL_LABELS.error, counts.error)}
					{toggle("warning", LEVEL_LABELS.warning, counts.warning)}
					{toggle("advice", LEVEL_LABELS.advice, counts.advice)}
				</div>
			</Toolbar>
			{props.checkError && (
				<div className={s.messageBar}>
					<MessageBar intent="error" layout="multiline">
						<MessageBarBody>
							<MessageBarTitle>Couldn't check with Dataverse</MessageBarTitle>
							{props.checkError}
						</MessageBarBody>
						<MessageBarActions
							containerAction={
								<Button appearance="transparent" icon={<Dismiss16Regular />} onClick={props.onDismissCheckError} aria-label="Dismiss" />
							}
						/>
					</MessageBar>
				</div>
			)}
			<div className={s.scroll}>
				{empty ? <div className={s.empty}>{empty}</div> : list.map((a) => <AdvisoryRow key={a.id} a={a} props={props} />)}
			</div>
			<Text size={200} className={s.footer}>
				Advice never blocks Execute. Select a location to show it in the tree.
			</Text>
		</div>
	);
}
