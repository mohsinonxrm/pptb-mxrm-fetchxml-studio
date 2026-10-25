/**
 * The scope chip (#227, variant A, shared with Dataverse Request Studio and Command Bar Studio):
 * one button above the entity picker that always shows the scope you're working in. It opens a
 * popover with the entity source, publishers, solutions and Advanced Find only; changes apply right
 * away. On first run without a scope, a teaching popover points at it.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import {
	Button,
	Caption1,
	Field,
	Popover,
	PopoverSurface,
	PopoverTrigger,
	Radio,
	RadioGroup,
	Switch,
	Tag,
	TagPicker,
	TagPickerControl,
	TagPickerGroup,
	TagPickerInput,
	TagPickerList,
	TagPickerOption,
	TagPickerOptionGroup,
	TeachingPopover,
	TeachingPopoverBody,
	TeachingPopoverFooter,
	TeachingPopoverHeader,
	TeachingPopoverSurface,
	TeachingPopoverTitle,
	Text,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { ChevronDown16Regular, Filter20Regular, LockClosed16Regular } from "@fluentui/react-icons";
import type { EntityScopeMode } from "../../model/displaySettings";
import { scopeChipText, solutionDetail, solutionMatchesSearch, solutionTagLabel } from "../../model/scopeChip";
import type { Publisher, Solution } from "../../api/pptbClient";

const useStyles = makeStyles({
	chip: {
		flex: "1 1 auto",
		minWidth: 0,
		maxWidth: "420px",
		justifyContent: "flex-start",
	},
	chipLabel: {
		flex: 1,
		minWidth: 0,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
		textAlign: "start",
	},
	chevron: {
		flexShrink: 0,
		marginLeft: tokens.spacingHorizontalXS,
	},
	tooltipDetail: {
		whiteSpace: "pre-line",
	},
	surface: {
		width: "380px",
		maxWidth: "calc(100vw - 32px)",
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
	},
	muted: {
		color: tokens.colorNeutralForeground3,
	},
	locked: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
		color: tokens.colorNeutralForeground3,
	},
	empty: {
		padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
		color: tokens.colorNeutralForeground3,
	},
	footer: {
		display: "flex",
		alignItems: "center",
		justifyContent: "space-between",
		gap: tokens.spacingHorizontalM,
	},
});

export const SCOPE_TIP_TITLE = "Choose where entities come from";
export const SCOPE_TIP_BODY =
	"Pick a publisher and one or more solutions to list their entities, or switch to All entities to see everything.";

export interface ScopeChipProps {
	/** The effective entity source (the preference, clamped to the user's privileges) */
	mode: EntityScopeMode;
	onModeChange: (mode: EntityScopeMode) => void;
	/** Privileges: Publisher → solution needs prvReadPublisher + prvReadSolution; Solution only needs prvReadSolution */
	canUsePublisherSolution: boolean;
	canUseSolutionOnly: boolean;
	publishers: readonly Publisher[];
	selectedPublisherIds: readonly string[];
	publishersLoading: boolean;
	publishersError?: string | null;
	onPublishersChange: (ids: string[]) => void;
	/** The solutions to choose from in this mode (for the selected publishers in Publisher → solution) */
	solutions: readonly Solution[];
	selectedSolutionIds: readonly string[];
	solutionsLoading: boolean;
	solutionsError?: string | null;
	onSolutionsChange: (ids: string[]) => void;
	advancedFindOnly: boolean;
	onAdvancedFindOnlyChange: (value: boolean) => void;
	/** Show the first-run teaching popover (no scope yet, not dismissed in this environment) */
	showTip: boolean;
	onTipDismiss: () => void;
}

export function ScopeChip(props: ScopeChipProps) {
	const {
		mode,
		publishers,
		selectedPublisherIds,
		solutions,
		selectedSolutionIds,
		showTip,
		onTipDismiss,
	} = props;
	const styles = useStyles();
	const [open, setOpen] = useState(false);
	const [chip, setChip] = useState<HTMLElement | null>(null);

	const selectedPublishers = useMemo(
		() => publishers.filter((p) => selectedPublisherIds.includes(p.publisherid)),
		[publishers, selectedPublisherIds],
	);
	const selectedSolutions = useMemo(
		() => solutions.filter((s) => selectedSolutionIds.includes(s.solutionid)),
		[solutions, selectedSolutionIds],
	);
	const text = scopeChipText(mode, selectedPublishers, selectedSolutions);

	// The full scope as a tooltip, only when the label is cut off
	const labelRef = useRef<HTMLSpanElement>(null);
	const [truncated, setTruncated] = useState(false);
	const [tooltipVisible, setTooltipVisible] = useState(false);
	useEffect(() => {
		const label = labelRef.current;
		if (!label || typeof ResizeObserver === "undefined") return;
		const check = () => setTruncated(label.scrollWidth > label.clientWidth);
		check();
		const observer = new ResizeObserver(check);
		observer.observe(label);
		return () => observer.disconnect();
	}, [text.label]);

	const chipButton = (
		<Button
			ref={setChip}
			className={styles.chip}
			appearance={text.configured ? "outline" : "primary"}
			size="small"
			icon={<Filter20Regular />}
			aria-label={text.ariaLabel}
		>
			<span ref={labelRef} className={styles.chipLabel}>
				{text.label}
			</span>
			<ChevronDown16Regular className={styles.chevron} />
		</Button>
	);

	return (
		<>
			<Popover open={open} onOpenChange={(_, d) => setOpen(d.open)} positioning="below-start" trapFocus withArrow>
				<PopoverTrigger disableButtonEnhancement>
					<Tooltip
						content={<span className={styles.tooltipDetail}>{text.detail}</span>}
						relationship="description"
						visible={truncated && tooltipVisible && !open}
						onVisibleChange={(_, d) => setTooltipVisible(d.visible)}
					>
						{chipButton}
					</Tooltip>
				</PopoverTrigger>
				<PopoverSurface className={styles.surface} aria-label="Entity scope">
					<Text weight="semibold" size={400}>
						Entity scope
					</Text>
					<ScopeEditor {...props} />
					<div className={styles.footer}>
						<Caption1 className={styles.muted}>Changes apply right away.</Caption1>
						<Button
							appearance="primary"
							onClick={() => {
								setOpen(false);
								// Back to the chip, as Escape does (a programmatic close doesn't restore focus)
								requestAnimationFrame(() => chip?.focus());
							}}
						>
							Done
						</Button>
					</div>
				</PopoverSurface>
			</Popover>

			{chip && (
				<TeachingPopover
					open={showTip && !open}
					onOpenChange={(_, d) => !d.open && onTipDismiss()}
					positioning={{ target: chip, position: "below", align: "start" }}
				>
					<TeachingPopoverSurface style={{ maxWidth: 320 }}>
						<TeachingPopoverHeader>Tip</TeachingPopoverHeader>
						<TeachingPopoverBody>
							<TeachingPopoverTitle>{SCOPE_TIP_TITLE}</TeachingPopoverTitle>
							<div>{SCOPE_TIP_BODY}</div>
						</TeachingPopoverBody>
						<TeachingPopoverFooter
							primary={{
								children: "Choose scope",
								onClick: () => {
									onTipDismiss();
									setOpen(true);
								},
							}}
							secondary={{ children: "Got it", onClick: onTipDismiss }}
						/>
					</TeachingPopoverSurface>
				</TeachingPopover>
			)}
		</>
	);
}

/** The popover's body: entity source, publishers, solutions, Advanced Find only */
function ScopeEditor({
	mode,
	onModeChange,
	canUsePublisherSolution,
	canUseSolutionOnly,
	publishers,
	selectedPublisherIds,
	publishersLoading,
	publishersError,
	onPublishersChange,
	solutions,
	selectedSolutionIds,
	solutionsLoading,
	solutionsError,
	onSolutionsChange,
	advancedFindOnly,
	onAdvancedFindOnlyChange,
}: ScopeChipProps) {
	const styles = useStyles();
	const [publisherQuery, setPublisherQuery] = useState("");
	const [solutionQuery, setSolutionQuery] = useState("");

	const publisherOptions = useMemo(() => {
		const q = publisherQuery.trim().toLowerCase();
		return publishers.filter(
			(p) =>
				!selectedPublisherIds.includes(p.publisherid) &&
				(!q || [p.friendlyname, p.uniquename, p.customizationprefix].some((v) => (v ?? "").toLowerCase().includes(q))),
		);
	}, [publishers, selectedPublisherIds, publisherQuery]);

	const { unmanaged, managed } = useMemo(() => {
		const open = solutions.filter((s) => !selectedSolutionIds.includes(s.solutionid) && solutionMatchesSearch(s, solutionQuery));
		return { unmanaged: open.filter((s) => !s.ismanaged), managed: open.filter((s) => s.ismanaged) };
	}, [solutions, selectedSolutionIds, solutionQuery]);

	const selectedPublishers = publishers.filter((p) => selectedPublisherIds.includes(p.publisherid));
	const selectedSolutions = solutions.filter((s) => selectedSolutionIds.includes(s.solutionid));
	const needsPublisher = mode === "publisher-solution" && selectedPublisherIds.length === 0;
	const solutionOption = (s: Solution) => (
		<TagPickerOption key={s.solutionid} value={s.solutionid} text={s.friendlyname} secondaryContent={solutionDetail(s)}>
			{s.friendlyname}
		</TagPickerOption>
	);

	return (
		<>
			<Field label="Show entities from">
				<RadioGroup value={mode} onChange={(_, d) => onModeChange(d.value as EntityScopeMode)}>
					<Tooltip
						content={canUsePublisherSolution ? "" : "Requires prvReadPublisher and prvReadSolution"}
						relationship="description"
						positioning="before"
					>
						<Radio value="publisher-solution" label="Publisher → solution (default)" disabled={!canUsePublisherSolution} />
					</Tooltip>
					<Tooltip content={canUseSolutionOnly ? "" : "Requires prvReadSolution"} relationship="description" positioning="before">
						<Radio value="solution-only" label="Solution only" disabled={!canUseSolutionOnly} />
					</Tooltip>
					<Radio value="all" label="All entities" />
				</RadioGroup>
			</Field>

			{mode !== "all" && (
				<Field
					label="Publishers"
					validationState={publishersError && mode === "publisher-solution" ? "error" : "none"}
					validationMessage={mode === "publisher-solution" ? publishersError : undefined}
				>
					<TagPicker
						selectedOptions={[...selectedPublisherIds]}
						onOptionSelect={(_, d) => {
							onPublishersChange(d.selectedOptions);
							setPublisherQuery("");
						}}
						disabled={mode !== "publisher-solution" || publishersLoading}
					>
						<TagPickerControl>
							<TagPickerGroup aria-label="Selected publishers">
								{mode === "publisher-solution" &&
									selectedPublishers.map((p) => (
										<Tag key={p.publisherid} value={p.publisherid} shape="rounded" size="small" dismissible>
											{p.friendlyname}
										</Tag>
									))}
							</TagPickerGroup>
							<TagPickerInput
								aria-label="Publishers"
								placeholder={
									!canUsePublisherSolution
										? "No access to publishers"
										: publishersLoading
											? "Loading…"
											: selectedPublisherIds.length
												? ""
												: "Search publishers…"
								}
								value={publisherQuery}
								onChange={(e) => setPublisherQuery(e.target.value)}
							/>
						</TagPickerControl>
						<TagPickerList>
							{publisherOptions.length === 0 ? (
								<div className={styles.empty}>
									{publisherQuery.trim() ? `No publishers match "${publisherQuery.trim()}"` : "No publishers found"}
								</div>
							) : (
								publisherOptions.map((p) => (
									<TagPickerOption
										key={p.publisherid}
										value={p.publisherid}
										text={p.friendlyname}
										secondaryContent={p.customizationprefix ? `Prefix ${p.customizationprefix}_` : p.uniquename}
									>
										{p.friendlyname}
									</TagPickerOption>
								))
							)}
						</TagPickerList>
					</TagPicker>
					{mode === "solution-only" && canUsePublisherSolution && (
						<Caption1 className={styles.locked}>
							<LockClosed16Regular /> Not used in Solution only mode
						</Caption1>
					)}
				</Field>
			)}

			{mode !== "all" && (
				<Field label="Solutions" validationState={solutionsError ? "error" : "none"} validationMessage={solutionsError}>
					<TagPicker
						selectedOptions={[...selectedSolutionIds]}
						onOptionSelect={(_, d) => {
							onSolutionsChange(d.selectedOptions);
							setSolutionQuery("");
						}}
						disabled={needsPublisher || solutionsLoading}
					>
						<TagPickerControl>
							<TagPickerGroup aria-label="Selected solutions">
								{selectedSolutions.map((s) => (
									<Tag key={s.solutionid} value={s.solutionid} shape="rounded" size="small" dismissible title={solutionDetail(s)}>
										{solutionTagLabel(s)}
									</Tag>
								))}
							</TagPickerGroup>
							<TagPickerInput
								aria-label="Solutions"
								placeholder={
									needsPublisher
										? "Select a publisher first"
										: solutionsLoading
											? "Loading…"
											: selectedSolutionIds.length
												? ""
												: "Search name, unique name or version…"
								}
								value={solutionQuery}
								onChange={(e) => setSolutionQuery(e.target.value)}
							/>
						</TagPickerControl>
						<TagPickerList>
							{unmanaged.length === 0 && managed.length === 0 ? (
								<div className={styles.empty}>
									{solutionQuery.trim() ? `No solutions match "${solutionQuery.trim()}"` : "No solutions found"}
								</div>
							) : (
								<>
									{unmanaged.length > 0 && <TagPickerOptionGroup label="Unmanaged">{unmanaged.map(solutionOption)}</TagPickerOptionGroup>}
									{managed.length > 0 && <TagPickerOptionGroup label="Managed">{managed.map(solutionOption)}</TagPickerOptionGroup>}
								</>
							)}
						</TagPickerList>
					</TagPicker>
				</Field>
			)}

			<div>
				<Switch
					checked={advancedFindOnly}
					onChange={(_, d) => onAdvancedFindOnlyChange(d.checked)}
					label="Advanced Find only"
				/>
				<Caption1 className={styles.muted} block>
					{advancedFindOnly
						? "Entities, attributes and relationships are limited to the ones Advanced Find shows."
						: "All entities, attributes and relationships are shown. A few system entities can't be queried; Dataverse returns an error for those."}
				</Caption1>
			</div>
		</>
	);
}
