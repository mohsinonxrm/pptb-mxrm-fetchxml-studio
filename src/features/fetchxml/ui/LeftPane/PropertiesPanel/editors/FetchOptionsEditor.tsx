/**
 * Property editor for the <fetch> element: every attribute in the FetchXML reference (#90b).
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/fetch
 */

import { useState } from "react";
import {
	Accordion,
	AccordionHeader,
	AccordionItem,
	AccordionPanel,
	Badge,
	Caption1,
	Checkbox,
	Dropdown,
	InfoLabel,
	Input,
	MessageBar,
	MessageBarBody,
	Option,
	Switch,
	Tag,
	TagGroup,
	TeachingPopover,
	TeachingPopoverBody,
	TeachingPopoverFooter,
	TeachingPopoverSurface,
	TeachingPopoverTitle,
	TeachingPopoverTrigger,
	type LabelProps,
} from "@fluentui/react-components";
import type { FetchNode } from "../../../../model/nodes";
import type { RuleFinding } from "../../../../model/fetchxmlRules";
import { getDataSourceIssue } from "../../../../model/attributeCapabilities";
import { planAggregateOff, hasAggregateOffChanges, type AggregateOffPlan } from "../../../../model/aggregateToggle";
import { useBuilder } from "../../../../state/builderContext";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoCheckbox, InfoField, NumberField, SectionHeading } from "../fields";
import { findingFor, worst } from "../ruleFindings";
import { useEntityInfo } from "../useEditorMetadata";
import { MetadataLoadWarning } from "../MetadataLoadWarning";
import { AggregateOffDialog } from "./AggregateOffDialog";

// Track if the teaching popover has been shown this session
let hasShownRetrieveAllPopover = false;

/** Documented query hint names (fetch reference › options) and the SQL Server hint each applies. */
const DOCUMENTED_HINTS: readonly { name: string; sql: string }[] = [
	{ name: "ForceOrder", sql: "FORCE ORDER" },
	{ name: "DisableRowGoal", sql: "USE HINT DISABLE_OPTIMIZER_ROWGOAL" },
	{ name: "EnableOptimizerHotfixes", sql: "USE HINT ENABLE_QUERY_OPTIMIZER_HOTFIXES" },
	{ name: "LoopJoin", sql: "LOOP JOIN" },
	{ name: "MergeJoin", sql: "MERGE JOIN" },
	{ name: "HashJoin", sql: "HASH JOIN" },
	{ name: "NO_PERFORMANCE_SPOOL", sql: "NO_PERFORMANCE_SPOOL" },
	{ name: "ENABLE_HIST_AMENDMENT_FOR_ASC_KEYS", sql: "USE HINT ENABLE_HIST_AMENDMENT_FOR_ASC_KEYS" },
];

interface FetchOptionsEditorProps {
	node: FetchNode;
	/** Query-rule findings for this node (fetchxmlRules.ts) */
	findings: readonly RuleFinding[];
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function FetchOptionsEditor({ node, findings, onUpdate }: FetchOptionsEditorProps) {
	const styles = useEditorStyles();
	const { options } = node;
	const { retrieveAllRecords, setRetrieveAllRecords, updateNode, removeNode } = useBuilder();
	const entityInfo = useEntityInfo(node.entity.name);
	const [showTeachingPopover, setShowTeachingPopover] = useState(false);
	const [aggregateOffPlan, setAggregateOffPlan] = useState<AggregateOffPlan | null>(null);
	// The fetch node whose UTC offset has been shown: clearing the field keeps it until another query (#98)
	const [utcOffsetShownFor, setUtcOffsetShownFor] = useState<string | null>(null);
	if (options.utcOffset !== undefined && utcOffsetShownFor !== node.id) setUtcOffsetShownFor(node.id);
	const showUtcOffset = options.utcOffset !== undefined || utcOffsetShownFor === node.id;

	const handleRetrieveAllChange = (_: unknown, data: { checked: boolean }) => {
		// Show the teaching popover the first time this is turned on in a session
		if (data.checked && !hasShownRetrieveAllPopover) {
			setShowTeachingPopover(true);
			hasShownRetrieveAllPopover = true;
		}
		setRetrieveAllRecords(data.checked);
	};

	const flag = (field: keyof FetchNode["options"]) => (checked: boolean) =>
		onUpdate({ [field]: checked ? true : undefined });

	// Turning aggregate off: offer to clean up aggregate-only settings (D4)
	const handleAggregateChange = (checked: boolean) => {
		if (checked) return onUpdate({ aggregate: true });
		const plan = planAggregateOff(node);
		if (hasAggregateOffChanges(plan)) return setAggregateOffPlan(plan);
		onUpdate({ aggregate: undefined });
	};
	const applyAggregateOff = (plan: AggregateOffPlan) => {
		for (const { nodeId, updates } of plan.updates) updateNode(nodeId, updates);
		for (const nodeId of plan.removeNodeIds) removeNode(nodeId);
		onUpdate({ aggregate: undefined, aggregateLimit: undefined });
		setAggregateOffPlan(null);
	};

	// top can't be combined with count, page or returntotalrecordcount (0x80040203)
	const hasTop = options.top !== undefined;
	const hasPaging =
		options.count !== undefined || options.page !== undefined || options.returnTotalRecordCount === true;
	const topConflict = findingFor(findings, "top");
	const notWithTop = "Not available with Maximum rows (top).";
	const notWithPaging = "Not available with paging (page size, page number or total record count).";

	const hints = options.queryHints ?? [];
	const documentedNames = new Set(DOCUMENTED_HINTS.map((h) => h.name.toLowerCase()));
	const otherHints = hints.filter((h) => !documentedNames.has(h.toLowerCase()));
	const hasHint = (name: string) => hints.some((h) => h.toLowerCase() === name.toLowerCase());
	const toggleHint = (name: string, on: boolean) => {
		const next = on ? [...hints, name] : hints.filter((h) => h.toLowerCase() !== name.toLowerCase());
		onUpdate({ queryHints: next.length > 0 ? next : undefined });
	};
	const performanceSet = [options.latematerialize, hints.length > 0, options.noLock].filter(Boolean).length;

	const retentionEnabled = entityInfo.metadata?.IsRetentionEnabled;
	const dataSourceIssue = getDataSourceIssue(options.dataSource, {
		logicalName: node.entity.name,
		isRetentionEnabled: retentionEnabled,
	});
	const dataSourceLabel =
		options.dataSource === "retained"
			? "Retained data (long-term retention)"
			: options.dataSource === "bin"
				? "bin (undocumented)"
				: (options.dataSource ?? "Active data (default)");

	return (
		<div className={styles.container}>
			{/* Metadata that couldn't load, with Retry (#201) */}
			<MetadataLoadWarning entityName={node.entity.name} />
			<div className={styles.section}>
				<SectionHeading>Execution</SectionHeading>
				<TeachingPopover open={showTeachingPopover} onOpenChange={(_, data) => setShowTeachingPopover(data.open)}>
					<TeachingPopoverTrigger>
						<Switch
							checked={retrieveAllRecords}
							onChange={handleRetrieveAllChange}
							label={{
								children: (_: unknown, props: LabelProps) => (
									<InfoLabel {...props} info="Fetches every page, one after another. Use with care on large entities.">
										Retrieve all records
									</InfoLabel>
								),
							}}
						/>
					</TeachingPopoverTrigger>
					<TeachingPopoverSurface>
						<TeachingPopoverBody>
							<TeachingPopoverTitle>Retrieve all records</TeachingPopoverTitle>
							<div>Fetches all pages progressively. Use with caution on large datasets.</div>
						</TeachingPopoverBody>
						<TeachingPopoverFooter primary={{ children: "Got it", onClick: () => setShowTeachingPopover(false) }} />
					</TeachingPopoverSurface>
				</TeachingPopover>
			</div>

			<div className={styles.section}>
				<SectionHeading>Query</SectionHeading>
				<InfoCheckbox
					label="Aggregate query"
					checked={options.aggregate === true}
					onChange={handleAggregateChange}
					validation={findingFor(findings, "aggregate")}
					info={
						<Info
							text="Groups records and calculates values (count, sum, average, minimum, maximum). Every attribute then needs Group by or an aggregate function, and an alias."
							href={DOCS.aggregate}
						/>
					}
				/>
				<InfoCheckbox
					label="Distinct rows"
					checked={options.distinct === true}
					onChange={flag("distinct")}
					validation={findingFor(findings, "distinct")}
					info={<Info text="Leaves out duplicate rows from the results." href={DOCS.distinctResults} />}
				/>
				<InfoCheckbox
					label="Return total record count"
					checked={options.returnTotalRecordCount === true}
					onChange={flag("returnTotalRecordCount")}
					disabled={hasTop && !options.returnTotalRecordCount}
					hint={hasTop && !options.returnTotalRecordCount ? notWithTop : undefined}
					validation={options.returnTotalRecordCount ? topConflict : undefined}
					info={
						<Info
							text="Returns how many records match the query, up to 5,000 (500 for elastic entities), shown with the results. When more match, the count shows as 5,000+."
							href={DOCS.countRows}
						/>
					}
				/>
				<InfoCheckbox
					label="Sort option sets by value"
					checked={options.useRawOrderBy === true}
					onChange={flag("useRawOrderBy")}
					info={
						<Info
							text="Sorts option set (choice) attributes by their integer value instead of their label, for every option set attribute the query sorts by."
							href={DOCS.orderRows}
						/>
					}
				/>
			</div>

			<div className={styles.section}>
				<SectionHeading>Paging and limits</SectionHeading>
				<NumberField
					label="Maximum rows (top)"
					value={options.top}
					min={1}
					max={5000}
					rangeSeverity="warning"
					placeholder="No limit"
					disabled={hasPaging && !hasTop}
					hint={hasPaging && !hasTop ? notWithPaging : undefined}
					validation={worst(topConflict, findingFor(findings, "top"))}
					onChange={(value) => onUpdate({ top: value })}
					info={<Info text="Returns at most this many records (up to 5,000). Can't be combined with paging." href={DOCS.limitRows} />}
				/>
				<NumberField
					label="Page size (count)"
					value={options.count}
					min={1}
					max={5000}
					rangeSeverity="warning"
					placeholder="5,000"
					disabled={hasTop && options.count === undefined}
					hint={hasTop && options.count === undefined ? notWithTop : undefined}
					validation={options.count !== undefined ? worst(topConflict, findingFor(findings, "count")) : undefined}
					onChange={(value) => onUpdate({ count: value })}
					info={<Info text="Number of records per page (up to 5,000)." href={DOCS.pageResults} />}
				/>
				<NumberField
					label="Page number"
					value={options.page}
					min={1}
					placeholder="1"
					disabled={hasTop && options.page === undefined}
					hint={hasTop && options.page === undefined ? notWithTop : undefined}
					validation={options.page !== undefined ? worst(topConflict, findingFor(findings, "page")) : undefined}
					onChange={(value) => onUpdate({ page: value })}
					info={<Info text="Which page to return, starting at 1." href={DOCS.pageResults} />}
				/>
				<InfoField
					label="Paging cookie"
					hint="Copy it from the previous page's results."
					info={<Info text="Makes retrieving the next page more efficient. It must come from the page before the one requested." href={DOCS.pageResults} />}
				>
					<Input
						className={styles.mono}
						value={options.pagingCookie ?? ""}
						placeholder="<cookie … />"
						onChange={(_, data) => onUpdate({ pagingCookie: data.value || undefined })}
					/>
				</InfoField>
			</div>

			{(options.aggregate === true || options.aggregateLimit !== undefined) && (
				<div className={styles.section}>
					<SectionHeading>Aggregation</SectionHeading>
					<NumberField
						label="Aggregate row limit"
						value={options.aggregateLimit}
						min={1}
						max={50000}
						placeholder="50,000"
						hint="Leave empty for the 50,000-row default."
						validation={findingFor(findings, "aggregatelimit")}
						onChange={(value) => onUpdate({ aggregateLimit: value })}
						info={
							<Info
								text="Aggregates over at most this many records, below the 50,000-record default. When the limit is reached there's no error: the result covers at most limit + 1 arbitrary records."
								href={DOCS.aggregateLimit}
							/>
						}
					/>
				</div>
			)}

			<div className={styles.section}>
				<InfoField
					label="Data source"
					validation={dataSourceIssue ? { state: "error", message: dataSourceIssue } : findingFor(findings, "datasource")}
					info={
						<Info
							text="Retained data queries records moved to Dataverse long-term retention. The entity must have long-term retention enabled."
							href={DOCS.dataRetention}
						/>
					}
				>
					<Dropdown
						value={dataSourceLabel}
						selectedOptions={[options.dataSource ?? "active"]}
						onOptionSelect={(_, data) =>
							onUpdate({ dataSource: data.optionValue === "active" ? undefined : data.optionValue })
						}
					>
						<Option value="active">Active data (default)</Option>
						<Option
							value="retained"
							text="Retained data (long-term retention)"
							disabled={retentionEnabled === false && options.dataSource !== "retained"}
						>
							<span className={styles.optionText}>
								Retained data (long-term retention)
								{retentionEnabled === false && (
									<span className={styles.optionSecondary}>Not enabled for {node.entity.name}</span>
								)}
							</span>
						</Option>
						{options.dataSource === "bin" && <Option value="bin">bin (undocumented)</Option>}
					</Dropdown>
				</InfoField>
			</div>

			<Accordion collapsible>
				<AccordionItem value="performance">
					<AccordionHeader>
						<span className={styles.accordionHeader}>
							Performance
							{performanceSet > 0 && (
								<Badge appearance="tint" size="small" color="informative">
									{performanceSet} set
								</Badge>
							)}
						</span>
					</AccordionHeader>
					<AccordionPanel>
						<div className={styles.accordionPanel}>
							<InfoCheckbox
								label="Late materialize"
								checked={options.latematerialize === true}
								onChange={flag("latematerialize")}
								info={
									<Info
										text="Breaks the query into smaller parts and reassembles the results. Can help queries with many joins or many lookup and calculated attributes; can make simple queries slower."
										href={DOCS.lateMaterialize}
									/>
								}
							/>
							<InfoField
								label="Query hints (options)"
								info={<Info text="SQL Server query hints, passed through the options attribute." href={DOCS.fetchOptions} />}
							>
								<div className={styles.checkboxGroup}>
									<MessageBar intent="warning" layout="multiline">
										<MessageBarBody>
											Use query hints only when Microsoft support recommends them. Incorrect use can damage query performance.
										</MessageBarBody>
									</MessageBar>
									{DOCUMENTED_HINTS.map((hint) => (
										<Checkbox
											key={hint.name}
											checked={hasHint(hint.name)}
											onChange={(_, data) => toggleHint(hint.name, data.checked === true)}
											label={
												<span className={styles.optionText}>
													<span className={styles.mono}>{hint.name}</span>
													<Caption1 className={styles.secondary}>{hint.sql}</Caption1>
												</span>
											}
										/>
									))}
								</div>
							</InfoField>
							{otherHints.length > 0 && (
								<InfoField
									label="Other hints (kept as written)"
									info="Not documented option names. Dataverse accepts them, but Microsoft doesn't document what they do."
								>
									<TagGroup onDismiss={(_, data) => toggleHint(String(data.value), false)} aria-label="Other query hints">
										{otherHints.map((hint) => (
											<Tag key={hint} value={hint} dismissible dismissIcon={{ "aria-label": `Remove ${hint}` }} appearance="outline">
												{hint}
											</Tag>
										))}
									</TagGroup>
								</InfoField>
							)}
							{options.noLock && (
								<InfoCheckbox
									label="No lock (legacy)"
									checked
									onChange={flag("noLock")}
									info={
										<Info
											text="Legacy setting, no longer necessary: Dataverse doesn't need it to avoid shared locks. Shown because the query sets it."
											href={DOCS.noLock}
										/>
									}
								/>
							)}
						</div>
					</AccordionPanel>
				</AccordionItem>
			</Accordion>

			{/* utc-offset isn't in the fetch element reference: shown only when the query sets it (#98),
			    and kept while it's being edited (cleared) so the field doesn't vanish mid-edit */}
			{showUtcOffset && (
				<div className={styles.section}>
					<SectionHeading>Advanced</SectionHeading>
					<NumberField
						label="UTC offset (undocumented)"
						value={options.utcOffset}
						min={-720}
						max={720}
						rangeSeverity="warning"
						placeholder="Not set"
						onChange={(value) => onUpdate({ utcOffset: value })}
						info={
							<Info
								text="Not in the fetch element reference; kept as written. Time zone offset in minutes for date and time filters, from −720 to 720. Clear it to remove it from the query."
								href={DOCS.fetch}
							/>
						}
					/>
				</div>
			)}

			<AggregateOffDialog
				plan={aggregateOffPlan}
				onConfirm={() => aggregateOffPlan && applyAggregateOff(aggregateOffPlan)}
				onTurnOffOnly={() => {
					onUpdate({ aggregate: undefined });
					setAggregateOffPlan(null);
				}}
				onCancel={() => setAggregateOffPlan(null)}
			/>
		</div>
	);
}
