/**
 * Property editor for <attribute>: every attribute in the FetchXML reference, driven by attribute
 * metadata (#90b). https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/attribute
 */

import { useEffect, useState } from "react";
import { Dropdown, Input, Option, Radio, RadioGroup } from "@fluentui/react-components";
import type { AttributeNode, EntityNode, FetchNode, LinkEntityNode } from "../../../../model/nodes";
import type { RuleFinding } from "../../../../model/fetchxmlRules";
import {
	AGGREGATE_FUNCTIONS,
	getAttributeCapabilities,
	getUnsupportedFunctionReason,
	type AggregateFunction,
} from "../../../../model/attributeCapabilities";
import { findOrdersUsingAlias } from "../../../../model/treeUtils";
import { sanitizeAlias } from "../../../../model/aliasUtils";
import { useBuilder } from "../../../../state/builderContext";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoCheckbox, InfoField, LeftoverSetting, SectionHeading, type Validation } from "../fields";
import { findingFor, worst } from "../ruleFindings";
import { useAttributeTypeInfo, useEntityInfo } from "../useEditorMetadata";

const FUNCTION_LABELS: Record<AggregateFunction, string> = {
	count: "Count records (count)",
	countcolumn: "Count values (countcolumn)",
	sum: "Sum (sum)",
	avg: "Average (avg)",
	min: "Minimum (min)",
	max: "Maximum (max)",
};

const DATE_PARTS: readonly [NonNullable<AttributeNode["dategrouping"]>, string][] = [
	["day", "Day"],
	["week", "Week"],
	["month", "Month"],
	["quarter", "Quarter"],
	["year", "Year"],
	["fiscal-period", "Fiscal period"],
	["fiscal-year", "Fiscal year"],
];
const NO_DATE_PART = "__none";

interface AttributeEditorProps {
	node: AttributeNode;
	/** The entity or link-entity that holds this attribute */
	container: EntityNode | LinkEntityNode;
	isRootEntity: boolean;
	isAggregateQuery: boolean;
	fetchQuery: FetchNode;
	/** Query-rule findings for this node (fetchxmlRules.ts) */
	findings: readonly RuleFinding[];
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function AttributeEditor({
	node,
	container,
	isRootEntity,
	isAggregateQuery,
	fetchQuery,
	findings,
	onUpdate,
}: AttributeEditorProps) {
	const styles = useEditorStyles();
	const { updateNode } = useBuilder();
	const entityName = container.name;
	const typeInfo = useAttributeTypeInfo(entityName, node.name || undefined);
	const caps = typeInfo ? getAttributeCapabilities(typeInfo) : undefined;
	const entityInfo = useEntityInfo(isRootEntity ? entityName : undefined);
	const primaryId = entityInfo.metadata?.PrimaryIdAttribute;

	// "Aggregate" chosen before a function: a UI-only state until a function is picked
	const [aggregatePending, setAggregatePending] = useState(false);
	useEffect(() => setAggregatePending(false), [node.id]);

	// Renaming an alias also updates the orders that sort by it, so the rename doesn't break them
	const ordersUsingAlias = node.alias ? findOrdersUsingAlias(fetchQuery, container, node.alias) : [];
	const handleAliasChange = (raw: string) => {
		const alias = sanitizeAlias(raw) || undefined;
		if (node.alias && alias) {
			for (const order of findOrdersUsingAlias(fetchQuery, container, node.alias)) updateNode(order.id, { alias });
		}
		onUpdate({ alias });
	};

	const aliasRequired = isAggregateQuery || node.rowAggregate !== undefined;
	const attributeField = (
		<InfoField
			label="Attribute"
			required
			info={<Info text="The logical name of the attribute to return." href={DOCS.attribute} />}
		>
			<AttributePicker
				entityLogicalName={entityName}
				value={node.name}
				onChange={(name) => onUpdate({ name })}
				placeholder="Select an attribute"
			/>
		</InfoField>
	);
	const aliasField = (
		<InfoField
			label="Alias"
			required={aliasRequired}
			hint={
				ordersUsingAlias.length > 0
					? `Sorting by this alias: renaming it also updates ${ordersUsingAlias.length === 1 ? "that order" : `those ${ordersUsingAlias.length} orders`}.`
					: undefined
			}
			validation={findingFor(findings, "alias")}
			info={
				<Info
					text="The name of the attribute in the results: letters, digits and underscores, starting with a letter or underscore. Unique across the query."
					href={DOCS.columnAliases}
				/>
			}
		>
			<Input value={node.alias ?? ""} onChange={(_, data) => handleAliasChange(data.value)} />
		</InfoField>
	);

	// ---------------------------------------------------------------- non-aggregate query
	if (!isAggregateQuery) {
		const isPrimaryKey = !!node.name && node.name === primaryId;
		const showCountChildren =
			node.rowAggregate !== undefined || (isRootEntity && isPrimaryKey && entityInfo.isHierarchical !== false);
		let countChildrenValidation: Validation = findingFor(findings, "rowaggregate");
		if (!countChildrenValidation && node.rowAggregate !== undefined) {
			if (entityInfo.isHierarchical === false) {
				countChildrenValidation = {
					state: "error",
					message: `${entityName} has no hierarchical relationship, so Dataverse rejects this (0x80047020).`,
				};
			} else if (primaryId && !isPrimaryKey) {
				countChildrenValidation = {
					state: "error",
					message: `Only works on the primary key (${primaryId}); Dataverse rejects it on other attributes (0x80040203).`,
				};
			}
		}
		const leftovers: { text: string; field: keyof AttributeNode; validation: NonNullable<Validation> }[] = [];
		const ignored = (text: string): NonNullable<Validation> => ({ state: "none", message: text });
		if (node.groupby) leftovers.push({ text: "Group by", field: "groupby", validation: findingFor(findings, "groupby") ?? ignored("") });
		if (node.aggregate) leftovers.push({ text: `Aggregate: ${node.aggregate}`, field: "aggregate", validation: findingFor(findings, "aggregate") ?? ignored("") });
		if (node.dategrouping) leftovers.push({ text: `Date grouping: ${node.dategrouping}`, field: "dategrouping", validation: ignored("Ignored outside aggregate queries.") });
		if (node.usertimezone !== undefined) leftovers.push({ text: "Time zone", field: "usertimezone", validation: ignored("Ignored outside aggregate queries.") });
		if (node.distinct) leftovers.push({ text: "Unique values only", field: "distinct", validation: ignored("Ignored outside aggregate queries.") });

		return (
			<div className={styles.container}>
				<div className={styles.section}>
					<SectionHeading>Attribute</SectionHeading>
					{attributeField}
					{aliasField}
				</div>

				{showCountChildren && (
					<div className={styles.section}>
						<SectionHeading>Hierarchy</SectionHeading>
						<InfoCheckbox
							label="Count child records"
							checked={node.rowAggregate !== undefined}
							onChange={(checked) => onUpdate({ rowAggregate: checked ? "CountChildren" : undefined })}
							hint={node.rowAggregate !== undefined ? "Returns the number of hierarchical child records in the alias attribute." : undefined}
							validation={countChildrenValidation}
							info={
								<Info
									text="Adds the number of child records under each record in a hierarchical relationship, such as parent accounts. The count includes records the user can't read. Only for entities with a hierarchical relationship."
									href={DOCS.rowAggregate}
								/>
							}
						/>
					</div>
				)}

				{leftovers.length > 0 && (
					<div className={styles.section}>
						<SectionHeading>Aggregate settings</SectionHeading>
						{leftovers.map((leftover) => (
							<LeftoverSetting
								key={leftover.field}
								text={leftover.text}
								validation={
									leftover.validation.message
										? leftover.validation
										: { state: "error", message: "The query isn't an aggregate query, so Dataverse rejects this." }
								}
								onRemove={() => onUpdate({ [leftover.field]: undefined })}
							/>
						))}
					</div>
				)}
			</div>
		);
	}

	// ---------------------------------------------------------------- aggregate query
	const bothRoles = !!node.groupby && !!node.aggregate;
	const role = bothRoles ? "" : node.groupby ? "groupby" : node.aggregate || aggregatePending ? "aggregate" : "";
	const canGroup = !caps || caps.groupby;
	const allowedFunctions = AGGREGATE_FUNCTIONS.filter((fn) => !caps || caps.functions.includes(fn));
	const canAggregate = allowedFunctions.length > 0;
	const functionUnsupported = !!node.aggregate && !!typeInfo && !allowedFunctions.includes(node.aggregate);

	const roleValidation = worst(
		findingFor(findings, undefined, "groupby"),
		node.groupby && !canGroup ? { state: "error", message: caps?.reason ?? "Dataverse can't group by this attribute." } : undefined,
		!role && !bothRoles && !canGroup && !canAggregate
			? { state: "error", message: `${caps?.reason} Remove this attribute from the aggregate query.` }
			: undefined,
	);
	const functionValidation: Validation = !node.aggregate
		? { state: "error", message: "Choose an aggregate function." }
		: functionUnsupported
			? { state: "error", message: getUnsupportedFunctionReason(node.aggregate, typeInfo!)! }
			: findingFor(findings, "aggregate");

	const isDateTime = caps?.groupbyNeedsDateGrouping === true;
	const typeUnknown = !typeInfo;
	const dateGroupingRelevant = role === "groupby" && (isDateTime || typeUnknown);
	const dateGroupingValidation: Validation =
		role === "groupby" && isDateTime && !node.dategrouping
			? { state: "error", message: "Choose a date part. Dataverse can't group a DateTime attribute without one (0x80041135)." }
			: undefined;
	const dateGroupingIgnored = node.dategrouping !== undefined && !dateGroupingRelevant;

	return (
		<div className={styles.container}>
			<div className={styles.section}>
				<SectionHeading>Attribute</SectionHeading>
				{attributeField}
				{aliasField}
			</div>

			<div className={styles.section}>
				<InfoField
					label="Role"
					required
					validation={roleValidation}
					info={<Info text="Group by makes a group per distinct value. Aggregate calculates a value for each group." href={DOCS.aggregate} />}
				>
					<RadioGroup
						layout="horizontal"
						value={role}
						onChange={(_, data) => {
							setAggregatePending(data.value === "aggregate");
							if (data.value === "groupby") onUpdate({ groupby: true, aggregate: undefined, distinct: undefined });
							else onUpdate({ groupby: undefined, dategrouping: undefined, usertimezone: undefined });
						}}
					>
						<Radio value="groupby" label="Group by" disabled={!canGroup && !node.groupby} />
						<Radio value="aggregate" label="Aggregate" disabled={!canAggregate && !node.aggregate} />
					</RadioGroup>
				</InfoField>

				{role === "aggregate" && (
					<>
						<InfoField
							label="Function"
							required
							validation={functionValidation}
							hint={caps && allowedFunctions.length < AGGREGATE_FUNCTIONS.length ? "Only functions this attribute type supports are listed." : undefined}
							info={
								<Info
									text="Count records counts every record; Count values counts records where this attribute has a value."
									href={DOCS.aggregateFunctions}
								/>
							}
						>
							<Dropdown
								placeholder="Choose a function"
								value={node.aggregate ? FUNCTION_LABELS[node.aggregate] : ""}
								selectedOptions={node.aggregate ? [node.aggregate] : []}
								onOptionSelect={(_, data) => onUpdate({ aggregate: data.optionValue })}
							>
								{functionUnsupported && node.aggregate && <Option value={node.aggregate}>{FUNCTION_LABELS[node.aggregate]}</Option>}
								{allowedFunctions.map((fn) => (
									<Option key={fn} value={fn}>
										{FUNCTION_LABELS[fn]}
									</Option>
								))}
							</Dropdown>
						</InfoField>
						{(node.aggregate === "countcolumn" || node.distinct) && (
							<InfoCheckbox
								label="Unique values only"
								checked={!!node.distinct}
								onChange={(checked) => onUpdate({ distinct: checked ? true : undefined })}
								validation={findingFor(findings, "distinct")}
								info={<Info text="With Count values, counts each distinct value once." href={DOCS.aggregateDistinct} />}
							/>
						)}
					</>
				)}

				{dateGroupingRelevant && (
					<InfoField
						label="Date grouping"
						required={isDateTime}
						validation={dateGroupingValidation}
						info={<Info text="The part of the date to group by." href={DOCS.aggregateDateGrouping} />}
					>
						<Dropdown
							placeholder="Choose a date part"
							value={DATE_PARTS.find(([value]) => value === node.dategrouping)?.[1] ?? ""}
							selectedOptions={node.dategrouping ? [node.dategrouping] : []}
							onOptionSelect={(_, data) =>
								onUpdate({ dategrouping: data.optionValue === NO_DATE_PART ? undefined : data.optionValue })
							}
						>
							{typeUnknown && <Option value={NO_DATE_PART}>None</Option>}
							{DATE_PARTS.map(([value, label]) => (
								<Option key={value} value={value}>
									{label}
								</Option>
							))}
						</Dropdown>
					</InfoField>
				)}
				{dateGroupingIgnored && (
					<LeftoverSetting
						text={`Date grouping: ${node.dategrouping}`}
						validation={{ state: "warning", message: "Ignored: date grouping only applies to Group by on DateTime attributes." }}
						onRemove={() => onUpdate({ dategrouping: undefined })}
					/>
				)}

				{(dateGroupingRelevant || node.usertimezone !== undefined) && (
					<InfoField
						label="Time zone"
						info={
							<Info
								text="The same date can fall on different days in different time zones. By default Dataverse groups using the user's time zone; choose UTC to group on the stored UTC value."
								href={DOCS.attribute}
							/>
						}
					>
						<RadioGroup
							value={node.usertimezone === false ? "utc" : "user"}
							onChange={(_, data) =>
								onUpdate({ usertimezone: data.value === "utc" ? false : node.usertimezone === true ? true : undefined })
							}
						>
							<Radio value="user" label="User's time zone (default)" />
							<Radio value="utc" label="UTC" />
						</RadioGroup>
					</InfoField>
				)}

				{node.rowAggregate !== undefined && (
					<LeftoverSetting
						text="Count child records"
						validation={findingFor(findings, "rowaggregate") ?? { state: "error", message: "Doesn't work in an aggregate query." }}
						onRemove={() => onUpdate({ rowAggregate: undefined })}
					/>
				)}
			</div>
		</div>
	);
}
