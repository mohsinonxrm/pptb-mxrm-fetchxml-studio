/**
 * Property editor for <order>: attribute (normal queries), alias (aggregate queries), descending
 * and entityname (#90b). https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/order
 *
 * Aggregate queries can only sort by an alias defined in the order's own entity / link-entity, or
 * in the link-entity named by `entityname` (verified live: anything else fails with 0x8004110a).
 */

import { Dropdown, Option, Radio, RadioGroup } from "@fluentui/react-components";
import type { AttributeNode, EntityNode, FetchNode, LinkEntityNode, OrderNode } from "../../../../model/nodes";
import type { RuleFinding } from "../../../../model/fetchxmlRules";
import { getOrderMetadataIssue } from "../../../../model/attributeCapabilities";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoField, LeftoverSetting, SectionHeading } from "../fields";
import { findingFor, worst } from "../ruleFindings";
import { useAttributeTypeInfo } from "../useEditorMetadata";
import { MetadataLoadWarning } from "../MetadataLoadWarning";

/** AttributeType values Dataverse can sort by: multi-select option sets, images and files can't (§9A). */
const ORDERABLE_TYPES = [
	"String", "Memo", "Integer", "BigInt", "Decimal", "Double", "Money", "DateTime", "Boolean",
	"Picklist", "State", "Status", "Lookup", "Customer", "Owner", "Uniqueidentifier", "EntityName",
];
const ROOT = "__root";

const joinedLinks = (container: EntityNode | LinkEntityNode): LinkEntityNode[] =>
	container.links.flatMap((link) => [link, ...joinedLinks(link)]);

const roleText = (attr: AttributeNode) =>
	attr.groupby
		? `Group by${attr.dategrouping ? ` · ${attr.dategrouping}` : ""}`
		: attr.aggregate
			? { count: "Count records", countcolumn: "Count values", sum: "Sum", avg: "Average", min: "Minimum", max: "Maximum" }[attr.aggregate]
			: "";

interface OrderEditorProps {
	node: OrderNode;
	/** The entity or link-entity that holds this order */
	container: EntityNode | LinkEntityNode;
	isRootEntity: boolean;
	fetchQuery: FetchNode;
	/** Query-rule findings for this node (fetchxmlRules.ts) */
	findings: readonly RuleFinding[];
	/** Findings on the <fetch> node, for aggregatelimit + order */
	fetchFindings: readonly RuleFinding[];
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function OrderEditor({ node, container, isRootEntity, fetchQuery, findings, fetchFindings, onUpdate }: OrderEditorProps) {
	const styles = useEditorStyles();
	const isAggregate = fetchQuery.options.aggregate === true;
	// Joined link-entities (not the any / not any links inside filters, which return no rows to sort)
	const links = joinedLinks(fetchQuery.entity);
	// Link-entities an order can target: by alias, or by name when the name is unique and there's no alias
	const nameCounts = links.reduce<Record<string, number>>((acc, l) => ({ ...acc, [l.name]: (acc[l.name] ?? 0) + 1 }), {});
	const targets = links
		.map((l) => ({ link: l, key: l.alias ?? (nameCounts[l.name] === 1 ? l.name : undefined) }))
		.filter((t): t is { link: LinkEntityNode; key: string } => !!t.key);
	const targetLink = node.entityname
		? (targets.find((t) => t.key === node.entityname)?.link ?? links.find((l) => l.name === node.entityname))
		: undefined;
	const scope: EntityNode | LinkEntityNode = isRootEntity ? (targetLink ?? container) : container;
	const scopeLabel = scope.type === "entity" ? `the root entity (${scope.name})` : `${scope.alias ?? scope.name} (${scope.name})`;

	const aliases = scope.attributes.filter((a) => a.alias && (a.groupby || a.aggregate));
	const aliasInScope = !!node.alias && aliases.some((a) => a.alias === node.alias);
	const typeInfo = useAttributeTypeInfo(scope.name, !isAggregate ? node.attribute : undefined);
	const orderTypeIssue = getOrderMetadataIssue(typeInfo);
	const limitConflict = findingFor(fetchFindings, "aggregatelimit");

	const entityValue = node.entityname
		? targetLink
			? `${targetLink.alias ?? targetLink.name} (${targetLink.name})`
			: node.entityname
		: `Root: ${fetchQuery.entity.name}`;

	return (
		<div className={styles.container}>
			{/* Metadata that couldn't load, with Retry (#201) */}
			<MetadataLoadWarning entityName={scope.name} />
			<div className={styles.section}>
				<SectionHeading>Sort</SectionHeading>

				{isRootEntity && targets.length > 0 && (
					<InfoField
						label="Entity"
						validation={findingFor(findings, "entityname")}
						info={
							<Info
								text="Choose a link-entity to apply its sort before the root entity's sorts. Otherwise link-entity sorts are applied last."
								href={DOCS.linkOrdersFirst}
							/>
						}
					>
						<Dropdown
							value={entityValue}
							selectedOptions={[node.entityname ?? ROOT]}
							onOptionSelect={(_, data) =>
								onUpdate({ entityname: data.optionValue === ROOT ? undefined : data.optionValue, attribute: undefined, alias: undefined })
							}
						>
							<Option value={ROOT}>{`Root: ${fetchQuery.entity.name}`}</Option>
							{targets.map(({ link, key }) => (
								<Option key={link.id} value={key}>{`${link.alias ?? link.name} (${link.name})`}</Option>
							))}
						</Dropdown>
					</InfoField>
				)}

				{!isAggregate ? (
					<InfoField
						label="Attribute"
						required
						validation={worst(
							!node.attribute && !node.alias ? { state: "error", message: "Choose an attribute to sort by." } : undefined,
							orderTypeIssue ? { state: "error", message: orderTypeIssue } : undefined,
							findingFor(findings, "attribute"),
						)}
						info={<Info text="The logical name of the attribute to sort by. It doesn't need to be a selected attribute." href={DOCS.orderRows} />}
					>
						<AttributePicker
							entityLogicalName={scope.name}
							filterByTypes={ORDERABLE_TYPES}
							value={node.attribute ?? ""}
							onChange={(attribute) => onUpdate({ attribute: attribute || undefined })}
							placeholder="Select an attribute"
						/>
					</InfoField>
				) : (
					<InfoField
						label="Sort by"
						required
						validation={worst(
							!node.alias
								? {
										state: "error",
										message: aliases.length === 0
											? `Add an alias to a Group by or Aggregate attribute of ${scopeLabel} first.`
											: "Choose an alias to sort by.",
									}
								: undefined,
							findingFor(findings, "alias"),
							limitConflict,
						)}
						info={
							<Info
								text="Aggregate queries sort by an attribute alias. Only aliases of Group by or Aggregate attributes of the chosen entity are listed."
								href={DOCS.order}
							/>
						}
					>
						<Dropdown
							placeholder="Choose an alias"
							value={node.alias ?? ""}
							selectedOptions={node.alias ? [node.alias] : []}
							onOptionSelect={(_, data) => onUpdate({ alias: data.optionValue, attribute: undefined })}
						>
							{node.alias && !aliasInScope && (
								<Option value={node.alias} text={node.alias}>
									<span className={styles.optionText}>
										{node.alias}
										<span className={styles.optionSecondary}>Not defined on this entity</span>
									</span>
								</Option>
							)}
							{aliases.map((attr) => (
								<Option key={attr.id} value={attr.alias!} text={attr.alias!}>
									<span className={styles.optionText}>
										{attr.alias}
										<span className={styles.optionSecondary}>
											{attr.name} · {roleText(attr)}
										</span>
									</span>
								</Option>
							))}
						</Dropdown>
					</InfoField>
				)}

				{/* Values that don't apply in this mode (parsed XML or after a mode change) */}
				{!isAggregate && node.alias && (
					<LeftoverSetting
						text={`Alias: ${node.alias}`}
						validation={findingFor(findings, "alias") ?? { state: "error", message: "Only aggregate queries sort by alias." }}
						onRemove={() => onUpdate({ alias: undefined })}
					/>
				)}
				{isAggregate && node.attribute && (
					<LeftoverSetting
						text={`Attribute: ${node.attribute}`}
						validation={findingFor(findings, "attribute") ?? { state: "error", message: "Aggregate queries sort by alias." }}
						onRemove={() => onUpdate({ attribute: undefined })}
					/>
				)}
				{!isRootEntity && node.entityname && (
					<LeftoverSetting
						text={`Entity: ${node.entityname}`}
						validation={findingFor(findings, "entityname") ?? { state: "error", message: "Not allowed on an order inside a link-entity." }}
						onRemove={() => onUpdate({ entityname: undefined })}
					/>
				)}

				<InfoField
					label="Direction"
					info={
						<Info
							text="Option set attributes sort by label unless 'Sort option sets by value' is on for the query. Lookups sort by the related record's primary name."
							href={DOCS.orderRows}
						/>
					}
				>
					<RadioGroup
						value={node.descending ? "desc" : "asc"}
						onChange={(_, data) => onUpdate({ descending: data.value === "desc" ? true : undefined })}
					>
						<Radio value="asc" label="Ascending" />
						<Radio value="desc" label="Descending" />
					</RadioGroup>
				</InfoField>
			</div>
		</div>
	);
}
