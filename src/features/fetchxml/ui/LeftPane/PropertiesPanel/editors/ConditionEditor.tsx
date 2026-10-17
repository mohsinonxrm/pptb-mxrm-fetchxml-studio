/**
 * Property editor for <condition>: attribute, operator, value(s) or valueof, and entityname (#90b).
 * Operators and value pickers follow the attribute's metadata.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/condition
 */

import { useState, useEffect, useMemo } from "react";
import { Input, Dropdown, Option, Radio, RadioGroup } from "@fluentui/react-components";
import type { ConditionNode, FetchNode } from "../../../../model/nodes";
import type { RuleFinding } from "../../../../model/fetchxmlRules";
import { VALUEOF_OPERATORS, valueofScopes } from "../../../../model/fetchxmlRules";
import { getValueOfCompatibleTypes } from "../../../../model/attributeCapabilities";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { OptionSetValuePicker } from "../../../../../../shared/components/OptionSetValuePicker";
import { BooleanValuePicker } from "../../../../../../shared/components/BooleanValuePicker";
import { NumericInputPicker } from "../../../../../../shared/components/NumericInputPicker";
import { DateTimeInputPicker } from "../../../../../../shared/components/DateTimeInputPicker";
import { BetweenInputPicker } from "../../../../../../shared/components/BetweenInputPicker";
import { MultiValuePicker } from "../../../../../../shared/components/MultiValuePicker";
import {
	adaptValueForOperator,
	getOperatorsForAttribute,
	operatorRequiresValue,
	operatorValueShape,
} from "../../../../model/operators";
import { loadAllEntityAttributes, loadAttributeDetailedMetadata } from "../../../../api/dataverseMetadata";
import type { AttributeMetadata } from "../../../../api/pptbClient";
import { collectLinkEntityReferences, isConditionInRootEntityFilter } from "../../../../model/treeUtils";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoField, LeftoverSetting, NumberField, SectionHeading } from "../fields";
import { useEntityInfo } from "../useEditorMetadata";
import { findingFor } from "../ruleFindings";
import { ValueOfPicker } from "./ValueOfPicker";

const THIS_ENTITY = "__this";
const OPTION_SET_TYPES = new Set(["Picklist", "State", "Status"]);
const NUMERIC_TYPES = new Set(["Integer", "BigInt", "Decimal", "Double"]);
/** Whole number from a stored value (parsed FetchXML keeps numbers as numbers; strings are possible) */
const toInteger = (v: unknown): number | undefined => {
	const n = typeof v === "number" ? v : typeof v === "string" && /^-?\d+$/.test(v.trim()) ? Number(v) : NaN;
	return Number.isInteger(n) ? n : undefined;
};

interface ConditionEditorProps {
	node: ConditionNode;
	/** The entity / link-entity whose filter holds the condition */
	entityName: string;
	fetchQuery: FetchNode;
	/** Query-rule findings for this node (fetchxmlRules.ts) */
	findings: readonly RuleFinding[];
	onUpdate: (updates: Record<string, unknown>) => void;
	/**
	 * A results-grid column filter (#23): the column fixes the attribute, so only the operator and
	 * value(s) are shown. The condition isn't in the tree yet; entityName is the column's entity.
	 */
	columnFilter?: boolean;
}

export function ConditionEditor({ node, entityName, fetchQuery, findings, onUpdate, columnFilter = false }: ConditionEditorProps) {
	const styles = useEditorStyles();
	const [selectedAttribute, setSelectedAttribute] = useState<AttributeMetadata | null>(null);

	const linkEntityReferences = useMemo(() => collectLinkEntityReferences(fetchQuery), [fetchQuery]);
	// entityname applies to conditions in the root entity's filters (filters on link-entity, after the join)
	const isInRootFilter = useMemo(() => isConditionInRootEntityFilter(fetchQuery, node.id), [fetchQuery, node.id]);
	const entityTarget = node.entityname
		? linkEntityReferences.find((r) => r.identifier === node.entityname)
		: undefined;
	const effectiveEntityName = isInRootFilter && entityTarget ? entityTarget.entityName : entityName;

	// What "alias.attribute" can reach from where this condition is
	const valueofScope = useMemo(() => valueofScopes(fetchQuery).get(node.id), [fetchQuery, node.id]);
	// Hierarchy operators need a hierarchical entity (#87)
	const { isHierarchical } = useEntityInfo(effectiveEntityName);
	// Operators from the attribute's metadata: type, multi-select, primary key, date-only behaviour (#87)
	const availableOperators = useMemo(
		() => getOperatorsForAttribute(selectedAttribute ?? undefined, isHierarchical),
		[selectedAttribute, isHierarchical],
	);
	const compatibleTypes = useMemo(
		() => getValueOfCompatibleTypes(selectedAttribute?.AttributeType),
		[selectedAttribute?.AttributeType],
	);

	// The attribute's metadata decides the operators and the value picker
	useEffect(() => {
		let cancelled = false;
		if (!node.attribute || !effectiveEntityName) {
			setSelectedAttribute(null);
			return;
		}
		// A lookup, not a picker: every attribute, so hidden ones keep their operators (#57)
		loadAllEntityAttributes(effectiveEntityName)
			.then(async (attributes: AttributeMetadata[]) => {
				const attr = attributes.find((a) => a.LogicalName === node.attribute);
				const type = attr?.AttributeType;
				let detailed = attr;
				// Ranges, precision and date formats come from the type-cast metadata
				if (attr && type && (NUMERIC_TYPES.has(type) || type === "DateTime")) {
					try {
						detailed = await loadAttributeDetailedMetadata(effectiveEntityName, node.attribute, type);
					} catch (error) {
						console.error(`Failed to load detailed metadata for ${node.attribute}:`, error);
					}
				}
				if (cancelled) return;
				setSelectedAttribute(detailed ?? null);
			})
			.catch((error) => {
				console.error("Failed to load attribute metadata for operator filtering:", error);
				if (cancelled) return;
				setSelectedAttribute(null);
			});
		return () => {
			cancelled = true;
		};
	}, [node.attribute, effectiveEntityName]);

	const requiresValue = operatorRequiresValue(node.operator);
	const operatorDef = availableOperators.find((op) => op.value === node.operator);
	// The operator decides the value's shape first, then the attribute's type (#87)
	const shape = operatorValueShape(node.operator);
	const requiresTwoValues = shape === "two";
	const requiresMultipleValues = shape === "multiple";
	const isSingleValue = shape === "single" || (shape === "two" && !selectedAttribute) || (shape === "multiple" && !selectedAttribute);
	const comparesToAttribute = node.valueof !== undefined;
	const operatorComparesAttributes = VALUEOF_OPERATORS.has(node.operator);
	// Only type pairs verified to work are offered (getValueOfCompatibleTypes)
	const canCompareToAttribute = operatorComparesAttributes && compatibleTypes.length > 0;
	const valueofFinding = findingFor(findings, "valueof");
	const attributeType = selectedAttribute?.AttributeType;

	const showEntity = !columnFilter && ((isInRootFilter && linkEntityReferences.length > 0) || !!node.entityname);
	const entityLabel = node.entityname
		? (entityTarget?.displayLabel ?? node.entityname)
		: `This entity (${entityName})`;

	const renderLiteralValue = () => {
		if (attributeType && OPTION_SET_TYPES.has(attributeType)) {
			return (
				<OptionSetValuePicker
					entityLogicalName={effectiveEntityName}
					attributeLogicalName={node.attribute}
					value={typeof node.value === "number" ? node.value : undefined}
					onChange={(val) => onUpdate({ value: val })}
					placeholder="Select a value"
				/>
			);
		}
		if (attributeType === "Boolean") {
			return (
				<BooleanValuePicker
					entityLogicalName={effectiveEntityName}
					attributeLogicalName={node.attribute}
					value={
						typeof node.value === "boolean"
							? node.value
							: typeof node.value === "string"
								? node.value === "true" || node.value === "1"
								: undefined
					}
					onChange={(val) => onUpdate({ value: val })}
					placeholder="Select true or false"
				/>
			);
		}
		if (attributeType && NUMERIC_TYPES.has(attributeType) && selectedAttribute) {
			return (
				<NumericInputPicker
					attribute={selectedAttribute}
					value={typeof node.value === "number" ? node.value : undefined}
					onChange={(val) => onUpdate({ value: val })}
					placeholder="Enter a number"
				/>
			);
		}
		if (attributeType === "DateTime" && selectedAttribute) {
			return (
				<DateTimeInputPicker
					attribute={selectedAttribute}
					value={typeof node.value === "string" ? node.value : node.value instanceof Date ? node.value : undefined}
					onChange={(val) => onUpdate({ value: val?.toISOString() })}
					placeholder="Select a date"
				/>
			);
		}
		return (
			<Input
				value={typeof node.value === "string" || typeof node.value === "number" ? String(node.value) : ""}
				onChange={(_, data) => onUpdate({ value: data.value || undefined })}
				placeholder="Value"
			/>
		);
	};

	return (
		<div className={styles.container}>
			<div className={styles.section}>
				{!columnFilter && <SectionHeading>Condition</SectionHeading>}

				{showEntity && (
					<InfoField
						label="Entity"
						validation={findingFor(findings, "entityname")}
						info={
							<Info
								text="Filter on a link-entity's attribute after the join instead of within it, usually with an outer link-entity (for example, records with no related record). Uses the link-entity's alias, or its name when it has none."
								href={DOCS.filtersOnLinkEntity}
							/>
						}
					>
						<Dropdown
							value={entityLabel}
							selectedOptions={[node.entityname ?? THIS_ENTITY]}
							onOptionSelect={(_, data) =>
								// The attribute belongs to the chosen entity, so start over
								onUpdate({
									entityname: data.optionValue === THIS_ENTITY ? undefined : data.optionValue,
									attribute: undefined,
									operator: undefined,
									value: undefined,
									valueof: undefined,
								})
							}
						>
							<Option value={THIS_ENTITY}>{`This entity (${entityName})`}</Option>
							{node.entityname && !entityTarget && (
								<Option value={node.entityname} text={node.entityname}>
									<span className={styles.optionText}>
										{node.entityname}
										<span className={styles.optionSecondary}>Not a link-entity in this query</span>
									</span>
								</Option>
							)}
							{isInRootFilter &&
								linkEntityReferences.map((ref) => (
									<Option key={ref.nodeId} value={ref.identifier}>
										{ref.displayLabel}
									</Option>
								))}
						</Dropdown>
					</InfoField>
				)}

				{!columnFilter && (
					<InfoField
						label="Attribute"
						required
						validation={!node.attribute ? { state: "error", message: "Choose an attribute to filter on." } : undefined}
						info={<Info text={`The logical name of the ${effectiveEntityName} attribute to filter on.`} href={DOCS.condition} />}
					>
						<AttributePicker
							entityLogicalName={effectiveEntityName}
							value={node.attribute}
							onChange={(logicalName) =>
								// A different attribute can have a different type: reset operator and value, keep the mode
								onUpdate({
									attribute: logicalName || undefined,
									operator: undefined,
									value: undefined,
									valueof: comparesToAttribute ? "" : undefined,
								})
							}
							placeholder="Select an attribute"
						/>
					</InfoField>
				)}

				<InfoField
					label="Operator"
					required
					info={<Info text="The operators listed depend on the attribute's type." href={DOCS.operators} />}
				>
					<Dropdown
						value={operatorDef?.label ?? node.operator ?? ""}
						selectedOptions={node.operator ? [node.operator] : []}
						onOptionSelect={(_, data) => {
							const operator = data.optionValue ?? "";
							// Keep the value only when the new operator takes the same shape (#87); only eq / ne /
							// gt / ge / lt / le compare to another attribute (0x80048415)
							onUpdate({
								operator,
								value: adaptValueForOperator(node.operator, operator, node.value),
								...(comparesToAttribute && !VALUEOF_OPERATORS.has(operator) ? { valueof: undefined } : {}),
							});
						}}
						placeholder="Select an operator"
					>
						{node.operator && !operatorDef && availableOperators.length > 0 && (
							<Option value={node.operator} text={node.operator}>
								<span className={styles.optionText}>
									{node.operator}
									<span className={styles.optionSecondary}>Not listed for this attribute type</span>
								</span>
							</Option>
						)}
						{availableOperators.map((op) => (
							<Option key={op.value} value={op.value}>
								{op.label}
							</Option>
						))}
					</Dropdown>
				</InfoField>

				{(shape === "integer" || shape === "year") && (
					<NumberField
						label={shape === "year" ? "Fiscal year" : "Number"}
						value={toInteger(node.value)}
						min={shape === "year" ? 1000 : 1}
						max={shape === "year" ? 9999 : undefined}
						onChange={(value) => onUpdate({ value })}
						validation={toInteger(node.value) === undefined ? { state: "error", message: shape === "year" ? "Enter a four-digit fiscal year." : "Enter a whole number above 0." } : undefined}
						info={
							<Info
								text={
									shape === "year"
										? "A four-digit fiscal year, per the environment's fiscal year settings."
										: "A whole number above 0 (for example 7 for the last 7 days). Dataverse rejects 0, negative numbers and dates."
								}
								href={DOCS.operators}
							/>
						}
					/>
				)}

				{shape === "fiscalPeriodYear" && (
					<>
						<NumberField
							label="Fiscal period"
							value={Array.isArray(node.value) ? toInteger(node.value[0]) : undefined}
							min={1}
							onChange={(period) =>
								onUpdate({ value: [period ?? "", Array.isArray(node.value) ? (node.value[1] ?? "") : ""] })
							}
							validation={!Array.isArray(node.value) || toInteger(node.value[0]) === undefined ? { state: "error", message: "Enter the fiscal period (1 or more)." } : undefined}
							info={
								<Info
									text="The period number within the fiscal year: a month, quarter or semester, depending on the environment's fiscal year settings. A period that doesn't exist returns no records."
									href={DOCS.operators}
								/>
							}
						/>
						<NumberField
							label="Fiscal year"
							value={Array.isArray(node.value) ? toInteger(node.value[1]) : undefined}
							min={1000}
							max={9999}
							onChange={(year) =>
								onUpdate({ value: [Array.isArray(node.value) ? (node.value[0] ?? "") : "", year ?? ""] })
							}
							validation={!Array.isArray(node.value) || toInteger(node.value[1]) === undefined ? { state: "error", message: "Enter a four-digit fiscal year." } : undefined}
							info={<Info text="A four-digit fiscal year." href={DOCS.operators} />}
						/>
					</>
				)}

				{requiresValue && requiresMultipleValues && selectedAttribute && (
					<InfoField
						label="Values"
						info={<Info text="Each value is written as a <value> element." href={DOCS.condition} />}
					>
						<MultiValuePicker
							entityLogicalName={effectiveEntityName}
							attributeLogicalName={node.attribute}
							attribute={selectedAttribute}
							value={Array.isArray(node.value) ? node.value : undefined}
							onChange={(val) => onUpdate({ value: val })}
							placeholder="Enter values (comma-separated)"
						/>
					</InfoField>
				)}

				{requiresValue && requiresTwoValues && !requiresMultipleValues && selectedAttribute && (
					<InfoField
						label="Values"
						info={<Info text="The two values the operator needs, written as <value> elements." href={DOCS.operators} />}
					>
						<BetweenInputPicker
							attribute={selectedAttribute}
							value={
								Array.isArray(node.value) && node.value.length === 2
									? [node.value[0], node.value[1]]
									: [undefined, undefined]
							}
							onChange={(val) => onUpdate({ value: val })}
							placeholder1="From"
							placeholder2="To"
						/>
					</InfoField>
				)}

				{isSingleValue && !columnFilter && (operatorComparesAttributes || comparesToAttribute) && (
					<InfoField
						label="Compare to"
						validation={valueofFinding?.state === "error" ? valueofFinding : undefined}
						hint={
							!canCompareToAttribute && !comparesToAttribute && selectedAttribute
								? `Comparing a ${selectedAttribute.AttributeTypeName?.Value?.replace(/Type$/, "") ?? selectedAttribute.AttributeType} attribute to another attribute isn't verified to work, so it isn't offered.`
								: undefined
						}
						info={
							<Info
								text="A fixed value, or another attribute of the same record or of a link-entity. Comparing attributes works only with eq, ne, gt, ge, lt and le. The docs require the same type; FetchXML Studio also offers the type pairs verified to work (for example Money with Decimal)."
								href={DOCS.filterOnColumnValues}
							/>
						}
					>
						<RadioGroup
							layout="horizontal"
							value={comparesToAttribute ? "attribute" : "value"}
							onChange={(_, data) =>
								onUpdate(data.value === "attribute" ? { value: undefined, valueof: "" } : { valueof: undefined, value: undefined })
							}
						>
							<Radio value="value" label="A value" />
							<Radio value="attribute" label="Another attribute" disabled={!canCompareToAttribute && !comparesToAttribute} />
						</RadioGroup>
					</InfoField>
				)}

				{isSingleValue &&
					(comparesToAttribute ? (
						<ValueOfPicker
							fetchQuery={fetchQuery}
							rootEntityName={fetchQuery.entity.name}
							effectiveEntityName={effectiveEntityName}
							scope={valueofScope}
							sameAlias={isInRootFilter ? node.entityname : valueofScope?.selfAlias}
							valueof={node.valueof ?? ""}
							compatibleTypes={compatibleTypes}
							onChange={(valueof) => onUpdate({ valueof: valueof ?? "" })}
						/>
					) : (
						<InfoField
							label="Value"
							info={<Info text="The value to compare the attribute with." href={DOCS.filterRows} />}
						>
							{renderLiteralValue()}
						</InfoField>
					))}

				{/* Settings Dataverse ignores here (parsed XML or after an operator change) */}
				{!requiresValue && node.valueof !== undefined && (
					<LeftoverSetting
						text={`Compare to: ${node.valueof || "(empty)"}`}
						validation={{
							state: "warning",
							message: valueofFinding?.message ?? `The '${node.operator}' operator doesn't compare to a value.`,
						}}
						onRemove={() => onUpdate({ valueof: undefined })}
					/>
				)}
				{node.aggregate && (
					<LeftoverSetting
						text={`Aggregate: ${node.aggregate}`}
						validation={{
							state: "warning",
							message:
								findingFor(findings, "aggregate")?.message ??
								"Dataverse ignores aggregate on a condition. FetchXML can't filter on aggregate values.",
						}}
						onRemove={() => onUpdate({ aggregate: undefined })}
					/>
				)}
			</div>
		</div>
	);
}
