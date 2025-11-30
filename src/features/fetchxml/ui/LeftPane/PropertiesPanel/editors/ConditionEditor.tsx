/**
 * Property editor for Condition nodes
 * Controls filter criteria: attribute, operator, value
 */

import { useState, useEffect, useMemo } from "react";
import {
	Field,
	Input,
	Dropdown,
	Option,
	Label,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { ConditionNode, FetchNode } from "../../../../model/nodes";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { OptionSetValuePicker } from "../../../../../../shared/components/OptionSetValuePicker";
import { BooleanValuePicker } from "../../../../../../shared/components/BooleanValuePicker";
import { NumericInputPicker } from "../../../../../../shared/components/NumericInputPicker";
import { DateTimeInputPicker } from "../../../../../../shared/components/DateTimeInputPicker";
import { BetweenInputPicker } from "../../../../../../shared/components/BetweenInputPicker";
import { MultiValuePicker } from "../../../../../../shared/components/MultiValuePicker";
import {
	getOperatorsForAttributeType,
	operatorRequiresValue,
	type OperatorDefinition,
} from "../../../../model/operators";
import {
	loadEntityAttributes,
	loadAttributeDetailedMetadata,
} from "../../../../api/dataverseMetadata";
import type { AttributeMetadata } from "../../../../api/pptbClient";
import {
	collectLinkEntityReferences,
	isConditionInRootEntityFilter,
} from "../../../../model/treeUtils";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
		padding: tokens.spacingVerticalM,
	},
	section: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
	},
	fieldWithTooltip: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
	},
	tooltipIcon: {
		color: tokens.colorNeutralForeground3,
		cursor: "help",
	},
});

interface ConditionEditorProps {
	node: ConditionNode;
	entityName: string;
	onUpdate: (updates: Record<string, unknown>) => void;
	isAggregateQuery?: boolean;
	fetchQuery?: FetchNode | null;
}

export function ConditionEditor({
	node,
	entityName,
	onUpdate,
	isAggregateQuery = false,
	fetchQuery = null,
}: ConditionEditorProps) {
	const styles = useStyles();

	// State for filtered operators based on selected attribute type
	const [availableOperators, setAvailableOperators] = useState<OperatorDefinition[]>([]);
	// State for selected attribute metadata
	const [selectedAttribute, setSelectedAttribute] = useState<AttributeMetadata | null>(null);

	// Collect available link-entity references for entityname picker
	const linkEntityReferences = useMemo(
		() => collectLinkEntityReferences(fetchQuery),
		[fetchQuery]
	);

	// Check if this condition is in the root entity's filter (where entityname is applicable)
	const isInRootFilter = useMemo(
		() => isConditionInRootEntityFilter(fetchQuery, node.id),
		[fetchQuery, node.id]
	);

	// Determine the effective entity name for attribute loading
	// If entityname is set, use the linked entity's name; otherwise use the parent entity
	const effectiveEntityName = useMemo(() => {
		if (node.entityname && isInRootFilter) {
			// Find the link-entity reference and get its entity name
			const ref = linkEntityReferences.find(
				(r) => r.identifier === node.entityname
			);
			if (ref) {
				return ref.entityName;
			}
		}
		return entityName;
	}, [node.entityname, isInRootFilter, linkEntityReferences, entityName]);

	// Load attribute metadata when attribute changes to get its type
	useEffect(() => {
		if (!node.attribute || !effectiveEntityName) {
			// No attribute selected - show all operators
			setAvailableOperators(getOperatorsForAttributeType(undefined));
			setSelectedAttribute(null);
			return;
		}

		// Load entity attributes to find the selected attribute's type
		loadEntityAttributes(effectiveEntityName)
			.then(async (attributes: AttributeMetadata[]) => {
				const attr = attributes.find((a) => a.LogicalName === node.attribute);
				const type = attr?.AttributeType;

				// For numeric and datetime types, load detailed metadata
				const needsDetailedMetadata =
					type === "Integer" ||
					type === "BigInt" ||
					type === "Decimal" ||
					type === "Double" ||
					type === "DateTime";

				let detailedAttr = attr;
				if (needsDetailedMetadata && attr) {
					try {
						// Load detailed metadata with type casting to get MinValue, MaxValue, Precision, Format
						detailedAttr = await loadAttributeDetailedMetadata(entityName, node.attribute, type);
						console.log(`Loaded detailed metadata for ${node.attribute}:`, detailedAttr);
					} catch (error) {
						console.error(`Failed to load detailed metadata for ${node.attribute}:`, error);
						// Fall back to basic metadata
						detailedAttr = attr;
					}
				}

				// Store the full attribute metadata (with detailed properties if available)
				setSelectedAttribute(detailedAttr || null);

				// Filter operators based on attribute type
				let operators = getOperatorsForAttributeType(type);

				// For DateTime with DateOnly behavior, exclude operators that require hours/minutes
				if (type === "DateTime" && detailedAttr) {
					const dateTimeBehavior = detailedAttr.DateTimeBehavior?.Value;
					const format = detailedAttr.Format;

					// DateOnly format or UserLocal behavior with DateOnly format should exclude hour/minute operators
					const isDateOnly = format === "DateOnly" || dateTimeBehavior === "DateOnly";

					if (isDateOnly) {
						const excludedOperators = [
							"last-x-hours",
							"next-x-hours",
							"olderthan-x-hours",
							"olderthan-x-minutes",
						];
						operators = operators.filter((op) => !excludedOperators.includes(op.value));
					}
				}

				setAvailableOperators(operators);
			})
			.catch((error) => {
				console.error("Failed to load attribute metadata for operator filtering:", error);
				// Fallback to all operators on error
				setAvailableOperators(getOperatorsForAttributeType(undefined));
				setSelectedAttribute(null);
			});
	}, [node.attribute, effectiveEntityName]);

	const handleTextChange = (field: string) => (_: unknown, data: { value: string }) => {
		onUpdate({ [field]: data.value || undefined });
	};

	const handleDropdownChange = (field: string) => (_: unknown, data: { optionValue?: string }) => {
		onUpdate({ [field]: data.optionValue });
	};

	const handleAttributeChange = (logicalName: string | undefined) => {
		// When attribute is cleared or changed, reset operator and value
		if (!logicalName) {
			onUpdate({ attribute: undefined, operator: undefined, value: undefined });
		} else {
			onUpdate({ attribute: logicalName, operator: undefined, value: undefined });
		}
	};

	const handleEntityNameChange = (identifier: string | undefined) => {
		// When entityname changes, also clear attribute, operator, and value
		// since the attribute must come from the selected entity
		onUpdate({
			entityname: identifier || undefined,
			attribute: undefined,
			operator: undefined,
			value: undefined,
		});
	};
	// Check if the current operator requires a value (using the smart function)
	const requiresValue = operatorRequiresValue(node.operator);

	// Get the operator definition to check for multi-value requirements
	const operatorDef = availableOperators.find((op) => op.value === node.operator);
	const requiresTwoValues = operatorDef?.requiresTwoValues ?? false;
	const requiresMultipleValues = operatorDef?.requiresMultipleValues ?? false;

	return (
		<div className={styles.container}>
			{/* Basic Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Condition Properties</Label>
				<Field label="Attribute Name" required>
					<div className={styles.fieldWithTooltip}>
						<AttributePicker
							entityLogicalName={effectiveEntityName}
							value={node.attribute}
							onChange={handleAttributeChange}
							placeholder="Select or type attribute name"
						/>
						<Tooltip
							content={
								node.entityname
									? `Logical name of the attribute from the linked entity "${node.entityname}".`
									: "Logical name of the attribute to filter on. Must exist in the parent entity."
							}
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
				<Field label="Operator" required>
					<div className={styles.fieldWithTooltip}>
						<Dropdown
							value={node.operator}
							selectedOptions={[node.operator]}
							onOptionSelect={handleDropdownChange("operator")}
							placeholder="Select operator"
						>
							{availableOperators.map((op) => (
								<Option key={op.value} value={op.value}>
									{op.label}
								</Option>
							))}
						</Dropdown>
						<Tooltip
							content="Comparison operator. Available operators depend on attribute type (string, number, datetime, lookup, etc.)."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
				{requiresValue && (
					<Field label="Value">
						<div className={styles.fieldWithTooltip}>
							{/* Multi-value operators (in, not-in, contain-values) */}
							{requiresMultipleValues && selectedAttribute ? (
								<MultiValuePicker
									entityLogicalName={effectiveEntityName}
									attributeLogicalName={node.attribute}
									attribute={selectedAttribute}
									value={Array.isArray(node.value) ? node.value : undefined}
									onChange={(val) => onUpdate({ value: val })}
									placeholder="Enter values (comma-separated)"
								/>
							) : /* Dual-value operators (between, not-between, fiscal-period-and-year) */
							requiresTwoValues && selectedAttribute ? (
								<BetweenInputPicker
									attribute={selectedAttribute}
									value={
										Array.isArray(node.value) && node.value.length === 2
											? [node.value[0], node.value[1]]
											: [undefined, undefined]
									}
									onChange={(val) => onUpdate({ value: val })}
									placeholder1="Start value"
									placeholder2="End value"
								/>
							) : /* Single-value operators - type-specific pickers */
							selectedAttribute?.AttributeType === "Picklist" ||
							  selectedAttribute?.AttributeType === "State" ||
							  selectedAttribute?.AttributeType === "Status" ? (
								<OptionSetValuePicker
									entityLogicalName={effectiveEntityName}
									attributeLogicalName={node.attribute}
									value={typeof node.value === "number" ? node.value : undefined}
									onChange={(val) => onUpdate({ value: val })}
									placeholder="Select a value"
								/>
							) : selectedAttribute?.AttributeType === "Boolean" ? (
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
							) : selectedAttribute?.AttributeType === "Integer" ||
							  selectedAttribute?.AttributeType === "BigInt" ||
							  selectedAttribute?.AttributeType === "Decimal" ||
							  selectedAttribute?.AttributeType === "Double" ? (
								<NumericInputPicker
									attribute={selectedAttribute}
									value={typeof node.value === "number" ? node.value : undefined}
									onChange={(val) => onUpdate({ value: val })}
									placeholder="Enter a number"
								/>
							) : selectedAttribute?.AttributeType === "DateTime" ? (
								<DateTimeInputPicker
									attribute={selectedAttribute}
									value={
										typeof node.value === "string"
											? node.value
											: node.value instanceof Date
											? node.value
											: undefined
									}
									onChange={(val) => onUpdate({ value: val?.toISOString() })}
									placeholder="Select a date"
								/>
							) : (
								<Input
									value={
										typeof node.value === "string" || typeof node.value === "number"
											? String(node.value)
											: ""
									}
									onChange={handleTextChange("value")}
									placeholder="Filter value"
								/>
							)}
							<Tooltip
								content={
									selectedAttribute?.AttributeType === "Picklist" ||
									selectedAttribute?.AttributeType === "State" ||
									selectedAttribute?.AttributeType === "Status"
										? "Select a value from the available options."
										: selectedAttribute?.AttributeType === "Boolean"
										? "Select true or false."
										: selectedAttribute?.AttributeType === "Integer" ||
										  selectedAttribute?.AttributeType === "BigInt" ||
										  selectedAttribute?.AttributeType === "Decimal" ||
										  selectedAttribute?.AttributeType === "Double"
										? "Enter a numeric value. Min/max constraints from metadata apply."
										: selectedAttribute?.AttributeType === "DateTime"
										? "Select a date and optionally a time."
										: "Value to compare against. For 'in' operator, use comma-separated values. For 'between', use two values separated by 'and'."
								}
								relationship="description"
							>
								<Info16Regular className={styles.tooltipIcon} />
							</Tooltip>
						</div>
					</Field>
				)}
			</div>

			{/* Advanced Section */}
			<div className={styles.section}>
				<Label weight="semibold">Advanced</Label>

				{/* Entity Name - only show for conditions in root entity filters where there are link-entities */}
				{isInRootFilter && linkEntityReferences.length > 0 && (
					<Field
						label="Entity Name (optional)"
						hint="Filter on a column from a linked entity (outer join scenario)."
					>
						<div className={styles.fieldWithTooltip}>
							<Dropdown
								value={node.entityname ?? ""}
								selectedOptions={node.entityname ? [node.entityname] : []}
								onOptionSelect={(_ev, data) => {
									const value = data.optionValue;
									handleEntityNameChange(value === "" ? undefined : value);
								}}
								placeholder="Select linked entity..."
							>
								<Option value="">None (current entity)</Option>
								{linkEntityReferences.map((ref) => (
									<Option key={ref.identifier} value={ref.identifier}>
										{ref.displayLabel}
									</Option>
								))}
							</Dropdown>
							<Tooltip
								content="When filtering on an outer-joined link-entity, specify which entity's column to filter. Use the alias if defined, otherwise the entity name."
								relationship="description"
							>
								<Info16Regular className={styles.tooltipIcon} />
							</Tooltip>
						</div>
					</Field>
				)}

				{isAggregateQuery && (
					<Field label="Aggregate Function (optional)">
						<div className={styles.fieldWithTooltip}>
							<Dropdown
								value={node.aggregate ?? "none"}
								selectedOptions={[node.aggregate ?? "none"]}
								onOptionSelect={handleDropdownChange("aggregate")}
								placeholder="None"
							>
								<Option value="none">None</Option>
								<Option value="count">Count</Option>
								<Option value="countcolumn">Count Column</Option>
								<Option value="sum">Sum</Option>
								<Option value="avg">Average</Option>
								<Option value="min">Minimum</Option>
								<Option value="max">Maximum</Option>
							</Dropdown>
							<Tooltip
								content="Apply aggregate function before comparing. Used in HAVING clauses for aggregate queries."
								relationship="description"
							>
								<Info16Regular className={styles.tooltipIcon} />
							</Tooltip>
						</div>
					</Field>
				)}

				<Field
					label="Value Of (optional)"
					hint="Compare to another attribute instead of a literal value."
				>
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.valueof ?? ""}
							onChange={handleTextChange("valueof")}
							placeholder="e.g., modifiedon"
						/>
						<Tooltip
							content="Compare this attribute to another attribute's value (cross-column comparison)."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
			</div>
		</div>
	);
}
