/**
 * Property editor for Attribute nodes
 * Controls column selection, aliases, aggregation, and grouping
 */

import {
	Field,
	Input,
	Dropdown,
	Option,
	Checkbox,
	Label,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { AttributeNode } from "../../../../model/nodes";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";

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

interface AttributeEditorProps {
	node: AttributeNode;
	entityName: string;
	onUpdate: (updates: Record<string, unknown>) => void;
	isAggregateQuery?: boolean; // Whether parent fetch has aggregate=true
}

export function AttributeEditor({
	node,
	entityName,
	onUpdate,
	isAggregateQuery = false,
}: AttributeEditorProps) {
	const styles = useStyles();

	const handleTextChange = (field: string) => (_: unknown, data: { value: string }) => {
		let value = data.value;
		// For alias field, remove spaces (aliases cannot contain spaces)
		if (field === "alias") {
			value = value.replace(/\s/g, "");
		}
		onUpdate({ [field]: value || undefined });
	};

	const handleCheckboxChange =
		(field: string) => (_: unknown, data: { checked: boolean | "mixed" }) => {
			onUpdate({ [field]: data.checked === true ? true : undefined });
		};

	const handleDropdownChange = (field: string) => (_: unknown, data: { optionValue?: string }) => {
		onUpdate({ [field]: data.optionValue || undefined });
	};

	return (
		<div className={styles.container}>
			{/* Basic Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Attribute Properties</Label>

				<Field label="Attribute Name" required>
					<div className={styles.fieldWithTooltip}>
						<AttributePicker
							entityLogicalName={entityName}
							value={node.name}
							onChange={(logicalName) => onUpdate({ name: logicalName })}
							placeholder="Select or type attribute name"
						/>
						<Tooltip
							content="Logical name of the attribute/column to retrieve from the entity."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Alias (optional)">
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.alias ?? ""}
							onChange={handleTextChange("alias")}
							placeholder="e.g., account_name, total_value"
						/>
						<Tooltip
							content="Optional alias for the column in results. Useful for aggregate queries or when joining multiple entities."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
			</div>

			{/* Aggregation Section (only show if parent fetch has aggregate=true) */}
			{isAggregateQuery && (
				<div className={styles.section}>
					<Label weight="semibold">Aggregation</Label>

					<Field label="Aggregate Function">
						<div className={styles.fieldWithTooltip}>
							<Dropdown
								value={node.aggregate ?? "none"}
								selectedOptions={[node.aggregate ?? "none"]}
								onOptionSelect={handleDropdownChange("aggregate")}
								placeholder="None"
							>
								<Option value="none">None</Option>
								<Option value="count">Count</Option>
								<Option value="countcolumn">Count Column (non-null)</Option>
								<Option value="sum">Sum</Option>
								<Option value="avg">Average</Option>
								<Option value="min">Minimum</Option>
								<Option value="max">Maximum</Option>
								<Option value="rowaggregate">Row Aggregate</Option>
							</Dropdown>
							<Tooltip
								content="Aggregate function to apply. SUM/AVG only work with numeric types. COUNT works with all types."
								relationship="description"
							>
								<Info16Regular className={styles.tooltipIcon} />
							</Tooltip>
						</div>
					</Field>

					<Field>
						<div className={styles.fieldWithTooltip}>
							<Checkbox
								checked={node.groupby ?? false}
								onChange={handleCheckboxChange("groupby")}
								label="Group By"
							/>
							<Tooltip
								content="Include this attribute in GROUP BY clause. Required for non-aggregated columns in aggregate queries."
								relationship="description"
							>
								<Info16Regular className={styles.tooltipIcon} />
							</Tooltip>
						</div>
					</Field>
				</div>
			)}

			{/* DateTime Grouping Section */}
			<div className={styles.section}>
				<Label weight="semibold">DateTime Options</Label>

				<Field label="Date Grouping (for datetime attributes)">
					<div className={styles.fieldWithTooltip}>
						<Dropdown
							value={node.dategrouping ?? "none"}
							selectedOptions={[node.dategrouping ?? "none"]}
							onOptionSelect={handleDropdownChange("dategrouping")}
							placeholder="None"
						>
							<Option value="none">None</Option>
							<Option value="day">Day</Option>
							<Option value="week">Week</Option>
							<Option value="month">Month</Option>
							<Option value="quarter">Quarter</Option>
							<Option value="year">Year</Option>
							<Option value="fiscal-period">Fiscal Period</Option>
							<Option value="fiscal-year">Fiscal Year</Option>
						</Dropdown>
						<Tooltip
							content="Group datetime values by time period. Only applicable to datetime attributes."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={node.usertimezone ?? false}
							onChange={handleCheckboxChange("usertimezone")}
							label="Convert to User's Timezone"
						/>
						<Tooltip
							content="Convert datetime values from UTC to the user's timezone. Only for datetime attributes."
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
