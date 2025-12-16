/**
 * Property editor for Filter nodes
 * Controls logical grouping (AND/OR) of conditions
 */

import {
	Field,
	Dropdown,
	Option,
	Input,
	Label,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { FilterNode } from "../../../../model/nodes";

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

interface FilterEditorProps {
	node: FilterNode;
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function FilterEditor({ node, onUpdate }: FilterEditorProps) {
	const styles = useStyles();

	const handleDropdownChange = (field: string) => (_: unknown, data: { optionValue?: string }) => {
		onUpdate({ [field]: data.optionValue });
	};

	const handleTextChange = (field: string) => (_: unknown, data: { value: string }) => {
		onUpdate({ [field]: data.value || undefined });
	};

	return (
		<div className={styles.container}>
			{/* Basic Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Filter Properties</Label>

				<Field label="Logical Operator" required>
					<div className={styles.fieldWithTooltip}>
						<Dropdown
							value={node.conjunction}
							selectedOptions={[node.conjunction]}
							onOptionSelect={handleDropdownChange("conjunction")}
						>
							<Option value="and">AND (all conditions must match)</Option>
							<Option value="or">OR (any condition can match)</Option>
						</Dropdown>
						<Tooltip
							content="AND: all child conditions/subfilters must be true. OR: at least one must be true."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
			</div>

			{/* Advanced Section */}
			<div className={styles.section}>
				<Label weight="semibold">Advanced</Label>

				<Field label="Query Hint (optional)" hint="Performance optimization hint. Rarely needed.">
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.hint ?? ""}
							onChange={handleTextChange("hint")}
							placeholder="e.g., union"
						/>
						<Tooltip
							content="Query optimizer hint. 'union' can improve performance for complex multi-table filters."
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
