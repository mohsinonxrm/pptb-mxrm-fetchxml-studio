/**
 * Property editor for Order nodes
 * Controls sorting/ordering of query results
 */

import {
	Field,
	Input,
	Radio,
	RadioGroup,
	Label,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { OrderNode } from "../../../../model/nodes";
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

interface OrderEditorProps {
	node: OrderNode;
	entityName: string;
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function OrderEditor({ node, entityName, onUpdate }: OrderEditorProps) {
	const styles = useStyles();

	const handleTextChange = (field: string) => (_: unknown, data: { value: string }) => {
		onUpdate({ [field]: data.value || undefined });
	};

	const handleRadioChange = (_: unknown, data: { value: string }) => {
		onUpdate({ descending: data.value === "desc" });
	};

	return (
		<div className={styles.container}>
			{/* Basic Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Order By Properties</Label>

				<Field label="Attribute Name" required>
					<div className={styles.fieldWithTooltip}>
						<AttributePicker
							entityLogicalName={entityName}
							value={node.attribute}
							onChange={(logicalName) => onUpdate({ attribute: logicalName })}
							placeholder="Select or type attribute name"
						/>
						<Tooltip
							content="Logical name of the attribute to sort by. Must exist in the parent entity or link-entity."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Sort Direction">
					<div className={styles.fieldWithTooltip}>
						<RadioGroup value={node.descending ? "desc" : "asc"} onChange={handleRadioChange}>
							<Radio value="asc" label="Ascending (A-Z, 0-9, oldest first)" />
							<Radio value="desc" label="Descending (Z-A, 9-0, newest first)" />
						</RadioGroup>
						<Tooltip
							content="Sort order: ascending (default) or descending. Multiple order nodes are applied in sequence."
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

				<Field
					label="Entity Name (optional)"
					hint="Only needed when sorting by a link-entity attribute. Use the link-entity alias."
				>
					<Input
						value={node.entityname ?? ""}
						onChange={handleTextChange("entityname")}
						placeholder="e.g., contact_link, parent_account"
					/>
				</Field>
			</div>
		</div>
	);
}
