/**
 * Property editor for Order nodes
 * Controls sorting/ordering of query results
 */

import {
	Field,
	Dropdown,
	Option,
	Radio,
	RadioGroup,
	Label,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { OrderNode, FetchNode } from "../../../../model/nodes";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { collectLinkEntityReferences } from "../../../../model/treeUtils";
import { useMemo } from "react";

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
	fetchQuery: FetchNode | null;
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function OrderEditor({ node, entityName, fetchQuery, onUpdate }: OrderEditorProps) {
	const styles = useStyles();

	// Collect all link-entity references for the entity dropdown
	const linkEntityOptions = useMemo(() => {
		const refs = collectLinkEntityReferences(fetchQuery);
		return refs;
	}, [fetchQuery]);

	// Determine which entity's attributes to show
	// If entityname is set, find the matching link-entity's logical name
	const attributeEntityName = useMemo(() => {
		if (!node.entityname) {
			// No entityname = root entity
			return entityName;
		}
		// Find the link-entity with matching alias/identifier
		const linkRef = linkEntityOptions.find(
			(ref) => ref.identifier === node.entityname || ref.alias === node.entityname
		);
		return linkRef?.entityName ?? entityName;
	}, [node.entityname, entityName, linkEntityOptions]);

	const handleRadioChange = (_: unknown, data: { value: string }) => {
		onUpdate({ descending: data.value === "desc" });
	};

	const handleEntityChange = (_: unknown, data: { optionValue?: string }) => {
		const value = data.optionValue;
		if (value === "__root__") {
			// Root entity selected - clear entityname and attribute
			onUpdate({ entityname: undefined, attribute: "" });
		} else {
			// Link-entity selected - set entityname and clear attribute
			onUpdate({ entityname: value || undefined, attribute: "" });
		}
	};

	// Get selected entity option value
	const selectedEntityValue = node.entityname ?? "__root__";

	return (
		<div className={styles.container}>
			{/* Basic Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Order By Properties</Label>

				<Field label="Entity" hint="Select root entity or a link-entity to sort by">
					<div className={styles.fieldWithTooltip}>
						<Dropdown
							value={
								selectedEntityValue === "__root__"
									? `Root: ${entityName}`
									: linkEntityOptions.find((o) => o.identifier === selectedEntityValue)
											?.displayLabel ?? selectedEntityValue
							}
							selectedOptions={[selectedEntityValue]}
							onOptionSelect={handleEntityChange}
						>
							<Option text={`Root: ${entityName}`} value="__root__">
								Root: {entityName}
							</Option>
							{linkEntityOptions.map((ref) => (
								<Option key={ref.identifier} text={ref.displayLabel} value={ref.identifier}>
									{ref.displayLabel}
								</Option>
							))}
						</Dropdown>
						<Tooltip
							content="Select the root entity or a link-entity to sort by its attributes."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Attribute Name" required>
					<div className={styles.fieldWithTooltip}>
						<AttributePicker
							entityLogicalName={attributeEntityName}
							value={node.attribute}
							onChange={(logicalName) => onUpdate({ attribute: logicalName })}
							placeholder="Select or type attribute name"
						/>
						<Tooltip content="Logical name of the attribute to sort by." relationship="description">
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
		</div>
	);
}
