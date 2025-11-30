/**
 * ValueOf Picker Component
 * Allows selecting a column for valueof comparison, including cross-entity columns
 * Implements Option B: select entity/alias first, then attribute
 */

import { useState, useEffect, useMemo } from "react";
import {
	Field,
	Dropdown,
	Option,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { FetchNode } from "../../../../model/nodes";
import { collectLinkEntityReferences } from "../../../../model/treeUtils";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";

const useStyles = makeStyles({
	container: {
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

interface ValueOfPickerProps {
	fetchQuery: FetchNode | null;
	rootEntityName: string;
	effectiveEntityName: string;
	valueof: string;
	compatibleTypes: string[];
	onChange: (valueof: string | undefined) => void;
}

/**
 * Parse a valueof string into entity alias and attribute parts
 * - "createdon" -> { alias: undefined, attribute: "createdon" }
 * - "contact_alias.modifiedon" -> { alias: "contact_alias", attribute: "modifiedon" }
 */
function parseValueOf(valueof: string): { alias: string | undefined; attribute: string } {
	if (!valueof) {
		return { alias: undefined, attribute: "" };
	}
	const dotIndex = valueof.indexOf(".");
	if (dotIndex === -1) {
		return { alias: undefined, attribute: valueof };
	}
	return {
		alias: valueof.substring(0, dotIndex),
		attribute: valueof.substring(dotIndex + 1),
	};
}

/**
 * Build a valueof string from entity alias and attribute
 */
function buildValueOf(alias: string | undefined, attribute: string): string {
	if (!attribute) return "";
	if (!alias) return attribute;
	return `${alias}.${attribute}`;
}

export function ValueOfPicker({
	fetchQuery,
	rootEntityName,
	effectiveEntityName,
	valueof,
	compatibleTypes,
	onChange,
}: ValueOfPickerProps) {
	const styles = useStyles();

	// Collect all available entities (root + link-entities)
	const linkEntityReferences = useMemo(
		() => collectLinkEntityReferences(fetchQuery),
		[fetchQuery]
	);

	// Build entity options: root entity + all link-entities
	const entityOptions = useMemo(() => {
		const options: Array<{ value: string; label: string; entityName: string }> = [];

		// Add root entity (empty alias means root)
		if (rootEntityName) {
			options.push({
				value: "", // empty means root entity
				label: `${rootEntityName} (root)`,
				entityName: rootEntityName,
			});
		}

		// Add link-entities
		for (const ref of linkEntityReferences) {
			options.push({
				value: ref.identifier,
				label: ref.displayLabel,
				entityName: ref.entityName,
			});
		}

		return options;
	}, [rootEntityName, linkEntityReferences]);

	// Parse current valueof to get alias and attribute
	const parsed = useMemo(() => parseValueOf(valueof), [valueof]);

	// Track selected entity alias (empty string = root entity)
	const [selectedAlias, setSelectedAlias] = useState<string>(parsed.alias || "");

	// Get the entity name for the selected alias
	const selectedEntityName = useMemo(() => {
		if (selectedAlias === "") {
			return effectiveEntityName || rootEntityName;
		}
		const ref = linkEntityReferences.find((r) => r.identifier === selectedAlias);
		return ref?.entityName || "";
	}, [selectedAlias, effectiveEntityName, rootEntityName, linkEntityReferences]);

	// Sync selectedAlias when valueof changes externally (e.g., parsing FetchXML)
	useEffect(() => {
		setSelectedAlias(parsed.alias || "");
	}, [parsed.alias]);

	const handleEntityChange = (_ev: unknown, data: { optionValue?: string }) => {
		const newAlias = data.optionValue ?? "";
		setSelectedAlias(newAlias);
		// Clear the attribute when entity changes
		onChange(undefined);
	};

	const handleAttributeChange = (logicalName: string) => {
		const newValueof = buildValueOf(selectedAlias || undefined, logicalName);
		onChange(newValueof || undefined);
	};

	return (
		<div className={styles.container}>
			{/* Entity/Alias selector */}
			<Field label="Compare From Entity">
				<div className={styles.fieldWithTooltip}>
					<Dropdown
						value={
							entityOptions.find((o) => o.value === selectedAlias)?.label ||
							"Select entity..."
						}
						selectedOptions={[selectedAlias]}
						onOptionSelect={handleEntityChange}
						placeholder="Select entity..."
					>
						{entityOptions.map((opt) => (
							<Option key={opt.value} value={opt.value}>
								{opt.label}
							</Option>
						))}
					</Dropdown>
					<Tooltip
						content="Select which entity to pick the comparison column from. Use root entity for same-row comparison, or a linked entity for cross-table comparison."
						relationship="description"
					>
						<Info16Regular className={styles.tooltipIcon} />
					</Tooltip>
				</div>
			</Field>

			{/* Attribute picker for selected entity */}
			<Field label="Compare To Column">
				<div className={styles.fieldWithTooltip}>
					<AttributePicker
						entityLogicalName={selectedEntityName}
						value={parsed.attribute}
						onChange={handleAttributeChange}
						placeholder={
							selectedEntityName
								? "Select column to compare"
								: "Select an entity first"
						}
						disabled={!selectedEntityName}
						filterByTypes={compatibleTypes.length > 0 ? compatibleTypes : undefined}
					/>
					<Tooltip
						content={
							compatibleTypes.length > 0
								? `Select a column of compatible type (${compatibleTypes.join(", ")}) to compare against.`
								: "Select another column to compare against."
						}
						relationship="description"
					>
						<Info16Regular className={styles.tooltipIcon} />
					</Tooltip>
				</div>
			</Field>
		</div>
	);
}
