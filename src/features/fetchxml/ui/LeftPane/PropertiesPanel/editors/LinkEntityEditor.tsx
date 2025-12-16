/**
 * Property editor for Link-Entity nodes
 * Controls joins/relationships to related entities
 */

import { useState, useEffect } from "react";
import {
	Field,
	Input,
	Dropdown,
	Option,
	Checkbox,
	Label,
	Tooltip,
	MessageBar,
	MessageBarBody,
	Dialog,
	DialogSurface,
	DialogTitle,
	DialogContent,
	DialogBody,
	DialogActions,
	Button,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import { debugLog } from "../../../../../../shared/utils/debug";
import type { LinkEntityNode, LinkType } from "../../../../model/nodes";
import { RelationshipPicker } from "../../../../../../shared/components/RelationshipPicker";

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

const LINK_TYPES: { value: LinkType; label: string; description: string }[] = [
	{
		value: "inner",
		label: "Inner Join",
		description: "Only records with matching related records",
	},
	{
		value: "outer",
		label: "Left Outer Join",
		description: "All parent records, even without related records",
	},
	{
		value: "any",
		label: "Any (EXISTS)",
		description: "At least one related record matches conditions",
	},
	{
		value: "not any",
		label: "Not Any (NOT EXISTS)",
		description: "No related records match conditions",
	},
	{ value: "all", label: "All", description: "All related records match conditions" },
	{ value: "not all", label: "Not All", description: "Not all related records match" },
	{ value: "exists", label: "Exists", description: "Semi-join existence check" },
	{ value: "in", label: "In", description: "Similar to exists" },
	{
		value: "matchfirstrowusingcrossapply",
		label: "Match First Row (CROSS APPLY)",
		description: "Performance: match only first related record",
	},
];

interface LinkEntityEditorProps {
	node: LinkEntityNode;
	parentEntityName: string;
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function LinkEntityEditor({ node, parentEntityName, onUpdate }: LinkEntityEditorProps) {
	const styles = useStyles();

	// Dialog state for confirming link-type change
	const [showConfirmDialog, setShowConfirmDialog] = useState(false);
	const [pendingLinkType, setPendingLinkType] = useState<LinkType | null>(null);

	// Determine if we should show the relationship picker
	// Show it if the node has a placeholder name or hasn't been properly configured
	const shouldShowPicker = node.name === "new_entity" || (!node.name && !node.from && !node.to);
	const [useRelationshipPicker, setUseRelationshipPicker] = useState(shouldShowPicker);

	// Reset picker state when switching to a different node or when node properties change
	useEffect(() => {
		debugLog("linkEntityEditor", "Node changed or props updated:", {
			nodeId: node.id,
			nodeName: node.name,
			from: node.from,
			to: node.to,
			parentEntityName,
			shouldShowPicker,
		});
		setUseRelationshipPicker(shouldShowPicker);
	}, [node.id, node.name, node.from, node.to, parentEntityName, shouldShowPicker]);

	const handleTextChange = (field: string) => (_: unknown, data: { value: string }) => {
		onUpdate({ [field]: data.value || undefined });
	};

	// Special handler for link-type changes that may require confirmation
	const handleLinkTypeChange = (_: unknown, data: { optionValue?: string }) => {
		const newLinkType = data.optionValue as LinkType | undefined;
		if (!newLinkType) return;

		// Check if switching to "not any" or "not all" and there are children that would need to be removed
		const isRestrictiveLinkType = newLinkType === "not any" || newLinkType === "not all";
		const hasChildren =
			node.attributes.length > 0 ||
			node.filters.length > 0 ||
			node.links.length > 0 ||
			node.allAttributes?.enabled;

		if (isRestrictiveLinkType && hasChildren) {
			// Show confirmation dialog
			setPendingLinkType(newLinkType);
			setShowConfirmDialog(true);
		} else {
			// Safe to change directly
			onUpdate({ linkType: newLinkType });
		}
	};

	// Handle dialog confirmation - proceed with change and remove children
	const handleConfirmLinkTypeChange = () => {
		if (pendingLinkType) {
			// Update link type and clear children
			onUpdate({
				linkType: pendingLinkType,
				attributes: [],
				allAttributes: undefined, // Remove all-attributes completely
				filters: [],
				links: [],
			});
		}
		setShowConfirmDialog(false);
		setPendingLinkType(null);
	};

	// Handle dialog cancellation - revert to current link type
	const handleCancelLinkTypeChange = () => {
		setShowConfirmDialog(false);
		setPendingLinkType(null);
		// No need to update anything - just close the dialog
	};

	const handleCheckboxChange =
		(field: string) => (_: unknown, data: { checked: boolean | "mixed" }) => {
			onUpdate({ [field]: data.checked === true ? true : undefined });
		};

	const handleRelationshipSelect = (relationship: {
		schemaName: string;
		referencedEntity: string;
		referencingEntity: string;
		referencedAttribute: string;
		referencingAttribute: string;
		type: "1N" | "N1" | "NN";
	}) => {
		debugLog("linkEntityEditor", "Relationship selected:", {
			nodeId: node.id,
			parentEntityName,
			relationship,
		});

		// Auto-populate from/to based on relationship type and direction
		const updates: Record<string, unknown> = {
			name: "",
			from: "",
			to: "",
			relationshipType: relationship.type,
		};

		// Determine which entity is the related/target entity
		// For 1:N - parent is ReferencedEntity, child is ReferencingEntity
		// For N:1 - parent is ReferencedEntity, child is ReferencingEntity
		// For N:N - both are referenced in the relationship

		if (relationship.type === "1N") {
			// Parent (current) → Child (related)
			// from = child's FK attribute, to = parent's PK attribute
			if (relationship.referencedEntity === parentEntityName) {
				// We're the referenced (parent), linking to referencing (child)
				updates.name = relationship.referencingEntity;
				updates.from = relationship.referencingAttribute;
				updates.to = relationship.referencedAttribute;
			} else {
				// Reverse direction
				updates.name = relationship.referencedEntity;
				updates.from = relationship.referencedAttribute;
				updates.to = relationship.referencingAttribute;
			}
		} else if (relationship.type === "N1") {
			// Child (current) → Parent (related)
			// from = parent's PK attribute, to = child's FK attribute
			if (relationship.referencingEntity === parentEntityName) {
				// We're the referencing (child), linking to referenced (parent)
				updates.name = relationship.referencedEntity;
				updates.from = relationship.referencedAttribute;
				updates.to = relationship.referencingAttribute;
			} else {
				// Reverse direction
				updates.name = relationship.referencingEntity;
				updates.from = relationship.referencingAttribute;
				updates.to = relationship.referencedAttribute;
			}
		} else if (relationship.type === "NN") {
			// Many-to-Many
			// Determine target entity (the one that's not parentEntityName)
			const targetEntity =
				relationship.referencedEntity === parentEntityName
					? relationship.referencingEntity
					: relationship.referencedEntity;

			updates.name = targetEntity;
			updates.from = relationship.referencingAttribute;
			updates.to = relationship.referencedAttribute;
			updates.intersect = true;
		}

		debugLog("linkEntityEditor", "Applying updates:", updates);
		onUpdate(updates);
		setUseRelationshipPicker(false); // Switch to manual mode after selection
	};

	return (
		<div className={styles.container}>
			{/* Relationship Picker Section */}
			{useRelationshipPicker && parentEntityName && (
				<div className={styles.section}>
					<Label weight="semibold">Select Relationship</Label>
					<Field
						label="Choose a relationship to auto-populate link details"
						hint="Or skip this and manually configure the link below"
					>
						<RelationshipPicker
							entityLogicalName={parentEntityName}
							onChange={handleRelationshipSelect}
							placeholder="Select a relationship..."
						/>
					</Field>
				</div>
			)}

			{/* Basic Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Link-Entity Properties</Label>

				<Field label="Related Entity Name" required>
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.name}
							onChange={handleTextChange("name")}
							placeholder="e.g., contact, account, systemuser"
						/>
						<Tooltip
							content="Logical name of the related entity to join."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="From Attribute" required>
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.from}
							onChange={handleTextChange("from")}
							placeholder="e.g., contactid, accountid"
						/>
						<Tooltip
							content="Attribute on the related entity (the entity being joined)."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="To Attribute" required>
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.to}
							onChange={handleTextChange("to")}
							placeholder="e.g., parentcustomerid, createdby"
						/>
						<Tooltip
							content="Attribute on the parent entity (the entity being joined from)."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Join Type" required>
					<div className={styles.fieldWithTooltip}>
						<Dropdown
							value={node.linkType}
							selectedOptions={[node.linkType]}
							onOptionSelect={handleLinkTypeChange}
						>
							{LINK_TYPES.map((type) => (
								<Option key={type.value} value={type.value} text={type.description}>
									{type.label}
								</Option>
							))}
						</Dropdown>
						<Tooltip
							content="Type of join. Inner = matching only. Outer = include parent records without match. EXISTS variations for existence checks."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				{/* Warning for "not any" and "not all" link types */}
				{(node.linkType === "not any" || node.linkType === "not all") && (
					<MessageBar intent="warning">
						<MessageBarBody>
							<strong>Note:</strong> The "{node.linkType}" link type does not support attributes.
							Attributes added to this link-entity will cause query execution errors.
						</MessageBarBody>
					</MessageBar>
				)}
			</div>

			{/* Optional Properties */}
			<div className={styles.section}>
				<Label weight="semibold">Optional Properties</Label>

				<Field label="Alias">
					<div className={styles.fieldWithTooltip}>
						<Input
							value={node.alias ?? ""}
							onChange={handleTextChange("alias")}
							placeholder="e.g., parent_account, created_by_user"
						/>
						<Tooltip
							content="Alias for referencing this link in conditions, order-by, or nested links."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={node.intersect ?? false}
							onChange={handleCheckboxChange("intersect")}
							label="Intersect (Many-to-Many)"
						/>
						<Tooltip
							content="Use intersect entity for N:N relationships. Dataverse automatically handles the intersect table."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={node.visible ?? true}
							onChange={handleCheckboxChange("visible")}
							label="Visible in Results"
						/>
						<Tooltip
							content="Include linked entity data in query results. Set to false to use link only for filtering."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
			</div>

			{/* Confirmation Dialog for Link Type Change */}
			<Dialog
				open={showConfirmDialog}
				onOpenChange={(_, data) => !data.open && handleCancelLinkTypeChange()}
			>
				<DialogSurface>
					<DialogBody>
						<DialogTitle>Confirm Link Type Change</DialogTitle>
						<DialogContent>
							<p>
								Changing the link type to <strong>"{pendingLinkType}"</strong> will remove all
								attributes, filters, and nested link-entities from this link-entity because these
								link types do not support child elements.
							</p>
							<p>This action cannot be undone. Do you want to proceed?</p>
						</DialogContent>
						<DialogActions>
							<Button appearance="primary" onClick={handleConfirmLinkTypeChange}>
								Yes, Change Link Type
							</Button>
							<Button appearance="secondary" onClick={handleCancelLinkTypeChange}>
								Cancel
							</Button>
						</DialogActions>
					</DialogBody>
				</DialogSurface>
			</Dialog>
		</div>
	);
}
