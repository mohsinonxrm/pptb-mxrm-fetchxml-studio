/**
 * Property editor for <link-entity>: name, from / to, link-type, alias and intersect (#90b, #86).
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/link-entity
 *
 * Verified live (2026-10-14): a link-entity inside a filter must be any / not any / all / not all;
 * joined link-entities accept all nine types; exists, in and the four filter types return nothing,
 * so attribute, all-attributes and order fail there (0x80041121). Filters and nested link-entities
 * are fine with every type.
 */

import { useState, useEffect } from "react";
import {
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	Dropdown,
	Input,
	Option,
	OptionGroup,
} from "@fluentui/react-components";
import { debugLog } from "../../../../../../shared/utils/debug";
import type { LinkEntityNode, LinkType } from "../../../../model/nodes";
import type { RuleFinding } from "../../../../model/fetchxmlRules";
import { NON_RETURNING_LINK_TYPES } from "../../../../model/fetchxmlRules";
import { sanitizeAlias } from "../../../../model/aliasUtils";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { RelationshipPicker, type RelationshipOption } from "../../../../../../shared/components/RelationshipPicker";
import { planManyToManyJoin } from "../../../../model/manyToMany";
import { loadEntityMetadata } from "../../../../api/dataverseMetadata";
import { useBuilder } from "../../../../state/builderContext";
import { collectLinkEntityAliases, generateUniqueAlias } from "../../../../model/aliasUtils";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoCheckbox, InfoField, LeftoverSetting, SectionHeading, type Validation } from "../fields";
import { findingFor } from "../ruleFindings";

/** Descriptions follow the link-entity reference and filter-rows (including the inverted all / not all). */
const LINK_TYPES: Record<LinkType, { label: string; description: string }> = {
	inner: { label: "Inner", description: "Default. Only records with a matching related record." },
	outer: { label: "Outer", description: "Also records without a matching related record." },
	exists: { label: "Exists", description: "Like inner, using EXISTS: each record once, no related attributes." },
	in: { label: "In", description: "Like inner, using IN: each record once, no related attributes." },
	matchfirstrowusingcrossapply: {
		label: "Match first row",
		description: "Like inner, matching only one related record: each record once.",
	},
	any: { label: "Any", description: "Records with at least one related record matching this link-entity's filter." },
	"not any": { label: "Not any", description: "Records with no related record matching this link-entity's filter." },
	all: {
		label: "All",
		description: "Records with related records, none of which match the filter. Invert the filter to require every related record to match.",
	},
	"not all": { label: "Not all", description: "The same as Any, despite the name." },
	// Kept when parsed, not offered: not in the reference, but Dataverse accepts it
	natural: { label: "Natural", description: "Not in the FetchXML reference; Dataverse accepts it." },
};
const JOIN_TYPES: LinkType[] = ["inner", "outer", "exists", "in", "matchfirstrowusingcrossapply"];
const RELATED_TYPES: LinkType[] = ["any", "not any", "all", "not all"];

interface LinkEntityEditorProps {
	node: LinkEntityNode;
	/** The entity / link-entity this link-entity joins from */
	parentEntityName: string;
	/** The link-entity is inside a filter (any / not any / all / not all only) */
	inFilter: boolean;
	/** Query-rule findings for this node (fetchxmlRules.ts) */
	findings: readonly RuleFinding[];
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function LinkEntityEditor({ node, parentEntityName, inFilter, findings, onUpdate }: LinkEntityEditorProps) {
	const styles = useEditorStyles();
	const { addLinkEntityWithConfig, fetchQuery } = useBuilder();
	const [pendingLinkType, setPendingLinkType] = useState<LinkType | null>(null);
	const [relationshipError, setRelationshipError] = useState<string | null>(null);

	// Offer the relationship picker until the link-entity has been set up
	const shouldShowPicker = node.name === "new_entity" || (!node.name && !node.from && !node.to);
	const [useRelationshipPicker, setUseRelationshipPicker] = useState(shouldShowPicker);
	useEffect(() => setUseRelationshipPicker(shouldShowPicker), [node.id, shouldShowPicker]);

	const returned = [
		node.attributes.length > 0 && `${node.attributes.length} attribute${node.attributes.length === 1 ? "" : "s"}`,
		node.allAttributes?.enabled && "all attributes",
		node.orders.length > 0 && `${node.orders.length} sort${node.orders.length === 1 ? "" : "s"}`,
	].filter((x): x is string => !!x);
	const returnedCount = node.attributes.length + (node.allAttributes?.enabled ? 1 : 0) + node.orders.length;

	const handleLinkTypeChange = (linkType: LinkType) => {
		if (NON_RETURNING_LINK_TYPES.has(linkType) && returned.length > 0) {
			setPendingLinkType(linkType);
		} else {
			onUpdate({ linkType });
		}
	};

	const confirmLinkTypeChange = (removeReturned: boolean) => {
		if (pendingLinkType) {
			onUpdate(
				removeReturned
					? { linkType: pendingLinkType, attributes: [], allAttributes: undefined, orders: [] }
					: { linkType: pendingLinkType },
			);
		}
		setPendingLinkType(null);
	};

	/**
	 * N:N (#85): join through the intersect entity with two nested link-entities, as documented
	 * (join-tables#many-to-many-relationships). This link-entity becomes the intersect entity and the
	 * related entity is added inside it, with the alias, so its attributes go there.
	 */
	const applyManyToMany = async (relationship: RelationshipOption) => {
		setRelationshipError(null);
		try {
			if (!relationship.intersectEntity) throw new Error("the relationship has no intersect entity.");
			const target =
				relationship.referencingEntity === parentEntityName ? relationship.referencedEntity : relationship.referencingEntity;
			const [parentMetadata, targetMetadata] = await Promise.all([
				loadEntityMetadata(parentEntityName),
				loadEntityMetadata(target),
			]);
			const join = planManyToManyJoin(
				parentEntityName,
				parentMetadata.PrimaryIdAttribute,
				{
					Entity1LogicalName: relationship.referencingEntity,
					Entity2LogicalName: relationship.referencedEntity,
					Entity1IntersectAttribute: relationship.referencingAttribute,
					Entity2IntersectAttribute: relationship.referencedAttribute,
					IntersectEntityName: relationship.intersectEntity,
				},
				() => targetMetadata.PrimaryIdAttribute,
			);
			if (!join) throw new Error(`${parentEntityName} isn't part of ${relationship.schemaName}.`);
			onUpdate({
				name: join.intersect.name,
				from: join.intersect.from,
				to: join.intersect.to,
				intersect: true,
				alias: undefined,
				relationshipType: "NN",
			});
			addLinkEntityWithConfig(node.id, join.target.name, join.target.from, join.target.to, "inner", "NN");
			setUseRelationshipPicker(false);
		} catch (error) {
			setRelationshipError(
				`Couldn't build the many-to-many join: ${error instanceof Error ? error.message : String(error)}`,
			);
		}
	};

	const handleRelationshipSelect = (relationship: RelationshipOption) => {
		if (relationship.type === "NN") {
			void applyManyToMany(relationship);
			return;
		}
		debugLog("linkEntityEditor", "Relationship selected:", { nodeId: node.id, parentEntityName, relationship });
		const updates: Record<string, unknown> = { name: "", from: "", to: "", relationshipType: relationship.type };

		// from = the related entity's attribute, to = the parent's attribute
		if (relationship.type === "1N") {
			if (relationship.referencedEntity === parentEntityName) {
				updates.name = relationship.referencingEntity;
				updates.from = relationship.referencingAttribute;
				updates.to = relationship.referencedAttribute;
			} else {
				updates.name = relationship.referencedEntity;
				updates.from = relationship.referencedAttribute;
				updates.to = relationship.referencingAttribute;
			}
		} else if (relationship.type === "N1") {
			if (relationship.referencingEntity === parentEntityName) {
				updates.name = relationship.referencedEntity;
				updates.from = relationship.referencedAttribute;
				updates.to = relationship.referencingAttribute;
			} else {
				updates.name = relationship.referencingEntity;
				updates.from = relationship.referencingAttribute;
				updates.to = relationship.referencedAttribute;
			}
		}
		// A new link-entity's placeholder alias ("new_entity…") becomes one for the related entity
		if (node.alias?.startsWith("new_entity") && typeof updates.name === "string" && updates.name && fetchQuery) {
			const otherAliases = collectLinkEntityAliases(fetchQuery);
			otherAliases.delete(node.alias);
			updates.alias = generateUniqueAlias(updates.name, otherAliases);
		}
		onUpdate(updates);
		setUseRelationshipPicker(false);
	};

	const offered = inFilter ? RELATED_TYPES : [...JOIN_TYPES, ...RELATED_TYPES];
	const linkTypeOption = (type: LinkType, secondary = LINK_TYPES[type]?.description) => (
		<Option key={type} value={type} text={LINK_TYPES[type]?.label ?? type}>
			<span className={styles.optionText}>
				{LINK_TYPES[type]?.label ?? type}
				<span className={styles.optionSecondary}>{secondary}</span>
			</span>
		</Option>
	);

	const fromToAdvice: Validation =
		!node.from && !node.to
			? {
					state: "warning",
					message: "Neither attribute is set: Dataverse uses the system many-to-many relationship between the entities, and fails (0x80041102) if there isn't one.",
				}
			: !node.from || !node.to
				? {
						state: "warning",
						message: "Set both attributes. With only one, Dataverse tries to work out the other from the relationship definitions.",
					}
				: undefined;

	return (
		<div className={styles.container}>
			{useRelationshipPicker && parentEntityName && (
				<div className={styles.section}>
					<SectionHeading>Relationship</SectionHeading>
					<InfoField
						label="Relationship"
						hint="Fills in the entity and the attributes to match. A many-to-many relationship adds its intersect entity with the related entity inside it. Or set them below."
						validation={relationshipError ? { state: "error", message: relationshipError } : undefined}
						info={<Info text={`The relationships of ${parentEntityName}.`} href={DOCS.joinTables} />}
					>
						<RelationshipPicker
							entityLogicalName={parentEntityName}
							onChange={handleRelationshipSelect}
							placeholder="Select a relationship…"
						/>
					</InfoField>
				</div>
			)}

			<div className={styles.section}>

				<InfoField
					label="Entity"
					required
					validation={!node.name ? { state: "error", message: "Enter the logical name of the related entity." } : undefined}
					info={<Info text="The logical name of the related entity." href={DOCS.linkEntity} />}
				>
					<Input
						className={styles.mono}
						value={node.name}
						onChange={(_, data) => onUpdate({ name: data.value })}
						placeholder="e.g. contact"
					/>
				</InfoField>

				<InfoField
					label="From attribute"
					validation={fromToAdvice}
					info={
						<Info
							text={`The ${node.name || "related entity"} attribute that matches the To attribute. Not technically required, but usually set. Both attributes should have the same type.`}
							href={DOCS.linkEntityFromTo}
						/>
					}
				>
					<AttributePicker
						entityLogicalName={node.name}
						value={node.from ?? ""}
						onChange={(logicalName) => onUpdate({ from: logicalName || undefined })}
						placeholder="Select an attribute"
						disabled={!node.name}
					/>
				</InfoField>

				<InfoField
					label="To attribute"
					info={
						<Info
							text={`The ${parentEntityName || "parent"} attribute that matches the From attribute. Not technically required, but usually set.`}
							href={DOCS.linkEntityFromTo}
						/>
					}
				>
					<AttributePicker
						entityLogicalName={parentEntityName}
						value={node.to ?? ""}
						onChange={(logicalName) => onUpdate({ to: logicalName || undefined })}
						placeholder="Select an attribute"
						disabled={!parentEntityName}
					/>
				</InfoField>

				<InfoField
					label="Link type"
					validation={findingFor(findings, "link-type")}
					hint={inFilter ? "Inside a filter, a link-entity checks related records and returns nothing." : undefined}
					info={<Info text={LINK_TYPES[node.linkType]?.description ?? "How the related records are matched."} href={DOCS.linkEntity} />}
				>
					<Dropdown
						value={LINK_TYPES[node.linkType]?.label ?? node.linkType}
						selectedOptions={[node.linkType]}
						onOptionSelect={(_, data) => data.optionValue && handleLinkTypeChange(data.optionValue as LinkType)}
					>
						{!offered.includes(node.linkType) &&
							linkTypeOption(node.linkType, inFilter ? "Not allowed inside a filter" : undefined)}
						{inFilter ? (
							RELATED_TYPES.map((t) => linkTypeOption(t))
						) : (
							<>
								<OptionGroup label="Join">{JOIN_TYPES.map((t) => linkTypeOption(t))}</OptionGroup>
								<OptionGroup label="Filter on related records">{RELATED_TYPES.map((t) => linkTypeOption(t))}</OptionGroup>
							</>
						)}
					</Dropdown>
				</InfoField>

				<InfoField
					label="Alias"
					validation={findingFor(findings, "alias")}
					info={
						<Info
							text="The name to use for this link-entity in conditions (entity, compare to) and sorts. Without one, Dataverse generates a name ({name}{N}) that the query can't reference. Letters, digits and underscores; starts with a letter or underscore."
							href={DOCS.linkEntity}
						/>
					}
				>
					<Input
						className={styles.mono}
						value={node.alias ?? ""}
						onChange={(_, data) => onUpdate({ alias: sanitizeAlias(data.value) || undefined })}
						placeholder="e.g. parent_account"
					/>
				</InfoField>

				<InfoCheckbox
					label="Intersect"
					checked={!!node.intersect}
					onChange={(checked) => onUpdate({ intersect: checked ? true : undefined })}
					info={
						<Info
							text="Shows that this link-entity only joins (typically the intersect entity of a many-to-many relationship) and returns no attributes on purpose. It doesn't change how the query runs."
							href={DOCS.linkEntity}
						/>
					}
				/>

				{node.visible !== undefined && (
					<LeftoverSetting
						text={`Visible: ${node.visible}`}
						validation={{
							state: "none",
							message: findingFor(findings, "visible")?.message ?? "Dataverse ignores visible.",
						}}
						onRemove={() => onUpdate({ visible: undefined })}
					/>
				)}
			</div>

			<Dialog modalType="alert" open={pendingLinkType !== null} onOpenChange={(_, data) => !data.open && setPendingLinkType(null)}>
				<DialogSurface>
					<DialogBody>
						<DialogTitle>Change the link type to {pendingLinkType ? LINK_TYPES[pendingLinkType].label : ""}?</DialogTitle>
						<DialogContent>
							The {pendingLinkType ? LINK_TYPES[pendingLinkType].label : ""} link type returns no attributes, so
							Dataverse rejects this link-entity&apos;s {returned.join(" and ")}. Remove {returnedCount === 1 ? "it" : "them"}{" "}
							now? Its filters and nested link-entities are kept.
						</DialogContent>
						<DialogActions position="start">
							<Button onClick={() => confirmLinkTypeChange(false)}>Change only</Button>
						</DialogActions>
						<DialogActions position="end">
							<Button onClick={() => setPendingLinkType(null)}>Cancel</Button>
							<Button appearance="primary" onClick={() => confirmLinkTypeChange(true)}>
								Change and remove
							</Button>
						</DialogActions>
					</DialogBody>
				</DialogSurface>
			</Dialog>
		</div>
	);
}
