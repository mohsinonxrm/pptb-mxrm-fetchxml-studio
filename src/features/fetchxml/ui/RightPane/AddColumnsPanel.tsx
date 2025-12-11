/**
 * Add Columns Panel (Drawer)
 * Shows available attributes with support for related entities (lookups)
 * Matches Power Apps Model-Driven Apps UX
 */

import { useState, useMemo, useCallback, useEffect } from "react";
import {
	DrawerBody,
	DrawerHeader,
	DrawerHeaderTitle,
	OverlayDrawer,
	Button,
	makeStyles,
	tokens,
	Text,
	Input,
	Checkbox,
	Tree,
	TreeItem,
	TreeItemLayout,
	Spinner,
	TabList,
	Tab,
	Combobox,
	Option,
	type SelectTabData,
	type SelectTabEvent,
} from "@fluentui/react-components";
import { Dismiss24Regular, Search20Regular } from "@fluentui/react-icons";
import type { AttributeMetadata, RelationshipMetadata } from "../../api/pptbClient";

const useStyles = makeStyles({
	drawer: {
		width: "420px",
	},
	searchContainer: {
		display: "flex",
		gap: tokens.spacingHorizontalS,
		marginBottom: tokens.spacingVerticalM,
	},
	searchInput: {
		flex: 1,
	},
	filterCombo: {
		minWidth: "120px",
	},
	tabContent: {
		marginTop: tokens.spacingVerticalM,
	},
	attributeList: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXXS,
		maxHeight: "calc(100vh - 300px)",
		overflowY: "auto",
	},
	attributeItem: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
		borderRadius: tokens.borderRadiusSmall,
		cursor: "pointer",
		"&:hover": {
			backgroundColor: tokens.colorNeutralBackground1Hover,
		},
	},
	attributeItemSelected: {
		backgroundColor: tokens.colorNeutralBackground1Selected,
	},
	relationshipTree: {
		maxHeight: "calc(100vh - 300px)",
		overflowY: "auto",
	},
	relationshipItem: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
	relationshipLabel: {
		display: "flex",
		flexDirection: "column",
	},
	relationshipName: {
		fontWeight: tokens.fontWeightSemibold,
	},
	relationshipEntity: {
		fontSize: tokens.fontSizeBase200,
		color: tokens.colorNeutralForeground3,
	},
	loadingContainer: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		padding: tokens.spacingVerticalXXL,
	},
	footer: {
		display: "flex",
		justifyContent: "flex-end",
		gap: tokens.spacingHorizontalS,
		padding: tokens.spacingVerticalM,
		borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
		marginTop: "auto",
	},
	emptyState: {
		display: "flex",
		flexDirection: "column",
		alignItems: "center",
		justifyContent: "center",
		padding: tokens.spacingVerticalXXL,
		color: tokens.colorNeutralForeground3,
		textAlign: "center",
	},
	selectedCount: {
		marginRight: "auto",
		color: tokens.colorNeutralForeground2,
		display: "flex",
		alignItems: "center",
	},
});

type FilterType = "all" | "standard" | "custom";

export interface RelatedEntityColumn {
	/** The relationship used to access this entity */
	relationship: RelationshipMetadata;
	/** The attribute on the related entity */
	attribute: AttributeMetadata;
	/** Display name including relationship context, e.g. "First Name (Primary Contact)" */
	displayName: string;
	/** The lookup attribute name on the root entity */
	lookupAttribute: string;
}

export interface AddColumnSelection {
	/** Root entity attributes to add */
	rootAttributes: string[];
	/** Related entity columns to add (will create link-entity if needed) */
	relatedColumns: RelatedEntityColumn[];
}

export interface AddColumnsPanelProps {
	/** Whether the panel is open */
	open: boolean;
	/** Entity display name */
	entityDisplayName?: string;
	/** Entity logical name */
	entityLogicalName?: string;
	/** Available attributes for the root entity */
	availableAttributes?: AttributeMetadata[];
	/** Currently selected attribute names (already in the query) */
	selectedAttributes?: string[];
	/** Lookup relationships (many-to-one) for related entity columns */
	lookupRelationships?: RelationshipMetadata[];
	/** Whether relationship data is loading */
	isLoadingRelationships?: boolean;
	/** Callback to load attributes for a related entity */
	onLoadRelatedAttributes?: (entityLogicalName: string) => Promise<AttributeMetadata[]>;
	/** Called when panel is closed */
	onClose: () => void;
	/** Called when columns are selected and Apply is clicked */
	onApply: (selection: AddColumnSelection) => void;
}

export function AddColumnsPanel({
	open,
	entityDisplayName,
	entityLogicalName,
	availableAttributes = [],
	selectedAttributes = [],
	lookupRelationships = [],
	isLoadingRelationships,
	onLoadRelatedAttributes,
	onClose,
	onApply,
}: AddColumnsPanelProps) {
	const styles = useStyles();
	const [activeTab, setActiveTab] = useState<"entity" | "related">("entity");
	const [searchText, setSearchText] = useState("");
	const [filterType, setFilterType] = useState<FilterType>("all");
	const [selectedRootAttrs, setSelectedRootAttrs] = useState<Set<string>>(new Set());
	const [selectedRelatedCols, setSelectedRelatedCols] = useState<RelatedEntityColumn[]>([]);
	const [expandedRelationships, setExpandedRelationships] = useState<Set<string>>(new Set());
	const [relatedAttributesCache, setRelatedAttributesCache] = useState<
		Map<string, AttributeMetadata[]>
	>(new Map());
	const [loadingEntities, setLoadingEntities] = useState<Set<string>>(new Set());

	// Reset state when panel opens
	useEffect(() => {
		if (open) {
			setSelectedRootAttrs(new Set());
			setSelectedRelatedCols([]);
			setSearchText("");
			setFilterType("all");
			setActiveTab("entity");
		}
	}, [open]);

	// Filter attributes based on search and filter type
	const filteredAttributes = useMemo(() => {
		let attrs = availableAttributes.filter(
			(attr) => !selectedAttributes.includes(attr.LogicalName)
		);

		// Apply filter type
		if (filterType === "standard") {
			attrs = attrs.filter(
				(attr) => !attr.LogicalName.startsWith("new_") && !attr.SchemaName.includes("_")
			);
		} else if (filterType === "custom") {
			attrs = attrs.filter(
				(attr) => attr.LogicalName.startsWith("new_") || attr.SchemaName.includes("_")
			);
		}

		// Apply search
		if (searchText) {
			const search = searchText.toLowerCase();
			attrs = attrs.filter(
				(attr) =>
					attr.LogicalName.toLowerCase().includes(search) ||
					(attr.DisplayName?.UserLocalizedLabel?.Label || "").toLowerCase().includes(search)
			);
		}

		// Sort by display name
		return attrs.sort((a, b) => {
			const aName = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
			const bName = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
			return aName.localeCompare(bName);
		});
	}, [availableAttributes, selectedAttributes, searchText, filterType]);

	// Filter lookup relationships
	const filteredRelationships = useMemo(() => {
		let rels = lookupRelationships;

		// Apply search to relationships
		if (searchText) {
			const search = searchText.toLowerCase();
			rels = rels.filter(
				(rel) =>
					rel.ReferencedEntity.toLowerCase().includes(search) ||
					rel.ReferencingAttribute.toLowerCase().includes(search) ||
					rel.SchemaName.toLowerCase().includes(search)
			);
		}

		return rels.sort((a, b) => a.ReferencingAttribute.localeCompare(b.ReferencingAttribute));
	}, [lookupRelationships, searchText]);

	const handleTabSelect = useCallback((_e: SelectTabEvent, data: SelectTabData) => {
		setActiveTab(data.value as "entity" | "related");
	}, []);

	const handleRootAttributeToggle = useCallback((logicalName: string) => {
		setSelectedRootAttrs((prev) => {
			const next = new Set(prev);
			if (next.has(logicalName)) {
				next.delete(logicalName);
			} else {
				next.add(logicalName);
			}
			return next;
		});
	}, []);

	const handleExpandRelationship = useCallback(
		async (relationship: RelationshipMetadata) => {
			const relKey = relationship.SchemaName;

			setExpandedRelationships((prev) => {
				const next = new Set(prev);
				if (next.has(relKey)) {
					next.delete(relKey);
				} else {
					next.add(relKey);
				}
				return next;
			});

			// Load related entity attributes if not already loaded
			if (!relatedAttributesCache.has(relationship.ReferencedEntity) && onLoadRelatedAttributes) {
				setLoadingEntities((prev) => new Set(prev).add(relationship.ReferencedEntity));
				try {
					const attrs = await onLoadRelatedAttributes(relationship.ReferencedEntity);
					setRelatedAttributesCache((prev) =>
						new Map(prev).set(relationship.ReferencedEntity, attrs)
					);
				} finally {
					setLoadingEntities((prev) => {
						const next = new Set(prev);
						next.delete(relationship.ReferencedEntity);
						return next;
					});
				}
			}
		},
		[relatedAttributesCache, onLoadRelatedAttributes]
	);

	const handleRelatedAttributeToggle = useCallback(
		(relationship: RelationshipMetadata, attribute: AttributeMetadata) => {
			const lookupAttr = relationship.ReferencingAttribute;
			const relDisplayName =
				availableAttributes.find((a) => a.LogicalName === lookupAttr)?.DisplayName
					?.UserLocalizedLabel?.Label || lookupAttr;
			const attrDisplayName =
				attribute.DisplayName?.UserLocalizedLabel?.Label || attribute.LogicalName;

			const relatedCol: RelatedEntityColumn = {
				relationship,
				attribute,
				displayName: `${attrDisplayName} (${relDisplayName})`,
				lookupAttribute: lookupAttr,
			};

			setSelectedRelatedCols((prev) => {
				// Check if already selected
				const existingIndex = prev.findIndex(
					(col) =>
						col.relationship.SchemaName === relationship.SchemaName &&
						col.attribute.LogicalName === attribute.LogicalName
				);

				if (existingIndex >= 0) {
					// Remove it
					return prev.filter((_, i) => i !== existingIndex);
				} else {
					// Add it
					return [...prev, relatedCol];
				}
			});
		},
		[availableAttributes]
	);

	const isRelatedAttributeSelected = useCallback(
		(relationship: RelationshipMetadata, attribute: AttributeMetadata) => {
			return selectedRelatedCols.some(
				(col) =>
					col.relationship.SchemaName === relationship.SchemaName &&
					col.attribute.LogicalName === attribute.LogicalName
			);
		},
		[selectedRelatedCols]
	);

	const handleApply = useCallback(() => {
		onApply({
			rootAttributes: Array.from(selectedRootAttrs),
			relatedColumns: selectedRelatedCols,
		});
	}, [selectedRootAttrs, selectedRelatedCols, onApply]);

	const totalSelected = selectedRootAttrs.size + selectedRelatedCols.length;

	return (
		<OverlayDrawer
			open={open}
			onOpenChange={(_e, data) => !data.open && onClose()}
			position="end"
			size="medium"
			className={styles.drawer}
		>
			<DrawerHeader>
				<DrawerHeaderTitle
					action={
						<Button
							appearance="subtle"
							aria-label="Close"
							icon={<Dismiss24Regular />}
							onClick={onClose}
						/>
					}
				>
					Add columns
				</DrawerHeaderTitle>
			</DrawerHeader>
			<DrawerBody>
				<TabList selectedValue={activeTab} onTabSelect={handleTabSelect}>
					<Tab value="entity">{entityDisplayName || entityLogicalName || "Entity"}</Tab>
					<Tab value="related">Related</Tab>
				</TabList>

				<div className={styles.tabContent}>
					<div className={styles.searchContainer}>
						<Input
							className={styles.searchInput}
							contentBefore={<Search20Regular />}
							placeholder="Search columns..."
							value={searchText}
							onChange={(_e, data) => setSearchText(data.value)}
						/>
						<Combobox
							className={styles.filterCombo}
							value={
								filterType === "all" ? "All" : filterType === "standard" ? "Standard" : "Custom"
							}
							onOptionSelect={(_e, data) =>
								setFilterType((data.optionValue as FilterType) || "all")
							}
						>
							<Option value="all">All</Option>
							<Option value="standard">Standard</Option>
							<Option value="custom">Custom</Option>
						</Combobox>
					</div>

					{activeTab === "entity" && (
						<div className={styles.attributeList}>
							{filteredAttributes.length === 0 ? (
								<div className={styles.emptyState}>
									<Text>
										{searchText ? "No matching attributes found" : "All attributes already added"}
									</Text>
								</div>
							) : (
								filteredAttributes.map((attr) => (
									<div
										key={attr.LogicalName}
										className={`${styles.attributeItem} ${
											selectedRootAttrs.has(attr.LogicalName) ? styles.attributeItemSelected : ""
										}`}
										onClick={() => handleRootAttributeToggle(attr.LogicalName)}
										role="checkbox"
										aria-checked={selectedRootAttrs.has(attr.LogicalName)}
										tabIndex={0}
										onKeyDown={(e) => {
											if (e.key === "Enter" || e.key === " ") {
												e.preventDefault();
												handleRootAttributeToggle(attr.LogicalName);
											}
										}}
									>
										<Checkbox
											checked={selectedRootAttrs.has(attr.LogicalName)}
											onChange={() => handleRootAttributeToggle(attr.LogicalName)}
										/>
										<Text>{attr.DisplayName?.UserLocalizedLabel?.Label || attr.LogicalName}</Text>
									</div>
								))
							)}
						</div>
					)}

					{activeTab === "related" && (
						<div className={styles.relationshipTree}>
							{isLoadingRelationships ? (
								<div className={styles.loadingContainer}>
									<Spinner size="small" label="Loading relationships..." />
								</div>
							) : filteredRelationships.length === 0 ? (
								<div className={styles.emptyState}>
									<Text>No lookup relationships found</Text>
								</div>
							) : (
								<Tree aria-label="Related entities">
									{filteredRelationships.map((rel) => {
										const lookupAttr = availableAttributes.find(
											(a) => a.LogicalName === rel.ReferencingAttribute
										);
										const lookupDisplayName =
											lookupAttr?.DisplayName?.UserLocalizedLabel?.Label ||
											rel.ReferencingAttribute;
										const isExpanded = expandedRelationships.has(rel.SchemaName);
										const relatedAttrs = relatedAttributesCache.get(rel.ReferencedEntity);
										const isLoading = loadingEntities.has(rel.ReferencedEntity);

										return (
											<TreeItem key={rel.SchemaName} itemType="branch" open={isExpanded}>
												<TreeItemLayout onClick={() => handleExpandRelationship(rel)}>
													<div className={styles.relationshipLabel}>
														<span className={styles.relationshipName}>{lookupDisplayName}</span>
														<span className={styles.relationshipEntity}>
															{rel.ReferencedEntity}
														</span>
													</div>
												</TreeItemLayout>
												{isExpanded && (
													<Tree>
														{isLoading ? (
															<TreeItem itemType="leaf">
																<TreeItemLayout>
																	<Spinner size="tiny" label="Loading..." />
																</TreeItemLayout>
															</TreeItem>
														) : relatedAttrs && relatedAttrs.length > 0 ? (
															relatedAttrs
																.filter((attr) => attr.AttributeType !== "Virtual")
																.sort((a, b) => {
																	const aName =
																		a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
																	const bName =
																		b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
																	return aName.localeCompare(bName);
																})
																.map((attr) => (
																	<TreeItem key={attr.LogicalName} itemType="leaf">
																		<TreeItemLayout
																			onClick={() => handleRelatedAttributeToggle(rel, attr)}
																		>
																			<div className={styles.attributeItem}>
																				<Checkbox
																					checked={isRelatedAttributeSelected(rel, attr)}
																					onChange={() => handleRelatedAttributeToggle(rel, attr)}
																				/>
																				<Text>
																					{attr.DisplayName?.UserLocalizedLabel?.Label ||
																						attr.LogicalName}
																				</Text>
																			</div>
																		</TreeItemLayout>
																	</TreeItem>
																))
														) : (
															<TreeItem itemType="leaf">
																<TreeItemLayout>
																	<Text style={{ color: tokens.colorNeutralForeground3 }}>
																		No attributes available
																	</Text>
																</TreeItemLayout>
															</TreeItem>
														)}
													</Tree>
												)}
											</TreeItem>
										);
									})}
								</Tree>
							)}
						</div>
					)}
				</div>

				<div className={styles.footer}>
					{totalSelected > 0 && (
						<Text className={styles.selectedCount}>
							{totalSelected} column{totalSelected !== 1 ? "s" : ""} selected
						</Text>
					)}
					<Button appearance="secondary" onClick={onClose}>
						Close
					</Button>
					<Button appearance="primary" onClick={handleApply} disabled={totalSelected === 0}>
						Add
					</Button>
				</div>
			</DrawerBody>
		</OverlayDrawer>
	);
}
