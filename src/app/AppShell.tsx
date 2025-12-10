/**
 * Main application shell with Fluent UI theming and layout
 */

import { useState, useEffect, useMemo } from "react";
import {
	FluentProvider,
	webLightTheme,
	webDarkTheme,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { usePptbContext } from "../shared/hooks/usePptbContext";
import { useLazyMetadata } from "../shared/hooks/useLazyMetadata";
import { PreviewTabs } from "../features/fetchxml/ui/RightPane/PreviewTabs";
import { EntitySelector } from "../features/fetchxml/ui/Toolbar/EntitySelector";
import { SaveViewButton } from "../features/fetchxml/ui/Toolbar/SaveViewButton";
import { TreeView } from "../features/fetchxml/ui/LeftPane/TreeView";
import { PropertiesPanel } from "../features/fetchxml/ui/LeftPane/PropertiesPanel";
import { BuilderProvider, useBuilder } from "../features/fetchxml/state/builderStore";
import { ThemeProvider } from "../shared/contexts/ThemeContext";
import {
	executeFetchXml,
	executeSystemView,
	executePersonalView,
	whoAmI,
	isDataverseAvailable,
} from "../features/fetchxml/api/pptbClient";
import type {
	AttributeMetadata,
	LoadedViewInfo,
	EntityMetadata,
} from "../features/fetchxml/api/pptbClient";
import { generateFetchXml } from "../features/fetchxml/model/fetchxml";
import { generateLayoutXml } from "../features/fetchxml/model/layoutxml";
import { collectAttributesFromFetchXml } from "../features/fetchxml/model/layoutxml";
import type { QueryResult } from "../features/fetchxml/ui/RightPane/ResultsGrid";

// ⚠️ IMPORTANT: makeStyles must be called OUTSIDE the component
// but tokens will automatically update when FluentProvider theme changes
// because Fluent UI v9 uses CSS variables under the hood
const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "row",
		height: "100%",
		width: "100%",
		overflow: "hidden",
		position: "absolute",
		top: 0,
		left: 0,
		right: 0,
		bottom: 0,
	},
	leftPane: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minWidth: "300px",
		maxWidth: "800px",
		height: "100%",
	},
	leftPaneTop: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minHeight: 0, // Important for flex child with overflow
	},
	horizontalResizeHandle: {
		height: "6px",
		cursor: "ns-resize",
		backgroundColor: tokens.colorNeutralBackground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		position: "relative",
		flexShrink: 0,
		":hover": {
			backgroundColor: tokens.colorNeutralBackground3Hover,
		},
		":active": {
			backgroundColor: tokens.colorBrandBackgroundPressed,
		},
	},
	horizontalGrip: {
		fontSize: "10px",
		color: tokens.colorNeutralForeground3,
		userSelect: "none",
		pointerEvents: "none",
		display: "flex",
		gap: "3px",
		alignItems: "center",
	},
	verticalResizeHandle: {
		width: "6px",
		cursor: "ew-resize",
		backgroundColor: tokens.colorNeutralBackground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		position: "relative",
		flexShrink: 0,
		":hover": {
			backgroundColor: tokens.colorNeutralBackground3Hover,
		},
		":active": {
			backgroundColor: tokens.colorBrandBackgroundPressed,
		},
	},
	verticalGrip: {
		fontSize: "10px",
		color: tokens.colorNeutralForeground3,
		userSelect: "none",
		pointerEvents: "none",
		display: "flex",
		flexDirection: "column",
		gap: "2px",
		lineHeight: "0.5",
	},
	leftPaneBottom: {
		display: "flex",
		flexDirection: "column",
		minHeight: "100px",
		overflow: "hidden",
		minWidth: 0, // Important for flex child with overflow
	},
	propertiesContent: {
		flex: 1,
		overflow: "hidden",
	},
	rightPane: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minWidth: 0,
	},
	toolbar: {
		padding: "12px 16px",
		borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		display: "flex",
		gap: "8px",
		alignItems: "center",
		flexShrink: 0,
	},
	placeholder: {
		padding: "16px",
		color: tokens.colorNeutralForeground3,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
	},
});

export function AppShell() {
	const { theme } = usePptbContext();
	// HARDCODED: Force dark theme
	const isDark = true;

	console.log("🎨 AppShell rendering:");
	console.log("  - theme from context:", theme);
	console.log("  - isDark (HARDCODED):", isDark);
	console.log("  - will use Fluent theme:", isDark ? "webDarkTheme" : "webLightTheme");
	console.log("  - tokens.colorNeutralBackground1:", tokens.colorNeutralBackground1);
	console.log("  - tokens.colorNeutralForeground1:", tokens.colorNeutralForeground1);

	return (
		<FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
			<ThemeProvider isDark={isDark}>
				<BuilderProvider>
					<AppContent />
				</BuilderProvider>
			</ThemeProvider>
		</FluentProvider>
	);
}

/**
 * Collect all unique entity logical names from a FetchXML query
 * Includes root entity and all link-entities (recursively)
 */
function collectEntitiesFromFetchQuery(
	fetchQuery: import("../features/fetchxml/model/nodes").FetchNode | null
): string[] {
	if (!fetchQuery?.entity?.name) return [];

	const entities = new Set<string>();
	entities.add(fetchQuery.entity.name);

	const collectFromLinks = (
		links: import("../features/fetchxml/model/nodes").LinkEntityNode[] | undefined
	) => {
		links?.forEach((link) => {
			if (link.name) {
				entities.add(link.name);
			}
			if (link.links) {
				collectFromLinks(link.links);
			}
		});
	};

	collectFromLinks(fetchQuery.entity.links);

	return Array.from(entities);
}

function AppContent() {
	const styles = useStyles();
	const builder = useBuilder();
	const { loadAttributes, loadEntityMetadata } = useLazyMetadata();

	// State for query execution
	const [queryResult, setQueryResult] = useState<QueryResult | null>(null);
	const [isExecuting, setIsExecuting] = useState(false);
	// Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>>
	const [attributeMetadata, setAttributeMetadata] = useState<
		Map<string, Map<string, AttributeMetadata>>
	>(new Map());
	// State for entity metadata (for layoutxml generation)
	const [entityMetadata, setEntityMetadata] = useState<EntityMetadata | null>(null);

	// State for resizable split
	const [topHeight, setTopHeight] = useState(58); // Percentage of left pane height for tree/properties
	const [leftPaneWidth, setLeftPaneWidth] = useState(480); // Width of left pane in pixels
	const [isDraggingHorizontal, setIsDraggingHorizontal] = useState(false); // Tree/Properties resize
	const [isDraggingVertical, setIsDraggingVertical] = useState(false); // Left/Right pane resize

	// Handle horizontal resize (tree/properties split)
	useEffect(() => {
		if (!isDraggingHorizontal) return;

		const handleMouseMove = (e: MouseEvent) => {
			const leftPane = document.querySelector("[data-left-pane]") as HTMLElement;
			if (!leftPane) return;

			const rect = leftPane.getBoundingClientRect();
			const newTopHeight = ((e.clientY - rect.top) / rect.height) * 100;

			// Constrain between 30% and 80%
			if (newTopHeight >= 30 && newTopHeight <= 80) {
				setTopHeight(newTopHeight);
			}
		};

		const handleMouseUp = () => {
			setIsDraggingHorizontal(false);
		};

		document.addEventListener("mousemove", handleMouseMove);
		document.addEventListener("mouseup", handleMouseUp);

		return () => {
			document.removeEventListener("mousemove", handleMouseMove);
			document.removeEventListener("mouseup", handleMouseUp);
		};
	}, [isDraggingHorizontal]);

	// Handle vertical resize (left/right pane split)
	useEffect(() => {
		if (!isDraggingVertical) return;

		const handleMouseMove = (e: MouseEvent) => {
			const appRoot = document.querySelector("[data-app-root]") as HTMLElement;
			if (!appRoot) return;

			const rect = appRoot.getBoundingClientRect();
			const newLeftPaneWidth = e.clientX - rect.left;

			// Constrain between 300px and 800px
			if (newLeftPaneWidth >= 300 && newLeftPaneWidth <= 800) {
				setLeftPaneWidth(newLeftPaneWidth);
			}
		};

		const handleMouseUp = () => {
			setIsDraggingVertical(false);
		};

		document.addEventListener("mousemove", handleMouseMove);
		document.addEventListener("mouseup", handleMouseUp);

		return () => {
			document.removeEventListener("mousemove", handleMouseMove);
			document.removeEventListener("mouseup", handleMouseUp);
		};
	}, [isDraggingVertical]);

	// Check Dataverse API on mount
	useEffect(() => {
		console.log("=== PPTB FetchXML Builder - Dataverse API Check ===");
		console.log("Dataverse API Available:", isDataverseAvailable());

		if (isDataverseAvailable()) {
			console.log("window.dataverseAPI methods:", Object.keys(window.dataverseAPI || {}));

			// Call WhoAmI to verify connection
			whoAmI()
				.then((result) => {
					if (result) {
						console.log("✅ WhoAmI Success:", {
							UserId: result.UserId,
							BusinessUnitId: result.BusinessUnitId,
							OrganizationId: result.OrganizationId,
						});
					} else {
						console.warn("⚠️ WhoAmI returned null - API may not be fully initialized");
					}
				})
				.catch((error) => {
					console.error("❌ WhoAmI Error:", error);
				});
		} else {
			console.error("❌ Dataverse API not available on window object");
		}

		console.log("===================================================");
	}, []);

	// Collect entities from FetchXML and memoize to avoid unnecessary reloads
	const entitiesInQuery = useMemo(() => {
		return collectEntitiesFromFetchQuery(builder.fetchQuery);
	}, [builder.fetchQuery]);

	// Create a stable key for the entities list to minimize effect re-runs
	const entitiesKey = useMemo(() => entitiesInQuery.sort().join(","), [entitiesInQuery]);

	// Load attribute metadata for all entities in the FetchXML query
	// This includes the root entity and all link-entities
	useEffect(() => {
		if (entitiesInQuery.length === 0) {
			setAttributeMetadata(new Map());
			setEntityMetadata(null);
			return;
		}

		const rootEntityName = entitiesInQuery[0];

		// Load attributes for all entities in the query
		Promise.all(
			entitiesInQuery.map(async (entityName) => {
				try {
					const attributes = await loadAttributes(entityName);
					const attrMap = new Map<string, AttributeMetadata>();
					attributes.forEach((attr) => {
						attrMap.set(attr.LogicalName, attr);
					});
					return { entityName, attrMap };
				} catch (error) {
					console.error(`Failed to load attributes for ${entityName}:`, error);
					return { entityName, attrMap: new Map<string, AttributeMetadata>() };
				}
			})
		).then((results) => {
			const multiEntityMap = new Map<string, Map<string, AttributeMetadata>>();
			results.forEach(({ entityName, attrMap }) => {
				multiEntityMap.set(entityName, attrMap);
			});
			setAttributeMetadata(multiEntityMap);
		});

		// Load entity metadata (for ObjectTypeCode and PrimaryIdAttribute)
		loadEntityMetadata(rootEntityName)
			.then((entity) => {
				setEntityMetadata(entity);
			})
			.catch((error) => {
				console.error("Failed to load entity metadata:", error);
				setEntityMetadata(null);
			});
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [entitiesKey, loadAttributes, loadEntityMetadata]);

	// Clear query results when the fetch query structure changes (entity change, view clear, etc.)
	// We use the entity node id as a proxy - it changes when entity is re-selected or view is cleared
	useEffect(() => {
		setQueryResult(null);
	}, [builder.fetchQuery?.entity?.id]);

	// Sync layout with FetchXML when needed (e.g., after adding/removing attributes)
	// Also sync when columnConfig is null but there are attributes (ensures columns button works)
	useEffect(() => {
		const hasAttributes = (builder.fetchQuery?.entity?.attributes?.length ?? 0) > 0;
		const needsSync = builder.layoutNeedsSync || (hasAttributes && !builder.columnConfig);

		if (needsSync && builder.fetchQuery) {
			// Build attribute type map from metadata for better column widths
			// Use root entity's attribute metadata
			const attributeTypeMap = new Map<string, string>();
			const rootEntityName = builder.fetchQuery.entity?.name;
			if (attributeMetadata && rootEntityName) {
				const rootEntityAttrs = attributeMetadata.get(rootEntityName);
				rootEntityAttrs?.forEach((attr, name) => {
					attributeTypeMap.set(name, attr.AttributeType || "");
				});
			}
			builder.syncLayoutWithFetchXml(attributeTypeMap);
		}
	}, [builder.layoutNeedsSync, builder.fetchQuery, builder.columnConfig, attributeMetadata]);

	// Generate FetchXML from builder state
	const fetchXml = builder.fetchQuery ? generateFetchXml(builder.fetchQuery) : "";

	// Generate LayoutXML from column config for saving
	const layoutXml = useMemo(() => {
		if (!builder.columnConfig || !entityMetadata) {
			return "";
		}
		return generateLayoutXml({
			...builder.columnConfig,
			objectTypeCode: entityMetadata.ObjectTypeCode,
			primaryIdAttribute: entityMetadata.PrimaryIdAttribute,
		});
	}, [builder.columnConfig, entityMetadata]);

	// Handle save view completion - update loaded view state
	const handleSaveViewComplete = (
		viewId: string,
		viewType: "system" | "personal",
		viewName: string
	) => {
		// Update the builder's loaded view state so subsequent saves overwrite the same view
		if (entityMetadata) {
			builder.setLoadedView({
				id: viewId,
				type: viewType,
				entitySetName: entityMetadata.EntitySetName,
				name: viewName,
			});
		}
		console.log(`✅ View saved: ${viewName} (${viewType}) - ${viewId}`);
	};

	const handleExecute = async () => {
		if (!fetchXml) return;

		setIsExecuting(true);
		setQueryResult(null);

		try {
			// Measure execution time
			const startTime = performance.now();

			// Determine execution method based on loaded view state
			let result;
			const loadedView = builder.loadedView;

			if (loadedView) {
				// We have a loaded view - check if it's still unmodified
				// by comparing current generated fetchXml with the original
				const isUnmodified =
					fetchXml.replace(/\s+/g, "") === loadedView.originalFetchXml.replace(/\s+/g, "");

				if (isUnmodified) {
					// Execute using optimized view query (savedQuery/userQuery)
					console.log(
						`📋 Executing ${loadedView.type} view "${loadedView.name}" via ${
							loadedView.type === "system" ? "savedQuery" : "userQuery"
						}=${loadedView.id}`
					);

					if (loadedView.type === "system") {
						result = await executeSystemView(loadedView.entitySetName, loadedView.id);
					} else {
						result = await executePersonalView(loadedView.entitySetName, loadedView.id);
					}
				} else {
					// View was modified - fall back to fetchXml execution
					console.log(`📝 View "${loadedView.name}" was modified - executing via fetchXmlQuery`);
					result = await executeFetchXml(fetchXml);
				}
			} else {
				// No loaded view - use standard fetchXml execution
				result = await executeFetchXml(fetchXml);
			}

			const executionTimeMs = Math.round(performance.now() - startTime);

			// Convert FetchXmlResult to QueryResult format for DataGrid
			// IMPORTANT: Derive columns from FetchXML query, not from result data
			// Dataverse doesn't return keys for null/empty values, so we'd miss columns
			// if we relied only on Object.keys(result.records[0])
			let columns: string[];

			// Start with columns from the result data (these have actual values)
			const resultKeys = result.records.length > 0 ? Object.keys(result.records[0]) : [];
			const resultKeySet = new Set(resultKeys.filter((k) => !k.includes("@")));

			if (builder.fetchQuery) {
				// Get columns from FetchXML - this includes all requested attributes
				const fetchXmlColumns = collectAttributesFromFetchXml(builder.fetchQuery);
				const fetchXmlColumnNames = fetchXmlColumns.map((col) => col.name);

				// Build the final column list, handling lookup field naming conventions
				// FetchXML uses: primarycontactid
				// Dataverse returns: _primarycontactid_value
				const columnSet = new Set<string>();
				columns = [];

				for (const colName of fetchXmlColumnNames) {
					// Check if this column exists in result keys directly
					if (resultKeySet.has(colName)) {
						columns.push(colName);
						columnSet.add(colName);
					}
					// Check if it's a lookup field (Dataverse returns _xxx_value for lookups)
					else if (resultKeySet.has(`_${colName}_value`)) {
						// Use the Dataverse naming convention for lookup fields
						columns.push(`_${colName}_value`);
						columnSet.add(`_${colName}_value`);
						columnSet.add(colName); // Mark original name as handled too
					}
					// Column requested but not in results (null for all records)
					else {
						columns.push(colName);
						columnSet.add(colName);
					}
				}

				// Add any result columns not already handled (e.g., extra columns from view)
				for (const key of resultKeys) {
					if (!columnSet.has(key) && !key.includes("@")) {
						columns.push(key);
						columnSet.add(key);
					}
				}
			} else {
				// Fallback: use result keys if no FetchXML query available
				columns = resultKeys.filter((k) => !k.includes("@"));
			}

			const rows = result.records.map((record) => ({
				...record,
			}));

			// Extract entity name from builder state
			const entityLogicalName = builder.fetchQuery?.entity.name;

			setQueryResult({
				columns,
				rows,
				totalRecordCount: result.totalRecordCount,
				moreRecords: result.moreRecords,
				pagingCookie: result.pagingCookie,
				entityLogicalName, // NEW: for CommandBar actions
				executionTimeMs, // NEW: for timing display
			});
		} catch (error) {
			console.error("Failed to execute FetchXML:", error);
			// TODO Phase 7: Show error notification to user
			setQueryResult({
				columns: [],
				rows: [],
			});
		} finally {
			setIsExecuting(false);
		}
	};

	const handleExport = () => {
		// TODO: Implement export functionality in Phase 7
	};

	return (
		<div className={styles.root} data-app-root>
			{/* Left Pane */}
			<div className={styles.leftPane} data-left-pane style={{ width: `${leftPaneWidth}px` }}>
				{/* Top: Tree View */}
				<div className={styles.leftPaneTop} style={{ height: `${topHeight}%` }}>
					<EntitySelector
						selectedEntity={builder.fetchQuery?.entity.name || null}
						onEntityChange={builder.setEntity}
						onNewQuery={builder.newQuery}
						onViewLoad={(viewInfo: LoadedViewInfo) => {
							// Load the view's FetchXML into the tree while preserving view info
							// Pass layoutxml for column configuration if available
							builder.loadView(
								viewInfo.originalFetchXml,
								{
									id: viewInfo.id,
									type: viewInfo.type,
									entitySetName: viewInfo.entitySetName,
									name: viewInfo.name,
								},
								viewInfo.layoutxml
							);
						}}
					/>
					{builder.fetchQuery ? (
						<TreeView
							fetchQuery={builder.fetchQuery}
							selectedNodeId={builder.selectedNodeId}
							onNodeSelect={builder.selectNode}
							onAddAttribute={builder.addAttribute}
							onAddAllAttributes={builder.addAllAttributes}
							onAddOrder={builder.addOrder}
							onAddFilter={builder.addFilter}
							onAddSubfilter={builder.addSubfilter}
							onAddCondition={builder.addCondition}
							onAddLinkEntity={builder.addLinkEntity}
							onRemoveNode={builder.removeNode}
						/>
					) : (
						<div className={styles.placeholder}>Select an entity to begin</div>
					)}
				</div>
				{/* Resize Handle */}
				<div
					className={styles.horizontalResizeHandle}
					onMouseDown={() => setIsDraggingHorizontal(true)}
					title="Drag to resize"
				>
					<span className={styles.horizontalGrip}>
						<span>•</span>
						<span>•</span>
						<span>•</span>
					</span>
				</div>{" "}
				{/* Bottom: Properties Panel */}
				<div className={styles.leftPaneBottom} style={{ height: `${100 - topHeight}%` }}>
					<div className={styles.toolbar}>
						<span style={{ fontWeight: 600 }}>Properties</span>
					</div>
					<div className={styles.propertiesContent}>
						<PropertiesPanel
							selectedNode={builder.selectedNode}
							fetchQuery={builder.fetchQuery}
							onNodeUpdate={builder.updateNode}
						/>
					</div>
				</div>
			</div>
			{/* Vertical Resize Handle */}
			<div
				className={styles.verticalResizeHandle}
				onMouseDown={() => setIsDraggingVertical(true)}
				title="Drag to resize"
			>
				<span className={styles.verticalGrip}>
					<span>•</span>
					<span>•</span>
					<span>•</span>
				</span>
			</div>{" "}
			{/* Right Pane */}
			<div className={styles.rightPane}>
				<PreviewTabs
					xml={fetchXml}
					result={queryResult}
					isExecuting={isExecuting}
					onExecute={handleExecute}
					onExport={handleExport}
					onParseToTree={builder.loadFetchXml}
					attributeMetadata={attributeMetadata}
					fetchQuery={builder.fetchQuery}
					columnConfig={builder.columnConfig}
					onColumnResize={builder.updateColumnWidth}
					saveViewButton={
						<SaveViewButton
							fetchXml={fetchXml}
							layoutXml={layoutXml}
							entityLogicalName={builder.fetchQuery?.entity?.name || ""}
							objectTypeCode={entityMetadata?.ObjectTypeCode || 0}
							primaryIdAttribute={entityMetadata?.PrimaryIdAttribute || ""}
							loadedView={builder.loadedView}
							onSaveComplete={handleSaveViewComplete}
							disabled={!fetchXml || !entityMetadata}
						/>
					}
					onReorderColumns={(columns) => {
						// Set the column config with new order
						if (builder.columnConfig) {
							builder.setColumnConfig({
								...builder.columnConfig,
								columns,
							});
						}
					}}
					onAddColumn={(attributeName) => {
						// Add the attribute to the root entity
						if (builder.fetchQuery?.entity.id) {
							builder.addAttributeByName(builder.fetchQuery.entity.id, attributeName);
						}
					}}
					onSortChange={(data) => {
						// Extract attribute name from columnId (may have entityName prefix or be an alias)
						let attribute = data.columnId;
						let entityName = data.entityName;

						// If columnId contains a dot, it's a link-entity column (alias.attribute)
						if (attribute.includes(".")) {
							const parts = attribute.split(".");
							entityName = parts[0];
							attribute = parts[1];
						}

						// Handle lookup field naming (_attributename_value -> attributename)
						if (attribute.startsWith("_") && attribute.endsWith("_value")) {
							attribute = attribute.slice(1, -6);
						}

						// Check if this is an aliased column on root entity - need to find original attribute name
						// FetchXML order-by uses the original attribute name, not the alias
						if (!entityName && builder.fetchQuery?.entity?.attributes) {
							const aliasedAttr = builder.fetchQuery.entity.attributes.find(
								(a) => a.alias === attribute
							);
							if (aliasedAttr) {
								attribute = aliasedAttr.name;
							}
						}

						builder.setSort(
							attribute,
							data.direction === "descending",
							data.isMultiSort,
							entityName
						);
					}}
				/>
			</div>
		</div>
	);
}
