/**
 * Main application shell with Fluent UI theming and layout
 */

import { useState, useEffect } from "react";
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
import type { AttributeMetadata, LoadedViewInfo } from "../features/fetchxml/api/pptbClient";
import { generateFetchXml } from "../features/fetchxml/model/fetchxml";
import { collectAttributesFromFetchXml } from "../features/fetchxml/model/layoutxml";
import type { QueryResult } from "../features/fetchxml/ui/RightPane/ResultsGrid";

// ⚠️ IMPORTANT: makeStyles must be called OUTSIDE the component
// but tokens will automatically update when FluentProvider theme changes
// because Fluent UI v9 uses CSS variables under the hood
const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "row",
		height: "98%",
		width: "100%",
		overflow: "hidden",
		position: "absolute",
		top: 0,
		left: 0,
		right: 0,
	},
	leftPane: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
		minWidth: "300px",
		maxWidth: "800px",
	},
	leftPaneTop: {
		display: "flex",
		flexDirection: "column",
		overflow: "hidden",
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

function AppContent() {
	const styles = useStyles();
	const builder = useBuilder();
	const { loadAttributes } = useLazyMetadata();

	// State for query execution
	const [queryResult, setQueryResult] = useState<QueryResult | null>(null);
	const [isExecuting, setIsExecuting] = useState(false);
	const [attributeMetadata, setAttributeMetadata] = useState<Map<string, AttributeMetadata>>(
		new Map()
	);

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

	// Load attribute metadata when entity changes
	useEffect(() => {
		const entityName = builder.fetchQuery?.entity?.name;
		if (!entityName) {
			setAttributeMetadata(new Map());
			return;
		}

		loadAttributes(entityName)
			.then((attributes) => {
				const map = new Map<string, AttributeMetadata>();
				attributes.forEach((attr) => {
					map.set(attr.LogicalName, attr);
				});
				setAttributeMetadata(map);
			})
			.catch((error) => {
				console.error("Failed to load attribute metadata:", error);
				setAttributeMetadata(new Map());
			});
	}, [builder.fetchQuery?.entity?.name, loadAttributes]);

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
			const attributeTypeMap = new Map<string, string>();
			if (attributeMetadata) {
				attributeMetadata.forEach((attr, name) => {
					attributeTypeMap.set(name, attr.AttributeType || "");
				});
			}
			builder.syncLayoutWithFetchXml(attributeTypeMap);
		}
	}, [builder.layoutNeedsSync, builder.fetchQuery, builder.columnConfig, attributeMetadata]);

	// Generate FetchXML from builder state
	const fetchXml = builder.fetchQuery ? generateFetchXml(builder.fetchQuery) : "";

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
				/>
			</div>
		</div>
	);
}
