/**
 * Main application shell with Fluent UI theming and layout
 */

import { useState, useEffect, useMemo, useCallback } from "react";
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
import type { RelatedEntityColumn } from "../features/fetchxml/ui/RightPane/AddColumnsPanel";
import { EntitySelector } from "../features/fetchxml/ui/Toolbar/EntitySelector";
import { SaveViewButton } from "../features/fetchxml/ui/Toolbar/SaveViewButton";
import { TreeView } from "../features/fetchxml/ui/LeftPane/TreeView";
import { PropertiesPanel } from "../features/fetchxml/ui/LeftPane/PropertiesPanel";
import { BuilderProvider, useBuilder } from "../features/fetchxml/state/builderStore";
import { ThemeProvider } from "../shared/contexts/ThemeContext";
import { DeleteConfirmDialog } from "../features/fetchxml/ui/Dialogs/DeleteConfirmDialog";
import { BulkDeleteDialog } from "../features/fetchxml/ui/Dialogs/BulkDeleteDialog";
import { WorkflowPickerDialog } from "../features/fetchxml/ui/Dialogs/WorkflowPickerDialog";
import {
	executeFetchXml,
	executeSystemView,
	executePersonalView,
	whoAmI,
	isDataverseAvailable,
	exportToExcel,
	downloadBase64File,
	checkPrivilegeByName,
	getEnvironmentUrl,
	buildRecordUrl,
	deleteRecord,
	deleteRecordsBatch,
	checkBulkDeletePrivilege,
	submitBulkDelete,
	submitBulkDeleteFromFetchXml,
	checkWorkflowPrivileges,
	getOnDemandWorkflows,
	executeWorkflowBatch,
	checkDeletePrivilege,
} from "../features/fetchxml/api/pptbClient";
import { exportToExcelLocal, downloadExcelFile } from "../features/fetchxml/api/excelExport";
import type {
	AttributeMetadata,
	LoadedViewInfo,
	EntityMetadata,
	RelationshipMetadata,
	WorkflowInfo,
	BatchDeleteProgress,
	BatchDeleteResult,
	WorkflowBatchProgress,
} from "../features/fetchxml/api/pptbClient";
import { generateFetchXml, addPagingToFetchXml } from "../features/fetchxml/model/fetchxml";
import { generateLayoutXml } from "../features/fetchxml/model/layoutxml";
import { collectAttributesFromFetchXml } from "../features/fetchxml/model/layoutxml";
import type { QueryResult } from "../features/fetchxml/ui/RightPane/ResultsGrid";
import { SettingsDrawer } from "../features/fetchxml/ui/Settings/SettingsDrawer";
import {
	defaultDisplaySettings,
	type DisplaySettings,
} from "../features/fetchxml/model/displaySettings";

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
	// Use theme from PPTB host context
	const isDark = theme === "dark";

	console.log("🎨 AppShell rendering:");
	console.log("  - theme from context:", theme);
	console.log("  - isDark:", isDark);
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
	const { loadAttributes, loadEntityMetadata, loadRelationships } = useLazyMetadata();

	// State for query execution
	const [queryResult, setQueryResult] = useState<QueryResult | null>(null);
	const [isExecuting, setIsExecuting] = useState(false);
	const [isLoadingMore, setIsLoadingMore] = useState(false);
	// Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>>
	const [attributeMetadata, setAttributeMetadata] = useState<
		Map<string, Map<string, AttributeMetadata>>
	>(new Map());
	// State for entity metadata (for layoutxml generation)
	const [entityMetadata, setEntityMetadata] = useState<EntityMetadata | null>(null);
	// State for lookup relationships (many-to-one) and one-to-many relationships
	const [lookupRelationships, setLookupRelationships] = useState<RelationshipMetadata[]>([]);
	const [oneToManyRelationships, setOneToManyRelationships] = useState<RelationshipMetadata[]>([]);
	const [isLoadingRelationships, setIsLoadingRelationships] = useState(false);

	// Paging state for infinite scroll / Retrieve All
	// Kept separate from queryResult to track progress across page fetches
	const [pagingState, setPagingState] = useState<{
		currentPage: number;
		pagingCookie?: string;
		moreRecords: boolean;
		isRetrieveAllInProgress: boolean;
	} | null>(null);

	// Export to Excel state
	const [exportStatus, setExportStatus] = useState<{
		isExporting: boolean;
		error?: string;
		hasPrivilege: boolean;
		privilegeChecked: boolean;
	}>({
		isExporting: false,
		hasPrivilege: false,
		privilegeChecked: false,
	});

	// Record action privileges state
	const [recordActionPrivileges, setRecordActionPrivileges] = useState<{
		canDelete: boolean;
		canBulkDelete: boolean;
		canRunWorkflow: boolean;
		privilegesChecked: boolean;
	}>({
		canDelete: false,
		canBulkDelete: false,
		canRunWorkflow: false,
		privilegesChecked: false,
	});

	// Dialog states for record actions
	const [deleteDialogState, setDeleteDialogState] = useState<{
		open: boolean;
		recordIds: string[];
		recordName?: string;
		isBatchDelete?: boolean; // Using batch DELETE for 4-100 records
	}>({ open: false, recordIds: [] });

	// Progress tracking for batch delete
	const [deleteProgress, setDeleteProgress] = useState<BatchDeleteProgress | null>(null);

	const [bulkDeleteDialogState, setBulkDeleteDialogState] = useState<{
		open: boolean;
		recordIds: string[];
		isAllRecords?: boolean;
		totalViewRecords?: number;
	}>({ open: false, recordIds: [] });

	const [workflowDialogState, setWorkflowDialogState] = useState<{
		open: boolean;
		recordIds: string[];
		preSelectedWorkflow?: WorkflowInfo;
	}>({ open: false, recordIds: [] });

	// Currently selected record IDs (from ResultsGrid)
	const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

	// Settings drawer state
	const [settingsDrawerOpen, setSettingsDrawerOpen] = useState(false);
	const [displaySettings, setDisplaySettings] = useState<DisplaySettings>(defaultDisplaySettings);

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

	// Check export privilege on mount
	useEffect(() => {
		const checkExportPrivilege = async () => {
			try {
				const user = await whoAmI();
				if (user) {
					const hasPrivilege = await checkPrivilegeByName(user.UserId, "prvExportToExcel");
					setExportStatus((prev) => ({
						...prev,
						hasPrivilege,
						privilegeChecked: true,
					}));
					console.log(`📊 Export to Excel privilege: ${hasPrivilege ? "GRANTED" : "DENIED"}`);
				} else {
					setExportStatus((prev) => ({
						...prev,
						hasPrivilege: false,
						privilegeChecked: true,
					}));
				}
			} catch (error) {
				console.error("Failed to check export privilege:", error);
				setExportStatus((prev) => ({
					...prev,
					hasPrivilege: false,
					privilegeChecked: true,
				}));
			}
		};

		if (isDataverseAvailable()) {
			checkExportPrivilege();
		}
	}, []);

	// Check record action privileges when entity changes
	const entityLogicalName = builder.fetchQuery?.entity.name;
	useEffect(() => {
		if (!entityLogicalName || !isDataverseAvailable()) {
			setRecordActionPrivileges({
				canDelete: false,
				canBulkDelete: false,
				canRunWorkflow: false,
				privilegesChecked: false,
			});
			return;
		}

		const checkRecordPrivileges = async () => {
			try {
				// Check delete privilege for this entity
				const canDelete = await checkDeletePrivilege(entityLogicalName);

				// Check bulk delete privilege
				const canBulkDelete = await checkBulkDeletePrivilege();

				// Check workflow privileges
				const canRunWorkflow = await checkWorkflowPrivileges();

				setRecordActionPrivileges({
					canDelete,
					canBulkDelete,
					canRunWorkflow,
					privilegesChecked: true,
				});

				console.log(`🔐 Record action privileges for ${entityLogicalName}:`, {
					canDelete,
					canBulkDelete,
					canRunWorkflow,
				});
			} catch (error) {
				console.error("Failed to check record action privileges:", error);
				setRecordActionPrivileges({
					canDelete: false,
					canBulkDelete: false,
					canRunWorkflow: false,
					privilegesChecked: true,
				});
			}
		};

		checkRecordPrivileges();
	}, [entityLogicalName]);

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

		// Load relationships (many-to-one and one-to-many) for the Add Columns panel
		setIsLoadingRelationships(true);
		loadRelationships(rootEntityName)
			.then((relationships) => {
				// Store lookup relationships (many-to-one)
				setLookupRelationships(relationships.manyToOne);
				// Store one-to-many relationships for 1-N column support
				setOneToManyRelationships(relationships.oneToMany);
			})
			.catch((error) => {
				console.error("Failed to load relationships:", error);
				setLookupRelationships([]);
				setOneToManyRelationships([]);
			})
			.finally(() => {
				setIsLoadingRelationships(false);
			});
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [entitiesKey, loadAttributes, loadEntityMetadata, loadRelationships]);

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

	// Check if current query is an aggregate query (disables delete/workflow buttons)
	const isAggregateQuery = builder.fetchQuery?.options?.aggregate === true;

	// Check if query has 1-N or N-N relationships that cause row duplication
	// When these relationships exist, record commands should be disabled
	const hasOneToManyRelationship = useMemo(() => {
		if (!builder.fetchQuery?.entity?.links) return false;

		const checkLinks = (links: typeof builder.fetchQuery.entity.links): boolean => {
			for (const link of links) {
				// Check if this link-entity is a 1-N or N-N relationship
				if (link.relationshipType === "1N" || link.relationshipType === "NN") {
					return true;
				}
				// Check nested link-entities
				if (link.links && checkLinks(link.links)) {
					return true;
				}
			}
			return false;
		};

		return checkLinks(builder.fetchQuery.entity.links);
	}, [builder.fetchQuery]);

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

	/**
	 * Build columns array from FetchXML and result data
	 * Handles lookup field naming conventions and ensures all requested columns are present
	 */
	const buildColumnsFromResult = (
		records: Record<string, unknown>[],
		fetchQuery: typeof builder.fetchQuery
	): string[] => {
		// Start with columns from the result data (these have actual values)
		const resultKeys = records.length > 0 ? Object.keys(records[0]) : [];
		const resultKeySet = new Set(resultKeys.filter((k) => !k.includes("@")));

		if (fetchQuery) {
			// Get columns from FetchXML - this includes all requested attributes
			const fetchXmlColumns = collectAttributesFromFetchXml(fetchQuery);
			const fetchXmlColumnNames = fetchXmlColumns.map((col) => col.name);

			// Build the final column list, handling lookup field naming conventions
			const columnSet = new Set<string>();
			const columns: string[] = [];

			for (const colName of fetchXmlColumnNames) {
				if (resultKeySet.has(colName)) {
					columns.push(colName);
					columnSet.add(colName);
				} else if (resultKeySet.has(`_${colName}_value`)) {
					columns.push(`_${colName}_value`);
					columnSet.add(`_${colName}_value`);
					columnSet.add(colName);
				} else {
					columns.push(colName);
					columnSet.add(colName);
				}
			}

			// Add any result columns not already handled
			for (const key of resultKeys) {
				if (!columnSet.has(key) && !key.includes("@")) {
					columns.push(key);
					columnSet.add(key);
				}
			}

			return columns;
		} else {
			return resultKeys.filter((k) => !k.includes("@"));
		}
	};

	/**
	 * Execute the query and handle Retrieve All if enabled
	 */
	const handleExecute = async () => {
		if (!fetchXml) return;

		setIsExecuting(true);
		setQueryResult(null);
		setPagingState(null);

		const startTime = performance.now();
		const entityLogicalName = builder.fetchQuery?.entity.name;
		const retrieveAll = builder.retrieveAllRecords;

		// Check if user has set a 'top' limit in FetchXML - if so, we respect it and don't enable paging
		const hasTopLimit = builder.fetchQuery?.options?.top !== undefined;
		// Get page size from fetch options (count), default to 5000 if not set
		const pageSize = builder.fetchQuery?.options?.count;

		try {
			// Determine execution method based on loaded view state
			let result;
			const loadedView = builder.loadedView;
			let useViewExecution = false;

			if (loadedView) {
				const isUnmodified =
					fetchXml.replace(/\s+/g, "") === loadedView.originalFetchXml.replace(/\s+/g, "");

				if (isUnmodified) {
					useViewExecution = true;
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
				}
			}

			// If not using view execution (either no view or view was modified), use FetchXML
			if (!result) {
				console.log(
					loadedView
						? `📝 View "${loadedView.name}" was modified - executing via fetchXmlQuery`
						: "📡 Executing FetchXML query"
				);
				result = await executeFetchXml(fetchXml);
			}

			const executionTimeMs = Math.round(performance.now() - startTime);
			const columns = buildColumnsFromResult(result.records, builder.fetchQuery);
			const rows = result.records.map((record) => ({ ...record }));

			// Set initial result
			setQueryResult({
				columns,
				rows,
				totalRecordCount: result.totalRecordCount,
				moreRecords: result.moreRecords,
				pagingCookie: result.pagingCookie,
				entityLogicalName,
				executionTimeMs,
			});

			// Initialize paging state
			setPagingState({
				currentPage: 1,
				pagingCookie: result.pagingCookie,
				moreRecords: result.moreRecords ?? false,
				isRetrieveAllInProgress: retrieveAll && (result.moreRecords ?? false) && !hasTopLimit,
			});

			setIsExecuting(false);

			// If Retrieve All is enabled and there are more records, start progressive loading
			// Only for FetchXML execution, not for view execution (which doesn't support paging)
			// Don't do Retrieve All if user set a 'top' limit - Dataverse handles the limit
			if (retrieveAll && result.moreRecords && !useViewExecution && !hasTopLimit) {
				await loadAllPages(
					fetchXml,
					columns,
					rows,
					result.pagingCookie,
					2,
					entityLogicalName,
					pageSize
				);
			}
		} catch (error) {
			console.error("Failed to execute FetchXML:", error);
			setQueryResult({ columns: [], rows: [] });
			setIsExecuting(false);
		}
	};

	/**
	 * Progressive loading of all pages (Retrieve All)
	 * Updates results as each page is fetched
	 */
	const loadAllPages = async (
		baseFetchXml: string,
		columns: string[],
		initialRows: Record<string, unknown>[],
		pagingCookie: string | undefined,
		startPage: number,
		entityLogicalName: string | undefined,
		pageSize?: number
	) => {
		let allRows = [...initialRows];
		let currentPagingCookie = pagingCookie;
		let page = startPage;
		let hasMore = true;

		setIsLoadingMore(true);

		try {
			while (hasMore) {
				console.log(`📄 Retrieve All: Fetching page ${page}...`);

				// Add paging parameters to FetchXML (page, paging-cookie, and count for page size)
				const pagedFetchXml = addPagingToFetchXml(
					baseFetchXml,
					page,
					currentPagingCookie,
					pageSize
				);
				const result = await executeFetchXml(pagedFetchXml);

				// Append new rows
				allRows = [...allRows, ...result.records];

				// Update the grid progressively
				setQueryResult((prev) => ({
					columns: prev?.columns || columns,
					rows: allRows,
					totalRecordCount: prev?.totalRecordCount,
					moreRecords: result.moreRecords,
					pagingCookie: result.pagingCookie,
					entityLogicalName,
					executionTimeMs: prev?.executionTimeMs,
				}));

				// Update paging state
				currentPagingCookie = result.pagingCookie;
				hasMore = result.moreRecords ?? false;

				setPagingState({
					currentPage: page,
					pagingCookie: currentPagingCookie,
					moreRecords: hasMore,
					isRetrieveAllInProgress: hasMore,
				});

				page++;
			}

			console.log(`✅ Retrieve All complete: ${allRows.length} total records`);
		} catch (error) {
			console.error("Failed during Retrieve All:", error);
		} finally {
			setIsLoadingMore(false);
			setPagingState((prev) => (prev ? { ...prev, isRetrieveAllInProgress: false } : null));
		}
	};

	/**
	 * Load more records (for infinite scroll when Retrieve All is OFF)
	 */
	const handleLoadMore = async () => {
		if (!fetchXml || !pagingState || !pagingState.moreRecords || isLoadingMore) return;
		if (pagingState.isRetrieveAllInProgress) return; // Don't allow manual load during Retrieve All

		// Don't load more if user has set a 'top' limit
		if (builder.fetchQuery?.options?.top !== undefined) return;

		setIsLoadingMore(true);

		// Get page size from fetch options to maintain consistent page sizes
		const pageSize = builder.fetchQuery?.options?.count;

		try {
			const nextPage = pagingState.currentPage + 1;
			console.log(
				`📄 Loading page ${nextPage}${pagingState.pagingCookie ? " with paging cookie" : ""}...`
			);

			// Add paging parameters: page number, paging cookie (required for reliable paging), and count (page size)
			const pagedFetchXml = addPagingToFetchXml(
				fetchXml,
				nextPage,
				pagingState.pagingCookie,
				pageSize
			);
			const result = await executeFetchXml(pagedFetchXml);

			// Append new rows to existing result
			setQueryResult((prev) => {
				if (!prev) return prev;
				return {
					...prev,
					rows: [...prev.rows, ...result.records],
					moreRecords: result.moreRecords,
					pagingCookie: result.pagingCookie,
				};
			});

			setPagingState({
				currentPage: nextPage,
				pagingCookie: result.pagingCookie,
				moreRecords: result.moreRecords ?? false,
				isRetrieveAllInProgress: false,
			});

			console.log(`✅ Page ${nextPage} loaded: ${result.records.length} records`);
		} catch (error) {
			console.error("Failed to load more records:", error);
		} finally {
			setIsLoadingMore(false);
		}
	};

	/**
	 * Export to Excel using Dataverse ExportToExcel API
	 * Requires a saved view (system or personal) and prvExportToExcel privilege
	 */
	const handleExport = async () => {
		// Check privilege first
		if (!exportStatus.hasPrivilege) {
			setExportStatus((prev) => ({
				...prev,
				error:
					"You don't have permission to export to Excel. Contact your administrator to request the 'Export to Excel' privilege.",
			}));
			return;
		}

		// Check if we have a saved view
		if (!builder.loadedView) {
			setExportStatus((prev) => ({
				...prev,
				error: "Export to Excel requires a saved view. Please save your query as a view first.",
			}));
			return;
		}

		if (!fetchXml || !layoutXml) {
			setExportStatus((prev) => ({
				...prev,
				error: "Export to Excel requires FetchXML and LayoutXML",
			}));
			return;
		}

		// Clear any previous error and set exporting state
		setExportStatus((prev) => ({
			...prev,
			isExporting: true,
			error: undefined,
		}));

		try {
			console.log(
				`📤 Exporting to Excel via ${builder.loadedView.type} view "${builder.loadedView.name}"...`
			);

			// Use view name for filename
			const viewName = builder.loadedView.name;

			const result = await exportToExcel(
				builder.loadedView.id,
				builder.loadedView.type,
				fetchXml,
				layoutXml,
				viewName
			);

			// Trigger download
			downloadBase64File(result.excelFile, result.filename);

			console.log(`✅ Export complete: ${result.filename}`);

			// Clear exporting state on success
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
			}));
		} catch (error) {
			console.error("Export to Excel failed:", error);
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
				error: error instanceof Error ? error.message : "Export to Excel failed. Please try again.",
			}));
		}
	};

	/**
	 * Export to Excel locally using exceljs
	 * Exports the current result set with native Excel data types
	 * No view required - works directly with query results
	 */
	const handleExportLocal = async () => {
		if (!queryResult || !queryResult.rows.length) {
			setExportStatus((prev) => ({
				...prev,
				error: "No data to export. Execute a query first.",
			}));
			return;
		}

		setExportStatus((prev) => ({
			...prev,
			isExporting: true,
			error: undefined,
		}));

		try {
			console.log(`📤 Exporting ${queryResult.rows.length} records locally...`);

			// Build column display names map
			const columnDisplayNames = new Map<string, string>();

			// Use actual column names from queryResult (includes _value format for lookups)
			const columns = queryResult.columns;

			for (const col of columns) {
				// Try to get display name from attribute metadata
				let displayName = col;

				if (attributeMetadata && entityLogicalName) {
					const entityAttrs = attributeMetadata.get(entityLogicalName);
					if (entityAttrs) {
						// Direct attribute lookup
						const attr = entityAttrs.get(col);
						if (attr?.DisplayName?.UserLocalizedLabel?.Label) {
							displayName = attr.DisplayName.UserLocalizedLabel.Label;
						}
						// Handle lookup columns (_primarycontactid_value -> primarycontactid)
						if (col.startsWith("_") && col.endsWith("_value")) {
							const baseAttr = col.slice(1, -6);
							const lookupAttr = entityAttrs.get(baseAttr);
							if (lookupAttr?.DisplayName?.UserLocalizedLabel?.Label) {
								displayName = lookupAttr.DisplayName.UserLocalizedLabel.Label;
							}
						}
					}
				}

				columnDisplayNames.set(col, displayName);
			}

			// Generate filename from entity or view name
			const fileName =
				builder.loadedView?.name ||
				entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
				entityLogicalName ||
				"export";

			const { buffer, fileName: finalFileName } = await exportToExcelLocal({
				records: queryResult.rows,
				columns,
				columnDisplayNames,
				attributeMetadata,
				entityName: entityLogicalName,
				fileName,
				displaySettings,
			});

			// Trigger download
			downloadExcelFile(buffer, finalFileName);

			console.log(`✅ Local export complete: ${finalFileName}`);

			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
			}));
		} catch (error) {
			console.error("Local export failed:", error);
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
				error: error instanceof Error ? error.message : "Local export failed. Please try again.",
			}));
		}
	};

	// Clear export error after 10 seconds
	useEffect(() => {
		if (exportStatus.error) {
			const timer = setTimeout(() => {
				setExportStatus((prev) => ({ ...prev, error: undefined }));
			}, 10000);
			return () => clearTimeout(timer);
		}
	}, [exportStatus.error]);

	// ============ RECORD ACTION HANDLERS ============

	/**
	 * Handle selection change from ResultsGrid
	 */
	const handleSelectionChange = useCallback((recordIds: string[]) => {
		setSelectedRecordIds(recordIds);
	}, []);

	/**
	 * Get selected record IDs - called by command bar actions
	 */
	const getSelectedRecordIds = useCallback(() => {
		return selectedRecordIds;
	}, [selectedRecordIds]);

	/**
	 * Open record in a new browser tab
	 */
	const handleOpenRecord = useCallback(
		async (recordIds: string[]) => {
			if (recordIds.length === 0 || !entityLogicalName) return;

			const envUrl = await getEnvironmentUrl();
			if (!envUrl) {
				console.error("Failed to get environment URL");
				return;
			}

			// Open each selected record in a new tab
			for (const recordId of recordIds) {
				const url = buildRecordUrl(entityLogicalName, recordId, envUrl);
				window.open(url, "_blank");
			}
		},
		[entityLogicalName]
	);

	/**
	 * Copy record URL(s) to clipboard
	 */
	const handleCopyRecordUrl = useCallback(
		async (recordIds: string[]) => {
			if (recordIds.length === 0 || !entityLogicalName) return;

			const envUrl = await getEnvironmentUrl();
			if (!envUrl) {
				console.error("Failed to get environment URL");
				return;
			}

			const urls = recordIds.map((id) => buildRecordUrl(entityLogicalName, id, envUrl));
			const textToCopy = urls.join("\n");

			try {
				await navigator.clipboard.writeText(textToCopy);
				console.log(`📋 Copied ${urls.length} record URL(s) to clipboard`);
			} catch (error) {
				console.error("Failed to copy to clipboard:", error);
			}
		},
		[entityLogicalName]
	);

	/**
	 * Open delete confirmation dialog
	 * Handles 3 scenarios:
	 * 1. No selection → Bulk delete ALL records from view (with warning)
	 * 2. 1-100 records → Batch DELETE requests (with progress)
	 * 3. >100 records → Bulk delete with FetchXML IN operator
	 */
	const handleDeleteRecords = useCallback(
		(recordIds: string[]) => {
			if (!entityLogicalName) return;

			// Scenario 1: No selection - delete ALL records from view
			if (recordIds.length === 0) {
				if (!recordActionPrivileges.canBulkDelete) {
					console.warn("Bulk delete privilege required to delete all records from view");
					return;
				}
				const totalRecords = queryResult?.rows.length || 0;
				setBulkDeleteDialogState({
					open: true,
					recordIds: [],
					isAllRecords: true,
					totalViewRecords: totalRecords,
				});
				return;
			}

			// Scenario 3: More than 100 records - use bulk delete with IN operator
			if (recordIds.length > 100 && recordActionPrivileges.canBulkDelete) {
				setBulkDeleteDialogState({
					open: true,
					recordIds,
					isAllRecords: false,
				});
				return;
			}

			// Scenario 2: 1-100 records - use batch DELETE or direct delete
			const recordName =
				recordIds.length === 1
					? (queryResult?.rows.find(
							(row) => row[entityMetadata?.PrimaryIdAttribute || ""] === recordIds[0]
					  )?.[entityMetadata?.PrimaryNameAttribute || ""] as string | undefined)
					: undefined;

			// Use batch delete for 4+ records (more efficient than sequential)
			const isBatchDelete = recordIds.length >= 4;
			setDeleteDialogState({ open: true, recordIds, recordName, isBatchDelete });
		},
		[entityLogicalName, recordActionPrivileges.canBulkDelete, queryResult, entityMetadata]
	);

	/**
	 * Explicitly trigger bulk delete dialog (from command bar menu)
	 * Bypasses the normal auto-routing logic and always uses bulk delete
	 */
	const handleBulkDeleteRecords = useCallback(
		(recordIds: string[]) => {
			if (!entityLogicalName || !recordActionPrivileges.canBulkDelete) return;

			const isAllRecords = recordIds.length === 0;
			const totalRecords = queryResult?.rows.length || 0;

			setBulkDeleteDialogState({
				open: true,
				recordIds,
				isAllRecords,
				totalViewRecords: isAllRecords ? totalRecords : undefined,
			});
		},
		[entityLogicalName, recordActionPrivileges.canBulkDelete, queryResult]
	);

	/**
	 * Execute delete for records (called from DeleteConfirmDialog)
	 * Supports both sequential delete (1-3 records) and batch delete (4-100 records)
	 */
	const handleConfirmDelete = useCallback(async (): Promise<BatchDeleteResult | void> => {
		if (!entityLogicalName || deleteDialogState.recordIds.length === 0) return;

		const recordIds = deleteDialogState.recordIds;
		setDeleteProgress(null);

		// For batch delete (4+ records), use $batch endpoint
		if (deleteDialogState.isBatchDelete) {
			const result = await deleteRecordsBatch(
				entityLogicalName, // Use logical name, API will pluralize
				recordIds,
				(progress) => setDeleteProgress(progress)
			);

			if (result.succeeded > 0) {
				console.log(`✅ Batch deleted ${result.succeeded} of ${recordIds.length} records`);
				// Re-execute query to refresh results
				handleExecute();
			}

			return result;
		}

		// Sequential delete for 1-3 records
		let successCount = 0;
		const errors: string[] = [];

		for (const recordId of recordIds) {
			try {
				await deleteRecord(entityLogicalName, recordId);
				successCount++;
			} catch (error) {
				errors.push(
					`Failed to delete ${recordId}: ${error instanceof Error ? error.message : String(error)}`
				);
			}
		}

		if (successCount > 0) {
			console.log(`✅ Deleted ${successCount} of ${recordIds.length} records`);
			// Re-execute query to refresh results
			handleExecute();
		}

		if (errors.length > 0) {
			console.error("Delete errors:", errors);
			return { succeeded: successCount, failed: errors.length, errors };
		}
	}, [
		entityLogicalName,
		deleteDialogState.recordIds,
		deleteDialogState.isBatchDelete,
		handleExecute,
	]);

	/**
	 * Submit bulk delete job (called from BulkDeleteDialog)
	 * Supports both selected records (IN operator) and all records from view
	 */
	const handleConfirmBulkDelete = useCallback(
		async (jobName: string) => {
			if (!entityLogicalName || !entityMetadata) {
				throw new Error("Missing entity information for bulk delete");
			}

			// If deleting ALL records from view (no selection)
			if (bulkDeleteDialogState.isAllRecords) {
				// Use the current FetchXML directly from builder state
				if (!builder.fetchQuery) {
					throw new Error("No FetchXML query available");
				}
				const currentFetchXml = generateFetchXml(builder.fetchQuery);
				const result = await submitBulkDeleteFromFetchXml(currentFetchXml, jobName);
				console.log(`📤 Bulk delete job (all records) submitted: ${result.asyncOperationId}`);
				return result;
			}

			// If deleting specific selected records
			if (bulkDeleteDialogState.recordIds.length === 0) {
				throw new Error("No records selected for bulk delete");
			}

			const recordIds = bulkDeleteDialogState.recordIds;
			const primaryIdAttribute = entityMetadata.PrimaryIdAttribute;

			const result = await submitBulkDelete(
				entityLogicalName,
				primaryIdAttribute,
				recordIds,
				jobName
			);

			console.log(`📤 Bulk delete job submitted: ${result.asyncOperationId}`);
			return result;
		},
		[
			entityLogicalName,
			entityMetadata,
			bulkDeleteDialogState.recordIds,
			bulkDeleteDialogState.isAllRecords,
			builder.fetchQuery,
		]
	);

	/**
	 * Run a specific workflow directly (from menu button)
	 * Opens the workflow dialog with the workflow pre-selected for confirmation
	 */
	const handleRunSpecificWorkflow = useCallback((workflow: WorkflowInfo, recordIds: string[]) => {
		if (recordIds.length === 0) return;
		// Open the workflow dialog with the selected workflow pre-selected
		setWorkflowDialogState({ open: true, recordIds, preSelectedWorkflow: workflow });
	}, []);

	/**
	 * Fetch available workflows for the entity
	 */
	const handleFetchWorkflows = useCallback(async (): Promise<WorkflowInfo[]> => {
		if (!entityLogicalName) return [];
		return getOnDemandWorkflows(entityLogicalName);
	}, [entityLogicalName]);

	/**
	 * Execute workflow on selected records
	 */
	const handleExecuteWorkflow = useCallback(
		async (
			workflowId: string,
			recordIds: string[],
			onProgress: (progress: WorkflowBatchProgress) => void
		): Promise<{ succeeded: number; failed: number; errors: string[] }> => {
			if (!entityLogicalName) {
				return { succeeded: 0, failed: recordIds.length, errors: ["Entity not selected"] };
			}
			return executeWorkflowBatch(workflowId, recordIds, entityLogicalName, onProgress);
		},
		[entityLogicalName]
	);

	// ============ END RECORD ACTION HANDLERS ============

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
					layoutXml={layoutXml}
					result={queryResult}
					isExecuting={isExecuting}
					isLoadingMore={isLoadingMore}
					onExecute={handleExecute}
					onExport={handleExport}
					onExportLocal={handleExportLocal}
					canExport={!!builder.loadedView && exportStatus.hasPrivilege}
					isExporting={exportStatus.isExporting}
					exportError={exportStatus.error}
					onDismissExportError={() => setExportStatus((prev) => ({ ...prev, error: undefined }))}
					exportDisabledReason={
						!builder.loadedView
							? "Save as a view first to enable export"
							: !exportStatus.hasPrivilege
							? "You don't have the prvExportToExcel privilege"
							: undefined
					}
					onParseToTree={builder.loadFetchXml}
					attributeMetadata={attributeMetadata}
					fetchQuery={builder.fetchQuery}
					columnConfig={builder.columnConfig}
					onColumnResize={builder.updateColumnWidth}
					onLoadMore={handleLoadMore}
					entityDisplayName={
						entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityMetadata?.LogicalName
					}
					lookupRelationships={lookupRelationships}
					oneToManyRelationships={oneToManyRelationships}
					isLoadingRelationships={isLoadingRelationships}
					onLoadRelatedAttributes={loadAttributes}
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
					onRemoveColumn={(columnName) => {
						// Remove the attribute from the FetchXML
						if (builder.fetchQuery?.entity.id) {
							// Find the attribute node by name and remove it
							const attr = builder.fetchQuery.entity.attributes.find((a) => a.name === columnName);
							if (attr) {
								builder.removeNode(attr.id);
							}
						}
					}}
					onAddColumn={(attributeName) => {
						// Add the attribute to the root entity
						if (builder.fetchQuery?.entity.id) {
							builder.addAttributeByName(builder.fetchQuery.entity.id, attributeName);
						}
					}}
					onAddRelatedColumns={(relatedColumns: RelatedEntityColumn[]) => {
						// Add related entity columns - this requires creating link-entities
						if (!builder.fetchQuery?.entity.id) return;

						// Group by relationship to create link-entities
						const byRelationship = new Map<string, RelatedEntityColumn[]>();
						for (const col of relatedColumns) {
							const key = col.relationship.SchemaName;
							if (!byRelationship.has(key)) {
								byRelationship.set(key, []);
							}
							byRelationship.get(key)!.push(col);
						}

						// For each relationship, find or create link-entity and add attributes
						for (const [, cols] of byRelationship) {
							const rel = cols[0].relationship;
							const relType = cols[0].relationshipType;

							let fromAttr: string;
							let toAttr: string;
							let relatedEntity: string;
							let linkType: "inner" | "outer";
							let relTypeForBuilder: "N1" | "1N";

							if (relType === "N1") {
								// For N:1 (Many-to-One) lookup relationships:
								// - ReferencingAttribute = FK on current entity (e.g., primarycontactid on account)
								// - ReferencedAttribute = PK on related entity (e.g., contactid on contact)
								// In FetchXML link-entity:
								// - from = attribute on the LINKED entity (the PK)
								// - to = attribute on the PARENT entity (the FK)
								fromAttr = rel.ReferencedAttribute; // PK on linked (contact.contactid)
								toAttr = rel.ReferencingAttribute; // FK on parent (account.primarycontactid)
								relatedEntity = rel.ReferencedEntity; // contact
								linkType = "outer"; // N:1 typically use outer join to not filter parent records
								relTypeForBuilder = "N1";
							} else {
								// For 1:N (One-to-Many) relationships:
								// - ReferencingEntity = the related entity (the "many" side that has the FK)
								// - ReferencedEntity = the current entity (the "one" side)
								// - ReferencingAttribute = FK on the related entity
								// - ReferencedAttribute = PK on the current entity
								// In FetchXML link-entity:
								// - from = FK on the LINKED entity (the FK to parent)
								// - to = PK on the PARENT entity
								fromAttr = rel.ReferencingAttribute; // FK on linked entity
								toAttr = rel.ReferencedAttribute; // PK on parent (e.g., accountid)
								relatedEntity = rel.ReferencingEntity; // The related entity (e.g., contact)
								linkType = "outer"; // 1:N use outer join to include parent records without children
								relTypeForBuilder = "1N";
							}

							// Check if link-entity already exists for this relationship
							const existingLinkEntity = builder.fetchQuery.entity.links.find(
								(le) => le.from === fromAttr && le.to === toAttr && le.name === relatedEntity
							);

							let linkEntityId: string;

							if (existingLinkEntity) {
								linkEntityId = existingLinkEntity.id;
							} else {
								// Create the link-entity with config - returns the ID immediately
								linkEntityId = builder.addLinkEntityWithConfig(
									builder.fetchQuery.entity.id,
									relatedEntity,
									fromAttr,
									toAttr,
									linkType,
									relTypeForBuilder
								);
							}

							// Add attributes to the link-entity using the ID
							for (const col of cols) {
								builder.addAttributeByName(linkEntityId, col.attribute.LogicalName);
							}
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
					// Record action handlers
					onOpenRecord={handleOpenRecord}
					onCopyRecordUrl={handleCopyRecordUrl}
					onDeleteRecords={handleDeleteRecords}
					onBulkDeleteRecords={handleBulkDeleteRecords}
					onRunSpecificWorkflow={handleRunSpecificWorkflow}
					canDelete={recordActionPrivileges.canDelete}
					canBulkDelete={recordActionPrivileges.canBulkDelete}
					canRunWorkflow={recordActionPrivileges.canRunWorkflow}
					onFetchWorkflows={handleFetchWorkflows}
					isAggregateQuery={isAggregateQuery}
					hasOneToManyRelationship={hasOneToManyRelationship}
					onSelectionChange={handleSelectionChange}
					getSelectedRecordIds={getSelectedRecordIds}
					onOpenSettings={() => setSettingsDrawerOpen(true)}
					displaySettings={displaySettings}
				/>
			</div>
			{/* Delete Confirmation Dialog */}
			<DeleteConfirmDialog
				open={deleteDialogState.open}
				recordName={deleteDialogState.recordName}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityLogicalName || "record"
				}
				recordCount={deleteDialogState.recordIds.length}
				onClose={() => {
					setDeleteDialogState({ open: false, recordIds: [] });
					setDeleteProgress(null);
				}}
				onConfirm={handleConfirmDelete}
				onProgress={deleteProgress ?? undefined}
			/>
			{/* Bulk Delete Dialog */}
			<BulkDeleteDialog
				open={bulkDeleteDialogState.open}
				recordCount={bulkDeleteDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityLogicalName || "records"
				}
				isAllRecords={bulkDeleteDialogState.isAllRecords}
				totalViewRecords={bulkDeleteDialogState.totalViewRecords}
				onClose={() => setBulkDeleteDialogState({ open: false, recordIds: [] })}
				onConfirm={handleConfirmBulkDelete}
			/>
			{/* Workflow Picker Dialog */}
			<WorkflowPickerDialog
				open={workflowDialogState.open}
				recordCount={workflowDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label || entityLogicalName || "records"
				}
				preSelectedWorkflow={workflowDialogState.preSelectedWorkflow}
				onClose={() => setWorkflowDialogState({ open: false, recordIds: [] })}
				onFetchWorkflows={handleFetchWorkflows}
				onExecute={(workflow, onProgress) =>
					handleExecuteWorkflow(workflow.workflowid, workflowDialogState.recordIds, onProgress)
				}
			/>
			{/* Settings Drawer */}
			<SettingsDrawer
				open={settingsDrawerOpen}
				settings={displaySettings}
				onClose={() => setSettingsDrawerOpen(false)}
				onSettingsChange={setDisplaySettings}
			/>
		</div>
	);
}
