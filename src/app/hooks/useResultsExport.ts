/**
 * Export to Excel for the results (#301): the prvExportToExcel check on mount, the server export of
 * a loaded view (Dataverse ExportToExcel), the local export of the loaded rows (exceljs, loaded on
 * first use, #187), and the export error, which clears itself after 10 seconds.
 */
import { useEffect, useState } from "react";
import {
	whoAmI,
	isDataverseAvailable,
	exportToExcel as exportViewToExcel,
	downloadBase64File,
	checkPrivilegeByName,
	loadFormattingMetadata,
	loadUserTimeZone,
} from "../../features/fetchxml/api/pptbClient";
import type { AttributeMetadata, EntityMetadata } from "../../features/fetchxml/api/pptbClient";
import { formatDataverseError } from "../../features/fetchxml/api/dataverseErrors";
import { orderByLayout, resultColumns } from "../../features/fetchxml/model/resultColumns";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import type { LayoutXmlConfig } from "../../features/fetchxml/model/layoutxml";
import type { DisplaySettings } from "../../features/fetchxml/model/displaySettings";
import type { LoadedViewState } from "../../features/fetchxml/state/builderStore";
import type { QueryResult } from "../../features/fetchxml/ui/RightPane/ResultsGrid";
import type { ExportError } from "../../features/fetchxml/ui/RightPane/ResultsTabBars";
import { debugLog } from "../../shared/utils/debug";

export interface ExportStatus {
	isExporting: boolean;
	/** Shown in the Results tab; Try again repeats the same export when it can help (#277) */
	error?: ExportError;
	hasPrivilege: boolean;
	privilegeChecked: boolean;
}

export interface ResultsExportInputs {
	/** The view the query was loaded from or saved as; the server export needs one */
	loadedView: LoadedViewState | null;
	/** The builder's query and column layout */
	fetchQuery: FetchNode | null;
	columnConfig: LayoutXmlConfig | null;
	/** The query's FetchXML and LayoutXML, as AppShell generates them */
	fetchXml: string;
	layoutXml: string;
	/** The results shown in the grid */
	queryResult: QueryResult | null;
	/** Attribute metadata per entity, and the root entity's metadata */
	attributeMetadata: Map<string, Map<string, AttributeMetadata>>;
	entityMetadata: EntityMetadata | null;
	displaySettings: DisplaySettings;
}

export interface ResultsExport {
	exportStatus: ExportStatus;
	/** Whether the server export can run: a loaded view and the privilege */
	canExport: boolean;
	/** Why the server export is off, or undefined when it's on */
	exportDisabledReason: string | undefined;
	/** The server export of the loaded view */
	exportToExcel: () => Promise<void>;
	/** The local export of the loaded rows */
	exportLocal: () => Promise<void>;
	/** Repeats the export that failed: the server export or the local one (#277) */
	retryExport: () => void;
	dismissExportError: () => void;
}

export function useResultsExport({
	loadedView,
	fetchQuery,
	columnConfig,
	fetchXml,
	layoutXml,
	queryResult,
	attributeMetadata,
	entityMetadata,
	displaySettings,
}: ResultsExportInputs): ResultsExport {
	const entityLogicalName = fetchQuery?.entity.name;

	// Export to Excel state
	const [exportStatus, setExportStatus] = useState<ExportStatus>({
		isExporting: false,
		hasPrivilege: false,
		privilegeChecked: false,
	});

	// Check export privilege on mount
	useEffect(() => {
		const checkExportPrivilege = async () => {
			try {
				const user = await whoAmI();
				if (user) {
					const hasPrivilege = await checkPrivilegeByName(
						user.UserId,
						"prvExportToExcel",
					);
					setExportStatus((prev) => ({
						...prev,
						hasPrivilege,
						privilegeChecked: true,
					}));
					debugLog(
						"app",
						`📊 Export to Excel privilege: ${hasPrivilege ? "GRANTED" : "DENIED"}`,
					);
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

	/**
	 * Export to Excel using Dataverse ExportToExcel API
	 * Requires a saved view (system or personal) and prvExportToExcel privilege
	 */
	const handleExport = async () => {
		// Check privilege first
		if (!exportStatus.hasPrivilege) {
			setExportStatus((prev) => ({
				...prev,
				error: {
					kind: "server",
					message:
						"You don't have permission to export to Excel. Contact your administrator to request the 'Export to Excel' privilege.",
					retryable: false,
				},
			}));
			return;
		}

		// Check if we have a saved view
		if (!loadedView) {
			setExportStatus((prev) => ({
				...prev,
				error: {
					kind: "server",
					message:
						"Export to Excel requires a saved view. Please save your query as a view first.",
					retryable: false,
				},
			}));
			return;
		}

		if (!fetchXml || !layoutXml) {
			setExportStatus((prev) => ({
				...prev,
				error: {
					kind: "server",
					message: "Export to Excel requires FetchXML and LayoutXML",
					retryable: false,
				},
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
			debugLog(
				"app",
				`📤 Exporting to Excel via ${loadedView.type} view "${loadedView.name}"...`,
			);

			// Use view name for filename
			const viewName = loadedView.name;

			const result = await exportViewToExcel(
				loadedView.id,
				loadedView.type,
				fetchXml,
				layoutXml,
				viewName,
			);

			// Trigger download
			await downloadBase64File(result.excelFile, result.filename);

			debugLog("app", `✅ Export complete: ${result.filename}`);

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
				error: {
					kind: "server",
					message: formatDataverseError(
						error,
						"Export to Excel failed. Please try again.",
					), // (#96)
					retryable: true,
				},
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
				error: {
					kind: "local",
					message: "No data to export. Execute a query first.",
					retryable: false,
				},
			}));
			return;
		}

		setExportStatus((prev) => ({
			...prev,
			isExporting: true,
			error: undefined,
		}));

		try {
			debugLog("app", `📤 Exporting ${queryResult.rows.length} records locally...`);

			// Build column display names map
			const columnDisplayNames = new Map<string, string>();

			// The grid's columns, in the grid's order (resultColumns.ts, #91)
			const getAttribute = (entityName: string, attributeName: string) =>
				attributeMetadata?.get(entityName)?.get(attributeName);
			const columns = orderByLayout(
				resultColumns(fetchQuery, queryResult.rows, getAttribute, queryResult.query),
				columnConfig,
			).map((c) => c.key);

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
				loadedView?.name ||
				entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
				entityLogicalName ||
				"export";

			// Date behaviours and formats, money precision and the user's time zone (#184); the export
			// still works without them (UTC, labelled; currency and decimals from formatted values)
			const entityNames = [...(attributeMetadata?.keys() ?? [])];
			// exceljs (~1 MB) loads on the first local export, not at startup (#187)
			const [formattingEntries, timeZone, { exportToExcelLocal, downloadExcelFile }] =
				await Promise.all([
					Promise.all(
						entityNames.map((name) =>
							loadFormattingMetadata(name)
								.then((formatting) => [name, formatting] as const)
								.catch((error: unknown) => {
									console.error(
										`Failed to load formatting metadata for ${name}:`,
										error,
									);
									return [name, new Map()] as const;
								}),
						),
					),
					loadUserTimeZone(),
					// The chunk is gone if the tool was updated while open; the browser caches the failure
					import("../../features/fetchxml/api/excelExport").catch(() => {
						throw new Error(
							"Couldn't load the Excel exporter. Reload the tool and try again.",
						);
					}),
				]);

			const { buffer, fileName: finalFileName } = await exportToExcelLocal({
				records: queryResult.rows,
				columns,
				columnDisplayNames,
				attributeMetadata,
				entityName: entityLogicalName,
				fileName,
				displaySettings,
				formattingMetadata: new Map(formattingEntries),
				timeZone,
			});

			// Trigger download
			await downloadExcelFile(buffer, finalFileName);

			debugLog("app", `✅ Local export complete: ${finalFileName}`);

			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
			}));
		} catch (error) {
			console.error("Local export failed:", error);
			setExportStatus((prev) => ({
				...prev,
				isExporting: false,
				error: {
					kind: "local",
					message:
						error instanceof Error
							? error.message
							: "Local export failed. Please try again.",
					retryable: true,
				},
			}));
		}
	};

	// A failed export is about the query and view it was for: a new run or another view clears it, so
	// Try again never repeats it on something else (#277). Paging keeps the same query.
	const resultQuery = queryResult?.query;
	const loadedViewId = loadedView?.id;
	useEffect(() => {
		setExportStatus((prev) => (prev.error ? { ...prev, error: undefined } : prev));
	}, [resultQuery, loadedViewId]);

	return {
		exportStatus,
		canExport: !!loadedView && exportStatus.hasPrivilege,
		exportDisabledReason: !loadedView
			? "Save as a view first to enable export"
			: !exportStatus.hasPrivilege
				? "You don't have the prvExportToExcel privilege"
				: undefined,
		exportToExcel: handleExport,
		exportLocal: handleExportLocal,
		retryExport: () => {
			if (exportStatus.error?.kind === "server") void handleExport();
			else if (exportStatus.error?.kind === "local") void handleExportLocal();
		},
		dismissExportError: () => setExportStatus((prev) => ({ ...prev, error: undefined })),
	};
}
