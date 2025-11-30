/**
 * LoadViewPicker - Grouped combobox for selecting System or Personal Views
 * Only visible when root entity is selected in tree
 * Fetches views for the selected entity and allows loading view's FetchXML into the tree
 */

import { useState, useEffect, useMemo, useCallback, type ChangeEvent } from "react";
import {
	Combobox,
	Option,
	OptionGroup,
	makeStyles,
	Spinner,
	tokens,
	type ComboboxProps,
} from "@fluentui/react-components";
import {
	getAllViews,
	type SavedView,
	type EntityMetadata,
	type LoadedViewInfo,
} from "../../api/pptbClient";
import { debugLog } from "../../../../shared/utils/debug";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: "4px",
		flex: 1,
		minWidth: "200px",
		maxWidth: "350px",
	},
	label: {
		fontSize: "14px",
		fontWeight: 600,
		color: tokens.colorNeutralForeground2,
	},
	combobox: {
		width: "100%",
	},
	loadingOption: {
		display: "flex",
		alignItems: "center",
		gap: "8px",
		padding: "8px",
	},
	noViewsOption: {
		fontStyle: "italic",
		color: tokens.colorNeutralForeground3,
	},
	defaultBadge: {
		fontSize: "10px",
		backgroundColor: tokens.colorBrandBackground,
		color: tokens.colorNeutralForegroundOnBrand,
		padding: "1px 4px",
		borderRadius: "4px",
		marginLeft: "6px",
	},
});

interface LoadViewPickerProps {
	/** Currently selected entity metadata (with ObjectTypeCode) */
	selectedEntityMetadata: EntityMetadata | null;
	/** Callback when a view is selected - provides the view info for execution optimization */
	onViewSelect: (viewInfo: LoadedViewInfo) => void;
	/** Optional: disable the picker */
	disabled?: boolean;
}

export function LoadViewPicker({
	selectedEntityMetadata,
	onViewSelect,
	disabled = false,
}: LoadViewPickerProps) {
	const styles = useStyles();

	// State for views
	const [systemViews, setSystemViews] = useState<SavedView[]>([]);
	const [personalViews, setPersonalViews] = useState<SavedView[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);

	// Combobox state - track both input value and selected view name
	const [inputValue, setInputValue] = useState("");
	const [selectedViewName, setSelectedViewName] = useState<string | null>(null);

	// Fetch views when entity changes
	useEffect(() => {
		if (!selectedEntityMetadata) {
			setSystemViews([]);
			setPersonalViews([]);
			setError(null);
			setSelectedViewName(null);
			setInputValue("");
			return;
		}

		const fetchViews = async () => {
			setLoading(true);
			setError(null);
			setSelectedViewName(null);
			setInputValue("");

			try {
				debugLog("viewAPI", `Loading views for ${selectedEntityMetadata.LogicalName}`);

				const { systemViews: sysViews, personalViews: persViews } = await getAllViews(
					selectedEntityMetadata.LogicalName
				);

				setSystemViews(sysViews);
				setPersonalViews(persViews);

				debugLog(
					"viewAPI",
					`Loaded ${sysViews.length} system views and ${persViews.length} personal views`
				);
			} catch (err) {
				const message = err instanceof Error ? err.message : "Failed to load views";
				setError(message);
				console.error("LoadViewPicker: Failed to fetch views:", err);
			} finally {
				setLoading(false);
			}
		};

		fetchViews();
	}, [selectedEntityMetadata]);

	// Combine all views for filtering
	const allViews = useMemo(() => {
		return [...systemViews, ...personalViews];
	}, [systemViews, personalViews]);

	// Filter options based on input value (only when typing, not when showing selected)
	const filterQuery = inputValue && inputValue !== selectedViewName ? inputValue : "";

	const filteredSystemViews = useMemo(() => {
		if (!filterQuery.trim()) return systemViews;
		const lowerQuery = filterQuery.toLowerCase();
		return systemViews.filter((v) => v.name.toLowerCase().includes(lowerQuery));
	}, [systemViews, filterQuery]);

	const filteredPersonalViews = useMemo(() => {
		if (!filterQuery.trim()) return personalViews;
		const lowerQuery = filterQuery.toLowerCase();
		return personalViews.filter((v) => v.name.toLowerCase().includes(lowerQuery));
	}, [personalViews, filterQuery]);

	// Handle selection
	const handleOptionSelect = useCallback<NonNullable<ComboboxProps["onOptionSelect"]>>(
		(_event, data) => {
			if (!data.optionValue || !selectedEntityMetadata) return;

			const view = allViews.find((v) => v.id === data.optionValue);
			if (view) {
				// Create LoadedViewInfo with all needed data for execution optimization
				const viewInfo: LoadedViewInfo = {
					id: view.id,
					type: view.type,
					originalFetchXml: view.fetchxml,
					entitySetName: selectedEntityMetadata.EntitySetName,
					name: view.name,
				};
				onViewSelect(viewInfo);
				debugLog("viewAPI", `Selected view: ${view.name} (${view.type}) - ID: ${view.id}`);
				// Set the selected view name to display in the combobox
				setSelectedViewName(view.name);
				setInputValue(view.name);
			}
		},
		[allViews, onViewSelect, selectedEntityMetadata]
	);

	// Handle input change for filtering
	const handleInputChange = useCallback(
		(event: ChangeEvent<HTMLInputElement>) => {
			setInputValue(event.target.value);
			// Clear selection when user starts typing something different
			if (event.target.value !== selectedViewName) {
				setSelectedViewName(null);
			}
		},
		[selectedViewName]
	);

	// Don't render if no entity selected
	if (!selectedEntityMetadata) {
		return null;
	}

	const hasViews = systemViews.length > 0 || personalViews.length > 0;
	const hasFilteredViews = filteredSystemViews.length > 0 || filteredPersonalViews.length > 0;

	return (
		<div className={styles.container}>
			<label className={styles.label}>Load View</label>
			<Combobox
				className={styles.combobox}
				placeholder={loading ? "Loading views..." : "Select a view to load"}
				disabled={disabled || loading || !hasViews}
				value={inputValue}
				onChange={handleInputChange}
				onOptionSelect={handleOptionSelect}
				freeform
			>
				{loading && (
					<Option key="loading" value="" text="Loading..." disabled>
						<div className={styles.loadingOption}>
							<Spinner size="tiny" />
							<span>Loading views...</span>
						</div>
					</Option>
				)}

				{!loading && error && (
					<Option key="error" value="" text="Error" disabled>
						<span style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</span>
					</Option>
				)}

				{!loading && !error && !hasViews && (
					<Option key="no-views" value="" text="No views available" disabled>
						<span className={styles.noViewsOption}>No views available</span>
					</Option>
				)}

				{!loading && !error && hasViews && !hasFilteredViews && filterQuery && (
					<Option key="no-match" value="" text="No matching views" disabled>
						<span className={styles.noViewsOption}>No matching views</span>
					</Option>
				)}

				{!loading && !error && filteredSystemViews.length > 0 && (
					<OptionGroup label="System Views">
						{filteredSystemViews.map((view) => (
							<Option key={view.id} value={view.id} text={view.name}>
								<span>
									{view.name}
									{view.isDefault && <span className={styles.defaultBadge}>Default</span>}
								</span>
							</Option>
						))}
					</OptionGroup>
				)}

				{!loading && !error && filteredPersonalViews.length > 0 && (
					<OptionGroup label="Personal Views">
						{filteredPersonalViews.map((view) => (
							<Option key={view.id} value={view.id} text={view.name}>
								{view.name}
							</Option>
						))}
					</OptionGroup>
				)}
			</Combobox>
		</div>
	);
}
