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
	toLoadedViewInfo,
	whoAmI,
	type SavedView,
	type EntityMetadata,
	type LoadedViewInfo,
} from "../../api/pptbClient";
import { debugLog } from "../../../../shared/utils/debug";
import { getViewSecondaryText } from "../../model/savedViews";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: "4px",
		flex: 1,
		// Shrinks with a narrow left pane (#53)
		minWidth: 0,
		maxWidth: "350px",
	},
	label: {
		fontSize: "14px",
		fontWeight: 600,
		color: tokens.colorNeutralForeground2,
	},
	combobox: {
		width: "100%",
		// Overrides Fluent's built-in Combobox minimum width (#53)
		minWidth: 0,
	},
	comboboxInput: {
		minWidth: 0,
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
	// Two-line options (#142), like the relationship picker (#25)
	option: {
		display: "flex",
		flexDirection: "column",
		minWidth: 0,
	},
	optionSecondary: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
	},
});

/**
 * A view option: its name, then "Shared by {owner}" for a personal view someone else owns, or the
 * query type / "Default view" for system views (#142, savedViews.ts getViewSecondaryText). Every
 * query type is listed rather than hidden (#93).
 */
function ViewOption({ view, currentUserId }: { view: SavedView; currentUserId: string | undefined }) {
	const styles = useStyles();
	const secondary = getViewSecondaryText(view, currentUserId);
	return (
		<Option value={view.id} text={view.name}>
			<span className={styles.option}>
				<span>{view.name}</span>
				{secondary && <span className={styles.optionSecondary}>{secondary}</span>}
			</span>
		</Option>
	);
}

interface LoadViewPickerProps {
	/** Currently selected entity metadata (with ObjectTypeCode) */
	selectedEntityMetadata: EntityMetadata | null;
	/** Callback when a view is selected - provides the view info for execution optimization */
	onViewSelect: (viewInfo: LoadedViewInfo) => void;
	/** Callback when the view selection is cleared */
	onViewClear?: () => void;
	/** Optional: disable the picker */
	disabled?: boolean;
	/** The name of the view the picker shows, or null (#25: the collapsed scope summary) */
	onSelectedViewChange?: (viewName: string | null) => void;
}

export function LoadViewPicker({
	selectedEntityMetadata,
	onViewSelect,
	onViewClear,
	disabled = false,
	onSelectedViewChange,
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

	// The connected user, to tell personal views shared with them from their own (#142). WhoAmI is
	// memoized per connection (sessionCache.ts).
	const [currentUserId, setCurrentUserId] = useState<string | undefined>(undefined);

	useEffect(() => {
		onSelectedViewChange?.(selectedViewName);
	}, [selectedViewName, onSelectedViewChange]);

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

				const [{ systemViews: sysViews, personalViews: persViews }, user] = await Promise.all([
					getAllViews(selectedEntityMetadata.LogicalName),
					whoAmI(),
				]);

				setSystemViews(sysViews);
				setPersonalViews(persViews);
				setCurrentUserId(user?.UserId);

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
			// Handle clear action - when clearable button is clicked, optionValue is empty
			if (!data.optionValue) {
				setSelectedViewName(null);
				setInputValue("");
				onViewClear?.();
				debugLog("viewAPI", "View selection cleared");
				return;
			}

			if (!selectedEntityMetadata) return;

			const view = allViews.find((v) => v.id === data.optionValue);
			if (view) {
				// Everything execution by ID, the layout and Save view need (shared with #197's prefill)
				const viewInfo: LoadedViewInfo = toLoadedViewInfo(view, selectedEntityMetadata.EntitySetName);
				onViewSelect(viewInfo);
				debugLog("viewAPI", `Selected view: ${view.name} (${view.type}) - ID: ${view.id}`);
				// Set the selected view name to display in the combobox
				setSelectedViewName(view.name);
				setInputValue(view.name);
			}
		},
		[allViews, onViewSelect, onViewClear, selectedEntityMetadata]
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
			<label className={styles.label}>Load view</label>
			<Combobox
				className={styles.combobox}
				input={{ className: styles.comboboxInput }}
				placeholder={loading ? "Loading views…" : "Select a view to load…"}
				disabled={disabled || loading || !hasViews}
				value={inputValue}
				onChange={handleInputChange}
				onOptionSelect={handleOptionSelect}
				clearable={!!selectedViewName}
				freeform
			>
				{loading && (
					<Option key="loading" value="" text="Loading…" disabled>
						<div className={styles.loadingOption}>
							<Spinner size="tiny" />
							<span>Loading views…</span>
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
						<span className={styles.noViewsOption}>No views match "{filterQuery.trim()}"</span>
					</Option>
				)}

				{!loading && !error && filteredSystemViews.length > 0 && (
					<OptionGroup label="System views">
						{filteredSystemViews.map((view) => (
							<ViewOption key={view.id} view={view} currentUserId={currentUserId} />
						))}
					</OptionGroup>
				)}

				{/* One group: views shared with the user say "Shared by {owner}" on a second line (#142) */}
				{!loading && !error && filteredPersonalViews.length > 0 && (
					<OptionGroup label="Personal views">
						{filteredPersonalViews.map((view) => (
							<ViewOption key={view.id} view={view} currentUserId={currentUserId} />
						))}
					</OptionGroup>
				)}
			</Combobox>
		</div>
	);
}
