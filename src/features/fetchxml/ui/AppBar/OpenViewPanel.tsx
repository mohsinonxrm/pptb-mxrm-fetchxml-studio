/**
 * Open view's popover content (#278, owner decision Q2): a search box over the entity's views, "My
 * views" above "System views". Each view shows its name, then "Default view", its query type or
 * "Shared by {owner}" (#142, #93); the view the query is gets the selected check. Mounted while the
 * popover is open, so the list is loaded fresh each time.
 *
 * Keyboard: the search box has focus first; Down moves into the list, where Up / Down move and Enter
 * opens a view (Fluent's Listbox).
 */
import { useEffect, useRef, useState, type KeyboardEvent } from "react";
import {
	Button,
	Listbox,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
	Option,
	OptionGroup,
	SearchBox,
	Spinner,
	Text,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import type { SavedView } from "../../api/pptbClient";
import { getViewSecondaryText } from "../../model/savedViews";
import { openViewGroups } from "./appBarRules";
import { useEntityViews } from "./useEntityViews";

const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
		width: "360px",
		maxWidth: "100%",
	},
	list: {
		maxHeight: "min(360px, 50vh)",
		overflowY: "auto",
		// Fluent's Listbox draws its own popup chrome; here it's part of the surface
		boxShadow: "none",
		padding: 0,
		minWidth: 0,
	},
	option: {
		display: "flex",
		flexDirection: "column",
		minWidth: 0,
	},
	secondary: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
	},
	status: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		color: tokens.colorNeutralForeground3,
		paddingBlock: tokens.spacingVerticalS,
	},
});

export interface OpenViewPanelProps {
	entityLogicalName: string;
	/** The entity's display name, for the empty state */
	entityDisplayName: string;
	/** The view the query is, checked in the list */
	currentViewId?: string;
	onPick: (view: SavedView) => void;
}

export function OpenViewPanel({
	entityLogicalName,
	entityDisplayName,
	currentViewId,
	onPick,
}: OpenViewPanelProps) {
	const styles = useStyles();
	const views = useEntityViews(entityLogicalName);
	const [search, setSearch] = useState("");
	const listRef = useRef<HTMLDivElement>(null);
	const searchRef = useRef<HTMLInputElement>(null);
	// The search box takes focus when the popover opens, so typing filters straight away
	useEffect(() => searchRef.current?.focus(), []);
	const { mine, system } = openViewGroups(views.systemViews, views.personalViews, search);
	const all = [...views.personalViews, ...views.systemViews];
	const hasViews = all.length > 0;

	const onSearchKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
		if (event.key === "ArrowDown" && listRef.current) {
			event.preventDefault();
			listRef.current.focus();
		}
	};

	const option = (view: SavedView) => {
		const secondary = getViewSecondaryText(view, views.currentUserId);
		return (
			<Option key={view.id} value={view.id} text={view.name}>
				<span className={styles.option}>
					<span>{view.name}</span>
					{secondary && <span className={styles.secondary}>{secondary}</span>}
				</span>
			</Option>
		);
	};

	let body;
	if (views.loading && !hasViews) {
		body = (
			<div className={styles.status} role="status">
				<Spinner size="tiny" />
				<span>Loading views…</span>
			</div>
		);
	} else if (views.error) {
		body = (
			<MessageBar intent="error">
				<MessageBarBody>Couldn't load the views. {views.error}</MessageBarBody>
				<MessageBarActions>
					<Button size="small" onClick={views.reload}>
						Try again
					</Button>
				</MessageBarActions>
			</MessageBar>
		);
	} else if (!hasViews) {
		body = (
			<Text className={styles.status}>{entityDisplayName} has no views you can open.</Text>
		);
	} else if (mine.length === 0 && system.length === 0) {
		body = (
			<Text className={styles.status} role="status">
				No views match "{search.trim()}".
			</Text>
		);
	} else {
		body = (
			<Listbox
				ref={listRef}
				className={styles.list}
				aria-label={`Views of ${entityDisplayName}`}
				selectedOptions={currentViewId ? [currentViewId] : []}
				onOptionSelect={(_, data) => {
					const view = all.find((v) => v.id === data.optionValue);
					if (view) onPick(view);
				}}
			>
				{mine.length > 0 && <OptionGroup label="My views">{mine.map(option)}</OptionGroup>}
				{system.length > 0 && (
					<OptionGroup label="System views">{system.map(option)}</OptionGroup>
				)}
			</Listbox>
		);
	}

	return (
		<div className={styles.root}>
			<SearchBox
				ref={searchRef}
				placeholder="Search views"
				aria-label="Search views"
				value={search}
				onChange={(_, data) => setSearch(data.value)}
				onKeyDown={onSearchKeyDown}
			/>
			{body}
		</div>
	);
}
