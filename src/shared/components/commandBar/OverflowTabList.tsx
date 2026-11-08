/**
 * The shared tab row (#232): a Fluent TabList inside Overflow, following Fluent's TabList overflow
 * pattern. Tabs that don't fit move into a "More" menu; the selected tab gets a higher priority so it
 * always stays visible, and choosing a tab from More selects it. Never wraps.
 */
import {
	Button,
	Menu,
	MenuItem,
	MenuList,
	MenuPopover,
	MenuTrigger,
	Overflow,
	OverflowItem,
	Tab,
	TabList,
	makeStyles,
	mergeClasses,
	useOverflowMenu,
	useOverflowVisibility,
	type TabListProps,
} from "@fluentui/react-components";
import { MoreHorizontal20Filled, MoreHorizontal20Regular, bundleIcon } from "@fluentui/react-icons";
import type { BarTab } from "./types";

const MoreIcon = bundleIcon(MoreHorizontal20Filled, MoreHorizontal20Regular);

/** The selected tab outranks the rest, so it's the last to overflow */
const SELECTED_PRIORITY = 2;
const TAB_PRIORITY = 1;

export interface OverflowTabListProps {
	/** The TabList's accessible name ("Query views") */
	ariaLabel: string;
	tabs: BarTab[];
	selectedValue: string;
	onTabSelect: (value: string) => void;
	size?: TabListProps["size"];
	appearance?: TabListProps["appearance"];
	/** Space Overflow keeps free, in px (Fluent's default when omitted) */
	padding?: number;
	className?: string;
}

const useStyles = makeStyles({
	tabList: {
		minWidth: 0,
		flexWrap: "nowrap",
		"& > *": { flexShrink: 0, whiteSpace: "nowrap" },
	},
});

function MoreTabs({
	tabs,
	onTabSelect,
	size,
}: {
	tabs: BarTab[];
	onTabSelect: (value: string) => void;
	size: TabListProps["size"];
}) {
	const { ref, isOverflowing, overflowCount } = useOverflowMenu<HTMLButtonElement>();
	const { itemVisibility } = useOverflowVisibility();
	if (!isOverflowing) return null;
	return (
		<Menu>
			<MenuTrigger disableButtonEnhancement>
				<Button
					ref={ref}
					appearance="transparent"
					size={size}
					icon={<MoreIcon />}
					role="tab"
					aria-label={`More tabs (${overflowCount})`}
				>
					More
				</Button>
			</MenuTrigger>
			<MenuPopover>
				<MenuList>
					{tabs
						.filter((t) => itemVisibility[t.value] === false)
						.map((t) => {
							const Icon = t.icon;
							return (
								<MenuItem
									key={t.value}
									icon={Icon ? <Icon /> : undefined}
									onClick={() => onTabSelect(t.value)}
								>
									{t.label}
								</MenuItem>
							);
						})}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
}

export function OverflowTabList({
	ariaLabel,
	tabs,
	selectedValue,
	onTabSelect,
	size = "medium",
	appearance,
	padding,
	className,
}: OverflowTabListProps) {
	const styles = useStyles();
	return (
		<Overflow minimumVisible={1} padding={padding}>
			<TabList
				aria-label={ariaLabel}
				size={size}
				appearance={appearance}
				selectedValue={selectedValue}
				onTabSelect={(_, data) => onTabSelect(String(data.value))}
				className={mergeClasses(styles.tabList, className)}
			>
				{tabs.map((t) => {
					const Icon = t.icon;
					return (
						<OverflowItem
							key={t.value}
							id={t.value}
							priority={t.value === selectedValue ? SELECTED_PRIORITY : TAB_PRIORITY}
						>
							<Tab value={t.value} icon={Icon ? <Icon /> : undefined}>
								{t.label}
								{t.content}
							</Tab>
						</OverflowItem>
					);
				})}
				<MoreTabs tabs={tabs} onTabSelect={onTabSelect} size={size} />
			</TabList>
		</Overflow>
	);
}
