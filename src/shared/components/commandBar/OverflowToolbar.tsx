/**
 * The shared command bar (#232): a Fluent Toolbar inside Overflow. Commands that don't fit drop
 * from the end by priority into a "More commands" menu, keeping their icon, disabled reason and
 * submenu there. Never wraps.
 *
 * Layout: `start` | Toolbar [commands by group, dividers between groups, More, `placement: "end"`
 * commands] | `end`. `start` and `end` hold non-command content (a label, a count, a Switch) and sit
 * outside the overflow region.
 *
 * Disabled commands are `disabledFocusable` with their reason in a Tooltip, on the bar and in More,
 * so keyboard and screen reader users can reach them and learn why.
 */
import {
	Children,
	Fragment,
	cloneElement,
	forwardRef,
	useEffect,
	useRef,
	type MutableRefObject,
	type ReactElement,
	type ReactNode,
	type RefObject,
} from "react";
import {
	Button,
	DATA_OVERFLOWING,
	Menu,
	MenuButton,
	MenuDivider,
	MenuItem,
	MenuItemCheckbox,
	MenuItemRadio,
	MenuList,
	MenuPopover,
	MenuTrigger,
	Overflow,
	OverflowDivider,
	OverflowItem,
	Spinner,
	SplitButton,
	Toolbar,
	ToolbarButton,
	ToolbarDivider,
	ToolbarRadioButton,
	ToolbarToggleButton,
	Tooltip,
	makeStyles,
	mergeClasses,
	tokens,
	useIsOverflowGroupVisible,
	useMergedRefs,
	useOverflowMenu,
	useOverflowVisibility,
	type MenuButtonProps,
	type MenuProps,
} from "@fluentui/react-components";
import { MoreHorizontal20Filled, MoreHorizontal20Regular, bundleIcon } from "@fluentui/react-icons";
import type { BarCommand } from "./types";

const MoreIcon = bundleIcon(MoreHorizontal20Filled, MoreHorizontal20Regular);

/** Group id for commands without one, so dividers and More groups still line up */
const NO_GROUP = "default";
const groupOf = (command: BarCommand) => command.group ?? NO_GROUP;

export interface OverflowToolbarProps {
	/** The Toolbar's accessible name ("FetchXML commands") */
	ariaLabel: string;
	commands: BarCommand[];
	/** Toolbar padding; buttons stay medium */
	size?: "small" | "medium";
	/** Before the toolbar, never overflows (a label, a picker) */
	start?: ReactNode;
	/** After the toolbar, never overflows (status text, a count, a Switch); truncates when space is short */
	end?: ReactNode;
	/** More's label Tooltip and accessible name */
	moreLabel?: string;
	/** MenuItems always at the bottom of More (More then always shows) */
	moreExtra?: ReactNode;
	/** Space Overflow keeps free, in px. Defaults to the Toolbar's own inline padding plus a little. */
	padding?: number;
	className?: string;
}

const useStyles = makeStyles({
	bar: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		minWidth: 0,
	},
	start: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		flexShrink: 0,
	},
	// A zero basis keeps the toolbar's size independent of how many commands are showing, so
	// overflowing an item never changes the space Overflow measures.
	toolbar: {
		flex: "1 1 0",
		minWidth: 0,
		flexWrap: "nowrap",
		"& > *": { flexShrink: 0, whiteSpace: "nowrap" },
	},
	end: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		flex: "0 1 auto",
		minWidth: 0,
		maxWidth: "50%",
		overflow: "hidden",
		whiteSpace: "nowrap",
		textOverflow: "ellipsis",
		color: tokens.colorNeutralForeground3,
	},
	hidden: { display: "none" },
	/** Destructive command: danger text, still subtle (forced colours keep the system text colour) */
	danger: {
		color: tokens.colorStatusDangerForeground1,
		":hover": { color: tokens.colorStatusDangerForeground1 },
		":hover:active": { color: tokens.colorStatusDangerForeground1 },
		"@media (forced-colors: active)": { color: "ButtonText" },
	},
});

/**
 * The accessible name override, only when set: an explicit `aria-label={undefined}` would also wipe
 * the one a label Tooltip gives (Tooltip spreads the child's own props over its aria props).
 */
const nameOf = (command: BarCommand) => (command.ariaLabel ? { "aria-label": command.ariaLabel } : {});

/** Can't be pressed: it has a reason, or it's running */
const isOff = (command: BarCommand) => !!command.disabledReason || !!command.busy;

function iconOf(command: BarCommand): ReactElement | undefined {
	if (command.busy) return <Spinner size="tiny" />;
	const Icon = command.icon;
	return Icon ? <Icon /> : undefined;
}

/** The bar control's Tooltip: icon-only commands are labelled by it ("Label: reason"), others described */
function barTooltip(command: BarCommand): { content: string; relationship: "label" | "description" } | undefined {
	if (command.iconOnly) {
		return { content: command.disabledReason ? `${command.label}: ${command.disabledReason}` : command.label, relationship: "label" };
	}
	const tip = command.disabledReason ?? command.hint;
	return tip ? { content: tip, relationship: "description" } : undefined;
}

/** Toolbar / MenuList checkedValues for the toggles and radios among the commands */
function checkedValuesOf(commands: BarCommand[]): Record<string, string[]> {
	const values: Record<string, string[]> = {};
	for (const c of commands) {
		if (c.kind !== "toggle") continue;
		const name = c.radioName ?? c.id;
		values[name] ??= [];
		if (c.checked) values[name].push(c.id);
	}
	return values;
}

/**
 * Turns a checked-value change (from the Toolbar or More) into the command's onClick. Disabled
 * commands are ignored here too: ToolbarRadioButton calls the Toolbar's handler even when it's
 * disabledFocusable.
 */
function onCheckedChange(commands: BarCommand[], name: string, checkedItems: string[]): void {
	const radio = commands.filter((c) => c.kind === "toggle" && c.radioName === name);
	if (radio.length) {
		const picked = radio.find((c) => c.id === checkedItems[0]);
		if (picked && !picked.checked && !isOff(picked)) picked.onClick?.();
		return;
	}
	const toggle = commands.find((c) => c.kind === "toggle" && !c.radioName && c.id === name);
	if (toggle && !isOff(toggle) && checkedItems.includes(toggle.id) !== !!toggle.checked) toggle.onClick?.();
}

const menuOpenChange =
	(command: BarCommand): MenuProps["onOpenChange"] =>
	(_, data) =>
		command.onMenuOpenChange?.(data.open);

/** One command on the bar. The ref reaches the element Overflow measures (OverflowItem passes it). */
const BarControl = forwardRef<HTMLButtonElement, { command: BarCommand }>(function BarControl({ command: c }, ref) {
	const styles = useStyles();
	const off = isOff(c);
	const icon = iconOf(c);
	const className = c.danger && !off ? styles.danger : undefined;
	const tooltip = barTooltip(c);
	const children = c.iconOnly ? undefined : c.label;
	const appearance = c.appearance ?? "subtle";
	const withTooltip = (element: ReactElement) =>
		tooltip ? (
			<Tooltip content={tooltip.content} relationship={tooltip.relationship}>
				{element}
			</Tooltip>
		) : (
			element
		);
	const common = { icon, className, disabledFocusable: off, ...nameOf(c) };

	if (c.kind === "menu") {
		return (
			<Menu onOpenChange={menuOpenChange(c)}>
				<MenuTrigger disableButtonEnhancement>
					{withTooltip(
						<MenuButton ref={ref} appearance={appearance} {...common}>
							{children}
						</MenuButton>,
					)}
				</MenuTrigger>
				<MenuPopover>{c.menu}</MenuPopover>
			</Menu>
		);
	}

	if (c.kind === "split") {
		// The reason belongs on the primary action, which is the focusable part a user lands on first
		const split = (menuButton: MenuButtonProps, primary: Record<string, unknown> = {}) => (
			<SplitButton
				ref={ref}
				appearance={appearance}
				icon={icon}
				className={className}
				disabledFocusable={off}
				menuButton={menuButton}
				primaryActionButton={{ ...primary, ...nameOf(c), onClick: off ? undefined : c.onClick }}
			>
				{children}
			</SplitButton>
		);
		return (
			<Menu onOpenChange={menuOpenChange(c)}>
				<MenuTrigger disableButtonEnhancement>
					{(menuButton: MenuButtonProps) =>
						tooltip ? (
							<Tooltip content={tooltip.content} relationship={tooltip.relationship}>
								{(primary) => split(menuButton, primary as Record<string, unknown>)}
							</Tooltip>
						) : (
							split(menuButton)
						)
					}
				</MenuTrigger>
				<MenuPopover>{c.menu}</MenuPopover>
			</Menu>
		);
	}

	if (c.kind === "toggle") {
		// Toolbar toggles have no secondary appearance either
		const toggleAppearance = appearance === "secondary" ? "subtle" : appearance;
		return withTooltip(
			c.radioName ? (
				<ToolbarRadioButton ref={ref} name={c.radioName} value={c.id} appearance={toggleAppearance} {...common}>
					{children}
				</ToolbarRadioButton>
			) : (
				<ToolbarToggleButton ref={ref} name={c.id} value={c.id} appearance={toggleAppearance} {...common}>
					{children}
				</ToolbarToggleButton>
			),
		);
	}

	const onClick = off ? undefined : c.onClick;
	// ToolbarButton has no secondary appearance; a plain Button behaves the same inside the Toolbar
	return withTooltip(
		appearance === "secondary" ? (
			<Button ref={ref} appearance="secondary" onClick={onClick} {...common}>
				{children}
			</Button>
		) : (
			<ToolbarButton ref={ref} appearance={appearance} onClick={onClick} {...common}>
				{children}
			</ToolbarButton>
		),
	);
});

/** The divider before group `next`: Overflow hides it with `prev`; it also hides once `next` is all in More */
function BarDivider({ prev, next }: { prev: string; next: string }) {
	const styles = useStyles();
	const nextVisibility = useIsOverflowGroupVisible(next);
	return (
		<OverflowDivider groupId={prev}>
			<ToolbarDivider className={nextVisibility === "hidden" ? styles.hidden : undefined} />
		</OverflowDivider>
	);
}

/** A split's menu with its primary action first */
function splitMenu(command: BarCommand, icon: ReactElement | undefined): ReactElement {
	const primary = (
		<MenuItem key="primary-action" icon={icon} onClick={command.onClick}>
			{command.label}
		</MenuItem>
	);
	if (!command.menu) return <MenuList>{primary}</MenuList>;
	const menu = command.menu as ReactElement<{ children?: ReactNode }>;
	return cloneElement(menu, undefined, primary, ...Children.toArray(menu.props.children));
}

/**
 * One overflowed command in More: the same command as a menu entry. Menu items take `disabled`, which
 * in Fluent v9 is already focusable (aria-disabled); their `disabledFocusable` is deprecated and does
 * nothing.
 */
function MoreEntry({ command: c }: { command: BarCommand }) {
	const styles = useStyles();
	const off = isOff(c);
	const icon = iconOf(c);
	const className = c.danger && !off ? styles.danger : undefined;
	const withReason = (element: ReactElement) =>
		c.disabledReason ? (
			<Tooltip content={c.disabledReason} relationship="description" positioning="after">
				{element}
			</Tooltip>
		) : (
			element
		);

	if (c.kind === "menu" || c.kind === "split") {
		return (
			<Menu onOpenChange={menuOpenChange(c)}>
				<MenuTrigger disableButtonEnhancement>
					{withReason(
						<MenuItem icon={icon} className={className} disabled={off} {...nameOf(c)}>
							{c.label}
						</MenuItem>,
					)}
				</MenuTrigger>
				<MenuPopover>{c.kind === "split" ? splitMenu(c, c.busy ? undefined : icon) : c.menu}</MenuPopover>
			</Menu>
		);
	}
	if (c.kind === "toggle") {
		return withReason(
			c.radioName ? (
				<MenuItemRadio name={c.radioName} value={c.id} icon={icon} disabled={off} {...nameOf(c)}>
					{c.label}
				</MenuItemRadio>
			) : (
				<MenuItemCheckbox name={c.id} value={c.id} icon={icon} disabled={off} {...nameOf(c)}>
					{c.label}
				</MenuItemCheckbox>
			),
		);
	}
	return withReason(
		<MenuItem icon={icon} className={className} disabled={off} {...nameOf(c)} onClick={off ? undefined : c.onClick}>
			{c.label}
		</MenuItem>,
	);
}

function More({
	commands,
	extra,
	label,
	buttonRef,
	pendingFocus,
}: {
	commands: BarCommand[];
	extra?: ReactNode;
	label: string;
	buttonRef: RefObject<HTMLButtonElement | null>;
	pendingFocus: MutableRefObject<boolean>;
}) {
	const { ref, isOverflowing } = useOverflowMenu<HTMLButtonElement>();
	const { itemVisibility } = useOverflowVisibility();
	const mergedRef = useMergedRefs(ref, buttonRef);

	// The focused command overflowed before More existed: focus More once it's mounted
	useEffect(() => {
		if (isOverflowing && pendingFocus.current) {
			pendingFocus.current = false;
			buttonRef.current?.focus();
		}
	}, [isOverflowing, pendingFocus, buttonRef]);

	if (!isOverflowing && !extra) return null;
	// Only items Overflow has measured and hidden; before the first layout nothing is hidden
	const overflowed = isOverflowing ? commands.filter((c) => itemVisibility[c.id] === false) : [];
	return (
		<Menu>
			<MenuTrigger disableButtonEnhancement>
				<Tooltip content={label} relationship="label">
					<ToolbarButton ref={mergedRef} appearance="subtle" icon={<MoreIcon />} />
				</Tooltip>
			</MenuTrigger>
			<MenuPopover>
				<MenuList checkedValues={checkedValuesOf(overflowed)} onCheckedValueChange={(_, data) => onCheckedChange(overflowed, data.name, data.checkedItems)}>
					{overflowed.map((c, i) => (
						<Fragment key={c.id}>
							{i > 0 && groupOf(overflowed[i - 1]) !== groupOf(c) && <MenuDivider />}
							<MoreEntry command={c} />
						</Fragment>
					))}
					{overflowed.length > 0 && extra ? <MenuDivider /> : null}
					{extra}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
}

/**
 * A command bar that never wraps: lower priority commands move into More first, pinned and
 * `placement: "end"` commands never do. Keep each group's commands next to each other.
 */
export function OverflowToolbar({
	ariaLabel,
	commands,
	size = "medium",
	start,
	end,
	moreLabel = "More commands",
	moreExtra,
	padding,
	className,
}: OverflowToolbarProps) {
	const styles = useStyles();
	const toolbarRef = useRef<HTMLDivElement>(null);
	const moreRef = useRef<HTMLButtonElement>(null);
	const pendingFocus = useRef(false);
	// The last command focused in this toolbar. When it overflows, the browser may already have moved
	// focus to the body, so the focused element alone can't tell.
	const lastFocused = useRef<HTMLElement | null>(null);

	const barCommands = commands.filter((c) => c.placement !== "end");
	const endCommands = commands.filter((c) => c.placement === "end");

	const onOverflowChange = () => {
		const focused = lastFocused.current;
		if (!focused || !focused.closest(`[${DATA_OVERFLOWING}]`)) return;
		const active = focused.ownerDocument.activeElement;
		if (active !== focused && active !== focused.ownerDocument.body && active !== null) return;
		lastFocused.current = null;
		if (moreRef.current) moreRef.current.focus();
		else pendingFocus.current = true;
	};

	const renderItem = (c: BarCommand) => (
		<OverflowItem key={c.id} id={c.id} groupId={groupOf(c)} {...(c.pinned || c.placement === "end" ? { pinned: true } : { priority: c.priority ?? 0 })}>
			<BarControl command={c} />
		</OverflowItem>
	);

	return (
		<div className={mergeClasses(styles.bar, className)}>
			{start !== undefined && start !== null && <div className={styles.start}>{start}</div>}
			<Overflow padding={padding ?? (size === "small" ? 16 : 24)} onOverflowChange={onOverflowChange}>
				<Toolbar
					ref={toolbarRef}
					aria-label={ariaLabel}
					size={size}
					className={styles.toolbar}
					checkedValues={checkedValuesOf(commands)}
					onCheckedValueChange={(_, data) => onCheckedChange(commands, data.name, data.checkedItems)}
					onFocus={(event) => {
						lastFocused.current = event.target as HTMLElement;
					}}
					onBlur={(event) => {
						// Focus left for somewhere else on the page (not "nowhere", as when the element hides)
						const next = event.relatedTarget as Node | null;
						if (next && !toolbarRef.current?.contains(next)) lastFocused.current = null;
					}}
				>
					{barCommands.map((c, i) => (
						<Fragment key={c.id}>
							{i > 0 && groupOf(barCommands[i - 1]) !== groupOf(c) && <BarDivider prev={groupOf(barCommands[i - 1])} next={groupOf(c)} />}
							{renderItem(c)}
						</Fragment>
					))}
					<More commands={barCommands} extra={moreExtra} label={moreLabel} buttonRef={moreRef} pendingFocus={pendingFocus} />
					{endCommands.length > 0 && barCommands.length > 0 && groupOf(endCommands[0]) !== groupOf(barCommands[barCommands.length - 1]) && (
						<OverflowDivider groupId={groupOf(endCommands[0])}>
							<ToolbarDivider />
						</OverflowDivider>
					)}
					{endCommands.map(renderItem)}
				</Toolbar>
			</Overflow>
			{end !== undefined && end !== null && <div className={styles.end}>{end}</div>}
		</div>
	);
}
