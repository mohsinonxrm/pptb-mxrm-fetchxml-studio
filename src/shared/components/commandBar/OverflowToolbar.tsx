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
	createContext,
	forwardRef,
	useContext,
	useEffect,
	useLayoutEffect,
	useRef,
	useState,
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
	Popover,
	PopoverSurface,
	PopoverTrigger,
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
	useIsOverflowItemVisible,
	useMergedRefs,
	useOverflowMenu,
	useOverflowVisibility,
	useRestoreFocusTarget,
	type MenuButtonProps,
	type MenuProps,
} from "@fluentui/react-components";
import { MoreHorizontal20Filled, MoreHorizontal20Regular, bundleIcon } from "@fluentui/react-icons";
import type { BarCommand } from "./types";
import { withReason } from "../WithReason";
import { StableTooltip } from "../StableTooltip";

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
	/**
	 * Plain MenuItems (actions) always at the bottom of More; More then always shows. More's MenuList
	 * controls checked state for the overflowed toggles only, so a checkbox or radio belongs in
	 * `commands` as a toggle, not here.
	 */
	moreExtra?: ReactNode;
	/** Space Overflow keeps free, in px. Defaults to the Toolbar's own inline padding plus a little. */
	padding?: number;
	/** Where the commands sit when there's room to spare. Default "start". */
	align?: "start" | "end";
	/**
	 * Put the `placement: "end"` commands at the far end of the bar, after the room to spare, as in a
	 * model-driven app's command bar (Settings at the right). Give them a group of their own: the
	 * divider before them is what moves.
	 */
	endAtEdge?: boolean;
	/**
	 * Every group the bar may show, including ones whose commands appear later (after a privilege
	 * check, say). Overflow learns them up front, so a late group doesn't start a fresh Overflow,
	 * which would close an open menu of the bar.
	 */
	groups?: string[];
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
		color: tokens.colorNeutralForeground3,
	},
	// Ellipsis needs a block-level box: a flex item, not the flex container itself
	endText: {
		minWidth: 0,
		overflow: "hidden",
		whiteSpace: "nowrap",
		textOverflow: "ellipsis",
	},
	alignEnd: { justifyContent: "flex-end" },
	// Takes the room to spare, so what follows sits at the far end; Overflow measures widths, not
	// margins, so it doesn't change what fits
	pushToEnd: { marginInlineStart: "auto" },
	hidden: { display: "none" },
	// A popover command's surface: its content sets the width; this keeps it inside small windows
	popoverSurface: {
		padding: tokens.spacingHorizontalM,
		maxWidth: "min(420px, calc(100vw - 32px))",
	},
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
const nameOf = (command: BarCommand) =>
	command.ariaLabel ? { "aria-label": command.ariaLabel } : {};

/** On the bar an icon-only command is named "Label: reason"; an `ariaLabel` replaces the label part only */
const barNameOf = (command: BarCommand) => ({
	...(command.ariaLabel && command.iconOnly && command.disabledReason
		? { "aria-label": `${command.ariaLabel}: ${command.disabledReason}` }
		: nameOf(command)),
	...(command.keyShortcuts ? { "aria-keyshortcuts": command.keyShortcuts } : {}),
});

/**
 * Checked-value names for the Toolbar and More's MenuList, namespaced so a toggle's id and a radio
 * group's name can't collide
 */
const checkedNameOf = (command: BarCommand) =>
	command.radioName ? `radio:${command.radioName}` : `toggle:${command.id}`;

/** Can't be pressed: it has a reason, or it's running */
const isOff = (command: BarCommand) => !!command.disabledReason || !!command.busy;

function iconOf(command: BarCommand): ReactElement | undefined {
	if (command.busy) return <Spinner size="tiny" />;
	const Icon = command.icon;
	return Icon ? <Icon /> : undefined;
}

/** The bar control's Tooltip: icon-only commands are labelled by it ("Label: reason"), others described */
function barTooltip(
	command: BarCommand,
): { content: string; relationship: "label" | "description" } | undefined {
	if (command.iconOnly) {
		return {
			content: command.disabledReason
				? `${command.label}: ${command.disabledReason}`
				: command.label,
			relationship: "label",
		};
	}
	const tip = command.disabledReason ?? command.hint;
	return tip ? { content: tip, relationship: "description" } : undefined;
}

/** Toolbar / MenuList checkedValues for the toggles and radios among the commands */
function checkedValuesOf(commands: BarCommand[]): Record<string, string[]> {
	const values: Record<string, string[]> = {};
	for (const c of commands) {
		if (c.kind !== "toggle") continue;
		const name = checkedNameOf(c);
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
	const named = commands.filter((c) => c.kind === "toggle" && checkedNameOf(c) === name);
	if (named.some((c) => c.radioName)) {
		const picked = named.find((c) => c.id === checkedItems[0]);
		if (picked && !picked.checked && !isOff(picked)) picked.onClick?.();
		return;
	}
	const toggle = named[0];
	if (toggle && !isOff(toggle) && checkedItems.includes(toggle.id) !== !!toggle.checked)
		toggle.onClick?.();
}

const menuOpenChange =
	(command: BarCommand): MenuProps["onOpenChange"] =>
	(_, data) =>
		command.onMenuOpenChange?.(data.open);

/** The More button: a popover command that overflowed opens anchored to it */
const MoreAnchorContext = createContext<RefObject<HTMLButtonElement | null> | null>(null);

/**
 * A popover command on the bar: a MenuButton (its chevron says something opens) as the Popover's
 * trigger. Once the command has overflowed, More opens the same Popover, anchored to More instead.
 */
const PopoverControl = forwardRef<
	HTMLButtonElement,
	{ command: BarCommand; common: Record<string, unknown>; children?: ReactNode }
>(function PopoverControl({ command: c, common, children }, ref) {
	const styles = useStyles();
	const visible = useIsOverflowItemVisible(c.id);
	const moreRef = useContext(MoreAnchorContext);
	const off = isOff(c);
	const tooltip = barTooltip(c);
	// Becoming unavailable closes it; say so, or it would reopen by itself once available again
	const { popoverOpen, onPopoverOpenChange } = c;
	useEffect(() => {
		if (off && popoverOpen) onPopoverOpenChange?.(false);
	}, [off, popoverOpen, onPopoverOpenChange]);
	const button = (
		<MenuButton
			ref={ref}
			appearance={c.appearance ?? "subtle"}
			aria-haspopup="dialog"
			{...common}
		>
			{children}
		</MenuButton>
	);
	return (
		<Popover
			open={!!c.popoverOpen && !off}
			onOpenChange={(_, data) => c.onPopoverOpenChange?.(data.open)}
			trapFocus
			positioning={
				visible || !moreRef?.current
					? { position: "below", align: "start" }
					: { target: moreRef.current, position: "below", align: "end" }
			}
		>
			<PopoverTrigger disableButtonEnhancement>
				<StableTooltip content={tooltip?.content} relationship={tooltip?.relationship}>
					{button}
				</StableTooltip>
			</PopoverTrigger>
			<PopoverSurface aria-label={c.label} className={styles.popoverSurface}>
				{c.popover}
			</PopoverSurface>
		</Popover>
	);
});

/** One command on the bar. The ref reaches the element Overflow measures (OverflowItem passes it). */
const BarControl = forwardRef<HTMLButtonElement, { command: BarCommand }>(function BarControl(
	{ command: c },
	ref,
) {
	const styles = useStyles();
	// Closing a dialog or drawer a command opened returns focus to that command (Fluent restores it
	// to the most recently focused target). One that overflowed meanwhile hides, so More takes over.
	const restoreFocus = useRestoreFocusTarget();
	const off = isOff(c);
	const icon = iconOf(c);
	const className = c.danger && !off ? styles.danger : undefined;
	const tooltip = barTooltip(c);
	const children = c.iconOnly ? undefined : c.label;
	const appearance = c.appearance ?? "subtle";
	// Always the same wrapper, so a command keeps its element when it gains or loses a hint or a
	// reason: a dialog it opened returns focus to it (#393)
	const withTooltip = (element: ReactElement) => (
		<StableTooltip content={tooltip?.content} relationship={tooltip?.relationship}>
			{element}
		</StableTooltip>
	);
	const common = { icon, className, disabledFocusable: off, ...restoreFocus, ...barNameOf(c) };

	if (c.kind === "popover") {
		return (
			<PopoverControl ref={ref} command={c} common={common}>
				{children}
			</PopoverControl>
		);
	}

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
				// The halves set their own colour, so danger goes on each of them
				menuButton={{
					...menuButton,
					className: mergeClasses(menuButton.className, className),
				}}
				primaryActionButton={{
					className,
					...primary,
					...restoreFocus,
					...barNameOf(c),
					onClick: off ? undefined : c.onClick,
				}}
			>
				{children}
			</SplitButton>
		);
		return (
			<Menu onOpenChange={menuOpenChange(c)}>
				<MenuTrigger disableButtonEnhancement>
					{(menuButton: MenuButtonProps) => (
						<StableTooltip
							content={tooltip?.content}
							relationship={tooltip?.relationship}
						>
							{(primary) => split(menuButton, primary as Record<string, unknown>)}
						</StableTooltip>
					)}
				</MenuTrigger>
				<MenuPopover>{c.menu}</MenuPopover>
			</Menu>
		);
	}

	if (c.kind === "toggle") {
		// Toolbar toggles have no secondary appearance either. They take the Toolbar's size unless told,
		// and the bar's buttons stay medium whatever its size.
		const toggleAppearance = appearance === "secondary" ? "subtle" : appearance;
		return withTooltip(
			c.radioName ? (
				<ToolbarRadioButton
					ref={ref}
					name={checkedNameOf(c)}
					value={c.id}
					appearance={toggleAppearance}
					size="medium"
					{...common}
				>
					{children}
				</ToolbarRadioButton>
			) : (
				<ToolbarToggleButton
					ref={ref}
					name={checkedNameOf(c)}
					value={c.id}
					appearance={toggleAppearance}
					size="medium"
					{...common}
				>
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

/**
 * The tree shape a command renders: a change here swaps the element under OverflowItem (Button for
 * ToolbarButton, radio for toggle). A Tooltip is always there (StableTooltip, #393), so gaining or
 * losing a hint or a reason isn't a new shape. Fluent's OverflowItem registers its
 * element only when id, priority, groupId or pinned change, so it would keep measuring the old,
 * detached element. Keying the OverflowItem by shape remounts it, and it registers the new one.
 */
const shapeOf = (command: BarCommand) =>
	[
		command.kind,
		command.kind === "button" && command.appearance === "secondary" ? "secondary" : "",
		command.radioName ? "radio" : "",
	].join(".");

/**
 * The divider after group `prev`. Overflow hides it with `prev`; it also hides when every group after
 * it is all in More. So with groups a | b | c and b all in More, this one (after a) stays between a
 * and c, and the one after b hides with b.
 */
function BarDivider({ prev, after }: { prev: string; after: string[] }) {
	const styles = useStyles();
	const { groupVisibility } = useOverflowVisibility();
	const nothingAfter = after.every((g) => groupVisibility[g] === "hidden");
	return (
		<OverflowDivider groupId={prev}>
			<ToolbarDivider className={nothingAfter ? styles.hidden : undefined} />
		</OverflowDivider>
	);
}

/**
 * Registers each group with the Overflow before it starts observing, then goes: Fluent's overflow
 * manager keeps a group once it has had an item, so an item of that group registering later is
 * safe (see the remount note in OverflowToolbar). Its placeholder items register in their layout
 * effects, which run before the Overflow's own (where it starts observing), and are gone again
 * before the browser paints.
 */
function GroupPrimer({ groups }: { groups: string[] }) {
	const styles = useStyles();
	const [primed, setPrimed] = useState(false);
	useLayoutEffect(() => setPrimed(true), []);
	if (primed) return null;
	return groups.map((groupId) => (
		<OverflowItem key={groupId} id={`__group:${groupId}`} groupId={groupId} priority={-1}>
			<span className={styles.hidden} aria-hidden />
		</OverflowItem>
	));
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
	const reasoned = (element: ReactElement) => withReason(element, c.disabledReason, "after");

	if (c.kind === "menu" || c.kind === "split") {
		return (
			<Menu onOpenChange={menuOpenChange(c)}>
				<MenuTrigger disableButtonEnhancement>
					{reasoned(
						<MenuItem icon={icon} className={className} disabled={off} {...nameOf(c)}>
							{c.label}
						</MenuItem>,
					)}
				</MenuTrigger>
				<MenuPopover>
					{c.kind === "split" && !c.menuHasPrimary
						? splitMenu(c, c.busy ? undefined : icon)
						: c.menu}
				</MenuPopover>
			</Menu>
		);
	}
	if (c.kind === "toggle") {
		return reasoned(
			c.radioName ? (
				<MenuItemRadio
					name={checkedNameOf(c)}
					value={c.id}
					icon={icon}
					disabled={off}
					{...nameOf(c)}
				>
					{c.label}
				</MenuItemRadio>
			) : (
				<MenuItemCheckbox
					name={checkedNameOf(c)}
					value={c.id}
					icon={icon}
					disabled={off}
					{...nameOf(c)}
				>
					{c.label}
				</MenuItemCheckbox>
			),
		);
	}
	// A popover opens once More's menu has closed and handed focus back to More, so the popover keeps
	// the focus it moves into itself
	const onClick =
		c.kind === "popover" ? () => setTimeout(() => c.onPopoverOpenChange?.(true), 0) : c.onClick;
	return reasoned(
		<MenuItem
			icon={icon}
			className={className}
			disabled={off}
			{...(c.kind === "popover" ? { "aria-haspopup": "dialog" as const } : {})}
			{...nameOf(c)}
			onClick={off ? undefined : onClick}
		>
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
	// The fallback target: a command chosen from More is gone once the menu closes, so focus returns here
	const restoreFocus = useRestoreFocusTarget();
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
					<ToolbarButton
						ref={mergedRef}
						appearance="subtle"
						icon={<MoreIcon />}
						{...restoreFocus}
					/>
				</Tooltip>
			</MenuTrigger>
			<MenuPopover>
				<MenuList
					checkedValues={checkedValuesOf(overflowed)}
					onCheckedValueChange={(_, data) =>
						onCheckedChange(overflowed, data.name, data.checkedItems)
					}
				>
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
	align = "start",
	endAtEdge = false,
	groups: declaredGroups,
	className,
}: OverflowToolbarProps) {
	const styles = useStyles();
	const toolbarRef = useRef<HTMLDivElement>(null);
	const moreRef = useRef<HTMLButtonElement>(null);
	const pendingFocus = useRef(false);
	// Each command's element on the bar (what Overflow measures), by command id
	const elements = useRef(new Map<string, HTMLElement>());
	// The last element focused in this toolbar, and its command. When it overflows, the browser may
	// already have moved focus to the body, so the focused element alone can't tell.
	const lastFocused = useRef<{ element: HTMLElement; id?: string } | null>(null);

	const barCommands = commands.filter((c) => c.placement !== "end");
	const endCommands = commands.filter((c) => c.placement === "end");
	const groups = barCommands.map(groupOf).filter((g, i, all) => all.indexOf(g) === i);

	// Fluent's overflow manager crashes when an item of a group it has never had registers after it
	// started observing (@fluentui/priority-overflow 9.4: it shows or hides the item before it adds
	// the item's group). A group it had keeps existing after its items go, so only a group new to
	// this Overflow needs a fresh one: Save coming back with its tab doesn't remount, tools found
	// after the first layout do, once. A remount closes an open menu of this bar.
	const known = useRef<{ groups: Set<string>; generation: number }>({
		groups: new Set(),
		generation: 0,
	});
	const allGroups = [...groups, ...endCommands.map(groupOf)];
	if (allGroups.some((g) => !known.current.groups.has(g))) {
		// The fresh Overflow learns the groups present now, and the declared ones (GroupPrimer)
		known.current = {
			groups: new Set([...allGroups, ...(declaredGroups ?? [])]),
			generation: known.current.generation + 1,
		};
	}

	const focusedIsGone = (element: HTMLElement) => {
		const active = element.ownerDocument.activeElement;
		return active === element || active === element.ownerDocument.body || active === null;
	};

	const onOverflowChange = () => {
		const focused = lastFocused.current?.element;
		if (!focused || !focused.closest(`[${DATA_OVERFLOWING}]`) || !focusedIsGone(focused))
			return;
		lastFocused.current = null;
		if (moreRef.current) moreRef.current.focus();
		else pendingFocus.current = true;
	};

	// A command whose shape changed (see shapeOf) is a new element: if the old one had focus, give
	// focus to the new one rather than leave it on the body
	useLayoutEffect(() => {
		const focused = lastFocused.current;
		if (!focused?.id || focused.element.isConnected) return;
		const next = elements.current.get(focused.id);
		if (!next || !focusedIsGone(focused.element)) {
			lastFocused.current = null;
			return;
		}
		// A SplitButton's element is its root; focus its primary action
		(next.matches("button") ? next : next.querySelector("button"))?.focus();
	});

	const renderItem = (c: BarCommand) => (
		<OverflowItem
			key={`${c.id}|${shapeOf(c)}`}
			ref={(element: HTMLElement | null) => {
				// React detaches the old ref before it attaches the new one in a commit
				if (element) elements.current.set(c.id, element);
				else elements.current.delete(c.id);
			}}
			id={c.id}
			groupId={groupOf(c)}
			{...(c.pinned || c.placement === "end"
				? { pinned: true }
				: { priority: c.priority ?? 0 })}
		>
			<BarControl command={c} />
		</OverflowItem>
	);

	return (
		<div className={mergeClasses(styles.bar, className)}>
			{start !== undefined && start !== null && <div className={styles.start}>{start}</div>}
			{/* With moreExtra, More always shows, so Overflow must keep room for it (hasHiddenItems) */}
			<Overflow
				key={known.current.generation}
				padding={padding ?? (size === "small" ? 16 : 24)}
				hasHiddenItems={!!moreExtra}
				onOverflowChange={onOverflowChange}
			>
				<Toolbar
					ref={toolbarRef}
					aria-label={ariaLabel}
					size={size}
					className={mergeClasses(styles.toolbar, align === "end" && styles.alignEnd)}
					checkedValues={checkedValuesOf(commands)}
					onCheckedValueChange={(_, data) =>
						onCheckedChange(commands, data.name, data.checkedItems)
					}
					onFocus={(event) => {
						const element = event.target as HTMLElement;
						const id = [...elements.current].find(([, item]) =>
							item.contains(element),
						)?.[0];
						lastFocused.current = { element, id };
					}}
					onBlur={(event) => {
						const element = event.target as HTMLElement;
						const next = event.relatedTarget as Node | null;
						// Focus went elsewhere on the page, or to nowhere while the element is still showing (a
						// click on empty space). Not when it hid (overflowed) or was replaced (shape change).
						const left = next
							? !toolbarRef.current?.contains(next)
							: element.isConnected && !element.closest(`[${DATA_OVERFLOWING}]`);
						if (left && lastFocused.current?.element === element)
							lastFocused.current = null;
					}}
				>
					<GroupPrimer groups={[...known.current.groups]} />
					<MoreAnchorContext.Provider value={moreRef}>
						{barCommands.map((c, i) => (
							<Fragment key={c.id}>
								{i > 0 && groupOf(barCommands[i - 1]) !== groupOf(c) && (
									<BarDivider
										prev={groupOf(barCommands[i - 1])}
										after={groups.slice(groups.indexOf(groupOf(c)))}
									/>
								)}
								{renderItem(c)}
							</Fragment>
						))}
					</MoreAnchorContext.Provider>
					<More
						commands={barCommands}
						extra={moreExtra}
						label={moreLabel}
						buttonRef={moreRef}
						pendingFocus={pendingFocus}
					/>
					{endCommands.length > 0 &&
						barCommands.length > 0 &&
						groupOf(endCommands[0]) !==
							groupOf(barCommands[barCommands.length - 1]) && (
							<OverflowDivider groupId={groupOf(endCommands[0])}>
								<ToolbarDivider
									className={endAtEdge ? styles.pushToEnd : undefined}
								/>
							</OverflowDivider>
						)}
					{endCommands.map(renderItem)}
				</Toolbar>
			</Overflow>
			{end !== undefined && end !== null && (
				<div className={styles.end}>
					<span className={styles.endText}>{end}</span>
				</div>
			)}
		</div>
	);
}
