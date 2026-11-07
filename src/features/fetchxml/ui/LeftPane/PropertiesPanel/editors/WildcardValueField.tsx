/**
 * Value field for like / not-like / begins-with / ends-with and their negations (#98): a Match
 * choice that writes the like pattern (Contains %x%, Begins with x%, Ends with %x, or a custom
 * pattern), and Match literally, which escapes % _ [ so they're matched as characters.
 * https://learn.microsoft.com/power-apps/developer/data-platform/wildcard-characters?tabs=fetchxml
 *
 * The Query Advisor's leading-wildcard advice (#97) still applies to the value written: Contains and
 * Ends with start with %, so they show it. It's advice, never a block.
 */

import { useEffect, useRef, useState } from "react";
import { Checkbox, Input, Radio, RadioGroup, makeStyles } from "@fluentui/react-components";
import {
	LIKE_OPERATORS,
	buildWildcardValue,
	hasWildcard,
	parseWildcardValue,
	type LikeMatch,
	type WildcardValue,
} from "../../../../model/wildcards";
import { DOCS } from "../docLinks";
import { Info, InfoField, type Validation } from "../fields";
import { worst } from "../ruleFindings";

const useStyles = makeStyles({
	// Four choices don't fit a narrow Properties panel on one line
	matches: { flexWrap: "wrap" },
});

const asText = (value: unknown): string | undefined =>
	value === undefined || value === null || value === "" ? undefined : String(value);

interface WildcardValueFieldProps {
	operator: string;
	value: unknown;
	/** Query Advisor advice on the value (leading wildcards, #97) */
	advice?: Validation;
	onChange: (value: string | undefined) => void;
}

export function WildcardValueField({ operator, value, advice, onChange }: WildcardValueFieldProps) {
	const styles = useStyles();
	const isLike = LIKE_OPERATORS.has(operator);
	const [state, setState] = useState<WildcardValue>(() => parseWildcardValue(operator, value));
	// The value this field last wrote: a different value (Parse to Tree, undo, an operator change)
	// is read back into the choices; the field's own writes keep its choices (e.g. Match literally
	// on for text without wildcard characters, which writes the same value as off)
	const written = useRef(asText(value));
	const operatorRef = useRef(operator);
	useEffect(() => {
		if (asText(value) === written.current && operator === operatorRef.current) return;
		written.current = asText(value);
		operatorRef.current = operator;
		setState(parseWildcardValue(operator, value));
	}, [operator, value]);

	const update = (next: Partial<WildcardValue>) => {
		const merged = { ...state, ...next };
		setState(merged);
		const built = merged.text === "" ? undefined : buildWildcardValue(operator, merged);
		written.current = built;
		onChange(built);
	};

	// To Custom pattern: the pattern written so far, to edit; from it: the text inside the wildcards
	const changeMatch = (match: LikeMatch) => {
		if (match === "custom") {
			return update({ match, text: state.text === "" ? "" : buildWildcardValue(operator, state), literal: false });
		}
		if (state.match !== "custom") return update({ match });
		const parsed = parseWildcardValue(operator, state.text);
		update(parsed.match === "custom" ? { match } : { match, text: parsed.text, literal: parsed.literal });
	};

	const custom = isLike && state.match === "custom";
	const strayWildcard = !isLike && !state.literal && hasWildcard(state.text);
	const noWildcard = custom && state.text !== "" && !hasWildcard(state.text);
	const helperWarning: Validation = strayWildcard
		? { state: "warning", message: "This value contains a wildcard character (% _ [). Turn on Match literally to search for it as text." }
		: noWildcard
			? { state: "warning", message: "No wildcard: this matches the whole value, like Equals." }
			: undefined;

	return (
		<>
			{isLike && (
				<InfoField
					label="Match"
					info={
						<Info
							text="Contains writes %text%, Begins with text% and Ends with %text. Custom pattern writes the value as you type it, wildcards included."
							href={DOCS.wildcards}
						/>
					}
				>
					<RadioGroup
						className={styles.matches}
						layout="horizontal"
						value={state.match}
						onChange={(_, data) => changeMatch(data.value as LikeMatch)}
					>
						<Radio value="contains" label="Contains" />
						<Radio value="begins" label="Begins with" />
						<Radio value="ends" label="Ends with" />
						<Radio value="custom" label="Custom pattern" />
					</RadioGroup>
				</InfoField>
			)}
			<InfoField
				label={custom ? "Pattern" : "Text"}
				validation={worst(helperWarning, advice)}
				info={
					<Info
						text={
							isLike
								? "like and not-like match only the wildcards in the value: % (any characters), _ (one character), [ ] (one character from a set). Pick a match above, or write the pattern yourself with Custom pattern."
								: "This operator adds the % wildcard itself, so don't type it. Any %, _ or [ in the value is still a wildcard unless Match literally is on."
						}
						href={DOCS.wildcards}
					/>
				}
			>
				<Input
					value={state.text}
					placeholder={custom ? "e.g. con_oso%" : "Text to find"}
					onChange={(_, data) => update({ text: data.value })}
				/>
			</InfoField>
			<Checkbox
				label="Match literally (% _ [ are text, not wildcards)"
				checked={state.literal && !custom}
				disabled={custom}
				onChange={(_, data) => update({ literal: data.checked === true })}
			/>
		</>
	);
}
