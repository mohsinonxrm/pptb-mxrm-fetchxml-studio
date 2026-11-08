import {
	countByLevel,
	LEVEL_LABELS,
	type Advisory,
	type AdvisoryLevel,
} from "../../model/advisories";

export interface ResultsTabLabel {
	/** Hidden before the first run */
	badge?: string;
	ariaLabel: string;
	menuLabel: string;
}

export type AdvisoriesBadgeColor = "danger" | "warning" | "informative";

export interface AdvisoriesTabLabel {
	/** Hidden when there are no advisories */
	badge?: string;
	badgeColor?: AdvisoriesBadgeColor;
	ariaLabel: string;
	menuLabel: string;
}

/** Results tab: the exact row count in the user's locale, no cap */
export function resultsTabLabel(rowCount: number | undefined, locale?: string): ResultsTabLabel {
	if (rowCount === undefined) return { ariaLabel: "Results", menuLabel: "Results" };
	const count = new Intl.NumberFormat(locale).format(rowCount);
	const noun = rowCount === 1 ? "row" : "rows";
	return {
		badge: count,
		ariaLabel: `Results, ${count} ${noun}`,
		menuLabel: `Results (${count} ${noun})`,
	};
}

const LEVELS: readonly AdvisoryLevel[] = ["error", "warning", "advice"];
const LEVEL_COLOR: Record<AdvisoryLevel, AdvisoriesBadgeColor> = {
	error: "danger",
	warning: "warning",
	advice: "informative",
};

/** Advisories tab: the total, coloured by the worst level present */
export function advisoriesTabLabel(advisories: readonly Advisory[]): AdvisoriesTabLabel {
	const counts = countByLevel(advisories);
	const total = advisories.length;
	if (total === 0) return { ariaLabel: "Advisories", menuLabel: "Advisories" };
	const worst = LEVELS.find((l) => counts[l] > 0)!;
	// The level labels read the same in singular and plural, so only the counts change
	const parts = LEVELS.filter((l) => counts[l] > 0).map(
		(l) => `${counts[l]} ${LEVEL_LABELS[l].toLowerCase()}`,
	);
	return {
		badge: String(total),
		badgeColor: LEVEL_COLOR[worst],
		ariaLabel: `Advisories, ${total}: ${parts.join(", ")}`,
		menuLabel: `Advisories (${total})`,
	};
}
