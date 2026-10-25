/**
 * The scope chip's text (#227): one button above the entity picker that always shows the scope
 * you're working in, and opens the publisher / solution pickers. Shared design with Dataverse
 * Request Studio and Command Bar Studio (variant A).
 */
import type { EntityScopeMode } from "./displaySettings";

/** What the chip needs to know about a publisher */
export interface ChipPublisher {
	friendlyname: string;
}

/** What the chip needs to know about a solution */
export interface ChipSolution {
	friendlyname: string;
	uniquename?: string;
	version?: string;
}

/** "Sales Core v1.2.0": the first three version parts tell a patch from its parent */
export function solutionTagLabel(solution: ChipSolution): string {
	const version = solution.version?.split(".").slice(0, 3).join(".");
	return version ? `${solution.friendlyname} v${version}` : solution.friendlyname;
}

/** "cso_salescore · v1.2.0.0": the solution option's second line, and its tag's tooltip */
export function solutionDetail(solution: ChipSolution): string {
	return [solution.uniquename, solution.version && `v${solution.version}`].filter(Boolean).join(" · ");
}

/** "Contoso", or "Contoso +1" when there's more than one */
const firstAndMore = (names: readonly string[]): string =>
	names.length > 1 ? `${names[0]} +${names.length - 1}` : (names[0] ?? "");

export interface ScopeChipText {
	/** False when the entity picker has nothing to list yet: the chip says "Choose scope" */
	configured: boolean;
	/** The chip's label */
	label: string;
	/** Accessible name: "Entity scope: …", or "Choose scope" */
	ariaLabel: string;
	/** Every publisher and solution, for the tooltip when the label is cut off */
	detail: string;
}

/**
 * The chip's text for the effective scope mode and the current selection.
 * - All entities: "All entities"
 * - Publisher → solution: "Contoso › Sales Core v1.2.0", with "+N" for more than one of either
 * - Solution only: "Sales Core v1.2.0 +1"
 * Without a solution in a scoped mode, there are no entities to list: "Choose scope".
 */
export function scopeChipText(
	mode: EntityScopeMode,
	publishers: readonly ChipPublisher[],
	solutions: readonly ChipSolution[],
): ScopeChipText {
	if (mode === "all") {
		return { configured: true, label: "All entities", ariaLabel: "Entity scope: All entities", detail: "All entities" };
	}
	if (solutions.length === 0) {
		return { configured: false, label: "Choose scope", ariaLabel: "Choose scope", detail: "" };
	}
	const solutionText = firstAndMore(solutions.map(solutionTagLabel));
	const label =
		mode === "publisher-solution" && publishers.length > 0
			? `${firstAndMore(publishers.map((p) => p.friendlyname))} › ${solutionText}`
			: solutionText;
	const detail = [
		mode === "publisher-solution" && publishers.length > 0 && `Publishers: ${publishers.map((p) => p.friendlyname).join(", ")}`,
		`Solutions: ${solutions.map((s) => [s.friendlyname, solutionDetail(s)].filter(Boolean).join(" — ")).join("; ")}`,
	]
		.filter(Boolean)
		.join("\n");
	return { configured: true, label, ariaLabel: `Entity scope: ${label}`, detail };
}

/** A solution matches the search by name, unique name or version */
export function solutionMatchesSearch(solution: ChipSolution, query: string): boolean {
	const q = query.trim().toLowerCase();
	if (!q) return true;
	return [solution.friendlyname, solution.uniquename, solution.version].some((v) => (v ?? "").toLowerCase().includes(q));
}
