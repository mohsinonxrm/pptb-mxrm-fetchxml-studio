/** Map query-rule findings (#107 rule engine) to Field validation messages in the editors. */
import type { RuleFinding } from "../../../model/fetchxmlRules";
import type { Validation } from "./fields";

const SEVERITY_ORDER = { error: 0, warning: 1, info: 2 } as const;

/** The most severe finding for any of the given FetchXML attributes of the node. */
export function findingFor(findings: readonly RuleFinding[], ...attributes: (string | undefined)[]): Validation {
	const match = findings
		.filter((f) => attributes.includes(f.attribute))
		.sort((a, b) => SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity])[0];
	if (!match) return undefined;
	return { state: match.severity === "info" ? "none" : match.severity, message: match.message };
}

/** Combine validations, keeping the most severe. */
export function worst(...validations: Validation[]): Validation {
	const rank = (v: Validation) => (v ? (v.state === "error" ? 0 : v.state === "warning" ? 1 : 2) : 3);
	return validations.filter(Boolean).sort((a, b) => rank(a) - rank(b))[0];
}
