/**
 * The text "Copy details" puts on the clipboard when part of the tool fails to render (#185): the
 * error, where it happened, the version and the time. Never record data or query values (#71): the
 * component stack holds component names only, and only its first lines are kept.
 */

export interface ErrorDetailsInput {
	/** Where it failed, e.g. "Results panel" or "Whole tool" */
	where: string;
	error: unknown;
	componentStack?: string | null;
	version: string;
	time: Date;
}

const STACK_LINES = 8;

export function errorName(error: unknown): string {
	return error instanceof Error ? error.name : "Error";
}

export function errorMessage(error: unknown): string {
	if (error instanceof Error) return error.message;
	return typeof error === "string" ? error : "Unknown error";
}

export function errorDetails({ where, error, componentStack, version, time }: ErrorDetailsInput): string {
	const stack = (componentStack ?? "")
		.split("\n")
		.map((line) => line.trim())
		.filter(Boolean)
		.slice(0, STACK_LINES);
	return [
		`FetchXML Studio ${version} · ${where}`,
		`${errorName(error)}: ${errorMessage(error)}`,
		stack.length > 0 ? `Component stack:\n  ${stack.join("\n  ")}` : "",
		`Time: ${time.toISOString()}`,
	]
		.filter(Boolean)
		.join("\n");
}
