/**
 * Error boundaries (#185). Without them, one render error unmounted the whole React tree and PPTB
 * showed an empty tab, losing the query (#217).
 *
 * - PaneErrorBoundary: around one panel. It shows an error MessageBar with Try again (re-mounts the
 *   panel) and Copy details; the rest of the tool keeps working. It also tries again by itself when
 *   `resetKey` changes (another node selected, new results, a different query).
 * - ToolErrorBoundary: around the whole tool. Reloading doesn't work in PPTB, so it offers Restart
 *   with this query (the tool is re-mounted with the last good FetchXML), Start a new query, Copy
 *   FetchXML and Copy details.
 *
 * Errors are logged once with console.error, without record data (#71). The look is variant A of
 * the reviewed prototype (.claude/mockups/error-boundary).
 */
import { Component, type ErrorInfo, type ReactNode } from "react";
import { errorDetails, errorMessage } from "../utils/errorDetails";
import { PaneErrorFallback, ToolErrorFallback } from "./ErrorFallbacks";

interface CaughtState {
	error: unknown;
	componentStack?: string | null;
	time?: Date;
}

const NOTHING_CAUGHT: CaughtState = { error: null };

// ── One panel ────────────────────────────────────────────────────────────────

interface PaneErrorBoundaryProps {
	/** For the details and the log, e.g. "Results panel" */
	where: string;
	/** Shown as the MessageBar title, e.g. "The results couldn't be shown" */
	title: string;
	/** When this changes after an error, the panel is tried again */
	resetKey?: unknown;
	children: ReactNode;
}

export class PaneErrorBoundary extends Component<PaneErrorBoundaryProps, CaughtState> {
	state: CaughtState = NOTHING_CAUGHT;

	static getDerivedStateFromError(error: unknown): CaughtState {
		return { error, time: new Date() };
	}

	componentDidCatch(error: unknown, info: ErrorInfo) {
		console.error(`[FetchXML Studio] ${this.props.where} failed to render: ${errorMessage(error)}`);
		this.setState({ componentStack: info.componentStack });
	}

	componentDidUpdate(previous: PaneErrorBoundaryProps) {
		if (this.state.error !== null && previous.resetKey !== this.props.resetKey) this.reset();
	}

	reset = () => this.setState(NOTHING_CAUGHT);

	render() {
		if (this.state.error === null) return this.props.children;
		return (
			<PaneErrorFallback
				title={this.props.title}
				message={errorMessage(this.state.error)}
				details={errorDetails({
					where: this.props.where,
					error: this.state.error,
					componentStack: this.state.componentStack,
					version: __FXS_VERSION__,
					time: this.state.time ?? new Date(),
				})}
				onTryAgain={this.reset}
			/>
		);
	}
}

// ── The whole tool ───────────────────────────────────────────────────────────

interface ToolErrorBoundaryProps {
	/** The last good query's FetchXML, if there is one */
	readFetchXml: () => string | undefined;
	/** Re-mount the tool, with the last good query or with a new one */
	onRestart: (withQuery: boolean) => void;
	children: ReactNode;
}

export class ToolErrorBoundary extends Component<ToolErrorBoundaryProps, CaughtState & { fetchXml?: string }> {
	state: CaughtState & { fetchXml?: string } = NOTHING_CAUGHT;

	static getDerivedStateFromError(error: unknown): CaughtState {
		return { error, time: new Date() };
	}

	componentDidCatch(error: unknown, info: ErrorInfo) {
		console.error(`[FetchXML Studio] The tool failed to render: ${errorMessage(error)}`);
		// Read once: the query the tool had before it failed
		this.setState({ componentStack: info.componentStack, fetchXml: this.props.readFetchXml() });
	}

	restart = (withQuery: boolean) => {
		this.props.onRestart(withQuery);
		this.setState({ ...NOTHING_CAUGHT, fetchXml: undefined });
	};

	render() {
		if (this.state.error === null) return this.props.children;
		return (
			<ToolErrorFallback
				fetchXml={this.state.fetchXml}
				details={errorDetails({
					where: "Whole tool",
					error: this.state.error,
					componentStack: this.state.componentStack,
					version: __FXS_VERSION__,
					time: this.state.time ?? new Date(),
				})}
				onRestart={this.restart}
			/>
		);
	}
}
