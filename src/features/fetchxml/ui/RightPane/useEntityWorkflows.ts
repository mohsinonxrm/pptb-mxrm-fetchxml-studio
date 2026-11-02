/**
 * The Run workflow menu's list, loaded when the menu first opens and kept for one entity (#182).
 * The command bar stays mounted while the entity changes, so a list kept regardless of entity
 * showed the previous entity's workflows, and picking one opened a dialog that couldn't find it.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import type { WorkflowInfo } from "../../api/pptbClient";

export interface EntityWorkflows {
	workflows: WorkflowInfo[];
	isLoading: boolean;
	/** The list was loaded for the current entity (an empty list means it has none) */
	loaded: boolean;
	/** Loads the list unless it's already loaded or loading for the current entity */
	load: () => Promise<void>;
}

export function useEntityWorkflows(
	entityName: string | undefined,
	onFetchWorkflows: (() => Promise<WorkflowInfo[]>) | undefined,
): EntityWorkflows {
	const [workflows, setWorkflows] = useState<WorkflowInfo[]>([]);
	const [isLoading, setIsLoading] = useState(false);
	const [loaded, setLoaded] = useState(false);
	// Bumped when the entity changes: a load started for the previous entity is dropped when it ends
	const requestRef = useRef(0);

	useEffect(() => {
		requestRef.current++;
		setWorkflows([]);
		setIsLoading(false);
		setLoaded(false);
	}, [entityName]);

	const load = useCallback(async () => {
		if (loaded || isLoading || !onFetchWorkflows) return;

		const request = ++requestRef.current;
		setIsLoading(true);
		try {
			const wfs = await onFetchWorkflows();
			if (request !== requestRef.current) return;
			setWorkflows(wfs);
			setLoaded(true);
		} catch (err) {
			if (request !== requestRef.current) return;
			console.error("Failed to load workflows:", err);
		} finally {
			if (request === requestRef.current) setIsLoading(false);
		}
	}, [loaded, isLoading, onFetchWorkflows]);

	return { workflows, isLoading, loaded, load };
}
