/**
 * Metadata loads that failed for the property editors (#201). The editors still fall back to free
 * text, but the panel says what couldn't load and why, with Retry. Failed loads aren't cached
 * (dataverseMetadata.ts), so a retry asks Dataverse again; loads that worked come from the cache.
 */
import { useSyncExternalStore } from "react";
import { formatDataverseError } from "../../../api/dataverseErrors";

/** What couldn't load: an entity's columns, or its details (entity metadata, hierarchy) */
export type MetadataKind = "columns" | "details";

const KINDS: readonly MetadataKind[] = ["columns", "details"];
const failures = new Map<string, string>(); // `${kind}:${entity}` → message
const listeners = new Set<() => void>();
let generation = 0;

function emit() {
	for (const listener of listeners) listener();
}

function subscribe(listener: () => void) {
	listeners.add(listener);
	return () => {
		listeners.delete(listener);
	};
}

/** Record the outcome of a load: an error, or undefined when it worked (clears an earlier failure) */
export function recordMetadataLoad(kind: MetadataKind, entityName: string, error?: unknown): void {
	const key = `${kind}:${entityName}`;
	if (error === undefined) {
		if (failures.delete(key)) emit();
		return;
	}
	const message = formatDataverseError(error, "The request failed.");
	if (failures.get(key) !== message) {
		failures.set(key, message);
		emit();
	}
}

/** Load every failed piece of metadata again (the hooks re-run on the new generation) */
export function retryMetadataLoads(): void {
	generation++;
	emit();
}

/** Forget every failure, for example when the connection changes */
export function clearMetadataLoadErrors(): void {
	if (failures.size === 0) return;
	failures.clear();
	emit();
}

/** Changes on every Retry; the metadata hooks list it as an effect dependency */
export function useMetadataRetryGeneration(): number {
	return useSyncExternalStore(subscribe, () => generation);
}

/** The first failed load for an entity, columns before details; undefined when nothing failed */
export function metadataLoadError(entityName: string | undefined): { kind: MetadataKind; message: string } | undefined {
	if (!entityName) return undefined;
	for (const kind of KINDS) {
		const message = failures.get(`${kind}:${entityName}`);
		if (message !== undefined) return { kind, message };
	}
	return undefined;
}

/** metadataLoadError as a hook: re-renders when a failure is recorded, cleared or retried */
export function useMetadataLoadError(entityName: string | undefined): { kind: MetadataKind; message: string } | undefined {
	// A string snapshot, so React sees no change until the failure itself changes
	const snapshot = useSyncExternalStore(subscribe, () => {
		const failure = metadataLoadError(entityName);
		return failure ? `${failure.kind}\n${failure.message}` : "";
	});
	return snapshot ? metadataLoadError(entityName) : undefined;
}
