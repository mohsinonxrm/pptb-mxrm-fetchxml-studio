/** The Bulk delete dialog's default job name: the entity and the time the dialog opened (#182) */
export function defaultBulkDeleteJobName(entityDisplayName: string | undefined, now: Date = new Date()): string {
	return `Bulk Delete ${entityDisplayName || "Records"} - ${now.toLocaleString()}`;
}
