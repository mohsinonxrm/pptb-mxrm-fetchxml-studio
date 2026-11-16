/**
 * The record actions' dialogs (#301): Delete, Bulk delete, Run workflow, Activate / Deactivate
 * (#143) and Close (#149). Their state and handlers are in useRecordActions.
 */
import type { EntityMetadata } from "../features/fetchxml/api/pptbClient";
import { DeleteConfirmDialog } from "../features/fetchxml/ui/Dialogs/DeleteConfirmDialog";
import { BulkDeleteDialog } from "../features/fetchxml/ui/Dialogs/BulkDeleteDialog";
import { WorkflowPickerDialog } from "../features/fetchxml/ui/Dialogs/WorkflowPickerDialog";
import { SetStateDialog } from "../features/fetchxml/ui/Dialogs/SetStateDialog";
import { CloseRecordsDialog } from "../features/fetchxml/ui/Dialogs/CloseRecordsDialog";
import type { RecordActions } from "./hooks/useRecordActions";

export interface RecordActionDialogsProps {
	/** The dialogs' state and handlers (useRecordActions) */
	recordActions: RecordActions;
	/** The root entity's logical name, or undefined without a query */
	entityLogicalName: string | undefined;
	/** The root entity's metadata, or null until it loads */
	entityMetadata: EntityMetadata | null;
	/** Whether the results have more pages (Bulk delete's record count) */
	moreRecords: boolean | undefined;
}

export function RecordActionDialogs({
	recordActions,
	entityLogicalName,
	entityMetadata,
	moreRecords,
}: RecordActionDialogsProps) {
	return (
		<>
			{/* Delete Confirmation Dialog */}
			<DeleteConfirmDialog
				open={recordActions.deleteDialogState.open}
				recordName={recordActions.deleteDialogState.recordName}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
					entityLogicalName ||
					"record"
				}
				entityLogicalName={entityLogicalName}
				recordCount={recordActions.deleteDialogState.recordIds.length}
				onClose={recordActions.dismissDeleteDialog}
				onConfirm={recordActions.confirmDelete}
				onProgress={recordActions.deleteProgress ?? undefined}
			/>
			{/* Bulk Delete Dialog */}
			<BulkDeleteDialog
				open={recordActions.bulkDeleteDialogState.open}
				recordCount={recordActions.bulkDeleteDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
					entityLogicalName ||
					"records"
				}
				entityLogicalName={entityLogicalName}
				isAllRecords={recordActions.bulkDeleteDialogState.isAllRecords}
				totalViewRecords={recordActions.bulkDeleteDialogState.totalViewRecords}
				moreRecords={moreRecords}
				onClose={recordActions.dismissBulkDeleteDialog}
				onConfirm={recordActions.confirmBulkDelete}
			/>
			{/* Workflow Picker Dialog */}
			<WorkflowPickerDialog
				open={recordActions.workflowDialogState.open}
				recordCount={recordActions.workflowDialogState.recordIds.length}
				entityDisplayName={
					entityMetadata?.DisplayName?.UserLocalizedLabel?.Label ||
					entityLogicalName ||
					"records"
				}
				entityLogicalName={entityLogicalName}
				preSelectedWorkflow={recordActions.workflowDialogState.preSelectedWorkflow}
				onClose={recordActions.dismissWorkflowDialog}
				onFetchWorkflows={recordActions.fetchWorkflows}
				onExecute={(workflow, onProgress) =>
					recordActions.executeWorkflow(
						workflow,
						recordActions.workflowDialogState.recordIds,
						onProgress,
					)
				}
			/>
			{/* Activate / Deactivate Dialog (#143) */}
			<SetStateDialog
				open={recordActions.setStateDialogState.open}
				action={recordActions.setStateDialogState.action}
				recordIds={recordActions.setStateDialogState.recordIds}
				entityLogicalName={entityLogicalName}
				entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
				primaryIdAttribute={entityMetadata?.PrimaryIdAttribute}
				closeAction={recordActions.setStateDialogState.closeAction}
				onClose={recordActions.dismissSetStateDialog}
				onExecute={recordActions.executeStateChange}
			/>
			{/* Close message Dialog (#149) */}
			<CloseRecordsDialog
				open={recordActions.closeDialogState.open}
				action={recordActions.closeDialogState.action}
				recordIds={recordActions.closeDialogState.recordIds}
				entityLogicalName={entityLogicalName}
				entityDisplayName={entityMetadata?.DisplayName?.UserLocalizedLabel?.Label}
				onClose={recordActions.dismissCloseDialog}
				onExecute={recordActions.executeClose}
			/>
		</>
	);
}
