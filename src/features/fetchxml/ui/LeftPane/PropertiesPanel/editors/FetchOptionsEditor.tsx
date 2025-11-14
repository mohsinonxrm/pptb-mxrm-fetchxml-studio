/**
 * Property editor for Fetch root node options
 * Controls query-level settings like aggregation, paging, distinct, etc.
 */

import {
	Field,
	Checkbox,
	Input,
	Label,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Info16Regular } from "@fluentui/react-icons";
import type { FetchNode } from "../../../../model/nodes";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
		padding: tokens.spacingVerticalM,
	},
	section: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
	},
	fieldWithTooltip: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
	},
	tooltipIcon: {
		color: tokens.colorNeutralForeground3,
		cursor: "help",
	},
});

interface FetchOptionsEditorProps {
	node: FetchNode;
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function FetchOptionsEditor({ node, onUpdate }: FetchOptionsEditorProps) {
	const styles = useStyles();
	const { options } = node;

	const handleCheckboxChange =
		(field: string) => (_: unknown, data: { checked: boolean | "mixed" }) => {
			onUpdate({ [field]: data.checked === true });
		};

	const handleNumberChange =
		(field: string, min?: number, max?: number) => (_: unknown, data: { value: string }) => {
			const value = data.value === "" ? undefined : Number(data.value);
			if (value !== undefined) {
				if ((min !== undefined && value < min) || (max !== undefined && value > max)) {
					return; // Ignore invalid values
				}
			}
			onUpdate({ [field]: value });
		};

	const handleTextChange = (field: string) => (_: unknown, data: { value: string }) => {
		onUpdate({ [field]: data.value || undefined });
	};

	return (
		<div className={styles.container}>
			{/* Aggregation & Distinct Section */}
			<div className={styles.section}>
				<Label weight="semibold">Query Options</Label>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={options.aggregate ?? false}
							onChange={handleCheckboxChange("aggregate")}
							label="Aggregate Query"
						/>
						<Tooltip
							content="Enable for grouping and aggregation (SUM, COUNT, AVG, MIN, MAX). Requires grouped or aggregated attributes."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={options.distinct ?? false}
							onChange={handleCheckboxChange("distinct")}
							label="Distinct Rows"
						/>
						<Tooltip
							content="Return only unique rows. Useful to eliminate duplicates from joins or multi-value fields."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={options.returnTotalRecordCount ?? false}
							onChange={handleCheckboxChange("returnTotalRecordCount")}
							label="Return Total Record Count"
						/>
						<Tooltip
							content="Include total count of records matching query (ignoring paging). May impact performance."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field>
					<div className={styles.fieldWithTooltip}>
						<Checkbox
							checked={options.noLock ?? false}
							onChange={handleCheckboxChange("noLock")}
							label="No Lock (Read Uncommitted)"
						/>
						<Tooltip
							content="Use NOLOCK hint for better read performance. May return uncommitted data. Use cautiously."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
			</div>

			{/* Paging Section */}
			<div className={styles.section}>
				<Label weight="semibold">Paging & Limits</Label>

				<Field label="Max Records (Top)">
					<div className={styles.fieldWithTooltip}>
						<Input
							type="number"
							min={1}
							value={options.top?.toString() ?? ""}
							onChange={handleNumberChange("top", 1)}
							placeholder="No limit"
						/>
						<Tooltip
							content="Maximum records to return (LIMIT clause). Leave empty for all records."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Page Size (Count)">
					<div className={styles.fieldWithTooltip}>
						<Input
							type="number"
							min={1}
							max={5000}
							value={options.count?.toString() ?? ""}
							onChange={handleNumberChange("count", 1, 5000)}
							placeholder="Default: 5000"
						/>
						<Tooltip
							content="Records per page for pagination. Max 5000. Used with Page Number."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Page Number">
					<div className={styles.fieldWithTooltip}>
						<Input
							type="number"
							min={1}
							value={options.page?.toString() ?? ""}
							onChange={handleNumberChange("page", 1)}
							placeholder="1"
						/>
						<Tooltip
							content="Which page to retrieve. Starts at 1. Used with Page Size."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>

				<Field label="Paging Cookie" hint="Copy from previous query results for efficient paging">
					<Input
						value={options.pagingCookie ?? ""}
						onChange={handleTextChange("pagingCookie")}
						placeholder="<cookie>...</cookie>"
					/>
				</Field>
			</div>

			{/* Advanced Section */}
			<div className={styles.section}>
				<Label weight="semibold">Advanced</Label>

				<Field label="UTC Offset (minutes)">
					<div className={styles.fieldWithTooltip}>
						<Input
							type="number"
							min={-720}
							max={720}
							value={options.utcOffset?.toString() ?? ""}
							onChange={handleNumberChange("utcOffset", -720, 720)}
							placeholder="0"
						/>
						<Tooltip
							content="Timezone offset for datetime filters. Range: -720 to 720 minutes from UTC."
							relationship="description"
						>
							<Info16Regular className={styles.tooltipIcon} />
						</Tooltip>
					</div>
				</Field>
			</div>
		</div>
	);
}
