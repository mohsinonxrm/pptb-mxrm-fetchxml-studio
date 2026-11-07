/**
 * The results info bar's total record count (#98), from returntotalrecordcount.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/count-rows?tabs=webapi
 *
 * - @Microsoft.Dynamics.CRM.totalrecordcount: "The total number of records up to 5,000; otherwise
 *   the value is -1" (-1 also when the query doesn't ask for it; verified live 2026-10-17).
 * - @Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded: true when more records match; the count
 *   is then the limit (5,000 for standard entities, 500 for elastic ones; verified live: 5000 + true
 *   for account), so it's shown as "5,000+".
 */

const format = (n: number) => n.toLocaleString("en-US");

/**
 * "1,234 total", or "5,000+ total" when the limit is exceeded; undefined when there's no count.
 * With the limit exceeded, the loaded rows can be more than the limit (Retrieve all), so the lower
 * bound is the larger of the two.
 */
export function totalCountText(
	totalRecordCount: number | undefined,
	limitExceeded: boolean | undefined,
	loadedRows: number,
): string | undefined {
	if (totalRecordCount === undefined || totalRecordCount < 0) return undefined;
	return limitExceeded ? `${format(Math.max(totalRecordCount, loadedRows))}+ total` : `${format(totalRecordCount)} total`;
}

/** The unfiltered entity total from RetrieveTotalRecordCount, worded so it isn't read as the query's count. */
export function entityTotalText(entityLogicalName: string, total: number): string {
	return `About ${format(total)} ${entityLogicalName} records in all: a snapshot from the last 24 hours that ignores this query's filters (RetrieveTotalRecordCount).`;
}
