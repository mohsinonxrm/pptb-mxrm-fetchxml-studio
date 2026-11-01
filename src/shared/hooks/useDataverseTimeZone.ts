/**
 * The user's Dataverse time zone (#207, #184), loaded once per connection; undefined until it has
 * loaded or when it can't be read.
 */
import { useEffect, useState } from "react";
import { loadUserTimeZone } from "../../features/fetchxml/api/pptbClient";
import type { DataverseTimeZone } from "../../features/fetchxml/model/dateTimeValues";

export function useDataverseTimeZone(): DataverseTimeZone | undefined {
	const [timeZone, setTimeZone] = useState<DataverseTimeZone | undefined>(undefined);
	useEffect(() => {
		let current = true;
		void loadUserTimeZone().then((tz) => {
			if (current) setTimeZone(tz);
		});
		return () => {
			current = false;
		};
	}, []);
	return timeZone;
}
