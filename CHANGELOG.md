# Changelog

All notable changes to FetchXML Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### 🔧 Fixed

- **A view opened from another tool loads like Load view (#197).**
  - **Before:** only the view's FetchXML was loaded. Its column layout was ignored, an unchanged view wasn't run by ID, and **Save** offered Save as rather than overwrite. A view that didn't exist, or a failed request, left an empty builder with no message.
  - **Now:** the view loads the way **Load view** loads it. When it can't be opened, a notification says why: "Couldn't open the view" / "…asked for a personal view (id) that doesn't exist or that you can't read", or Dataverse's error.

- **Run workflow: background workflows in batches, and no workflow runs twice (#264).**
  - **Background workflows** now go in `$batch` requests, like delete, Activate / Deactivate and Close, with the same per-record errors, progress, throttling back-off and connection guard.
  - **A batch that stops early** (PPTB doesn't report which of its operations ran): records whose workflow already started a new system job aren't sent again; the others are sent one at a time. System jobs are compared with a snapshot taken just before the batch, so earlier runs don't count. If the jobs can't be relied on (they can't be read, or the workflow deletes its completed jobs, so a fast run's job may already be gone), none of that batch is sent again, and each record says to check its system jobs. Another user starting the same workflow on the same record at that moment would make it count as started: it may be skipped, never run twice.
  - **Real-time workflows** run immediately and leave no system job ([Create real-time workflows](https://learn.microsoft.com/dynamics365/customerengagement/on-premises/developer/create-real-time-workflows)), so they stay one request per record.

- **Closing the tool with unsaved Save view edits asks first (#264).** While the Save view dialog has changes to the name, description, view type or default setting, PPTB asks before the tool closes, as it does while a view saves.

- **Advisories follow-ups (#170).**
  - **Paging without a unique order** stays advice. Live, Dataverse broke ties by the primary key in every case tried, but that isn't documented, and [Order rows](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows#ordering-and-paging) still asks for a unique column. It now also applies when the query's results are paged (more than 5,000 rows, loaded by scrolling or Retrieve All), not only when the query sets `count` or `page`.
  - **Stale check:** after the query changes, an earlier **Check with Dataverse** no longer shows its findings or "Dataverse agrees" tags, and they aren't counted on the tab; they could describe a query Dataverse never saw. The status line says when it was checked and to check again.
  - **Metadata loading:** the metadata for the metadata rules (attribute types, text lengths, the primary key) loads when the Advisories tab is first opened, not as soon as an entity is chosen (as #71 intended). The tab's count can grow when you open it.

- **Long operations wait out Dataverse throttling, and Retrieve All stops at 100,000 rows (#183).**
  - **Before:** a throttled request was retried 3 times after 2, 4 and 8 s, 14 s in all, while Dataverse's service protection limits are measured over 5 minutes, so long runs gave up early. Batched deletes and record actions had their own waits (5, 15 and 30 s). An operation that a batch answered with 429 was recorded as failed, not retried. Retrieve All kept loading with no limit, so a wide table could exhaust the tool's memory.
  - **Now:** one back-off for long operations: Retrieve All, deletes of 4 records or more, Activate / Deactivate, Close, workflows on many records, and the one-request-per-record path on older hosts. It waits about 5, 10, 20, 40, 80 and 120 s (± 20% jitter), about 4.5 minutes in all, as [Service protection API limits](https://learn.microsoft.com/power-apps/developer/data-platform/api-limits) suggests. The wait shows in Retrieve All's status line and in the Delete and Run workflow dialogs ("Dataverse is limiting requests: trying again in 40s"); Activate / Deactivate and Close wait the same way but don't show it yet (UI/UX pass). A request you're waiting on, such as the first page or a delete of 1 to 3 records (sent one at a time), keeps the short back-off.
  - **Batches:** only the operations a batch answered with a service protection error are sent again. Dataverse didn't run them, and the others aren't repeated.
  - **Retrieve All** stops once 100,000 rows or more are loaded (it can overshoot by up to one page), with a notice. More pages still load one at a time by scrolling, and Bulk delete (job) deletes every matching record without loading them. The limit isn't configurable yet; a Settings option is part of the UI/UX pass.

- **The local Excel export keeps each record's currency and shows dates as the grid does (#184).**
  - **Money:** amounts were always formatted with `$`, so a EUR, GBP or INR amount showed a dollar sign. Each cell now uses its record's currency symbol and decimals from Dataverse's formatted value (`€1.234,50` → `"€"#,##0.00`, `¥1,235` → no decimals), wherever your negative currency format puts the sign. A column with a fixed precision uses it.
  - **Dates:** date-times were written in UTC while the grid shows your Dataverse time zone. They now match the grid. Date Only columns are written as dates (`yyyy-mm-dd`, no `00:00:00`), and Time Zone Independent columns as stored. Raw values (raw mode, and the "(Raw)" columns) are in UTC, and the header says "(UTC)"; so does the header when your time zone couldn't be read. If a column's date behaviour couldn't be loaded, its formatted text is written.
  - **Raw mode** wrote formatted numbers. Raw values now have no number format, apart from dates.
  - **Decimal places:** decimal and floating point columns always got two decimals, because their precision wasn't loaded, and precision 0 would have given `0.`. They now use the column's precision (`0` for none).
  - **Grid fallbacks:** when a cell has no formatted value, Money no longer gets a `$`, and a Date Only value is no longer shown as the day before west of UTC.

- **Picked dates filter on the day and time you picked (#207, #198).**
  - **Before:** the date picker wrote every value as a UTC instant from the browser's time zone (`toISOString()`), whatever the column. Checked live on a sandbox, that went wrong whenever the computer's time zone differed from your Dataverse one (on the sandbox, Dataverse UTC and the computer Pacific):
    - **Date and time columns** were filtered in the computer's time zone, not your Dataverse one: picking 09:00 filtered on 16:00.
    - **on / on-or-before / on-or-after** take the date of that instant in your Dataverse time zone, so near midnight the date shifted. For example, a computer in Sydney with Dataverse set to UTC turned 14 March into `2024-03-13T23:00:00.000Z`, the 13th.
    - **Time Zone Independent columns** compare the time as written, so values were off by the computer's offset.
    - The picker showed Date Only values as the day before for anyone west of UTC (#198); the value written was right.
  - **Now:** values are written the way Dataverse reads them:
    - `2024-03-14` for Date Only columns, date-only formats and the **on** operators. Those take a date only, so the time picker is hidden.
    - The wall clock for Time Zone Independent columns: `2024-03-14T09:00:00`.
    - The wall clock with your Dataverse time zone's offset for other date and time columns: `2024-03-14T09:00:00+09:00`. It reads like what you picked, and it's exact.
  - **The pickers:** they show values in the same terms, so a Date Only value is never the day before. The time picker says which time zone applies ("Time in your Dataverse time zone (Tokyo Standard Time)"). The time zone is your Dataverse Personal Options setting, read once per connection, including its current daylight-saving rule (for dates in years when your time zone followed other rules, the offset can be an hour out).
  - **Loaded FetchXML** keeps its values as written until you pick a new date.

- **The generated C# QueryExpression compiles and means what the query does (#178).**
  - **Link-entities:** they were written with `ColumnSet =` and `Criteria =`, which `LinkEntity` doesn't have, so the code didn't compile (CS0117). They use `Columns` and `LinkCriteria` now, as do any/all filters on related rows (`AnyAllFilterLinkEntity`), which leave out the columns Dataverse ignores.
  - **Orders:** `new OrderExpression(...)` had the alias and entity name swapped. It's (attribute, order, alias, entity name) now, per the [SDK reference](https://learn.microsoft.com/dotnet/api/microsoft.xrm.sdk.query.orderexpression.-ctor).
  - **Values:** text that looked like a number (`"00123"`, `"12.50"`, a 20-digit code) was written as a number, changing the query or failing to compile. Each value now keeps the .NET type Dataverse reports, checked against live responses: strings stay quoted, decimals and money get `m` (`1000.5m`), whole numbers stay plain, GUIDs become `new Guid(...)`. Line breaks, tabs and other control characters in names and values are escaped.
  - **Columns:** a query with both plain and aliased columns (`<attribute name="name"/><attribute name="createdon" alias="created"/>`) lost the plain ones. Both are written now: `new ColumnSet("name") { AttributeExpressions = { ... } }`.
  - **Aggregates:** `hasGroupBy: true` isn't a constructor argument, and a date grouping without an alias picked the wrong constructor. Both are set as properties (`HasGroupBy = true`, `DateTimeGrouping = XrmDateTimeGrouping.Year`).
  - **Operators:** the numeric operator map didn't match the SDK from 39 on. It does now, and an unknown code is written as a cast that still compiles. (Dataverse returns operator names, so this only mattered for numeric codes.)
  - **Less noise:** the default `PageInfo` Dataverse always returns is left out unless the query pages. Query hints, `DataSource` (`datasource="retained"`), `FilterHint` and quick find filters are kept.
  - **Other snippets:** the Web API URL also encodes `! ' ( ) *`, so the cURL line no longer triggers bash history expansion ("event not found"). The JavaScript and Web API snippets find the entity name when `name` isn't the first attribute, and the offline fallback for `transactioncurrency` is `transactioncurrencies`.

- **Copied rows can't paste into a spreadsheet as live formulas (#208).**
  - **Before:** Copy wrote cell text as it was. A value such as `=HYPERLINK(...)` or `@SUM(...)`, often entered by users or integrations, became a working formula when the rows were pasted into Excel or Google Sheets.
  - **Now:** in copied rows (Ctrl+C on rows, the Copy menu, the row and selected-rows menu items), text starting with `=`, `+`, `-`, `@`, a tab or a carriage return gets a leading apostrophe, as [OWASP's CSV injection guidance](https://owasp.org/www-community/attacks/CSV_Injection) recommends, so the spreadsheet shows it as text. A formatted number's sign is left alone (`-5`, `-$5.00`), but a choice label such as `-Other` is protected too. Copying a single cell or a column (for example to paste into an `in` condition) still copies it exactly. The local Excel export wasn't affected: it writes text as text cells.

- **Record actions use the selected entity's privileges, not the previous entity's (#179).**
  - **Before:** switching entity started a new privilege check without stopping the previous one. If the previous entity's check finished last, Delete and Bulk delete used its result: enabled where you can't delete, or disabled where you can. Dataverse still refused a delete you weren't allowed to make, but the buttons and confirmations were wrong.
  - **Now:** an outdated check's result is dropped, and the record actions stay unavailable until the selected entity's check finishes. The three checks also run at the same time instead of one after another.

- **The Code tab no longer sends Dataverse requests on every keystroke (#199).**
  - **Before:** while the Code tab was open, each change to the query sent both `FetchXmlToQueryExpression` (C# QueryExpression) and `FetchXMLToSQL` (SQL), whichever format was selected. Typing a condition value sent a pair per keystroke. That counted against service protection limits, and half-typed queries flashed server errors.
  - **Now:** a format is converted only while it's selected, half a second after typing stops, and once per query and environment, so switching formats back and forth reuses the result. A failed conversion is tried again when you select its format again.

- **FetchXML editor suggestions and hovers no longer pile up (#194).**
  - **Before:** the editor registered its FetchXML completion and hover providers each time it mounted and never removed them. It remounts on every connection change, so after a few changes each suggestion and hover appeared several times, and memory grew over the session. The providers also answered in the read-only LayoutXML viewer.
  - **Now:** they're registered once, apply only to the FetchXML editor, and are removed when it goes.

- **Metadata that couldn't load is reported, with Retry (#201).**
  - **Before:** when an entity's columns or details failed to load (a missing privilege, throttling, the network), the property editors silently fell back to free text, the relationship picker showed only "Error loading relationships", and failed display names silently became logical names. "Nothing to show" and "the request failed" looked the same.
  - **Now:** the property editor shows a warning, "Couldn't load columns for account" (or entity details), with Dataverse's reason and **Retry**; the editor still works with free text. The relationship picker shows the reason under the box with **Retry**, and when only display names failed, the list works with logical names and a hint says why. A failed relationship load also no longer crashes the Properties panel, which is what happened before (the error box's Combobox had no selection list). Retry reloads what failed; what loaded comes from the cache. The design was reviewed with the owner.

- **Parse to tree and Save View keep values exactly as written (#181).**
  - **Lookup display names on `<value>`:** a saved view's `<value uiname="Contoso" uitype="account">{id}</value>` lost `uiname` and `uitype` when the query went through the tree, so loading a view and saving it again stripped the names model-driven apps show in the view's filter. They're kept now, for each value. Changing or removing a value drops only its own.
  - **Long decimals** kept only 10 decimal places (`0.12345678901234` became `0.123456789`). Every digit is kept now; only scientific notation (`1e-12`) is written out in full.
  - **Line breaks and tabs** in an attribute value turned into spaces. They're written as `&#10;`, `&#13;` and `&#9;` now.
  - **Spaces in `<value>` text** were trimmed, so `<value> </value>` became empty; the pac CLI and Web API snippets did the same. They're kept now.
  - **`top`, `count`, `page` and `utc-offset`** accepted partly numeric text: `top="10abc"` became `top="10"` without a word. Such values are now reported like any other invalid value; Dataverse rejects them (verified live: "Input string was not in a correct format"). A sign and surrounding spaces are still accepted, as Dataverse accepts them.

- **Save View: no duplicate views after a failed step, and no closing mid-save (#180).**
  - **Before:** saving a new system view creates it, adds it to the chosen solution, then publishes. If adding to the solution or publishing failed, the dialog said "Save failed" and stayed on "Save as new", so pressing Save again created a second view with the same name. A publish failure after an overwrite also said "Save failed", though the view was saved. And Esc, a click outside or the title-bar X closed the dialog while it was still saving, so the outcome was never shown.
  - **Now:** a warning says what was saved and which step failed ("Saved, but not added to the solution" or "Saved, but not published"), with Dataverse's message. The form locks, and the main button retries only that step (and publishing after it, if requested). **Close** loads the saved view, so the next Save overwrites it.
  - While saving, Esc and clicks outside are ignored and the X is disabled, like Cancel. The design was reviewed with the owner.

- **An apostrophe in a value no longer breaks Check with Dataverse or Save View validation (#176).**
  - **Before:** a condition value like `O'Brien` ended the OData string literal the FetchXML was sent in. **Check with Dataverse** failed with 0x80060888 "Syntax error", and Save View's validation quietly said it was unavailable. Verified live.
  - **Now:** the FetchXML is passed as a parameter alias, with quotes doubled and the value URL-encoded, as [Use Web API functions](https://learn.microsoft.com/power-apps/developer/data-platform/webapi/use-web-api-functions#passing-parameters-to-a-function) recommends.
  - **The Code tab (SQL, C# QueryExpression) and Bulk Delete's query conversion** now build the same request themselves instead of relying on the host to escape the FetchXML. PPTB Desktop did escape it; the PPTB VS Code extension doesn't (#173), so `'`, `&`, `#`, `%` or `+` in a value could break those requests or change the value there. Verified live on Desktop: the result is identical for a value with all five.
  - A view ID passed in by another tool must be a GUID before it's used in a query, and the privilege name in the privilege check is quoted like other values.

- **Deleting more than 100 selected records needs the Bulk Delete privilege (#177).**
  - **Before:** more than 100 selected records are deleted by a Bulk Delete job, which needs `prvBulkDelete`. Without it, the tool fell back to deleting every selected record itself, one request per record in batches, while the README said Delete covers 1–100 rows.
  - **Now:** without the privilege, Delete is unavailable for more than 100 selected records. "Delete unavailable" beside it says why and names the privilege, the same way as "Delete all unavailable" (#174). Selecting 100 or fewer works as before, and with the privilege nothing changes.

- **A render error no longer blanks the whole tool (#185).**
  - **Before:** an exception while drawing any part of FetchXML Studio unmounted everything. PPTB then showed an empty tab, and the query in progress was lost (as in #217).
  - **One panel fails:** the scope, query tree, properties, FetchXML editor, results, advisories, LayoutXML and Code panels each have their own error boundary. A failing panel shows an error message bar with **Try again** and **Copy details**, and the rest of the tool keeps working. The panel also tries again by itself when what it shows changes, for example another node, new results or a different query.
  - **The whole tool fails:** a page shows your query's FetchXML with **Restart with this query**, which re-mounts the tool and reloads the last good query, the same way as after a connection change. It also offers **Start a new query**, **Copy FetchXML** and **Copy details**. Reloading the page doesn't work in PPTB.
  - **Details:** "Copy details" gives the error, where it happened, the version and the time, never record data. Each error is logged once to the console.
  - The look was reviewed as a prototype (variant A).

- **The query tree shows which node is selected (#229).**
  - **Before:** the selected node drove the properties panel, but its row had no highlight once the mouse moved away, and screen readers weren't told it was selected.
  - **Now:** the selected row has a subtle selected background, semibold text and a brand indicator bar (as in Fluent Nav). It's marked `aria-selected`, and Windows high contrast shows it in the system highlight colours.
  - **Keyboard:** Enter selects the focused node, like a click.

- **Check with Dataverse no longer throws on an unfinished condition (#97, found by the #217 incomplete-node test).** Matching Dataverse's findings to the query read `condition.attribute.toLowerCase()`, so a condition whose attribute had just been cleared crashed the Advisories tab. A blank link-entity name had the same problem. The incomplete-node test now also runs the Advisories list, the metadata rules and the Dataverse merge.

### ✨ Added

- **Scope chip: where entities come from, in one button (#227).** It's the design shared with Dataverse Request Studio and Command Bar Studio (variant A).
  - **The chip** heads the left pane and always shows the current scope:
    - `Contoso › Sales Core v1.2.0`, with `+N` for more;
    - `All entities`;
    - a primary **Choose scope** when nothing is picked yet.

    A solution shows its first three version parts, so a patch differs from its parent. A tooltip shows every publisher and solution when the label is cut off.
  - **Its popover** holds the scope settings, and changes apply right away:
    - **Show entities from:** Publisher → solution, Solution only, All entities. Options your privileges don't allow are off.
    - **Publishers** and **Solutions**, as tag pickers. Solutions show the unique name and version, grouped Unmanaged / Managed, and search matches the name, unique name or version.
    - **Advanced Find only.**
    - **Done** returns focus to the chip.
  - **First run:** a tip points at the chip ("Choose where entities come from"). Its **Choose scope** opens the popover. It's shown once per environment, and only after any remembered scope has been restored.
  - **Entity picker:** says "Choose a scope to list entities" until there is one, and shows how many entities are in scope.
  - **Replaced:**
    - the always-visible Publishers / Solutions comboboxes;
    - the collapsible Query scope header and its summary (#25);
    - the "remembered for this environment" note (remembering is now simply how it works);
    - the Entity source and Advanced Find only settings in the Settings drawer.

- **Query tree icons say what each node is, and are coloured by what it does (#230).**
  - **Colours,** from Fluent palette tokens that adapt to light, dark and high contrast:

    | Colour | Nodes |
    |---|---|
    | Brand | the query |
    | Blue | tables: entity and link-entities |
    | Teal | columns: attributes |
    | Orange | criteria: filters and conditions |
    | Grape | sort: orders |

  - **Every node type has its own icon.** Fetch and Entity, and Filter and Condition, used to share one.
    - Fetch `</>`, Entity a table, link-entity a joined table.
    - Condition an operators icon, order an up or down arrow by direction, All attributes three columns.
  - **Attributes show their Dataverse column type**, from the `AttributeTypeName` FXS already loads, so there are no new requests. All 24 types on Microsoft Learn are covered, for example:
    - Text, Multiple lines, Whole number, Decimal, Currency, Yes/No, Date and time;
    - Choice, Choices, Status, Status reason;
    - Lookup, Customer, Owner, Image, File;
    - a key for the primary key.

    While metadata loads, a generic column icon shows. Calculated, rollup, formula and prompt columns get a badge.
  - **New Settings option, Display › Tree icons:**
    - **Tinted icons** (default): one colour per icon.
    - **Color icons:** Fluent's multicolour icons where they exist, filled tinted icons elsewhere. In Windows high contrast they switch to the plain filled icon, as Fluent's icon package advises.
  - **Add menus:** All attributes, Condition and Link-entity use the matching icons.

- **Density for the results grid and the query tree (#228), as in Dataverse Request Studio.**
  - **Where:** a **Density** menu in the results command bar: **Medium** (as before), **Small** and **Extra small**. Wide results with many columns fit more on screen.
  - **How it applies:**
    - One setting drives both. The grid uses Fluent's `DataGrid size`: rows of 42, 34 and 24 px.
    - The tree uses Fluent's `Tree size`. Small and Extra small both use the small tree, because Fluent's Tree has no extra-small size.
    - The Yes/No switch, choice badges and lookup cells shrink to fit the smaller rows.
  - Saved with the display settings. A saved value from an older version, or an unknown one, falls back to Medium.

- **Close dialogs for opportunity, case, quote, order and invoice (#149, verified live on temporary records):** in the results command bar, **Close** replaces Deactivate for these entities.
  - **Opportunity:** Close as won / Close as lost (`WinOpportunity` / `LoseOpportunity`), with actual revenue (left empty, each opportunity keeps its own), close date, description and, for lost, a competitor that suggests Out-Sold.
  - **Case:** Resolve case (`CloseIncident`), with resolution, resolution type, billable time and remarks. **Cancel case** is a plain state change to Cancelled; Merged isn't offered.
  - **Quote:** Win quote / Close quote (`WinQuote` / `CloseQuote`). Lost is preselected rather than the metadata default Revised.
  - **Order:** Fulfill order / Cancel order (`FulfillSalesOrder`, bound to the salesorder collection / `CancelSalesOrder`).
  - **Invoice:** Mark as paid (Complete or Partial) / Cancel invoice, plain state changes; invoices have no close message.
  - **How it runs:**
    - Status reasons come from the target state in metadata. Each dialog starts with the environment and scope check.
    - The same values go to every selected record. Each record gets one request, and its own close activity (`opportunityclose`, `incidentresolution`, `quoteclose`, `orderclose`). Requests are batched with progress, per-record errors, throttle retry and the connection guard, and the query re-runs afterwards.
    - A batch that stops early doesn't resend the records it already closed.
    - Items are off without Write on the entity or Create on activities.
- **Activate reopens** closed opportunities, cases, orders and invoices (back to statecode 0; verified live). It stays off for quotes and contracts.
- **Query Advisor: Advisories tab and Check with Dataverse (#97, verified live, read-only):**
  - **Advisories tab** between Results and Code, with a count on the tab. One list of what Dataverse will reject, might reject, and performance advice, with filters. Each row has an explanation, a Microsoft Learn link, and its location, which selects the node in the tree. Advice never blocks Execute. Validate's messages get **View all in Advisories**; the tree, the properties panel and Validate are unchanged.
  - **New checks that use metadata** (skipped while metadata isn't loaded):
    - order on a choice, status, status reason or Yes/No attribute (not with `useraworderby`);
    - any condition on a large text attribute: multiline text, or text longer than 850 characters;
    - a condition on a calculated or formula attribute (not rollup);
    - a paged query without the primary key in its orders, with the quick fix **Add {primary key} as the last order**. It isn't offered when link-entities have orders, because Dataverse applies those after the entity's orders.
  - **Check with Dataverse**, on demand only: sends the FetchXML you're looking at to `ValidateFetchXmlExpression`. Matching findings get **Dataverse agrees**; anything else gets a **From Dataverse** row quoting Dataverse's message. The status line says when the query has changed since the check. Off without a connection.
  - Attribute metadata now includes `SourceType`, and String `MaxLength` is read once per entity (cached).

### 🔧 Changed

- **Close dialogs say how to undo too (#69).** As in 1.5.0-beta.18, for the Close actions added in 1.6 (#149):
  - **Win / Lose opportunity, Resolve case, Fulfill / Cancel order:** "To undo, run Activate on the records this changes. The close activities stay." Activate reopens them; a reopened opportunity's close activity becomes Canceled.
  - **Win / Close quote:** "There's no undo: won or closed quotes can't be reactivated here."
  - **Cancel case, Mark as paid / Cancel invoice** (plain state changes): "To undo, run Activate…".
  - **Activate** on an entity with a Close menu says to undo with Close, not Deactivate.

### 🏗️ Technical

- **Relative asset URLs (#173).** The build used Vite's default root-absolute `base` (`/assets/…`). That works in PPTB Desktop, which serves each tool at the root of its own origin. The PPTB VS Code extension serves `dist/` from a sub-path, where those URLs miss the tool's folder; the same problem left Command Bar Studio as a blank tab there (Command Bar Studio #89). The build now uses `base: "./"`, so the page, its chunks, the stylesheet and the editor font load from the tool's own folder wherever it's served.
- **Faster startup: exceljs loads on the first local export (#187).** The local Excel export's library was in the main bundle, so every start parsed it, though most sessions never export. It's now a separate chunk, loaded when you first use Export (Local), alongside the formatting metadata the export already waits for. The main bundle shrinks from 2,576 kB to 1,631 kB (729 kB to 455 kB gzipped). The large-result-set items in #187 stay open.
- **Repo hygiene, part of #205.** Removed files nothing imports: `ColumnConfigDialog.tsx`, `XmlTextArea.tsx`, `editors/optionalValue.ts` and the Vite template's `src/assets/react.svg`. Also removed two unused exports: `LoadingCellRenderer` and `ValueNode`, which #181 didn't need. `engines` is now `^22.22.2 || ^24.15.0 || >=26`, the Node versions supported by jsdom 30 (the Vitest environment). The one-time Prettier reformat and the CI format check follow after 1.5.0 ships, so forward merges from `release/1.5` don't conflict. The ESLint 10 upgrade gets its own PR.
- **CI fails on high or critical runtime vulnerabilities (#72):** `npm audit --omit=dev --audit-level=high` runs before the other checks. The README now has CI and license badges.
- **Dependabot (#76):** grouped weekly npm and GitHub Actions updates into `dev`, with a 14-day cooldown so only versions published at least two weeks ago are proposed. React majors are reviewed manually.
- **`CONTRIBUTING.md` and the PR template (#80):** setup, the branch and PR workflow, what every PR includes, the product rules (ground in Microsoft Learn, never refuse supported FetchXML, confirmations name the environment and scope), the UI and wording rules, and the dependency rules. The PR template asks for the target branch, light and dark screenshots for UI changes, and the release-prep checks.

## [1.5.0-beta.18] - 2026-11-01

### 🔧 Changed

- **Confirmations say how to undo the change, or that it can't be undone (#69).** The PPTB [marketplace policy](https://docs.powerplatformtoolbox.com/policies/marketplace) asks a confirmation to offer a way back where feasible and, when there's none, to say so plainly. Delete and Bulk delete already said "This action cannot be undone." Now:
  - **Run workflow:** "There's no undo: whatever the workflow changes stays changed."
  - **Save view, overwriting a view:** "There's no undo: the view's current definition is replaced. To keep it, cancel and use Save as instead." For a managed system view, it says that removing the view's unmanaged layer in Power Apps brings the managed definition back.
  - **Activate / Deactivate:** "To undo, run Deactivate (or Activate) on the records this changes." For an entity that enforces status reason transitions, it adds "if this entity's status reason transitions allow it". When going back offers more than one status reason, it adds that the previous one isn't restored.
- **README:** "What this tool changes" says where each operation's way back is, and that the publisher and solution you pick are remembered per environment on your machine.

### 🏗️ Technical

- **Release workflow hardening (#188).**
  - **Separate jobs:** `release.yml` is split into `build`, `publish` and `github-release`. The job that installs and runs the dependencies (`npm ci --ignore-scripts`, no cache, tests, build) only reads the repository. Only `publish` can get the npm OIDC token, only for a tag push, and it installs no project dependencies: it publishes the tarball `build` packed. Only `github-release` can write to the repository.
  - **More checks before publishing:** the release build now also runs the runtime audit, lint and the package-contents check, so a tag on a commit that skipped CI can't publish.
  - **Pinned versions:** npm is pinned to 11.21.1 instead of `npm@latest`, and Node to 24.22.0. Actions are pinned by commit SHA (checkout v7.0.1, setup-node v7.1.0, upload-artifact v7.0.2, download-artifact v8.0.2). Each is at least 14 days old, the project's dependency rule.
  - **Rehearsal:** the workflow can run by hand as a rehearsal: every check and `npm publish --dry-run`, with no npm token and no GitHub Release. A manual run never publishes, even on a tag.
  - **`latest` can't move back:** a stable tag gets `latest` only for main's own version line, and a `support/X.Y` commit is checked first. Support branches merge forward into `main`, so before this, re-running a support release after that merge would have published it to `latest`.
  - **CI:** CI uses `npm ci --ignore-scripts` and the same pins, doesn't keep the checkout's token, and runs the runtime audit on this branch too.
  - **Minimatch advisory:** the moderate minimatch advisory (GHSA-g3hv-75wg-m249, through exceljs → archiver) stays open for now, tracked in #262. The only fix is minimatch 11.0.1, published 2026-10-27, under the 14-day rule. The copies exceljs uses are minimatch 3 and 5, whose callers break on 11's API. The browser build uses exceljs's prebuilt bundle, which doesn't load archiver at all.

## [1.5.0-beta.17] - 2026-10-31

### 🔧 Fixed

- **Saving a loaded view keeps its lookup display names and exact values (#181, from 1.4.6).** A saved view's `<value uiname="…" uitype="…">` attributes are kept for each value, so model-driven apps still show the names in the view's filter. Long decimals keep every digit, line breaks and tabs in attribute values are written as `&#10;`, `&#13;` and `&#9;`, and spaces in `<value>` text are kept, including in the pac CLI and Web API snippets.
- **A failed relationship load no longer crashes the tool (from 1.4.6).** When the relationship list for a new link-entity couldn't load, the tool crashed instead of showing "Error loading relationships".

### 🔧 Changed

- **Insider icon: the gold corner is bigger, deeper gold and flush with the corner.** The old stamp's right point sat inside the hexagon's edge, leaving a thin blue gap, and the light gold faded at small sizes. The stamp now fills the corner, about 1.5 times larger, in two tones of gold (`#CC8A00` and `#F2A100`) split on the same diagonal as the blue. Prerelease builds only; the release icon is unchanged.

## [1.5.0-beta.16] - 2026-10-30

### 🔧 Fixed

- **Bulk Delete jobs no longer fail with 0x80048d19 "Error identified in Payload provided by the user for Entity :''" (#224, from 1.4.5).** Dataverse rejected "Delete all", "Bulk delete (job)" and deleting more than 100 selected records whenever the query had a condition on a number, Yes/No, choice, status, money or decimal column, such as the "Active …" views. Condition values are now sent as strings, which is what the Web API's [`Object`](https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/object) type defines.
- **Bulk deleting more than about 530 selected records works (#177 first part, from 1.4.5).** The query for selected records is built in the tool instead of being converted through a GET request with every ID in the URL.

### 🔒 Security

- **The editor ships DOMPurify 3.4.17 (from 1.4.5).** Monaco Editor bundles its own copy, which npm `overrides` can't reach, so 1.5.0-beta.15 shipped DOMPurify 3.4.15. The build now uses the npm `dompurify` package at the version `overrides` sets, and CI and release check the built bundle (`.github/scripts/check-dompurify.mjs`).

## [1.5.0-beta.15] - 2026-10-20

### 🔧 Fixed

- **"Delete all" deletes what the query on screen matches (#174, from 1.4.4).** With nothing selected, Delete all and Bulk delete (job) submit the FetchXML that produced the visible results. A quick find view's test value is included. They're unavailable ("Delete all unavailable", with the reason in its tooltip) before a run, after a failed run, or when the tree has changed in a way that matches different records, until you run the query again. Adding, removing or sorting columns in the grid doesn't block it.
- **pac CLI snippet: PowerShell's curly single quotes are doubled too (#196, from 1.4.4),** so a value like `O’Neil` can't end the string.
- **The first page gets the same recovery as later pages (#183).**
  - The first page of results was requested once, with no retry after service protection errors (#96) and no smaller page after the ToolBox's 16 MB response limit (#68). That was the common case #68 was meant for: a 5,000-row first page of a wide Entity, which failed.
  - Now a too-large first page is requested again with half the count (from Dataverse's default 5,000 when the query sets none), and later pages keep that size.
  - Saved views are retried after throttling. Their page size is Dataverse's, so it can't be made smaller.
  - Queries with `top` or `aggregate` are retried but not shrunk.
- **No write resumes after the tool's connection changed (#175).** The connection guard (#66) was checked before each batch, but not in three places:
  - the 1–3 record delete loop;
  - the waits before a throttled `$batch` or request is retried (5–30 s);
  - the older-host delete path.
  
  A retry after an in-place connection change would have run in the new environment, where the same record IDs can exist. Those records now report "Not deleted: the tool's connection changed…" (or "Not changed…" for Activate/Deactivate) instead. Workflow retries stop the same way.

## [1.5.0-beta.14] - 2026-10-19

### 🔧 Fixed

- **Picking a condition's attribute no longer crashes the tool (#217, the 1.4.3 hotfix).** The 1.5 betas had the same bug as 1.4.1 and 1.4.2: choosing a condition's attribute, changing its entity or clearing its attribute blanked the tool with "Cannot read properties of undefined (reading 'replace')".
  - A missing value is now written as empty (`operator=""`). An unfinished condition says "Choose an attribute" / "Choose an operator", and the tree shows "name (no operator)".
  - The results grid's column filter chips no longer throw when a condition's operator is cleared.
  - A new test blanks every string field of every node and runs every pure consumer (91 cases).
  - Verified in Power Platform ToolBox Desktop. Merged forward from `main` (1.4.3).

## [1.5.0-beta.13] - 2026-10-18

### 🔧 Fixed

- **pac CLI snippet: the shell no longer expands FetchXML values (#166, found by CodeQL).** The XML was wrapped in double quotes with only `"` escaped, so bash and zsh expanded values such as `$100`, `$(…)` or backticks when the snippet was pasted, and PowerShell rejected `\"`. The snippet now single-quotes the XML on a **bash / zsh** line (`'` written as `'''`) and a **PowerShell 7.3+** line (`'` written as `''`), and points to `--xmlFile` for Windows PowerShell 5.1. It uses `pac env fetch`: `pac org` was renamed `pac env` in pac 1.32, and `org` still works.

### 🏗️ Technical

- **Release branches (`RELEASING.md`).** A minor release now stabilizes on its own `release/X.Y` branch, cut from `dev` at the feature freeze, so `dev` is never frozen. Betas (`-beta.N`, `-rc.N`) are tagged there and publish to npm `beta`. A stable release merges `release/X.Y` → `main` and is tagged there (npm `latest`). Fixes go to the oldest supported branch that has the bug and merge forward (`main` → `release/*` → `dev`).
- **Supported versions (`SECURITY.md`):** the current stable release gets fixes. The previous minor gets critical and security fixes for 30 days after the next minor ships, from a `support/X.Y` branch that publishes to npm `support-X.Y`, never `latest`.
- **`release.yml` chooses the channel from the tag:** `-beta.N` / `-rc.N` → `beta` (pre-release); a stable tag on `main` → `latest`; a stable tag on `support/X.Y` → `support-X.Y` (not marked as the latest GitHub Release). Any other tag fails before publishing. CI also runs on pushes to `release/**` and `support/**`.
- **Release process: stabilizing before a stable release (`RELEASING.md`).** Each stable release has a milestone and a "Release X.Y.Z" tracking issue. Before a beta is promoted to npm `latest`, the milestone is clear, the latest beta is smoke-tested in PPTB, it soaks for 7 days without a new P0/P1 bug, and CI and the runtime audit are clean. Stabilizing happens on the release branch, so `dev` isn't frozen. The same process applies to Command Bar Studio and Dataverse Request Studio.

## [1.5.0-beta.12] - 2026-10-18

### ✨ Added

- **Saved views follow-ups (#142):**
  - **Shared personal views:** Load view shows "Shared by {owner}" (or "Shared with you"). System views show their query type or "Default view" as a second line.
  - **Write access:** before overwriting a personal view you don't own, Save view checks your access (`RetrievePrincipalAccess`). Without write access, it offers **Save a copy…**; if access can't be determined, Dataverse decides.
  - **Managed system views:** overwriting one shows a warning naming its solution. It's a warning only, not a block.
  - **Quick find views:** enter a test value for `{0}`, with an optional trailing `%`, to execute them. The value is used only in the FetchXML that runs; the tree, the editor and saved views keep `{0}`. A loaded quick find view runs as FetchXML, not by ID, because by ID it matches nothing (verified live).
  - **Set as default view** for public system views.
  - **Description:** pre-filled when you overwrite a view, and clearing it clears the view's description.
- **Docs-driven helpers (#98):**
  - **Wildcard helper** for `like` / `not-like` values: Contains (`%x%`), Begins with (`x%`), Ends with (`%x`) or a custom pattern. **Match literally** escapes `%` → `[%]`, `_` → `[_]` and `[` → `[[]`, per "Wildcard characters".
    - begins-with and ends-with warn when a typed character would act as a wildcard.
    - Stored values read back into the helper, and switching operators keeps the text.
    - The Query Advisor's leading-wildcard advice ignores escaped characters, and now also appears in the results-grid column filter.
  - **Total record count:** with *Return total record count*, the info bar shows "Rows: 50 of 1,234 total", or "5,000+ total" (500+ on elastic entities) when the limit is exceeded. **Count all {entity} records (unfiltered snapshot)** calls `RetrieveTotalRecordCount` and says plainly that it ignores the query's filters.
  - **Hierarchy picker:** under / not-under / above / eq-or-under / eq-or-above on the primary key pick a record by name (`startswith` search, so no leading wildcard). The record's name is resolved back from its ID, and **Enter an ID instead** gives a GUID field with validation. The hint names the hierarchical relationship and the 100-level limit.
  - **GUID values:** Lookup, Customer, Owner and Uniqueidentifier condition values that aren't GUIDs are flagged. Verified live: Dataverse rejects them with 0x80040203.
  - **Undocumented settings:** "UTC offset (undocumented)" shows only when the query sets it, and the condition aggregate leftover reads "Aggregate (undocumented)".

### 🔧 Fixed

- **Record actions use the current selection after the results refresh (#158).** Previously, after a refresh (re-running the query, or the automatic refresh after Delete, Activate or Run workflow), the grid showed nothing selected, but the command bar kept the previous selection. Its actions stayed enabled and ran on the previously selected records. The command bar now clears its selection with the grid's.

### 🏗️ Technical

- **Release process (`RELEASING.md`):** `main` holds the current stable release (npm `latest`); betas are tagged on `dev` (npm `beta`, GitHub pre-release); hotfixes branch from `main` and merge back into `dev`. PRs add CHANGELOG entries under `[Unreleased]` and don't change the version, which is set when a release is cut. The same process applies to Command Bar Studio and Dataverse Request Studio.
- **`release.yml` refuses a stable tag whose commit isn't on `main`,** so a stray tag on `dev` can't replace npm `latest`. Beta tags (`-beta.N`) are allowed from `dev`.

## [1.5.0-beta.11] - 2026-10-17

### ✨ Added

- **Activate and Deactivate work, driven by metadata (#143, verified live on temporary records).**
  - **Simple entities** (Account, Contact…): the state's default status reason is applied.
  - **Multi-state entities** (Lead, activities, custom state models): you pick the status reason, grouped by state. When the entity enforces status-reason transitions, only targets allowed from every selected record's current status are offered.
  - **Lead:** Deactivate defaults to Disqualified. Choosing Qualified warns that this doesn't run QualifyLead; Dataverse accepts the plain update.
  - **Opportunity, Case, Quote, Order, Invoice and Contract** close through their own messages (WinOpportunity / LoseOpportunity, CloseIncident, WinQuote, FulfillSalesOrder…). For these the buttons are disabled, with a tooltip naming the documented action. Proper dialogs are tracked in #149.
  - **Also disabled, with a reason:** entities with no state, results with one-to-many joins, and missing Write privilege. The Write privilege name comes from metadata, e.g. `prvWriteActivity` for activities.
  - **How it runs:** a PATCH of `statecode` + `statuscode`, sent through `$batch` in chunks of 25, with per-record errors, throttle retry, progress and the close warning. The query runs again afterwards.
- **Confirmations name the environment and the scope (#69, Marketplace policy).** Delete, Bulk delete, Run workflow, Save view and Activate/Deactivate show:
  - the environment name, type and URL, with a warning for **Production**;
  - the entity's display and logical name;
  - the exact scope. A Bulk delete with no selection says "Every record the query matches" and never claims a total it doesn't know.

  **Type-to-confirm:** on Production, for more than 100 records, or when the total is unknown, you type the entity's logical name before the action enables.

### 🔧 Fixed

- **Bulk delete no longer shows a stray horizontal scrollbar (#152).** The "Danger zone" warning was 1px wider than the dialog, a grid item's default `min-width`. Its title now reads "Bulk delete all matching {entity}" or "Bulk delete N records".

### 🔧 Changed

- **Sentence case** in the delete, bulk delete, workflow and save view dialogs and the results command bar, completing the sweep from beta.10: Delete record, Run workflow, Save and publish, Personal view, Bulk delete (job), Delete all, Retrieve all stopped.

### 🏗️ Technical

- New `model/recordState.ts` (state planning from metadata, transition intersection, close-message entities) and `shared/components/TargetScope.tsx` with `targetScopeRules.ts`, all with tests.
- The `$batch` runner in `batchDelete.ts` is shared by delete and the state update.

## [1.5.0-beta.10] - 2026-10-17

### ✨ Added

- **Collapsible query scope section (#25).** It collapses to one line, "**Account** · View: Active accounts · Contoso Core", with **New** always visible. The collapsed state is remembered.
- **The publisher and solution selection is remembered per environment (#25),** keyed by environment URL in the ToolBox settings. It's restored once the lists load; publishers or solutions that no longer exist are dropped.
- **Right-click a tree row (#24).** Right-click, Shift+F10 or the context-menu key opens the row's Add items, a divider, then **Delete** in red. The root entity has no Delete.

### 🔧 Changed

- **Tree (#24, #25):**
  - The Add menus are reordered, renamed and given the tree's icons: All attributes, **Select attributes…** (now second), **Blank attribute**, Order by, Filter, Link-entity.
  - Deleting a node with items inside asks first: "Delete filter (and)? This also deletes the 2 items inside it."
  - Rows with nothing inside have no chevron.
  - A link-entity that can return attributes but has none shows "(no attributes)".
  - A link-entity without an alias is labelled by its lookup, e.g. "contact via primarycontactid (outer)". The lookup is `to` for N:1 and `from` for 1:N.
- **Relationship picker (#25):** each option has two lines, the related entity, then "via {lookup} ({logical}) · {schema name}". The groups are Many-to-one, One-to-many and Many-to-many, search matches any part, and the list is wide enough that options stay on two lines.
- **Publishers and solutions (#25)** show two lines: name, then `uniquename · vX` for solutions and `uniquename · prefix` for publishers. The empty states are worded consistently.
- **Sentence case and "…" across the UI,** matching Dataverse Request Studio and Command Bar Studio. For example: Load view, System views, Entity source, All entities, Advanced Find only, Use logical names, Editor mode, Parse to tree, Send to tool, Save as….

### 🏗️ Technical

- New `model/scopeSelection.ts`, `model/treeLabels.ts` and `model/relationshipOptions.ts`, and the `useRememberedScope` hook, with tests.

## [1.5.0-beta.9] - 2026-10-17

### 🏗️ Technical

- **monaco-editor 0.54 → 0.57 (#145)** clears the moderate `dompurify` advisories.
  - `@monaco-editor/react` stays at 4.7.0, because 4.8 is still a prerelease. An ambient type declaration maps its `editor.api` import to Monaco 0.57's `monaco-editor/editor` entry.
  - **Smaller bundle:** Monaco loads only the editor core, its features and the languages FXS shows (XML, C#, JavaScript, shell and SQL), using the tree-shakeable entry points added in Monaco 0.56. With the full 0.57 entry, the package would have exceeded its size budget. Now it's 13 files instead of 98, and slightly smaller overall.
  - **No worker console errors:** the editor worker runs on the main thread from the start, as before (PPTB's CSP blocks `blob:` workers), without the errors Monaco 0.57 would log first.
  - **Remaining audit findings:** two low `dompurify` advisories (fixed in 3.4.16). Monaco bundles its own copy and doesn't use the affected `IN_PLACE` mode.
- **`uuid` for exceljs pinned to 11.1.1 with an npm override (#145, GHSA-w5hq-g745-h8pq).** The browser build of exceljs bundles its own uuid and only calls `v4()`, so the advisory never applied to the tool.
- **A round-trip test for local Excel export:** headers, typed values, formats, dates, labels, the "both" (Raw) columns, the frozen header and auto-filter.

## [1.5.0-beta.8] - 2026-10-17

### ✨ Added

- **Query Advisor (#97).** Before you run a query, FetchXML Studio flags the performance anti-patterns Microsoft Learn documents, as notes in Validate, Parse to Tree and the property panel. **Advice never blocks Execute.** It covers:
  - leading wildcards in `like` / `begins-with` values, and `ends-with`;
  - `<all-attributes />` or no attributes selected;
  - sorting by a link-entity's attributes, which is slower and can lose the paging cookie;
  - `distinct` without an order;
  - simple paging past 50,000 records;
  - aggregate queries with no filter (the 50,000-record aggregate limit);
  - the legacy `no-lock`.
- **Query limits, verified live (#97):**
  - **More than 15 joined link-entities** blocks Execute (0x8004430D). Link-entities inside filters don't count.
  - **More than 500 conditions** blocks Execute (0x8004430C). Microsoft Learn says conditions and link-entities count together, but live, only conditions count: 500 conditions plus a link-entity run.
  - **URL-encoded FetchXML over 32,768 characters** gets a warning with "Execute anyway". Live, about 44,500 characters failed with HTTP 414, but the exact limit depends on the URL.

### 🏗️ Technical

- New `model/queryAdvisor.ts`, called from `validateQueryRules`, with 91 table-driven tests. Every advice case is also checked to block nothing.
- The condition Operator and Value fields, the aggregate and distinct options and the order Attribute field now show their advice notes, through the existing `findingFor`.

## [1.5.0-beta.7] - 2026-10-17

### 🔧 Fixed

- **Load View labels querytype 2048 as "Other"** (`SavedQueryTypeOther`). It used to read "Template", a label that belongs to the template types (#144).
- **PPTB events no longer log their payloads to the console (#71).** Connection events carry environment details. Event names are logged only when debug logging is on.

### 🏗️ Technical

- **Safe OData literals (#81).** Every value placed inside a single-quoted OData string literal now has its quotes doubled (`odataString`). That covers entity and attribute names in metadata queries, solution unique names, and the workflow and view filters. A value can no longer end the literal early.
- **The duplicate LayoutXML helpers are removed from `pptbClient.ts` (#81).** They were unused; `model/layoutxml.ts` is the only implementation.

## [1.5.0-beta.6] - 2026-10-17

### 🔧 Fixed

- **Column headers for one-to-many link-entities name the related entity (verified live).**
  - A 1:N link-entity's column read "Full Name (Account)": the header used the parent's primary key, which is the link's `to` attribute.
  - It now reads "Full Name (Contact)". The entity's display name comes from the cached entity lists, so it's there before the entity has been opened.
  - N:1 columns are unchanged and still name the lookup ("Full Name (Primary Contact)").

## [1.5.0-beta.5] - 2026-10-17

### 🏗️ Technical

- **Browser dev harness against a live org: `npm run dev:harness` (#78), modelled on Command Bar Studio #44.**
  - **What it runs:** the real UI, in a browser, against the Dataverse org in `DATAVERSE_URL`.
  - **How requests flow:**
    - A stub of the PPTB host makes requests the same way PPTB 1.2.7 does: the same entity-set names, `execute()` encoding, Prefer headers and error text, plus optional `CallerObjectId` impersonation.
    - A localhost-only Vite proxy attaches an Azure CLI token on the server side. The token is kept in memory only and is never logged or sent to the browser.
    - The proxy only accepts requests from localhost with a harness header, and forwards only to the org's Web API over https.
  - **Not shipped:** the harness is dev-only. The production build is byte-identical without it, and it isn't in the package. See `dev/README.md`.
- **Dev tooling updated, clearing every high npm audit finding (#77):**
  - versions: vite 7.3.6, @vitejs/plugin-react 5.2.0, typescript-eslint 8.71.0, eslint 9.39.5, plus transitive fixes (rollup, postcss and others);
  - audit: 25 findings (17 high) → 4 moderate, 0 high;
  - every changed package was published at least two weeks ago;
  - no runtime dependencies changed.
  - **Not bumped:** monaco-editor 0.57. Its `exports` map breaks `@monaco-editor/react` 4.7's types, so it needs its own change.

## [1.5.0-beta.4] - 2026-10-17

### ⚠️ Requires Power Platform ToolBox 1.2.7

- **PPTB 1.2.7 is now required (#82).**
  - `@pptb/types` is 1.2.7.
  - `features.connections: 1`, and `minAPI` is 1.2.7 (it was 1.2.2).
  - Older hosts and `1.2.7-dev` Insider builds show the tool as unsupported.

### ✨ Added

- **Changing the tool's connection in place resets everything that came from the old environment (#66).**
  - The metadata cache, entity lists, publishers and solutions, privileges, WhoAmI, results and the loaded view are all cleared, and the UI reloads for the new connection. The query is kept.
  - A notification says what happened.
  - Token refreshes and edits to other connections reset nothing. Deleting another connection no longer marks the tool as disconnected.
  - Workflow runs and multi-record deletes stop if the connection changes partway through.
- **The ToolBox warns before closing while work is running (#70):** Retrieve All, deletes, workflow runs and Save view (`preventClose`).
- **Record links open in the connection's browser profile (#70).** This covers Open record and the bulk delete job link (`openInConnectionBrowser`).

### 🔧 Fixed

- **Deleting records uses the entity's metadata entity set name (#61).** This fixes entities whose entity set name the host guesses wrongly.
  - Multi-record deletes are sent as real OData `$batch` requests of 25 (`executeBatch`).
  - Errors are reported per record: a batch Dataverse rejects is retried one record at a time, so one failure doesn't fail the rest.
- **Clipboard writes try the ToolBox first (#70),** then the browser clipboard. That applies to copying results, record URLs, FetchXML and code.

### 🏗️ Technical

- New `api/host.ts` adapter. Every host API is feature-detected, with a fallback.
- New `api/batchDelete.ts` and `shared/hooks/connectionIdentity.ts`.
- `metadataCache.clear()` now clears every cache, with a generation guard so requests still running for the old connection can't refill it.

## [1.5.0-beta.3] - 2026-10-17

### ✨ Added

- **Dataverse errors explain what to do (#96).** For documented error codes, the error bar, paging notices, export, delete, bulk delete, workflow and Save view messages start with what to do. Dataverse's own message follows after "Details:". Undocumented codes are shown unchanged. Covered:
  - service protection limits (429, 0x80072322/21/26);
  - the 50,000-record aggregate limit (0x8004E023);
  - the quick find record limit;
  - more than 500 conditions or 15 link-entities;
  - query throttling and timeouts caused by leading wildcards, calculated attributes or other anti-patterns;
  - aggregate mistakes;
  - invalid condition values;
  - missing attributes or entities;
  - entity set names Web API doesn't recognise;
  - records not found;
  - missing privileges.
- **Throttling back-off (#96).** When Dataverse service protection limits are hit, deletes, workflow runs and page loads wait and retry: per `Retry-After` when it's given, otherwise 2, 4 and 8 seconds. Each wait is capped at 60 seconds, with at most 3 retries.
- **Pages over Power Platform ToolBox's 16 MB response limit (#68).**
  - **Load more and Retrieve All:** the page is requested again at half the size, at most three times and never below 100 rows. The pages that follow keep the smaller size.
  - **The first page and saved views:** the error explains how to make the response smaller (fewer attributes, or a smaller `count`), and nothing is retried.

### 🔧 Changed

- **Fewer startup calls and a quiet console (#71).**
  - WhoAmI and each privilege check are sent once per connection, and concurrent callers share one request.
  - Raw query results, privilege responses, payloads and routine status no longer go to the console. They're behind the debug flags (`src/shared/utils/debug.ts`), which are now all off by default.
  - The Electron IPC prefix is removed from errors shown to users.

### 🏗️ Technical

- New modules, each with tests:
  - `api/dataverseErrors.ts`: reads the PPTB 1.2.7 host's error formats and maps documented codes to Microsoft Learn;
  - `api/throttleRetry.ts`;
  - `api/sessionCache.ts`: `resetDataverseSessionCache()`, to be called on a connection change (#66);
  - `api/pageRecovery.ts`.

## [1.5.0-beta.2] - 2026-10-17

### 📝 Documentation

- **README: "What this tool changes" (#69, Marketplace policy).** Lists every operation that writes to Dataverse: Delete, Bulk Delete, Run workflow, and saving personal and system views. For each it says when it runs, what it does and which privilege it needs. It also lists what doesn't write: both exports and the locally stored settings.
  - It states plainly that Delete with no rows selected submits a Bulk Delete job for every record the query matches, not just the loaded rows.

## [1.5.0-beta.1] - 2026-10-17

### 🔧 Fixed

- **Cross-apply columns stay filled when they're filtered (#91 follow-up, verified live).**
  - A `matchfirstrowusingcrossapply` link-entity normally returns its attributes by schema name (`FullName`).
  - A condition with `entityname` targeting that link-entity, such as a column filter from the grid, makes Dataverse join it normally. It then returns `mc.fullname` instead.
  - The grid now uses whichever key Dataverse returned, including for cross-apply `<all-attributes />`, so a filtered cross-apply column no longer shows as empty.
- **Two cross-apply queries Dataverse rejects are now flagged before they run (verified live 2026-10-17).**
  - An order inside a `matchfirstrowusingcrossapply` link-entity (0x80041121). Order from the entity with `entityname` instead; that works.
  - An aliased attribute in one, which fails with a generic SQL error (0x80044150) unless a condition with `entityname` targets the link-entity. With such a condition the alias works, so that case isn't flagged.
## [1.4.6] - 2026-10-31

Hotfix for 1.4.5: saved views keep their lookup display names and exact values, and a failed relationship load no longer crashes the tool.

### 🔧 Fixed

- **Saving a loaded view keeps its lookup display names and exact values (#181).**
  - A saved view's `<value uiname="Contoso" uitype="account">{id}</value>` lost `uiname` and `uitype` when the query went through the tree. Loading a view and saving it again stripped the names model-driven apps show in the view's filter. They're kept now, for each value; changing or removing a value drops only its own.
  - Long decimals kept only 10 decimal places (`0.12345678901234` became `0.123456789`). Every digit is kept now.
  - Line breaks and tabs in an attribute value turned into spaces. They're written as `&#10;`, `&#13;` and `&#9;` now.
  - Spaces in `<value>` text were trimmed, including in the pac CLI and Web API snippets. They're kept now.
- **A failed relationship load no longer crashes the tool (found testing #201).** When the relationship list for a new link-entity couldn't load (a missing privilege, throttling, the network), the tool crashed instead of showing "Error loading relationships".

### 🏗️ Technical

- `ConditionNode.valueAttributes` keeps each `<value>` element's own attributes; `compactFetchXml()` makes the one-line FetchXML for pac and the Web API without touching `<value>` text.
- `fetchxmlRoundTrip.test.ts` and `fetchxmlCodeGenerators.test.ts` cover each case.

## [1.4.5] - 2026-10-30

Hotfix for 1.4.4: Bulk Delete jobs are accepted again, and the editor ships a current DOMPurify.

### 🔧 Fixed

- **Bulk Delete jobs no longer fail with 0x80048d19 "Error identified in Payload provided by the user for Entity :''" (#224).**
  - "Delete all", "Bulk delete (job)" and deleting more than 100 selected records submit a `BulkDelete` job. Dataverse rejected the job whenever the query had a condition on a number, Yes/No, choice, status, money or decimal column, such as the common "Active …" views.
  - The cause: the query sent with the job had those condition values as JSON numbers and booleans. The Web API's [`Object`](https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/object) type, used for condition values, defines `Value` as a string.
  - Every condition value is now sent as a string, with its type unchanged. Checked live for whole numbers, Yes/No, decimals, `in` lists, dates, lookups and conditions on related tables.
- **Bulk deleting more than about 530 selected records works (#177, first part).** The job's query was converted from FetchXML by a GET request with every record ID in the URL, which went over the URL length limit. The query for selected records is now built in the tool, so the number of IDs no longer matters. (The second part of #177, about the fallback without the Bulk Delete privilege, is planned for 1.6.)

### 🔒 Security

- **The editor ships DOMPurify 3.4.17.** Monaco Editor bundles its own copy of DOMPurify, which npm `overrides` can't reach, so 1.4.4 shipped DOMPurify 3.1.7. The build now uses the npm `dompurify` package in its place, at the version `overrides` sets (3.4.17: GHSA-cg2g-xh64-v4q7, GHSA-mrv2-3pqr-j5jr, GHSA-c6gr-qvv4-5v3f). Same approach as Command Bar Studio 0.6.2 and Dataverse Request Studio.

### 🏗️ Technical

- `bulkDeleteQuery.ts`: `withStringConditionValues` (string condition values for converted queries) and `selectedRecordsQuery` (the selected-records query, built locally), with tests.
- `vite.config.ts`: the `monacoDompurify()` plugin resolves Monaco's DOMPurify import to the npm package.
- CI and release check the built bundle's DOMPurify version (`.github/scripts/check-dompurify.mjs`) and fail on a mismatch.

## [1.4.4] - 2026-10-19

Hotfix for 1.4.3: record actions act only on what you see, and the pac CLI snippet is quoted safely.

### 🔧 Fixed

- **"Delete all" deletes what the query on screen matches (#174).**
  - With nothing selected, "Delete all" and "Bulk delete (job)" submit a Bulk Delete job for a query. They used the query tree as it was at that moment. If you edited a filter in the left pane after running the query, the grid kept showing the old rows, but the job would delete what the edited query matched. It also worked before anything had run.
  - Now these actions submit the query that produced the results on screen. They're unavailable, with "Delete all unavailable" explaining why, until you run the query again whenever the tree has changed in a way that matches different records.
  - Changes you make in the results grid don't block it: adding or removing columns, adding related columns and sorting keep the same records, and column filters re-run the query.
  - Deleting selected records is unchanged.
- **Record actions use the current selection after more results load (#158, backported from 1.5).** When the next page loaded (scrolling, or Retrieve All), the grid cleared its selection, but Delete, Activate and Run workflow still had the previous record IDs. They now clear together.
- **The pac CLI snippet no longer lets the shell expand FetchXML values (#166, #196).**
  - It used to wrap the XML in double quotes, so bash and zsh expanded `$100`, `$(…)` or backticks, and PowerShell rejected `\"`.
  - It now single-quotes the XML, on a **bash / zsh** line and a **PowerShell 7.3+** line.
  - PowerShell's curly single quotes (`‘’‚‛`) are doubled too, so a value like `O’Neil` can't end the string.
  - It uses `pac env fetch`: pac 1.32 renamed `pac org` to `pac env`, and `pac org` still works. It points to `--xmlFile` for Windows PowerShell 5.1.

### 🏗️ Technical

- **Tests:**
  - `deleteScope.test.ts`: 16 scenarios of which edits change the records "Delete all" would delete.
  - `fetchxmlCodeGenerators.test.ts`: the pac snippet, with the quoted XML run through a real POSIX shell and PowerShell when available, including curly-quote injection attempts.

## [1.4.3] - 2026-10-19

Hotfix for 1.4.1 and 1.4.2.

### 🔧 Fixed

- **Picking a condition's attribute no longer crashes the tool.** In 1.4.1 and 1.4.2, choosing an attribute for a condition blanked the whole tool with "Cannot read properties of undefined (reading 'replace')". Changing a condition's entity or clearing its attribute did the same.
  - **Cause:** picking an attribute rightly clears the operator, because the new attribute's type may not support it, and changing the entity clears the attribute. The FetchXML serializer, which runs on every change for the preview, escaped those cleared values without checking for them.
  - **Fix:** a missing value is now written as empty (`operator=""`), so the preview shows the condition as unfinished, and its editor still asks for the attribute and operator.
  - Reproduced and verified in Power Platform ToolBox Desktop against a live environment.
- **An unfinished condition says what's missing.** It reads "Choose an attribute" or "Choose an operator", where it used to say "The 'undefined' operator needs a value". The tree shows it as "name (no operator)" instead of "name undefined".
- **Results grid:** the column filter chips no longer throw when a filter condition's operator has been cleared in Properties.

### 🏗️ Technical

- **New test:** every string field of every node in a rich query is blanked (undefined and empty), one at a time. Each variant then goes through the serializer, the query rules, the column and filter helpers, a re-parse and every code generator: 91 cases.

## [1.4.2] - 2026-10-17

The v1.4.2 milestone: the features other users asked for, a results grid that works from the query, saved views that behave as documented, operators that follow the docs, and the CI/CD and release setup shared with Command Bar Studio and Dataverse Request Studio. It was iterated through `1.4.2-beta.1` to `1.4.2-beta.7` (detailed below). Behaviour follows Microsoft Learn and the live operator matrix. Valid FetchXML is never refused, and a Dataverse error is shown as it is.

### ✨ Added

- **Copy results to the clipboard (#50).**
  - Right-click a cell to copy its value, its raw value (GUID, number, ISO date or option value), the row, the selected rows or the column.
  - **Ctrl+C** copies the selected rows, or the focused cell.
  - The command bar's **Copy** menu copies the selected or all loaded rows, as shown or raw.
  - Rows copy as tab-separated values with headers, so they paste into Excel as cells.
- **Column filters that write FetchXML (#23).**
  - Clicking a column header opens a menu: **Sort**, **Filter by…** and **Clear filter**. Shift+click still adds the column to a multi-column sort.
  - **Filter by…** uses the condition editor's operators and value pickers, adds a `<condition>` and runs the query again, so the filter covers every record.
  - Active filters show as removable tags above the grid.
  - Column filters are kept in their own `<filter type="and">`, apart from the conditions you built. Link-entity columns use `entityname`.
- **Quick search over the loaded rows (#23).** It filters the loaded rows as you type, in any column, using the values as shown, and highlights matches.
- **Operator help (#95).** The Operator info popover explains the selected operator and links to its doc page: the hierarchy operators, the hierarchy security operators, Contains Values and Equals User Language.

### 🔧 Fixed

- **Results grid columns come from the query (#91).** They used to come from the first row, which Dataverse returns without its null values.
  - Lookups that are empty in row 1 now show their values.
  - `<all-attributes />` columns from every loaded page are shown.
  - `matchfirstrowusingcrossapply` columns appear and can be sorted and filtered.
  - The local Excel export has exactly the grid's columns.
- **Saved views (#93).**
  - An unmodified loaded view runs by its view ID, so Retrieve All pages through it.
  - Personal views shared with you appear in Load View.
  - Every type of view is listed, each labelled.
  - Saved layoutxml uses the documented `jump`, `object` and row `id`.
  - System views are checked with `ValidateSavedQuery`, and views that aren't customizable can't be overwritten.
  - **Save and Publish** publishes the view's entity, the documented `PublishXml` form. The `<savedqueries>` node it used to send isn't in the publish request schema. This also publishes any other pending customizations of that entity (#93 h).
  - Load View reports views it can't load or had to change.
- **Operators per attribute type (#95).** Memo attributes gain Equals, Not Equals, In and Not In, as do choices (multi-select), which use the option picker. All are documented, verified live, or both. No operator was removed.
- **Tree (#24, #25).**
  - Selecting a node no longer scrolls the tree to the bottom.
  - A node added from an Add menu is selected and scrolled into view.
- **Header sort for aggregate and aliased link-entity columns (#92)** was confirmed fixed by #114, and is now covered by tests.

### 🏗️ Technical

- **CI/CD and releases (#130).**
  - CI runs typecheck, lint, tests, `pptb-validate`, the build and a package-contents check on every PR and on pushes to `dev`/`main`.
  - Releases come from `v*` tags: npm trusted publishing with provenance, and GitHub Releases with generated notes (see `RELEASING.md`).
  - Prettier config, `.gitattributes` (LF), `engines.node >=22`.
- **Security (#77).** `tmp` was updated to 0.2.7, which clears the high-severity runtime `npm audit` finding.
- **Tests.** 637 tests, up from 507. New suites:
  - copy (`copyResults`);
  - column filters (`columnFilters`);
  - search highlighting;
  - the operator matrix;
  - grid columns (`resultColumns`);
  - saved views (`savedViews`) and layoutxml.

## [1.4.2-beta.7] - 2026-10-17

### 🏗️ Technical

- **The shared CI/CD and release standard, same as Command Bar Studio and Dataverse Request Studio (#130).**
  - **CI** (`ci.yml`) runs on every PR and on pushes to `dev`/`main`: typecheck, lint, tests, `pptb-validate`, the build, then `check-pack.mjs`. That last step checks the package contains the required files (including `pptb.config.json` and `dist/icons/icon.svg`), contains no `src/`, `.map` or `.github/` files, and stays within a size budget of 3 MB packed / 10 MB unpacked.
  - **Releases** (`release.yml`) come from `v*` tags:
    - the tag must match `package.json`;
    - npm publishing uses trusted publishing with provenance, and `-beta.N` versions go to the `beta` dist-tag;
    - the GitHub Release has generated notes, and betas are marked as pre-releases.
    - See `RELEASING.md`.
  - **Repository tooling:**
    - Added Prettier (`.prettierrc`, `format`/`format:check`), `.gitattributes` (LF), a `typecheck` script (`tsc -b`) and `engines.node >=22`.
    - Removed the `coverage` and `finalize-package` scripts and `@vitest/coverage-v8`.
    - Shorter PR template; the issue forms are trimmed to Bug and Feature.
- **`tmp` updated to 0.2.7 in the lockfile, which clears the high-severity runtime `npm audit` finding (#77).** It wasn't exploitable in the tool: exceljs's browser build doesn't include tmp. The minimatch and brace-expansion findings were already fixed.

## [1.4.2-beta.6] - 2026-10-17

### 🔧 Fixed

- **Saved views (#93).**
  - **An unmodified loaded view runs by its view ID again**, so Retrieve All pages through it with `@odata.nextLink`. The stored FetchXML is now compared in the tree's generated form, so default attributes such as `descending="false"` no longer cause a mismatch. The comparison is exact; it used to strip whitespace, which hid edits inside attribute values.
  - **Personal views shared with you appear in Load View.** Before, only views owned by you or your teams were listed. Dataverse applies access rights when reading `userquery`, per Saved queries on Microsoft Learn.
  - **Load View lists every type of view**, each labelled. Public views come first (default on top), then Advanced find, Associated, Quick find, Lookup and the others. Overwriting keeps the view's query type; Save As creates a public view.
  - **Saved layoutxml now follows the documented form.**
    - `object` (the entity type code) and the row `id` (the primary key) come from metadata.
    - `jump` is the primary name attribute; it used to be the first column, which could be a lookup.
    - A loaded view keeps its own `jump` and `imageproviderfunctionname`.
    - A query without a column layout gets one generated, instead of an empty layoutxml.
    - Attribute values are escaped.
  - **Save View dialog:**
    - High and critical performance messages show as errors (saving still isn't blocked).
    - A validation call that fails now says so, instead of showing "0 validation messages".
    - "FetchXML is valid" now reads "No performance issues found".
  - **System views that aren't customizable** (`iscustomizable` false) can't be overwritten: the dialog explains this and points to Save As.
  - **Load View reports problems:** an error when a view's FetchXML can't be loaded (it used to fail silently), and a warning listing what the tree had to change.

### 🏗️ Technical

- System views are checked with `ValidateSavedQuery` before they're saved. It's bound to the `savedqueries` collection (`savedqueries/Microsoft.Dynamics.CRM.ValidateSavedQuery`, per the Web API reference); it used to be called unbound, and in fact was never called. A rejection shows in the existing Save failed bar. There's no `userquery` binding, so personal views aren't validated this way.
- New `model/savedViews.ts` (`isUnmodifiedView`, query type labels and order, managed property reading, Load View notices) and `buildViewLayoutXml` in `layoutxml.ts`, with tests.

## [1.4.2-beta.5] - 2026-10-17

### 🔧 Fixed

- **Grid columns come from the query, not from the first row (#91).** Dataverse leaves null values out of each row, so columns used to depend on which values row 1 happened to have.
  - **Lookups:** a lookup that's empty in the first row no longer shows "—" in every row. Raw mode, "both" mode and Copy raw value work too, because the column is keyed `_name_value` from metadata.
  - **Empty attributes:** every attribute in the query is a column, even when no loaded row has a value.
  - **All attributes:** `<all-attributes />`, and a root entity without attribute elements, show every attribute that has a value in any loaded row, including link-entity all-attributes. Later pages add new columns at the end, and existing columns never move.
  - **Cross-apply:** `matchfirstrowusingcrossapply` columns, which Dataverse returns by schema name, now appear, and can be sorted, filtered and removed.
  - **Exists and in:** attributes inside `exists`/`in` link-entities are never grid columns. Those link types return no columns.
  - **Excel export:** local Excel export has exactly the grid's columns, in the grid's order. It no longer adds the primary key on its own.

### 🏗️ Technical

- New `model/resultColumns.ts` (`resultColumns`, `orderByLayout`), with tests. The grid and the local Excel export both use it.
  - `QueryResult` carries the query that ran instead of a precomputed `columns` list.
  - `resolveColumn` maps cross-apply schema-name keys back to their attribute.
- #92 (header sort writing invalid orders) was already fixed by #114. New tests cover aggregate link-entity columns, an aggregate attribute without an alias, and sorting a root lookup from its `_name_value` column.

## [1.4.2-beta.4] - 2026-10-17

### 🔧 Fixed

- **Operators now follow the docs for every attribute type (#95).**
  - Multi-line text (Memo) attributes offer Equals, Not Equals, In and Not In. The operators reference covers them under "String", and they passed the live operator matrix.
  - Choices (multi-select) attributes offer Equals, Not Equals, In and Not In, as listed in the choices doc, alongside Contains Values / Does Not Contain Values. Equals and Not Equals use the option picker.
  - The Operator info popover explains the operator you picked, with a link to its page:
    - the hierarchy operators: primary key of a hierarchical entity only, up to 100 levels;
    - the hierarchy security operators: Hierarchy Depth;
    - Contains Values: full-text indexing, so new records can take a few seconds to match;
    - Equals User Language.

### 🏗️ Technical

- `operators.matrix.test.ts` asserts the operator list for 28 attribute cases. Each row cites the operators reference, the choices doc or the hierarchy doc, or the #87 live matrix.
  - Operators are only ever added on that evidence, never removed.
  - Kept on purpose: user and team operators on user lookups and Uniqueidentifier, because they're verified live.
  - Not offered: Equals Business on Owner (accepted, but it can never match).

## [1.4.2-beta.3] - 2026-10-17

### ✨ Added

- **Column filters that write FetchXML (#23).**
  - Clicking a column header opens a menu: Sort (A to Z / Z to A; smallest to largest for numbers; oldest to newest for dates), **Filter by…**, and **Clear filter** when the column is filtered. Shift+click still adds the column to a multi-column sort.
  - **Filter by…** opens the property panel's condition editor for that column: operators and values from the attribute's metadata.
  - **Apply** adds a `<condition>` to the query and runs it again, so the filter covers every record.
  - Filtered columns show a filter icon. Active filters show as tags above the grid ("Annual Revenue greater than 200000"; choice values by label). Dismiss a tag, or use **Clear all**, to remove filters and run the query again.
- **Quick search over the loaded rows (#23).**
  - The search box above the grid filters the loaded rows as you type, matching any column's value as shown, with matches highlighted.
  - The info bar shows how many loaded rows match.
  - The Copy menu's "All loaded rows" becomes "All matching rows".
  - Rows the search hides are deselected, and scrolling doesn't load more pages while searching.

### 🔧 Changed

- **Clicking a column header opens the header menu instead of sorting straight away.** Sort is the menu's first item. Shift+click still sorts directly.

### 🏗️ Technical

- New `model/columnFilters.ts`, with tests. Column filters live in their own `<filter type="and">`:
  - FetchXML allows one `<filter>` per entity, so the group nests inside the root filter. With no root filter, the group becomes the root filter.
  - An `or` root filter is wrapped in an `and` filter, and unwrapped again when the last column filter is removed. The store keeps the wrapper's ID (`columnFilterWrapperId`), so an `and` filter the user wrote is never unwrapped.
  - Link-entity columns use `entityname` (the alias, or the link-entity's name), per "Filters on link-entity" on Microsoft Learn.
- The builder store tracks the group in `columnFilterId`, with actions `ADD_COLUMN_FILTER`, `REMOVE_COLUMN_FILTER` and `CLEAR_COLUMN_FILTERS`.
  - A FetchXML load (editor sync) relocates the group by position when its conditions match.
  - Loading a view, changing the entity or starting a new query drops the group.
- AppShell runs the query once the filter change has committed.
- `ConditionEditor` takes a `columnFilter` option: the column fixes the attribute, so the Entity, Attribute and Compare-to fields are hidden.
- Search highlighting reads the search text from a context (`SearchHighlightContext`), so typing doesn't rebuild the grid's columns. The text cell renderers wrap their text in `HighlightText`.

## [1.4.2-beta.2] - 2026-10-17

### ✨ Added

- **Copy results to the clipboard (#50).**
  - **Right-click a cell** for: Copy value; Copy raw value (GUID, number, ISO date, option value or true/false, when it differs from what's shown); Copy row; Copy selected rows; Copy column. Copy column gives one value per line, e.g. to paste into an `in` condition.
  - **Ctrl+C** copies the selected rows, or the focused cell when no rows are selected. Text you've highlighted copies as usual.
  - The command bar's new **Copy** menu copies the selected rows or all loaded rows, with values **as shown** or **raw** (GUIDs, numbers, ISO dates).
  - Rows copy as tab-separated values with the grid's headers, in the grid's column order, so they paste into Excel as cells. Values with tabs, line breaks or quotes are quoted the way Excel writes them.
  - A toast confirms what was copied, or says the clipboard couldn't be written.

### 🏗️ Technical

- New `copyResults.ts` (TSV building, value text, clipboard write with a fallback to `toolboxAPI.utils.copyToClipboard`), with tests.
- `ResultsGrid` is now a `forwardRef` exposing `copyRows(scope, mode)`; `PreviewTabs` connects it to the command bar's Copy menu.
- The cell menu is a single controlled Fluent `Menu`, positioned at the mouse. When opened with the keyboard's context menu key, it's positioned at the cell.

## [1.4.2-beta.1] - 2026-10-17

### 🔧 Fixed

- **Selecting a node in the tree jumped to the bottom (#25).** After every selection, the tree scrolled to its *last* item, assuming the selected node was always the newest one. It now scrolls to the selected node, only within the tree, and only when the node is off-screen. Selecting a visible node no longer moves the tree.
- **A node added from the Add menu wasn't selected (#24).**
  - The builder selected the new attribute, order, filter or link-entity, but the menu click also reached the tree row behind it, which selected the row again.
  - The new node is now selected, its properties open, and the tree scrolls to it.

### 🏗️ Technical

- `TreeView.tsx`:
  - the scroll targets `[data-fui-tree-item-value=<node id>]` inside the tree container, with `block: "nearest"`, and its timer is cleared on the next selection;
  - the Add menus' `MenuPopover`s stop click propagation, because React events bubble through portals to the row's `onClick`.

## [1.4.1] - 2026-10-17

The v1.4.1 milestone: correctness against Dataverse, a complete properties panel, and the open bugs. It was iterated through `1.4.1-beta.1` to `1.4.1-beta.20` (detailed below). Every behaviour was checked against Microsoft Learn and a live Dataverse environment: valid FetchXML is never refused, only queries Dataverse rejects are blocked, and warnings are reserved for what can't be determined.

### ✨ Added

- **A complete, metadata-driven properties panel (#90b).** Every node type has an editor covering the FetchXML reference, with info popovers linking to the docs. Tree badges show link types, aggregates and other options at a glance.
- **Lossless Parse to Tree (#90a).** Every documented FetchXML attribute round-trips, as do unknown, namespaced and legacy attributes. All 60 Microsoft Learn FetchXML examples round-trip without loss. When Parse to Tree would remove or correct something, it asks first and lists the changes (#89).
- **Validation that matches Dataverse (#48, #89, #90b, #94, #87).**
  - Errors are only for queries Dataverse rejects; notes never block.
  - Every rule was verified live.
  - Valid but unusual FetchXML (e.g. SQL 4 CDS's `xmlns:generator`, or `link-type="natural"`) is accepted.
- **Condition operators per attribute (#87).**
  - The operators offered follow a live operator × attribute-type matrix.
  - Value inputs match each operator: whole numbers for "x" operators, a fiscal period plus year, `between`, lists.
  - The four documented fiscal "x" operators were added.
- **Many-to-many joins (#85).** An N:N relationship joins through its intersect entity, as documented. Intersect entities (`systemuserroles`, `teammembership`, …) can be joined to each side of their N:N relationship (#57).
- **Paging and Retrieve All (#58, #60).**
  - The documented paging cookie is used, with page-number paging when Dataverse returns none.
  - Views follow `@odata.nextLink` past 5,000 rows.
  - **Stop** keeps the rows already loaded, and a failed page shows a message with **Retry**.
- **Advanced Find Only applies everywhere (#57).** It applies to entities, attributes and relationships in every entity scope. With it off, FetchXML Studio builds queries Advanced Find can't: system and customization entities, intersect entities, and hidden attributes and relationships. Unreadable attributes are never offered.
- **Select Attributes (#52):** a wider dialog with resizable columns, one-line cells, a type filter, "Selected only" and a selection count.

### 🔧 Fixed

- **Root entity and record IDs (#54, #59).** The root entity is detected correctly, and FetchXML is never modified: no primary key is injected.
- **FetchXML runs against the entity's real Web API entity set name (#61).**
- **Editor mode (#102).** Execute, Save and Send to Tool run exactly what's in the editor.
- **Escaping and aliases (#63, #94).**
  - Every generated value is escaped, and apostrophes stay literal for Power Automate.
  - Corrected aliases update their references.
  - The Intellisense, Power Automate and Web API snippets follow the docs.
- **Results grid columns (#114).**
  - Edit / Add columns and header sorting write valid FetchXML for aliased and link-entity columns and aggregate queries.
  - Columns of link-entities without an alias show in the grid.
- **Parse to Tree, files and views keep their query (#57).** Changing the publisher or solution selection never clears the query.
- **Selecting a new link-entity no longer crashes the app.**
- **Dark theme (#67)** covers the whole tool and follows PPTB's theme live, including System.
- **Link-entities (#86, #88):** switching to not any / not all keeps filters, and "None" is never written as a value.
- **The "+ New" button** is no longer covered by the entity box in a narrow pane (#53).
- **The tool icon** is bundled under `dist/` and fills PPTB's icon slot (#62, #121). It shows in PPTB once PowerPlatformToolBox/desktop-app#738 ships.

### 🏗️ Technical

- **Dependencies:** Fluent UI React v9 upgraded to 9.74.9 (#104), and `package-lock.json` is committed.
- **Tests and lint:** Vitest with 507 tests (#64). Lint passes with zero warnings and runs as `eslint . --max-warnings=0` (#65).

## [1.4.1-beta.20] - 2026-10-17

### 🔧 Fixed

- **The tool icon looked small in PPTB (#121).** PPTB draws tool icons in a 20×20 box, but the hexagon only filled 56×48 of the icons' 64×64 canvas, so it showed at about 17×15 px.
  - Both icons (`icon.svg` and the Insider `icon-insider.svg`) now crop their `viewBox` to the hexagon: `4 8 56 48`. They draw at about 20×17 px.
  - The artwork is unchanged; no shape or stroke falls outside the new box.
  - The build still emits the right variant to `dist/icons/icon.svg`, and `pptb-validate` passes.

## [1.4.1-beta.19] - 2026-10-17

### 🔧 Fixed

Lint is clean: `npm run lint` now runs `eslint . --max-warnings=0` and passes with zero warnings (#65). The hook-dependency warnings were fixed rather than suppressed, and two of them were real bugs:
- **The Code tab's Web API snippet could show a stale entity set name.** The snippet didn't update when the entity set name finished loading.
- **Loading FetchXML from a file could report through a stale editor callback.** That handler didn't list the callback as a dependency.

### 🏗️ Technical

- `builderStore.tsx`: the builder's actions are created once (`dispatch` never changes), and the context value is memoized on the state. Consumers can list builder actions as hook dependencies, and they only re-render when the builder's state changes.
- `AppShell.tsx`:
  - the layout sync effect lists `syncLayoutWithFetchXml`, which is now stable;
  - the 1:N / N:N check uses `LinkEntityNode[]` instead of `typeof builder…`;
  - record deletes re-run the query through a ref to the latest `handleExecute`, called after the `await`, instead of a version captured when the callback was created.
- `CodePanel.tsx` / `FetchXmlEditor.tsx`: the missing dependencies above.
- `TreeView.tsx`: the one `eslint-disable` without a reason now explains why the effect runs only on selection changes.
- `eslint.config.js`: ignores `.claude/`, which holds local, git-ignored prototypes.
## [1.4.1-beta.18] - 2026-10-17

### 🔧 Fixed

- **Select Attributes: long names wrapped onto several lines and the columns couldn't be resized (#52).** Design option A, reviewed in a prototype:
  - The dialog is wider: up to 1000px, never wider than the window.
  - The Display name, Logical name and Type columns are resizable and sortable. Each cell stays on one line, long text ends with an ellipsis, and the full text shows in a tooltip.
  - Search matches display or logical name. A **type filter**, a **Selected only** switch and an "N of M selected" count help with long attribute lists.
  - Types use `AttributeTypeName`, so images, files and multi-select choices no longer all show as "Virtual".
- **"+ New" was overlapped by the entity box in a narrow left pane (#53).** Fluent's Combobox has a built-in minimum width (about 250px).
  - The entity, publisher, solution and view boxes now shrink with the pane.
  - Below 300px "New" keeps only its icon, with a "New query" tooltip and label.

### 🏗️ Technical

- `SelectAttributesDialog.tsx`: `resizableColumns` + `columnSizingOptions`, truncated cells with tooltips, type / selected-only filters, count in a fluid `DialogActions`.
- `EntitySelector.tsx` / `LoadViewPicker.tsx`: comboboxes and their inputs use `minWidth: 0`; the "New" button switches to icon-only with a container query (the toolbar is already an inline-size container).

## [1.4.1-beta.17] - 2026-10-16

### 🔧 Fixed

- **Changing the publisher or solution selection cleared the query** when the new selection didn't include its entity (#57 follow-up). Selecting the Default Solution, for example, wiped a `systemuserroles` query with no warning. Deselecting asked first, but still reset the tree.
  - The scope only filters the entity picker, so the query is now never cleared.
  - An entity outside the selection shows by its logical name, and Load View still works for it (from the full entity list).
  - The "Confirm Publisher / Solution Change" dialog, which warned that the tree would be reset, is removed.

### 🏗️ Technical

- `EntitySelector.tsx`: no auto-clear effect and no confirm dialog; `selectedEntityMetadata` falls back to the full entity list.

### 🧪 Test (PPTB, sandbox)

- Parse to Tree on `<fetch><entity name="systemuserroles"><attribute name="roleid" /><link-entity name="role" from="roleid" to="roleid" alias="r"><attribute name="name" /></link-entity></entity></fetch>`. Then select the Default Publisher and Default Solution, and deselect both: the query stays every time, with no dialog; the entity box shows `systemuserroles`; Load View is available.

## [1.4.1-beta.16] - 2026-10-16

### 🔧 Fixed

**Advanced Find Only** now does what it says, everywhere (#57). [`IsValidForAdvancedFind`](https://learn.microsoft.com/dotnet/api/microsoft.xrm.sdk.metadata.entitymetadata.isvalidforadvancedfind) means "shown in Advanced Find"; it isn't a query restriction. Turning it off is how FetchXML Studio builds queries Advanced Find can't. Checked against the sandbox:
- **Hidden entities were missing in the Publisher → Solution and Solution scopes** (`userquery`, the reported case). Those scopes always loaded Advanced Find entities only. The setting now applies in every scope.
  - Verified: 347 of 421 hidden entities query fine with FetchXML, including `userquery`, `savedquery`, `systemform`, `sdkmessageprocessingstep` and the intersect entities.
  - The rest return a Dataverse error, shown on Execute as usual.
- **Hidden attributes and relationships were never offered**, even with the setting off. With it off, every picker now offers them: the tree, Select attributes, conditions, orders, link-entities, Add columns and relationships.
  - Verified: 1,028 of 1,053 hidden attributes and all 1,593 hidden relationships work.
- **Attributes that can't be read are never offered.** `IsValidForRead = false` attributes fail in every query (`0x80041a08`, 25 of 25).
- **Operators for attribute types that only appear with the setting off** (verified on each):
  - label attributes such as `statuscodename` (Virtual): `null`, `not-null`, `eq`, `like`; they can be sorted, but not grouped or aggregated (`0x8004e01d` / `0x80044150`);
  - images, files, managed properties and entity names: `null` and `not-null` only (`eq` / `like` fail: `0x80040203`);
  - PartyList (`to`, `cc`, `from`): no conditions; every condition fails (`0x80040203`).
- **Attributes already in a query keep their metadata** (operators, value inputs, grid formatting) whatever the setting. A pasted query that uses a hidden attribute works either way.
- The setting's description and tooltip say what it covers.
- **Parse to Tree emptied the tree and the editor** in the Publisher → Solution and Solution scopes when the selected solutions didn't include the query's entity (reported with `systemuserroles`). The same happened to opened files, loaded views and queries from other tools, and toggling Advanced Find Only would have done it too.
  - The scope only filters the entity picker. The query is now cleared only after *you* change the publisher or solution selection, once the new entity list has loaded without the query's entity.
  - Otherwise the entity box shows the query's logical name.
- **Intersect entities can be joined from the relationship picker.** `systemuserroles`, `teammembership`, `listmember` and the rest have no many-to-one relationships for their attributes (plain Uniqueidentifiers), so nothing was offered. They now offer a join to each side of the many-to-many relationship they implement, e.g. `systemuserroles → role via roleid (systemuserroles_association)`, in the relationship picker and in Add columns → Related.
  - Verified on every intersect entity in the sandbox: 141 entities, 302 joins. Every join runs except on the few intersect entities that can't be queried at all.
  - Picking one adds `<link-entity name="role" from="roleid" to="roleid">`.
- **Selecting a new link-entity crashed the app** (blank screen). Its placeholder name has no attributes, and the attribute picker's fallback dropped a prop Fluent needs (`selectedOptions`).
- **A new link-entity kept its placeholder alias** (`new_entity`) after a 1:N or N:1 relationship was picked. It now gets an alias for the related entity (e.g. `systemuser`).

### 🏗️ Technical

- New `state/metadataScope.ts`: scope rules (`isEntityInScope`, `isAttributeInScope`, `isRelationshipInScope`) and the current setting (`metadataScope`, `useAdvancedFindOnly`).
- `dataverseMetadata.ts`:
  - attributes and relationships are fetched once, unfiltered, and filtered locally;
  - `loadAllEntityAttributes` for lookups, `loadEntityAttributes` for pickers;
  - `loadAllEntities` shares the Publisher / Solution entity cache.
- `pptbClient.ts`: `getAllEntitiesForScopes` (replaces the Advanced Find–only `getAllAdvancedFindEntities`); the unused `getAdvancedFindEntitiesByNames` is removed; attributes select `IsValidForRead`.
- `operators.ts`: Virtual attributes are typed by `AttributeTypeName` (multi-select, image, file, label).
- `attributeCapabilities.ts`: VirtualType can be sorted.
- `manyToMany.ts`: `planIntersectEntityJoins` (an intersect entity's joins to each side) and `relationshipKey` (the two joins share a schema name); `loadEntityRelationships` adds them for intersect entities.
- `EntitySelector.tsx`: the query is cleared only after a publisher / solution change made by the user.
- `aliasUtils.ts`: `collectLinkEntityAliases` and `generateUniqueAlias` moved here from the builder store.

### 🧪 Test (sandbox)

- With Advanced Find Only off, every scope: `<fetch top="5"><entity name="userquery"><attribute name="name" /><attribute name="returnedtypecode" /></entity></fetch>`.
- Intersect entity: `<fetch><entity name="systemuserroles"><attribute name="roleid" /><link-entity name="role" from="roleid" to="roleid" alias="r"><attribute name="name" /></link-entity></entity></fetch>`.
- Hidden attribute: `<fetch top="5"><entity name="account"><attribute name="name" /><attribute name="versionnumber" /><order attribute="versionnumber" descending="true" /></entity></fetch>`.
- Label attribute: `<fetch top="5"><entity name="contact"><attribute name="fullname" /><filter><condition attribute="statuscodename" operator="like" value="%Active%" /></filter><order attribute="statuscodename" /></entity></fetch>`.
- In PPTB, Publisher → Solution scope with a solution that doesn't include `systemuserroles`: Parse to Tree on the intersect query keeps the tree; Execute returns 407 rows. Added the `systemuser` join from the relationship picker and Full Name from Add columns → Related: `<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="systemuser"><attribute name="fullname" /></link-entity>`, 407 rows.

## [1.4.1-beta.15] - 2026-10-16

### 🔧 Fixed

Paging, **Retrieve All** and infinite scroll (#58, #60). Every case was run against the sandbox through the production code, following [Page results using FetchXml](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results).
- **The paging cookie was never used (#58).** FetchXML Studio read `@Microsoft.Dynamics.CRM.pagingcookie`, which the Web API doesn't send.
  - The cookie is now read from `@Microsoft.Dynamics.CRM.fetchxmlpagingcookie`: the `pagingcookie` attribute of that element, URL-decoded twice, as documented.
  - Every page after the first was fetched by page number only, which Microsoft doesn't recommend. Verified: pages fetched with the cookie match page-number paging, with no duplicates.
- **Queries Dataverse returns no cookie for** (sorted by a link-entity attribute, or `distinct`) page by page number. That's the documented fallback, and a note above the grid says so.
  - Dataverse reports page 2 as the next page on *every* page of these queries (verified). The next page is now always the page just fetched + 1, and an empty page ends paging, so a loop can't repeat a page.
- **Retrieve All could mix results (#60).**
  - Changing the entity, loading a view or starting a new query while pages were loading brought the old rows back. Pages from an older run are now dropped.
  - A record delete re-runs the query directly. That could start a second loop alongside the first; the first is now dropped.
  - A failed page was only logged to the console. Now a message above the grid says how many rows loaded and which page failed, with **Retry**.
- **Stop button (#60).** While Retrieve All runs, **Stop** appears at the start of the results command bar. The loaded rows stay; scrolling to the end loads more one page at a time.
- **System and personal views stopped at 5,000 rows.** Views return `@odata.nextLink` instead of a cookie (verified); FetchXML Studio now follows it, for Retrieve All and scrolling alike.
- **The query's own paging was ignored:**
  - `count` was read from the tree instead of the FetchXML that ran (editor mode). Each page now keeps the executed query's `count`.
  - `page="3"` in the query continued at page 2. It now continues at page 4.
  - `top` is read from the executed FetchXML; queries with `top` aren't paged, as documented.
  - `page`, `paging-cookie` and `count` in single quotes were added again instead of replaced, which made the XML invalid.
  - Cookie values containing `$&`, `$1` and similar were mangled when written into the query.
- **Each page's full response was logged to the console**, which kept every row of a large Retrieve All in memory. It's now debug-only.

### 🏗️ Technical

- New `api/paging.ts`: `parseFetchXmlPagingCookie`, `readFetchPagingOptions`, `nextLinkToQueryPath`, `advancePaging`, `hasNextPage`, and the `PagingState` / `PagingNotice` types.
- `pptbClient.ts`: one `pagingFromResponse` for FetchXML and view responses; new `executeNextLink`.
- `fetchxml.ts`: `addPagingToFetchXml` replaces either quote style and keeps the query's `count` unless one is given.
- `AppShell.tsx`: one paging state (in a ref, so the page loops don't read stale state), plus result and page run tokens.

### 🧪 Test FetchXML (sandbox)

- Cookie paging: `<fetch count="50"><entity name="systemuser"><attribute name="systemuserid" /><order attribute="fullname" /><order attribute="systemuserid" /></entity></fetch>`: 234 of 234 users, no duplicates.
- No cookie (distinct): `<fetch count="50" distinct="true"><entity name="systemuser"><attribute name="fullname" /><order attribute="fullname" /></entity></fetch>`: every distinct name once.
- No cookie (link-entity order): `<fetch count="40"><entity name="systemuser"><attribute name="systemuserid" /><order entityname="bu" attribute="name" /><order attribute="systemuserid" /><link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="bu" link-type="inner"><attribute name="name" /></link-entity></entity></fetch>`: 234 of 234.
- Starting page: `<fetch count="50" page="3"><entity name="systemuser"><attribute name="systemuserid" /><order attribute="systemuserid" /></entity></fetch>`: the users after the first 100.
- Default page size: `<fetch><entity name="account"><attribute name="accountid" /></entity></fetch>`: 3 pages, 15,000 unique accounts.
- View: **Active Accounts** (system view): 3 pages through `@odata.nextLink`, 15,000 unique.
- `<fetch top="10"><entity name="account"><attribute name="accountid" /></entity></fetch>`: 10 rows, no paging.

## [1.4.1-beta.14] - 2026-10-15

### 🔧 Fixed

Condition operators and their values (#87). Every operator in the [operators reference](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators) was run against every Advanced Find attribute type on the sandbox (1,640 queries), plus the value shapes each operator takes.
- **"x" and fiscal operators showed a date picker.** `last-x-days`, `olderthan-x-months`, `in-fiscal-period` and the rest wrote an ISO date, which Dataverse rejects (`0x80041139`).
  - They now take a whole number above 0. Verified: 0, negatives, decimals and dates fail.
  - `in-fiscal-year` takes a four-digit year (verified: `23` fails).
  - The three period-and-year operators take a **Fiscal period** and a **Fiscal year**, written as two `<value>` elements, as documented.
- **Four documented operators were missing:** `last-x-fiscal-years`, `last-x-fiscal-periods`, `next-x-fiscal-years` and `next-x-fiscal-periods` are now selectable.
- **Changing the operator kept a value of the wrong shape** (a date for a number, a list for a single value). The value is now kept only when the new operator takes the same shape.
- **Operators offered per attribute follow what Dataverse accepts:**
  - `contain-values` / `not-contain-values` only for multi-select choices (single choices fail: `0x80048415`); multi-select choices get their own list instead of every operator, and their option values are offered.
  - The user / team operators (`eq-userid`, `eq-userteams`, …) aren't offered for Customer lookups (`0x8004034b`).
  - The hierarchy operators (`under`, `above`, …) only for the primary key of a hierarchical entity (lookups fail: `0x80040203`).
  - Hour and minute operators are left out for date-only *behaviour* (`0x80048415`), but no longer for a date-only *format* with user-local behaviour, which accepts them (verified).
  - Strings get `gt` / `ge` / `lt` / `le`, as documented.
- **New checks (Validate / Execute), verified live:**
  - an "x" operator without a whole number above 0;
  - a fiscal year that isn't four digits;
  - `between` or a period-and-year operator without exactly two values;
  - `in` / `not-in` / `contain-values` without `<value>` elements;
  - a single-value operator with `<value>` elements or no value.
  - A value on an operator that takes none is a note: Dataverse ignores it.
  - None of the 60 Microsoft Learn examples trigger them.

### 🏗️ Technical

- `operators.ts`: `getOperatorsForAttribute` (type, multi-select, primary key, date-only behaviour, hierarchical entity), `operatorValueShape` and `adaptValueForOperator`; new value types `integer`, `year`, `fiscalPeriodYear`.
- `fetchxmlRules.ts`: `conditionValueIssue`.
- The option set request also casts to `MultiSelectPicklistAttributeMetadata`.

## [1.4.1-beta.13] - 2026-10-15

### 🔧 Fixed

Results grid **Edit columns / Add columns** and **header sorting** (#114). Each change was checked against the sandbox.
- **Removing a column only worked for un-aliased root attributes.** Aliased attributes and link-entity columns were silently kept.
  - Removing now finds the attribute behind any column.
  - It also removes sorts that use the attribute's alias (Dataverse rejects an order alias that no longer exists).
  - An **outer** link-entity left with nothing in it is removed too. A 1:N join would otherwise keep multiplying the rows without showing anything. Inner joins are kept, because they still filter.
- **All columns could be removed.** A query with no attributes returns every attribute, so Apply is now disabled until at least one column remains.
- **Add columns in aggregate queries** added plain attributes, which Dataverse rejects (every attribute needs Group by or an aggregate, and an alias). It's now unavailable there, with the reason.
- **The Related tab offered attributes the link-entity already returns.** It now leaves them out.
- **Columns of link-entities without an alias never showed in the grid.**
  - Dataverse names them `{name}{N}.attribute`, where N is the position among the joined link-entities (verified, and documented on the link-entity reference), not `name.attribute`.
  - The layout and the grid now use the same names. This affects queries loaded from views or pasted FetchXML.
- **Header sorting wrote orders Dataverse rejects:**
  - aggregate queries got `<order attribute>` instead of `<order alias>`;
  - an aliased link-entity attribute got `<order attribute="alias">` on the root (`0x80041103`); it's now `entityname` + attribute name;
  - multi-select option sets could be sorted (`0x80050224`); unsortable types now show a message instead.
  - **Single-column sort:** it replaces every order in the query, including orders inside link-entities, so the results are sorted by that column only.

### 🏗️ Technical

- `model/queryColumns.ts`: one model of the result columns (keys exactly as Dataverse returns them) shared by the layout, the grid, column removal and sorting. Tests in `queryColumns.test.ts`.
- `setSort` takes a planned order (root, or inside a link-entity when the root can't target it).

## [1.4.1-beta.12] - 2026-10-15

### 🔧 Fixed

- **Apostrophes stay literal in generated FetchXML (#63).** Values were written with `&apos;`. That's valid XML, but it broke Power Automate expressions pasted into the Fetch Xml Query field (e.g. `@{body('Parse_payload')?['Student']}`).
  - Attribute values are always in double quotes, so only `&`, `<`, `>` and `"` are escaped; `<value>` text escapes `&`, `<` and `>`.
  - Verified live: raw apostrophes are accepted in both forms.
- **Every value in generated FetchXML is escaped (#63, #94).** Entity, attribute and link-entity names, aliases, `from` / `to`, order references, `hint` and `operator` were written unescaped, so pasted or parsed values with `"`, `&` or `<` produced broken XML.
- **Corrected aliases no longer break their references (#94).** When Parse to Tree corrects an invalid alias, it now updates every reference to it as well: condition and order `entityname`, `valueof` (including aliases that contained dots), and aggregate order aliases. The fix is listed with the corrections.
  - Verified live: a link-entity alias with dots fails as soon as `valueof` refers to it (`0x80040216`); the corrected query runs.
- **New link-entity aliases can't collide with link-entities inside filters (#94).** Generated aliases now take any / not any / all / not all link-entities into account.
- **Intellisense follows the FetchXML reference (#94).**
  - Value suggestions now work for hyphenated attributes (`link-type`, `paging-cookie`, `no-lock`, `utc-offset`).
  - Added the documented attributes that were missing: `aggregatelimit`, `datasource`, `latematerialize`, `options` (with the documented hints) and `useraworderby` on fetch; `entityname` on order; `rowaggregate` on attribute; `hint` and both quick find overrides on filter.
  - Removed `order` as a child of `<fetch>` and the condition `aggregate`, which Dataverse ignores.
  - Undocumented attributes are suggested only where verified and say so: `utc-offset`, `version`, `mapping`, `output-format`, `min-active-row-version`, link-entity prefiltering, `natural`, and the saved-view `uiname` / `uitype` / `uihidden`.
  - The operator list is the parser's own list, so the two can't drift apart.
  - Condition `attribute` stays required: the reference lists it as optional, but Dataverse rejects a condition without it (`0x8004111e`).
- **Power Automate Top Count comes from `top` only (#94).** It used to take `count`, which is a page size, not a limit ([page-results](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results)).
- **Web API snippet (#94).** It adds the documented `Prefer: odata.include-annotations="*"` header (formatted values, paging cookie and more-records annotations) and the paging-cookie steps.
- **More than one filter in an entity is noted (#94).** The reference allows one. Dataverse accepts several and combines them with *and* (verified, same rows as one nested filter), so it's a note, never an error.

### 🏗️ Technical

- `KNOWN_OPERATORS` moved to `operators.ts`, shared by the parser and the intellisense.
- `collectLinkEntityAliases` uses `collectAllLinkEntities`.
- `correctness94.test.ts`: escaping round trips, alias references, the intellisense catalogue against the reference, and the code generators.

## [1.4.1-beta.11] - 2026-10-15

### 🔧 Fixed

- **Many-to-many relationships from the relationship picker return the right rows (#85).** Choosing an N:N relationship built a single link-entity straight to the related entity, using the intersect entity's attribute names. That silently returned 0 rows (or failed, depending on the side).
  - It now builds the [documented pattern](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables#many-to-many-relationships): the link-entity becomes the intersect entity (`intersect="true"`, joined on the parent's primary key), and the related entity is nested inside it, with the alias, so attributes are added there.
  - The attributes come from the relationship metadata (`IntersectEntityName`, `Entity1/2IntersectAttribute`) and both entities' primary keys. It works from either side, and for self-referential relationships.
  - **Verified live:** for every many-to-many relationship with data in the sandbox (47 checks from both sides, system, custom and self-referential), the join returned exactly the intersect entity's pairs.

- **Attribute pickers show their value even when it isn't in the attribute list.** For example, intersect entities usually have no Advanced Find attributes, so the From / To pickers on an N:N join looked empty.

### 🏗️ Technical

- `model/manyToMany.ts`: `planManyToManyJoin`, with tests.

## [1.4.1-beta.10] - 2026-10-15

### 🔧 Fixed

- **Dark theme covers the whole tool (#67).** The left pane (tree and properties) and any surface without its own background showed the page's white canvas behind dark-theme text, which made them unreadable in dark mode. The theme background now covers the whole tool. Native controls (scrollbars, date inputs) follow the theme too (`color-scheme`).
- **The tool follows the PPTB theme live, including System (#67).**
  - On `settings:updated`, the tool now always asks the host for the resolved theme (`getCurrentTheme()`). It used to read the raw setting, and "system" was treated as light.
  - The host sends no event when the OS theme changes while ToolBox is set to System. The tool now also listens for the OS colour scheme and asks the host again, so it switches without a reload. An explicit Light or Dark setting is kept.
  - Only the latest theme request applies.
  - Verified against the PPTB 1.2.7 host source: `getCurrentTheme()` resolves System from `nativeTheme`, and `settings:updated` carries only the changed settings.

### 🏗️ Technical

- `usePptbContext.test.ts`: the theme contract (settings:updated with System, OS changes, explicit settings, cleanup, standalone).
- Removed the theme debug logging from `AppShell`.

## [1.4.1-beta.9] - 2026-10-14

### ✨ Added

- **A complete, metadata-driven properties panel (#90b).** Every node type now has an editor that covers the FetchXML reference. The "coming in a later phase" placeholders are gone. Fields use Fluent UI v9 info popovers with a **Learn more** link to the matching Microsoft Learn article, sentence-case labels and Entity / Attribute / Logical Name terms.
  - **Fetch Query:**
    - **Query:** aggregate, distinct, total record count, sort option sets by value.
    - **Paging and limits:** top, count, page and paging cookie. `top` can't be combined with paging, so the conflicting fields explain why.
    - **Aggregation:** `aggregatelimit` (1–50,000).
    - **Data source:** Active / Retained. Retained is disabled when the entity's long-term retention is off.
    - **Performance:** late materialize, the documented query hints as checkboxes, and any other hints kept as tags.
    - **Advanced:** UTC offset.
    - **Turning off aggregate** lists what it would remove or change (Group by, aggregates, alias sorts) and asks first.
  - **Entity:** the entity's logical and display name, plus report prefiltering (`enableprefiltering`, `prefilterparametername`).
  - **All attributes:** what it does, Microsoft's guidance against it, and a Remove button.
  - **Attribute:**
    - **Role:** Group by or Aggregate.
    - **Function:** only the functions the attribute's type supports.
    - **Group by options:** date grouping (required for date attributes), time zone, and unique values for count of values.
    - **Count child records:** offered on the primary key of hierarchical entities.
    - **Renaming an alias** updates the sorts that use it.
  - **Sort:**
    - **Regular queries:** an attribute picker limited to sortable types.
    - **Aggregate queries:** a list of the aliases in scope.
    - **Entity:** sort a link-entity first.
  - **Filter:** all (and) / any (or), the union hint (only for `or`, one per query), quick find, and the quick find record limit overrides.
  - **Condition:** Entity (`entityname`, filter after an outer join), attribute, operator and value. **Compare to** chooses between a value and another attribute.
  - **Link-entity:**
    - entity, From / To attribute pickers, link type, alias and intersect;
    - link types are offered by placement: inside a filter only any / not any / all / not all;
    - descriptions follow the docs, including that `all` / `not all` mean the opposite of what the names suggest.
- **Tree badges:**
  - **Fetch Query:** Aggregate, Distinct, Retained, Hints.
  - **Attributes:** Group by (with date grouping), the aggregate function, Unique, Child count.
  - **Filters:** Union, Quick find.
  - **Any node with a rule error:** a red **Error** badge.
  - **Conditions that compare attributes** read `attribute eq → other`.
- **Editor mode follows the properties panel.** While editor mode is on, panel edits update the editor when the editor hasn't changed the query. If it has, the editor keeps your text and says so.

### 🔧 Fixed

- **Comparing to another attribute (`valueof`) only offers what Dataverse accepts (#94).**
  - The choice appears only for eq, ne, gt, ge, lt and le (others fail with `0x80048415`). Changing to another operator clears it.
  - **Which entities are offered.** The docs show the documented same-row and link-entity-alias cases; the picker offers only those, plus what was verified to work from where the condition is:
    - the root entity by its logical name;
    - from the root entity's filters: any joined link-entity alias;
    - from a link-entity's filter: only link-entities joined before it (its own children and later link-entities fail with `0x80040216`).
    - A link-entity name without an alias, or a link-entity inside a filter, is never offered.
  - **Which attributes are offered.** The docs say "the same type". FetchXML Studio offers the same type plus the pairs verified to work, with all six operators: Integer / option set / status / state / entity name; Decimal / Money; Lookup / Customer / Owner / unique identifier; String / Memo; Boolean / managed property; DateTime (any behaviour).
    - Verified with a live matrix over every Advanced Find attribute type: about 2,000 queries, including calculated, rollup, composite and date-only attributes.
    - Pairs across these groups fail (`0x80040203`), so they aren't offered.
    - Multi-select option sets, and any type that wasn't verified, aren't offered at all.
- **Conditions on a link-entity attribute (`entityname`) load that entity's metadata (#94).** Numeric limits and date-only rules previously came from the wrong entity.
- **Switching a link-entity to not any / not all no longer deletes its filters and nested link-entities (#86).**
  - Those filters are what these types evaluate.
  - Switching to any type that returns no attributes (exists, in, any, not any, all, not all) now lists the attributes, all-attributes and sorts that Dataverse would reject, and asks before removing them.
  - The tree disables adding attributes and sorts to those link-entities, and **Add Columns** creates a new link-entity instead of adding to one of them.
- **From / To on a link-entity are no longer marked required.** The docs say they're optional; the panel now explains what Dataverse does when one or both are missing.
- **Intersect** is described correctly: it only signals intent and doesn't change how the query runs.
- **"Visible in Results" couldn't be turned off.** `visible` isn't in the FetchXML reference, and Dataverse ignores it (verified). It's shown only when present, with a Remove button.
- **The condition "Aggregate function" option is gone.** Dataverse ignores `aggregate` on a condition and runs it as a plain row condition (verified): FetchXML can't filter on aggregate values. Parsed queries keep it, with a note and a Remove button.
- The serializer now escapes condition `attribute`, `entityname` and `valueof` values.
- Removed a leftover "Export not yet implemented" console fallback.

### 🔍 New checks

Each check was verified against a live Dataverse org.
- **Errors:**
  - `top` with `count`, `page` or `returntotalrecordcount` (moved from the parser to the query rules, so the panel shows it on the field);
  - `page` below 1;
  - `valueof` with an unsupported operator (`0x80048415`), or with a qualifier the condition can't reach: not a joined link-entity alias or the root entity, or a link-entity joined after the one whose filter holds the condition (`0x80040216`);
  - a link-entity inside a filter that isn't any / not any / all / not all (`0x8004111c`);
  - attributes, all-attributes or orders inside an exists, in, any, not any, all or not all link-entity (`0x80041121`).
- **Notes (never block):** condition `aggregate`, link-entity `visible`, and `valueof` on operators that take no value.

### 🏗️ Technical

- Shared panel building blocks: `fields.tsx` (`InfoField`, `InfoCheckbox`, `NumberField`, `LeftoverSetting`), `docLinks.ts`, `editorStyles.ts`, `ruleFindings.ts` and `useEditorMetadata.ts`.
- `aggregateToggle.ts` plans what turning off aggregate changes; `treeUtils.ts` adds `findNodeContainer`, `collectAllLinkEntities`, `collectAllFilters` and `findOrdersUsingAlias`.
- Metadata: `IsRetentionEnabled`, `IsLogical` and `IsPrimaryId` are read, and hierarchical entities are detected from their relationships (cached).
- `attributeCapabilities.ts` adds `getValueOfCompatibleTypes`; `fetchxmlRules.ts` exports `VALUEOF_OPERATORS`, `FILTER_LINK_TYPES`, `NON_RETURNING_LINK_TYPES` and `joinedLinkAliases`.

## [1.4.1-beta.8] - 2026-10-14

### 🔧 Fixed

- **The tool icon now shows in Power Platform ToolBox (#62).** PPTB resolves the top-level `icon` relative to `dist/`, but the SVGs were never copied there, so PPTB showed its default icon. A valid SVG bundled under `dist` is also a **required** item for Verified status ([publishing](https://docs.powerplatformtoolbox.com/tool-development/publishing), [maturity model](https://docs.powerplatformtoolbox.com/tool-development/maturity-model)).
  - The build now bundles the icon at `dist/icons/icon.svg`, and `package.json` points `icon` there.
  - Prerelease versions (e.g. `1.4.1-beta.8`) get the gold-stamped Insider icon; releases get the regular blue icon. There's no manual switch per release.

### 🏗️ Technical

- **`package-lock.json` is committed**, pinning every dependency with an integrity hash. npm's guidance: it's "intended to be committed into source repositories" and isn't published. PPTB users install the pre-built package, so this only affects development and CI.
  - It enables `npm ci` and `npm audit` (a Verified requirement) and CI (#72).
  - It pins exactly the versions tested so far (restored from the last `npm-shrinkwrap.json`), not a fresh resolution.
- `files` no longer lists `icons` (the icon ships inside `dist`) or the removed `npm-shrinkwrap.json`.

## [1.4.1-beta.7] - 2026-10-14

### 🔧 Fixed

- **Editor mode: actions now run exactly what's in the editor (#102).** In editor mode, **Execute**, **Save / Save As** and **Send to Tool** first load the editor's FetchXML into the tree, then act on the tree.
  - **What went wrong before:**
    - Execute ran the editor XML, but the results grid took its columns, layout, entity and record actions from the tree. With a different query in the editor, you saw the tree's (often empty) columns.
    - Save As and Send to Tool silently used the tree's XML, not the editor's.
    - Execute, Save As and Send to Tool were disabled while the tree was empty, even with a valid query in the editor.
  - **Now:**
    - Execute is enabled as soon as the editor has content.
    - The tree always matches what ran, was saved or was sent, so the results grid, column picker, sorting and record actions all match too.
    - Editor mode stays on, and your column widths and order are kept. Your text is kept as typed, unless loading changed the query; then the editor shows the cleaned FetchXML that ran, with an "Editor updated" note.
    - When the editor already matches the tree, nothing is reloaded.
  - **No silent changes:** if loading into the tree would change the query (for example, dropping a second `<entity>`), a confirmation lists the changes first ("Load and execute" / "Load and save" / "Load and send"). Invalid FetchXML is reported and nothing runs. Cancel leaves the editor and tree untouched.
  - Loading a different query clears the loaded-view link, as any tree edit does, so Save then saves a new view.
- **Results are no longer cleared by a fast response** that arrives before the tree update is rendered.

### 🏗️ Technical

- `editorSync.ts`: a context that lets toolbar buttons rendered inside the results/editor pane ask it to sync before acting.
- `ParseChangesDialog.tsx`: the Parse to Tree confirmation, shared by Parse to Tree and the editor-mode actions.
- `loadFetchXml(xml, { preserveLayout })` keeps the current column layout and merges new attributes into it.

## [1.4.1-beta.6] - 2026-10-14

### 🔧 Fixed

- **Validate and Execute now catch the query mistakes Dataverse rejects (#90b, part 1).** New query-level rules, each verified against a live Dataverse org. FXS's parser agreed with Dataverse on all 48 rule cases (0 mismatches).
  - **Aggregate queries:**
    - every attribute needs Group by or an aggregate (`0x8004111e`) and an alias (`0x8004110b`);
    - an attribute can't be both Group by and an aggregate (`0x80041137`);
    - Group by and aggregates outside aggregate queries are errors (`0x8004112e`, `0x8004111a`);
    - an order can only use aliases from its own entity or link-entity (`0x8004110a`);
    - `aggregatelimit` together with an order fails with a SQL error.
  - **Orders:** `entityname` must match a link-entity alias, a link-entity name or the root entity (`0x80041108`), and isn't allowed on orders inside a link-entity.
  - **Aliases:** must be unique across the query, case-insensitively, and can't reuse the logical name of another attribute selected without an alias. Link-entity aliases must be unique too (`0x80041130`).
  - **Filters:** the union hint must be exactly `union`, on an `or` filter, once per query (`0x80041122`). Other hint values were silently dropped by Parse to Tree; they're now reported.
  - **Count child records (`rowaggregate`):** needs an alias, works only on the root entity, and doesn't work in aggregate queries.
  - **Ignored, not errors:** settings Dataverse accepts but ignores are reported as information only. Examples: `aggregatelimit` or date grouping outside aggregate queries, and `distinct` on aggregates other than `countcolumn`.
- **Boolean spellings Dataverse accepts are read correctly.**
  - `usertimezone="1"` used to be read as **false** (inverted meaning after Parse to Tree).
  - `groupby="1"`, `descending="1"`, `distinct="1"`, `returntotalrecordcount="1"` and `enableprefiltering="1"` (the documented form) were dropped.
  - `descending="True"`, which Dataverse rejects (`0x80041119`), is now an error instead of silently sorting ascending.
  - `aggregate` only makes an aggregate query when it's exactly `"true"`, as Dataverse does.
- **Report prefiltering is written as `enableprefiltering="1"`**, the documented value ([Use data filtering](https://learn.microsoft.com/en-us/dynamics365/customerengagement/on-premises/analytics/improve-report-performance-by-using-filters?view=op-9-1)).
- **Choosing "None" no longer writes `aggregate="none"` / `dategrouping="none"` (#88).** The Attribute and Condition editors clear the value instead, and the serializer never writes values outside the documented sets. The invalid "Row Aggregate" aggregate option is removed (`0x8004111a`).

### 🏗️ Technical

- `model/fetchxmlRules.ts`: query-level rules shared by the parser and (next) the property editors. Findings carry the `nodeId` they refer to.
- `model/attributeCapabilities.ts`: the metadata-driven capability matrix for the property panel. It records which attribute types support sort, group by and each aggregate function. It came from a 912-query live matrix over every attribute type, identical through `fetchXmlQuery` and `queryData`.
  - Example: Money attributes can't be counted through the Web API.
  - Example: multi-select option sets can't be sorted or grouped.
- `AttributeNode.aggregate` no longer allows `"rowaggregate"`.
- 280 unit tests.
- `npm-shrinkwrap.json` removed, and `finalize-package` no longer runs `npm shrinkwrap` (npm 12 dropped that command). The local `.claude/` workspace folder is git-ignored.

## [1.4.1-beta.5] - 2026-10-13

### 🔧 Fixed

- **FetchXML now runs against the entity's real entity set name (#61).** Queries used to go through the host's `fetchXmlQuery`, which guesses the Web API entity set name with English plural rules. Measured on the sandbox org, the guess is wrong for **67 of 1,612** Advanced Find entities, so every query on them failed. Examples: `feedback` (host requested `feedbacks`) and `msdyn_conversationdata` (actual set `msdyn_conversationsdata`).
  - Execution now uses `dataverseAPI.queryData(`${EntitySetName}?fetchXml=…`)` with the `EntitySetName` from entity metadata (cached). Microsoft requires the entity set name for the entity in the FetchXML's `entity` element ([Retrieve data with FetchXML](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/retrieve-data?tabs=webapi)).
  - The request is otherwise identical: GET with `Prefer: odata.include-annotations="*"`. Results were the same as before across 1,047 audit queries on entities the host names correctly.
  - The entity set comes from the FetchXML being executed, so editor-mode XML and paged requests (Load more, Retrieve all) use the right entity too.
  - Two system entities (`archivecleanupinfo`, `delegatedauthorization`) report an `EntitySetName` the service doesn't route. When Dataverse answers "Resource not found for the segment", the query is retried once with the entity's `LogicalCollectionName`.
  - `fetchXmlQuery` remains only as a fallback when entity metadata can't be loaded.
- Record delete still goes through the host, which only accepts a logical name. The metadata-based fix needs PPTB 1.2.7's `executeBatch` and moves to the v1.5.0 upgrade (#82).

## [1.4.1-beta.4] - 2026-10-13

### 📦 Dependencies

- **Fluent UI React v9 upgraded (#104)** ahead of the property-panel work in #90b:
  - `@fluentui/react-components` 9.72.4 → 9.74.9;
  - `@fluentui/react-icons` 2.0.313 → 2.0.343;
  - `@fluentui/react-calendar-compat` 0.3.16 → 0.4.7 (required by the date picker);
  - `@fluentui/react-datepicker-compat` 0.6.19 → 0.6.39;
  - `@fluentui/react-timepicker-compat` 0.4.24 → 0.4.41;
  - `@fluentui-contrib/react-data-grid-react-window` 1.4.1 → 1.4.2.
- Every new version was published at least two weeks ago. `react-icons` is pinned to 2.0.343 instead of 2.0.344 (published 2026-10-09) for that reason.
- The dependency tree has a single copy of each Fluent package. `npm audit` drops from 29 findings to 26 (three fewer high); the rest are tracked in #77.

### 🏗️ Technical

- `npm-shrinkwrap.json` was regenerated with npm 11.20.0. npm 12 no longer has `npm shrinkwrap`, and it writes a `package-lock.json` (without integrity hashes) instead of updating the shrinkwrap. Use npm 11 for dependency changes until the release tooling in #75 handles this.

## [1.4.1-beta.3] - 2026-10-13

### 🔧 Fixed

- **Parse to Tree is now lossless (#90a).** Previously the tree model couldn't hold several documented FetchXML attributes, so Parse to Tree (and editing in tree mode) silently dropped them or changed the query's meaning. Measured on all 60 FetchXML examples in the Microsoft Learn FetchXML articles: 25 of them changed after Parse to Tree. Now all 60 round-trip without loss.
  - **`<fetch>`:** `aggregatelimit`, `datasource`, `latematerialize` (was parsed but never written), `options` (query hints, kept exactly as written) and `useraworderby`.
  - **`<attribute>`:** `distinct` (for `countcolumn`) and `rowaggregate="CountChildren"`.
  - **`<order>`:** `alias`, which aggregate queries require.
  - **`<filter>`:** `isquickfindfields`, `overridequickfindrecordlimitenabled` and `overridequickfindrecordlimitdisabled`.
  - **`<entity>`:** `enableprefiltering` and `prefilterparametername` were read but never written.
  - **`<link-entity>`:** empty `from` / `to` are now omitted instead of being written as `from="" to=""`, which made Dataverse silently return no rows. The documented "let Dataverse choose" join for system many-to-many relationships now works after Parse to Tree.
  - **Everything else that works is kept verbatim:** unknown attributes, namespaced attributes, namespace declarations (`xmlns:generator`) and legacy attributes (`version`, `output-format`, `mapping`, `uiname`, `uitype`). A default `xmlns`, which Dataverse rejects, is still removed.
- **New validation, verified against a live Dataverse environment:**
  - `aggregatelimit` must be 1–50,000;
  - `datasource` must be `retained` (documented) or `bin` (accepted by Dataverse);
  - `aggregate="rowaggregate"` is an error (use `rowaggregate="CountChildren"`);
  - aggregate queries must order by `alias` (`0x80041131`), and non-aggregate queries by `attribute` (`0x80041132`);
  - query hint names that aren't documented get a note, because Dataverse silently ignores them.

### 🏗️ Technical

- Model: `FetchNode.options` gained `aggregateLimit`, `dataSource`, `useRawOrderBy` and `queryHints` (replacing an unused `hints` object). `AttributeNode` gained `distinct` and `rowAggregate`. `OrderNode.attribute` is now optional, alongside a new `alias`. `FilterNode` gained the quick find flags. Every node gained `extraAttributes` for verbatim pass-through.
- New round-trip suite (`fetchxmlRoundTrip.test.ts`) over all Learn examples, committed as MIT-licensed fixtures with source URLs (`__fixtures__/learnFetchXml.ts`), plus real-world FetchXML (Advanced Find, saved views, SQL 4 CDS, report prefiltering). 194 unit tests.
- Live replay: every Learn example executed as written and as regenerated, through the PPTB `dataverseAPI` request path, with identical results for all 53 full queries.
- Small display fallbacks for alias-only orders: the tree label shows `alias (alias)`, and the grid shows sort arrows on aggregate columns ordered by alias. The full UI controls for the new attributes follow in #90b.

## [1.4.1-beta.2] - 2026-10-13

### 🔧 Fixed

- **Valid FetchXML is no longer blocked in the editor (#89, #48).** Previously *any* parser warning blocked Validate and Execute. That refused FetchXML Dataverse accepts, such as `<fetch xmlns:generator="MarkMpn.SQL4CDS">` from SQL 4 CDS, `useraworderby`, `aggregatelimit` and `<order alias>`.
  - **Principle:** everything the [FetchXML reference](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/) documents is supported. Undocumented behaviour verified to work is supported too. Only what is known not to work is blocked.
  - Every classification was checked against a live Dataverse environment through the PPTB `dataverseAPI` request path: 33 probes, 0 mismatches.
  - **Errors** (Execute is blocked) are things Dataverse is known to reject:
    - invalid attribute alias characters;
    - an unknown child element;
    - non-integer `top` / `count` / `page`;
    - `top` combined with `count`, `page` or `returntotalrecordcount`;
    - an unknown `link-type`;
    - an invalid filter `type`, aggregate or date grouping;
    - a missing `name` / `attribute` / `operator`;
    - an order with neither attribute nor alias;
    - a default `xmlns`.

    Parse to Tree can still build the tree and lists what it corrects or removes. Malformed XML, a missing `<entity>`, and a link-entity with empty `from=""` and `to=""` (Dataverse silently returns no rows) can't be parsed at all.
  - **Warnings** (with **Execute anyway**) are only for what can't be determined in the client: a link-entity without `from` / `to` (valid only for system many-to-many relationships), and an operator that isn't in the documented list.
  - **Notes** never block:
    - namespace declarations, namespaced attributes and unknown attributes (Dataverse ignores them);
    - documented attributes the visual builder doesn't support yet, including `datasource`;
    - `top` / `count` above 5,000 (Dataverse returns at most 5,000 rows);
    - the deprecated `neq` operator (works like `ne`);
    - extra `<entity>` elements (Dataverse uses the first).

    Common legacy attributes from Advanced Find and saved views (`version`, `output-format`, `mapping`, `uiname`, `uitype`) are accepted silently.
- **`link-type="natural"` is supported.** It's undocumented, but Dataverse accepts it; it's listed in Dataverse's own link-type error and was verified against a live org. It's preserved through Parse to Tree.
- **Parse to Tree no longer changes queries silently.** When it would remove or rewrite something, a confirmation dialog lists the changes under **Removed** and **Corrected**. Examples: `datasource="retained"`, which would otherwise read live data; `<order alias>`; an invalid alias.
  - Order elements that the tree can't represent are now dropped, with a note, instead of being rewritten as the invalid `<order attribute="" />`.
  - The misleading "Auto-corrected" label is gone.
- **Validate** shows at most one bar per level (Invalid FetchXML / "Dataverse might reject this query" / Notes) instead of one bar per finding. Lists are bulleted.

### 🏗️ Technical

- `ParseWarning` gained `severity` (`error` | `warning` | `info`) and `kind` (`invalid` | `ignored` | `dropped` | `corrected` | `advice`).
- New helpers: `getExecutionErrors()`, `getBlockingWarnings()` and `getParseToTreeChanges()`.
- Unknown-attribute checks now cover every element, using per-element attribute lists from the reference.
- New `ParseIssueList` component. The Parse to Tree confirmation uses a Fluent UI v9 `Dialog` (`modalType="alert"`, Cancel focused first).
- 44 new unit tests, 109 in total.

## [1.4.1-beta.1] - 2026-10-12

### 🔧 Fixed

- **Wrong attribute injected into the root entity (#54).** The root table was taken from the first entry of an alphabetically sorted list of the query's tables. That list was sorted *in place*, so a link-entity whose name sorts first (e.g. `bookableresource` before `msdyn_timeentry`) was treated as the root. Its primary key (`bookableresourceid`) was injected into the root `<entity>`, and its metadata also drove relationships, the layout and bulk delete.
  - The root now always comes from the query's `<entity>` element.
  - Metadata responses that arrive after the query changed are discarded.
  - Root metadata is only used when its `LogicalName` matches the current root.
- **FetchXML is no longer modified (#59).** The primary key used to be injected into the generated FetchXML. That broke aggregate queries (`0x8004111e: An attribute can not be requested when an aggregate operation has been specified…`), silently defeated `distinct`, and the injected attribute never appeared in the tree.
  - Validated on a live org through the PPTB `dataverseAPI` request path: Dataverse returns the root primary key for every non-aggregate, non-distinct query without it being selected ([Select columns](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/select-columns?tabs=webapi)). Distinct queries return none ([Ordering and paging](https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results?tabs=webapi#ordering-and-paging)).
  - What you see in the FetchXML pane, copy, save as a view, send to another tool or generate code from is now exactly what you built.
- **Record actions use real record IDs only.**
  - Record actions (Open, Copy URL, Delete, Bulk Delete, Run Workflow) are enabled only when the results contain the table's primary key.
  - For aggregate and distinct results they are disabled, and a "Record actions unavailable" indicator explains why. Previously, Open/Copy URL could be clicked for an aggregate row and used a random ID.
  - Grid row IDs are now unique and stable even when a 1:N join repeats a record. Synthetic row IDs are never passed to record actions.

### 🏗️ Technical

- `generateFetchXml(fetchNode)` no longer takes a `primaryIdAttribute` parameter.
- New `getQueryEntities()` (`model/treeUtils.ts`) and `buildRowIdentity()` (`ui/RightPane/rowIdentity.ts`), with unit tests: 65 tests in total.
- **Unit test baseline (#64).** Added Vitest (`npm test`, `npm run test:watch`, `npm run coverage`) with jsdom, plus 52 characterization tests for the FetchXML serializer, parser (including a parse → generate → parse round-trip of the #54 query), paging-attribute rewriting, alias utilities and the operator catalogue.
- **Lint cleanup, part 1 (#65).** ESLint errors reduced to zero with no behaviour change:
  - `useBuilder` / `BuilderContext` moved to `state/builderContext.ts`;
  - `useTheme` / `ThemeContext` moved to `shared/contexts/useTheme.ts`;
  - `getCellRenderer` moved to `RightPane/getCellRenderer.tsx` (React Fast Refresh: component modules export only components);
  - removed two `any` types in `ResultsGrid.tsx`;
  - `_`-prefixed bindings are now allowed as intentionally unused.
  - The remaining `react-hooks/exhaustive-deps` warnings can affect runtime behaviour and are handled separately.

## [1.4.0] - 2026-07-01

### ✨ Added

#### Tool-to-Tool (T2T) invocation

FetchXML Studio now participates in PPTB Tool-to-Tool (T2T) invocation in both directions. This feature was iterated through the `1.2.2-beta.*` and `1.2.3-beta.*` prereleases (detailed below); `1.4.0` is its first stable release. All T2T behavior is capability-gated and feature-detected, so standalone use and hosts without invocation/discovery support are unaffected. `features.minAPI` stays `1.2.2`.

- **Send to Tool (caller)** — Send the current FetchXML (plus root entity) to another PPTB tool. Targets are discovered automatically via the capability registry (`invocation.findToolsByCapability("fetchxml")`); the "Send to Tool" dropdown lists installed tools that declare the `fetchxml` capability, excluding FetchXML Studio itself (matched on the runtime tool id from `getToolContext()`). One-way, fire-and-forget handoff via `launchTool(..., { noReturn: true })`.
- **Callee (inbound prefill + return)** — When launched by another tool, FetchXML Studio consumes the inbound launch context to pre-populate the builder (resolution priority `fetchXml` → `viewRef` → `entityLogicalName`; the root entity is always derived from the FetchXML). A "Return FetchXML" button (shown only when launched as a callee) returns `{ fetchXml }` to the caller via `returnData()`.
- **Contract** — `pptb.config.json` declares `capabilities: ["fetchxml"]`, a `prefill` schema, and a `returnTopic`.

### 🏗️ Technical

- Depends on `@pptb/types` `1.2.3` (stable). New modules: `src/features/fetchxml/api/invocation.ts`, `src/shared/hooks/useLaunchContext.ts`, `src/shared/hooks/useSendToTools.ts`, `src/features/fetchxml/ui/Toolbar/SendToToolButton.tsx`. See the `1.2.x-beta.*` entries below for the full iteration history.

## [1.3.1] - 2026-05-30

### 🔧 Fixed

- **Double sort arrows on results grid column headers** (#22, #35) — The Fluent UI v9 `DataGrid` component renders its own sort indicator when `sortable` is set without a controlled `sortState`. This caused both Fluent's built-in arrow and the custom `renderHeaderCell` multi-sort arrow to render simultaneously. Fixed by passing a controlled `sortState` sentinel value (`sortColumn: "__sort_managed_externally__"`) that never matches any real column ID, preventing Fluent from rendering its internal indicator. The custom FetchXML-driven sort arrows in the column headers are now the sole source of truth.

- **"Parse to Tree" injects incorrect primary key for activity entities** (#37) — Standard entities follow the `${entityname}id` naming convention (e.g. `accountid`, `contactid`), but activity entities (`email`, `task`, `phonecall`, `fax`, `letter`, `appointment`, and custom activities) all inherit from `activitypointer` and use `activityid` as their primary key. The previous code hardcoded the `${entityname}id` pattern everywhere, causing `emailid`, `taskid`, etc. to be injected into the FetchXML query tree — attributes that do not exist in Dataverse. Fixed by threading `EntityMetadata.PrimaryIdAttribute` (the authoritative value from the Dataverse metadata API) through the full call chain. No fallback pattern or guessing — the primary key is only injected when the metadata is available.

### 🏗️ Technical

- `fetchxml.ts` — `generateFetchXml(fetchNode, primaryIdAttribute?)` accepts the PK from the caller; only injects the primary ID attribute when explicitly provided.
- `layoutxml.ts` — `generateLayoutFromFetchXml(fetchQuery, primaryIdAttribute?, attributeTypeMap?)` — PK is now optional; stored in `LayoutXmlConfig.primaryIdAttribute` when provided.
- `builderStore.tsx` — `ViewLoadInfo` now includes `originalFetchXml: string`; `SYNC_LAYOUT_WITH_FETCHXML` action and `syncLayoutWithFetchXml()` both accept `primaryIdAttribute?`; the store no longer derives or guesses the primary key itself.
- `ResultsGrid.tsx` — `primaryIdAttribute` prop replaces the old `${entityName}id` `useMemo`; `sortState` sentinel eliminates the double-arrow rendering.
- `PreviewTabs.tsx` — threads `primaryIdAttribute?` down to `ResultsGrid`.
- `AppShell.tsx` — passes `entityMetadata.PrimaryIdAttribute` to all call sites: `generateFetchXml`, `syncLayoutWithFetchXml`, `setLoadedView` (via `originalFetchXml`), and `<PreviewTabs>`.

---

## [1.3.0] - 2026-05-29

### ✨ Added

#### Code Generation Panel (#38)
- **Code tab** — New tab in the right pane (alongside FetchXML, LayoutXML, Results) with a tabbed panel that generates ready-to-use code from the current query in seven formats. Async formats (C# QueryExpression, SQL) run in parallel and cancel automatically when the query changes.

#### C# QueryExpression (#39, #40)
- Converts FetchXML to a full `QueryExpression` SDK object graph via the Dataverse `FetchXmlToQueryExpression` Web API function.
- Emits `ColumnSet`, `FilterExpression`, `ConditionExpression`, `LinkEntity`, `OrderExpression`, `XrmAttributeExpression`, `PagingInfo`, `TopCount`, `Distinct`, and `NoLock`.
- Full aggregate query support — `XrmAggregateType` (Count, CountColumn, Sum, Avg, Min, Max) and `XrmDateTimeGrouping` (Day, Week, Month, Quarter, Year, FiscalPeriod, FiscalYear).
- Correctly resolves all Dataverse API enum values returned as strings (`ConditionOperator`, `JoinOperator`, `LogicalOperator`, `OrderType`, `XrmAggregateType`, `XrmDateTimeGrouping`) via a shared `resolveEnum()` helper.
- Correctly unwraps .NET-serialized condition values (`{ "Type": "System.Int32", "Value": 0 }`) to C# integer / string / GUID literals via `unwrapODataValue()`.

#### C# FetchExpression (#43)
- Wraps the FetchXML in `new FetchExpression(@"...")` using a C# verbatim string literal for multi-line XML.

#### JavaScript (#43)
- Generates an `Xrm.WebApi.retrieveMultipleRecords` call for model-driven app JS, PCF components, and browser console.

#### pac CLI (#43)
- Generates a `pac org fetch --xml "..."` command for the Power Platform CLI.

#### Power Automate (#43)
- Field-by-field form mirroring the Dataverse "List rows" connector UI — not a plain code dump.
- Displays Table name, Select Columns, Filter Rows, Sort By, Top Count, and Fetch Xml Query fields with individual copy buttons and flash confirmation feedback.
- Table name resolved from `EntityDefinitions` metadata (already cached from entity selection) — accurate OData entity set name, not a heuristic pluralisation.

#### Web API (#43)
- Full OData URL with URL-encoded FetchXML, required request headers, a `curl` snippet, and a PowerShell `Invoke-RestMethod` example.
- Entity set name resolved from the same metadata cache as the Power Automate tab.

#### SQL (#39)
- T-SQL generated via the Dataverse `FetchXMLToSQL` Web API function (undocumented — labelled "Preview" with a tooltip warning).
- Robust response-property fallback parsing handles multiple possible response key names.

### 🏗️ Technical

- Added `src/features/fetchxml/engine/` — new code-generation module:
  - `queryExpressionTypes.ts` — TypeScript interfaces mirroring the Dataverse QE JSON schema; all enum fields typed as `number | string` to match actual API wire responses.
  - `queryExpressionCodegen.ts` — C# SDK code emitter with `resolveEnum()` and `unwrapODataValue()` helpers for robust handling of string-name enums and .NET-serialized value wrappers.
  - `fetchxmlCodeGenerators.ts` — Sync code generators (FetchExpression, JavaScript, pac CLI, Power Automate spec, Web API); pure functions with no API calls required.
- Added `src/features/fetchxml/ui/RightPane/PowerAutomatePane.tsx` — field-by-field form component (3-column grid: label | input | copy button) with "N of N populated" badge and "Copy all" action.
- Added `src/features/fetchxml/ui/RightPane/CodePanel.tsx` — 7-tab code generation panel; async tabs run in parallel with abort-on-change cancellation.
- `PreviewTabs.tsx` — Added Code tab with `Code20Regular` icon.
- `pptbClient.ts` — Added `fetchXmlToQueryExpression` (calls `FetchXmlToQueryExpression` Dataverse function) and `fetchXmlToSQL` (calls `FetchXMLToSQL`) with multi-property fallback response parsing.

## [1.2.3-beta.3] - 2026-07-01

### 🏗️ Technical

- Upgraded `@pptb/types` `1.2.3-beta.0` → `1.2.3` (stable) ahead of merging the T2T branch into `dev`. Types are compile-time only, so there is **no runtime change** — the API surface FetchXML Studio uses (capability discovery, `launchTool` `noReturn`, `getToolContext`) is identical in the stable release. `features.minAPI` stays at `1.2.2`; discovery remains feature-detected and degrades gracefully on older hosts.

## [1.2.3-beta.2] - 2026-07-01

### 🐛 Fixed

- **Send to Tool no longer spins forever.** The launch is now fire-and-forget — previously the button awaited `launchTool`, whose promise doesn't resolve until the callee window closes, so the spinner stayed up for the whole session (reported when sending FXS → Bulk Data Studio). This is purely a caller-side fix; a callee does **not** need to declare a `returnTopic` to avoid it.
- **Always shows a "Send to Tool" dropdown.** Previously a single discovered tool collapsed the toolbar control into a button labelled with that tool's name; it now consistently reads "Send to Tool" and lists the target(s) in the menu.
- **Each target shows its tool id on its own line** beneath the display name in the menu.
- **FetchXML Studio no longer lists itself** as a target. Self-exclusion now matches on the runtime tool id from `getToolContext()` (with the npm package id as a fallback), since discovery reports the runtime id (e.g. `npm-mohsinonxrm-pptb-fetchxml-studio`) rather than the npm package name.

## [1.2.3-beta.1] - 2026-06-17

### ✨ Added

#### Tool-to-Tool (T2T) — capability-based tool discovery

The "Send to Tool" target list is now populated automatically via the PPTB capability registry instead of a manually maintained list. Requires **`@pptb/types` ≥ 1.2.3-beta.0** and a host that exposes capability discovery.

- **Automatic discovery** — FetchXML Studio calls `toolboxAPI.invocation.findToolsByCapability("fetchxml")` to find installed tools that declare the `fetchxml` capability, excludes itself, and lists them in the Send to Tool button (single button for one match, dropdown for several). Hidden when no matches or when the host lacks discovery — feature-detected at runtime, so older hosts degrade gracefully.
- **One-way send** — Launches now pass `noReturn: true`, the dedicated flag for the "Send To" pattern, which suppresses the callee's "Return to [Caller]" banner since FXS isn't waiting for data back.

### 🔧 Changed

- **Removed the manual "Tool Integration" settings.** The Settings → Tool Integration section and the `targetTools` display setting are gone; discovery replaces them entirely.

### 🏗️ Technical

- Bumped `@pptb/types` `1.2.2-beta.1` → `1.2.3-beta.0`, which now types `invocation.findToolsByCapability` / `getKnownCapabilityTags`, the `capabilities` field on `InvocationConfig`, the `CapabilityTag`/`KnownCapabilityTag` types, and `launchTool`'s `noReturn` option. Removed the local `InvocationAPI` interface in `invocation.ts` in favor of the now-published host types.
- `src/features/fetchxml/api/invocation.ts` — Added `DiscoveredTool` + `discoverFetchXmlTools()` (feature-detected; filters out our own id; defaults `name` to `id`); added `noReturn: true` to `sendFetchXmlToTool`.
- `src/shared/hooks/useSendToTools.ts` — New hook; discovers send targets once on mount.
- `src/app/AppShell.tsx` — Uses `useSendToTools`; Send to Tool slot now gated on `isT2TSupported() && sendToTools.length > 0`.
- `src/features/fetchxml/ui/Toolbar/SendToToolButton.tsx` — Prop changed from `targetTools: string[]` to `tools: DiscoveredTool[]`; labels use the discovered tool name.
- `src/features/fetchxml/ui/Settings/SettingsDrawer.tsx` — Removed the Tool Integration section, its handlers/state/styles, and now-unused imports.
- `src/features/fetchxml/model/displaySettings.ts` — Removed `targetTools` from `DisplaySettings` and the defaults.

## [1.2.2-beta.2] - 2026-06-16

### ✨ Added

#### Tool-to-Tool (T2T) Invocation — Callee (prefill + return)

FetchXML Studio can now be *launched by* another PPTB tool, pre-populated with an inbound query, and hand a query back. Completes the round-trip alongside the existing caller support.

- **Inbound prefill** — On launch via T2T, FetchXML Studio reads its launch context and pre-populates the builder. Resolution priority is `fetchXml` → `viewRef` → `entityLogicalName`: a `fetchXml` string is loaded directly; a `viewRef` has its FetchXML retrieved from the referenced `savedquery`/`userquery`; and a bare `entityLogicalName` starts a fresh query on that entity. When `fetchXml` is present the root entity is **derived from the FetchXML itself** and any `entityLogicalName` hint is ignored, because a mismatched value would break this metadata-driven tool's lookups.
- **Return FetchXML button** — When launched as a callee, a "Return FetchXML" button appears in the FetchXML toolbar. It returns `{ fetchXml }` (the current editor-or-tree query) to the caller via `returnData()`, which the host resolves as the caller's `launchTool(...)` result. This is intentionally separate from the host-injected "Return to [Caller]" banner, which only navigates back and resolves the caller's promise with `null`.
- **`pptb.config.json`** — Added the `fetchxml` capability tag (for discovery via `findToolsByCapability`) and a `returnTopic` describing the `{ fetchXml }` return shape. Relaxed the prefill contract so `entityLogicalName` is optional/hint-only.

### 🏗️ Technical

- `src/features/fetchxml/api/invocation.ts` — Added callee helpers: `getLaunchContext()`, `resolveLaunchPrefill(context)` (returns a `{ kind: "fetchxml" | "entity" }` action following the `fetchXml` → `viewRef` → `entityLogicalName` priority; resolves `viewRef` by reusing the `dataverseAPI.queryData` view-read route filtered to a single record), `returnDataToInvokingTool(data)`, and `returnFetchXmlToInvokingTool(fetchXml)`. Added `FetchXmlStudioLaunchPrefill` type and `LaunchPrefillAction` union.
- `src/shared/hooks/useLaunchContext.ts` — New hook. Calls `getLaunchContext()` once on mount and exposes `{ loading, isCallee, context }` so detection happens in one place.
- `src/app/AppShell.tsx` — Consumes the prefill exactly once on launch (loads incoming FetchXML via `builder.loadFetchXml`), and passes `isCallee` to `PreviewTabs`.
- `src/features/fetchxml/ui/RightPane/PreviewTabs.tsx` — Added `isCallee?` prop and a `ReturnFetchXmlButton` (gated on `isCallee`) rendered next to Execute; uses a `getCurrentXml()` helper that returns the live editor buffer when in editor mode.

### 🔧 Changed

- **"Send to Tool" button visibility** — The caller button is now hidden entirely unless at least one target tool is configured under Settings → Tool Integration (previously it rendered as a disabled button with a tooltip). The two T2T toolbar actions are now strictly capability-gated and never change meaning by mode: **Return FetchXML** shows only when launched as a callee (`isCallee`); **Send to Tool** shows only when targets are configured.

## [1.2.2-beta.1] - 2026-05-25

### ✨ Added

#### Tool-to-Tool (T2T) Invocation — Caller

FetchXML Studio can now send the active query to any other PPTB tool that accepts a FetchXML prefill payload. Requires **PPTB host ≥ 1.2.2**.

- **Send to Tool button** — A new "Send to Tool" button appears in the toolbar (alongside Save View) whenever the PPTB host exposes the `toolboxAPI.invocation` API. Three interaction modes depending on how many target tools are configured:
  - *No tools configured* — Button is disabled with a tooltip pointing to Settings → Tool Integration.
  - *One tool configured* — Single button; click launches the tool directly.
  - *Multiple tools configured* — Dropdown menu listing all configured tools; click an item to launch.
- **Settings → Tool Integration section** — New section in the Settings drawer to manage the list of target tool npm package IDs (e.g. `@linked365/pptb-bulk-data-studio`). Supports add (text input + Enter/Add button) and remove (× button per item). The list is persisted via `toolboxAPI.settings` alongside other display preferences.
- **Active connection forwarding** — The caller automatically retrieves the currently active Dataverse connection and forwards its ID as `primaryConnectionId` when launching the target tool, so the callee opens against the same environment.
- **`pptb.config.json`** — Added PPTB callee contract file at the repository root. Declares the prefill schema this tool accepts when *receiving* a T2T invocation from another tool (see [Callee Contract](#-callee-contract-pptbconfigjson) in the README). `viewRef` is modelled as a Dataverse EntityReference: required fields `id` (UUID string) and `entityLogicalName` (enum: `"savedquery"` | `"userquery"`).

### 🏗️ Technical

- `src/features/fetchxml/api/invocation.ts` — New module. Defines a local `InvocationAPI` interface (mirrors the PPTB host API not yet in `@pptb/types`) and exposes `isT2TSupported()` and `sendFetchXmlToTool(targetToolId, { fetchXml, entityLogicalName })`.
- `src/features/fetchxml/ui/Toolbar/SendToToolButton.tsx` — New toolbar component. Handles the disabled / single / multi-tool rendering variants; calls `sendFetchXmlToTool`; surfaces errors via `toolboxAPI.utils.showNotification`.
- `src/features/fetchxml/model/displaySettings.ts` — Added `targetTools: string[]` (default `[]`) to `DisplaySettings` interface and `defaultDisplaySettings`.
- `src/features/fetchxml/ui/Settings/SettingsDrawer.tsx` — Added Tool Integration section with add/remove UI. Added `Input`, `Field`, `Add20Regular`, `Dismiss16Regular`, `PlugConnected20Regular` to imports.
- `src/features/fetchxml/ui/RightPane/PreviewTabs.tsx` — Added `sendToToolButton?: ReactNode` prop; rendered unconditionally (all tabs) next to `saveViewButton`.
- `src/app/AppShell.tsx` — Imports `SendToToolButton` + `isT2TSupported`; passes `sendToToolButton` prop to `PreviewTabs` gated on `isT2TSupported()`.
- `package.json` — `features.minAPI` bumped from `1.2.0` → `1.2.2` to reflect the new PPTB host requirement. `scheduler` added to `devDependencies` (required peer dep of `@fluentui/react-context-selector` that was missing from the install tree).

#### Packaging fixes (resolves #33)

- **`icon` path corrected** — `"icon"` field in `package.json` was missing the `icons/` directory prefix (`"icon-insider.svg"` → `"icons/icon-insider.svg"`). Icon was not resolvable by PPTB at install time.
- **`files` array expanded** — Added `"icons"`, `"index.html"`, `"LICENSE"`, `"README.md"`, `"CHANGELOG.md"`, `"SECURITY.md"` alongside the existing `"dist"` and `"npm-shrinkwrap.json"` entries. Previously the `icons/` folder (referenced by the `"icon"` field) and all root-level docs were absent from the published package.
- **Production source maps disabled** — Added `sourcemap: false` to `vite.config.ts` `build` options. Monaco worker source maps were silently adding ~17 MB to the package; disabling them brings the published size down to ~3–4 MB.

---

## [1.2.1] - 2026-04-07

### ✨ Added

- **Copy button in editor mode** — The FetchXML editor toolbar now includes a Copy button when in editor mode, making it consistent with the read-only toolbar. Previously Copy was only available when the editor was not in edit mode.

### 🔧 Fixed

- **Results grid column overlap** — Column widths are now estimated dynamically from header text (capped at 400 px) so adjacent columns no longer overlap on first render. All cells use `TableCellLayout truncate` for clean overflow.
- **Results grid header/row misalignment** — Removed a stray CSS class from `DataGridBody` that was causing header and row widths to drift out of sync when the grid had overflowing content.
- **Alias validation in property editors** — Typing an invalid alias character (space, hyphen, special character) in `AttributeEditor` or `LinkEntityEditor` now sanitizes the value in real time: invalid characters are stripped, and a leading digit is prefixed with `_`.
- **Alias validation in parser** — When parsing FetchXML (both *Parse to Tree* and *Execute*), invalid aliases are now automatically sanitized. A parse warning describes the correction (e.g. *"Alias 'this is id' had invalid characters and was corrected to 'thisisid'."*) instead of silently forwarding the invalid alias to Dataverse where it would cause an API error.
- **Execute blocked when alias warnings are present** — Clicking Execute in editor mode now validates the XML completely before sending it to Dataverse. If parse warnings exist (e.g. an invalid alias that wasn't auto-corrected via *Parse to Tree*), execution is blocked and the warning is surfaced in a message bar above the editor with a hint to use *Parse to Tree* to auto-correct. Previously the bad alias was silently sent to Dataverse and the error appeared in the Results tab.
- **Editor state preserved across tab switches** — Switching from the FetchXML tab to Results or LayoutXML and back no longer resets editor mode or discards the edited XML. The editor component is now always mounted (hidden via `display: none` when not active) so Monaco state — editor mode toggle, cursor position, edit buffer — is fully preserved.
- **Validate contradictory messages** — *Validate* no longer shows both a "Valid" success message and warnings simultaneously. When warnings are present, validate now shows each warning plus an "Action required" notice and stops — the "Valid" success message is only shown when there are zero warnings.
- **Parse to Tree corrections shown as informational** — Auto-corrected alias warnings from *Parse to Tree* are now shown with `intent="info"` (blue) rather than `intent="warning"` (yellow), correctly communicating that the tool has already resolved the problem.
- **Execute API errors surfaced** — Dataverse API errors that occur during query execution are now captured and displayed in a message bar in the Results tab. Previously they were only visible in the browser console.

### 🏗️ Technical

- Added `src/features/fetchxml/model/aliasUtils.ts` — `isValidAlias()` and `sanitizeAlias()` shared utilities enforcing the Dataverse alias constraint (`[A-Za-z_][A-Za-z0-9_]*`).
- `fetchxmlParser.ts` — `parseAttribute` and `parseLinkEntity` now import `sanitizeAlias`; invalid aliases are corrected in place and described in the parse warning rather than propagated as-is.
- `PreviewTabs.tsx` — `FetchXmlEditor` is always mounted; `handleExecute` blocks on parse warnings; navigating back to the FetchXML tab clears any stale editor validation error banner.
- `AppShell.tsx` — `handleExecute` accepts an optional `xmlOverride` (used when editor mode is active); `executeError` state threads Dataverse API errors back to the `PreviewTabs` message bar.

---

## [1.2.0] - 2026-04-03

### ✨ Added

#### Query Scope Settings
- **Entity Scope Mode** — New setting (persisted across sessions) to control which entities appear in the entity dropdown:
  - **Publisher + Solution** — Select a publisher, then narrow to specific solutions; only entities belonging to those solutions are shown
  - **Solution Only** — Skip the publisher step and pick solutions directly
  - **All Entities** — Show every entity in the environment (no solution filter)
- **Advanced Find Only toggle** — When enabled (default), entities and attributes are limited to those marked `IsValidForAdvancedFind = true`, matching the classic Advanced Find behavior. Disabling shows all entities and attributes — including system and developer-only tables — useful for building admin or integration queries. The filter is applied locally so toggling is instant with no additional API call.
- Scope mode and Advanced Find preference are persisted individually via `toolboxAPI.settings` and survive across sessions.

#### Monaco Editor — Local Bundling
- Monaco is now fully bundled in `dist/` instead of being fetched from a CDN at runtime.
- Eliminates CSP errors and ensures line numbers, syntax highlighting, and all editor features work correctly in the PPTB host environment without network access to external CDNs.

### 🔧 Fixed

- **Internal solutions in Solution-only mode** — Solutions from `isreadonly = true` publishers (Microsoft, system publishers) are now excluded from the solution picker, consistent with what Publisher + Solution mode shows. Previously some Microsoft-owned visible solutions appeared in the list.
- **Parse-to-tree no longer resets the entity** — Pasting FetchXML into the editor and clicking "Parse to Tree" no longer wipes the selected entity. The validation that auto-clears the entity selection now only fires when there is an active solution filter in place (i.e. specific solutions have been selected) and the entity has dropped out of scope — not when the entity was set programmatically via the XML editor.
- **Entity list cache race condition** — `loadAllEntities` no longer accepts an `advancedFindOnly` parameter. It always fetches all entities from the server (no `$filter=IsValidForAdvancedFind eq true`). The Advanced Find filter is now applied locally in the EntitySelector `memo`. This eliminates a bug where an earlier `advancedFindOnly=true` load could poison the cache and cause subsequent `advancedFindOnly=false` calls to silently return the wrong (filtered) set.
- **Privilege check console noise** — `checkPrivilegeByName` errors for privileges that don't exist in the environment (e.g. `prvDeleteGitorganization` for system tables) are now logged as `console.warn` instead of `console.error`. The UI behavior is unchanged.
- **CSP exception** — Removed the invalid `*.crm*.dynamics.com` wildcard from `cspExceptions` (browsers reject multi-depth subdomain wildcards). The only `connect-src` exception kept is `https://*.dynamics.com`.

### 🏗️ Technical

- `useAccessMode` preloads only the AF-valid entity metadata cache (`allEntityMetadataCache`) on startup — used by publisher-solution and solution-only modes. The all-entities cache is populated lazily the first time "All Entities" scope mode is used.
- `loadAllEntities` / `useLazyMetadata.loadEntities` signatures simplified (no `advancedFindOnly` param — always fetches everything).
- `vite.config.ts` — `manualChunks` splits Monaco into its own chunk (`monaco-editor-*.js`, ~3.7 MB; `monaco-editor-*.css`, ~145 kB).
- TypeScript strict compile: **zero errors** (`tsc -b --noEmit`).

---

## [1.1.0] - 2026-03-23

### ⬆️ Upgraded

- **`@pptb/types` 1.0.7 → 1.2.0** — now the authoritative source for all PPTB host API types (`window.dataverseAPI`, `window.toolboxAPI`). Added to `tsconfig.app.json` `types` array so global augmentations apply project-wide without per-file triple-slash references.
- **`features.minAPI`** set to `"1.2.0"` in `package.json` to enforce the correct PPTB host version at install time.

### ✨ Added

#### Native File Save Dialog
- Both export paths (local ExcelJS export and Dataverse `ExportToExcel` action) now use `window.toolboxAPI.fileSystem.saveFile` when running inside the PPTB desktop host, presenting a native OS save-file dialog instead of a forced browser blob download.
- Graceful fallback to browser blob/anchor download when running in standalone or dev context (i.e. when `window.toolboxAPI` is not available).

#### Connection Environment Awareness
- `PptbContext` now exposes `environment?: "Dev" | "Test" | "UAT" | "Production"` sourced from the active `DataverseConnection`.
- Subscribes to `connection:updated`, `connection:created`, and `connection:deleted` PPTB events to keep the field in sync dynamically.

#### Package Manifest Hardening
- Added top-level `"icon"` field pointing to `icons/icon.svg` (replacing deprecated `configurations.iconUrl`).
- `"features"` block added: `{ "multiConnection": "none", "minAPI": "1.2.0" }`.
- `cspExceptions` migrated from plain-string arrays to the required object format with `domain` + `exceptionReason` fields (+ `optional: true` on style-src).
- Added `"validate": "pptb-validate"` npm script for pre-publish manifest validation.
- Added `"finalize-package": "npm run build && npm shrinkwrap"` convenience script.

### ♻️ Refactored

#### `pptbClient.ts`
- Removed the entire hand-written `declare global { interface Window { dataverseAPI?: { ... } } }` block (~60 lines). The `window.dataverseAPI` global is now typed entirely by `@pptb/types`.
- Added local `DataverseFetchXmlResponse` interface to capture OData response annotations beyond what `@pptb/types` `FetchXmlResult` declares (`@Microsoft.Dynamics.CRM.totalrecordcount`, `morerecords`, `pagingcookie`, `totalrecordcountlimitexceeded`). These are OData wire annotations returned directly by Dataverse — not fabricated by PPTB — so the cast is safe at runtime.
- `downloadBase64File()` changed from `void` to `async Promise<void>`; uses `fileSystem.saveFile` when available.

#### `usePptbContext.ts`
- Removed the entire hand-written `declare global { interface Window { toolboxAPI?: { ... } } }` block. The `window.toolboxAPI` global is now typed entirely by `@pptb/types`.
- `PptbContext` interface: removed legacy `organizationId?: string` (not present in `DataverseConnection` v1.2.0 types); replaced with `environment?: "Dev" | "Test" | "UAT" | "Production"`.

#### `excelExport.ts`
- `downloadExcelFile()` changed from `void` to `async Promise<void>`; uses `fileSystem.saveFile` when available.

#### `AppShell.tsx`
- Both download call sites updated with `await` to match the new async signatures.

### 🔧 Technical

- TypeScript strict compile: **zero errors** (`tsc -b --noEmit`).
- `pptb-validate`: **✔ Validation passed**.

---

## [1.0.6] - 2026-01-14

### ✨ Added

#### Select Attributes Dialog (#12)
- **Bulk Attribute Selection**: New "Select Attributes" modal dialog for efficient multi-attribute selection
  - Accessible from entity and link-entity context menus in the tree view
  - Works on root entity, link-entities, and nested link-entities at any level
- **Smart Selection Grid**: DataGrid showing all available entity attributes
  - Three sortable columns: Logical Name, Display Name, and Data Type
  - Pre-selects currently selected attributes with checkboxes
  - Multi-select support for bulk add/remove operations
- **Search & Filter**: Integrated SearchBox for quick attribute filtering
  - Filters across all columns (logical name, display name, data type)
  - Real-time filtering as you type
- **Fixed Headers**: Column headers remain visible while scrolling through attributes
- **Confirmation Required**: Changes only applied when user clicks Apply button
  - Cancel button discards changes without affecting the tree
- **Smart Updates**: Intelligently adds new attributes and removes deselected ones
  - Preserves existing attributes that remain selected
  - Maintains attribute order and properties
- **Performance**: Leverages metadata cache for fast attribute loading
  - Falls back to API loading if attributes not cached
  - Shows only attributes valid for Advanced Find

### 🔧 Technical

- Added `SelectAttributesDialog` component with FluentUI DataGrid
- New `UPDATE_ATTRIBUTES` action in builderStore for bulk attribute operations
- Enhanced TreeView with dialog state management and attribute loading
- Integrated with existing metadata caching system

## [1.0.0] - 2024-12-15

### 🎉 Initial Release

FetchXML Studio v1.0.0 - A powerful FetchXML query builder for Power Platform ToolBox.

### ✨ Features

#### Query Builder
- Tree-based visual query builder with hierarchical FetchXML structure
- Full entity, attribute, filter, order, and link-entity support
- Context-aware properties panel for each node type
- Nested filter groups with AND/OR logic
- Support for all FetchXML operators (60+ operators)
- Aggregate queries with groupby and distinct
- Query hints for performance optimization

#### Relationships & Joins
- Relationship picker with 1:N, N:1, N:N categorization
- Inner and outer join support
- Filter link-entity for any/all/not any/not all scenarios
- Value-of conditions for cross-entity comparisons
- Entity name conditions for outer join filtering

#### Results Grid
- Virtualized DataGrid for large datasets (react-window)
- Rich cell rendering for lookups, option sets, dates, currency
- Multi-column sorting (click + Shift+click)
- Row selection (single and multi-select)
- Resizable and reorderable columns
- Formatted values with raw value toggle

#### FetchXML Editor
- Monaco editor with XML syntax highlighting
- Bi-directional editing (visual ↔ XML)
- Parse FetchXML to visual builder
- Copy to clipboard
- LayoutXML preview tab

#### Views
- Load system and personal views
- Optimized view execution via SavedQuery/UserQuery APIs
- Save queries as personal views
- Update existing views
- Solution-aware save (add to solution)

#### Data Operations
- Export to Excel with native data types
- Delete selected records
- Bulk delete jobs (async)
- Run on-demand workflows on selection
- Batch progress tracking with ETA

#### User Experience
- Dark/Light theme support (follows PPTB theme)
- Lazy metadata loading
- Intelligent caching (per-session)
- Resizable split panes
- Keyboard shortcuts (Ctrl+Enter to execute)

#### Security
- Privilege-aware operations
- Export/Delete privilege checks
- Bulk delete privilege validation
- Workflow execution privilege checks

### 🏗️ Technical

- Built with React 19 and TypeScript 5.9
- Fluent UI v9 components
- Monaco Editor integration
- ExcelJS for native Excel export
- Vite build tooling
- Power Platform ToolBox integration via @pptb/types

### 📚 Documentation

- Comprehensive README with feature overview
- Debug logging system documentation
- Project structure guide
- Contributing guidelines

---

## Development History

### Pre-release commits (dev branch)

- `feat: Beta release enhancements - Settings, Export, and UX improvements`
- `feat(preview): add LayoutXML tab to preview pane`
- `feat(commands): add record action commands - delete, bulk delete, workflow`
- `feat(columns): Add/Edit Columns panels with related entity support`
- `feat(export): add Export to Excel with status feedback and privilege check`
- `feat(paging): fix paging with cookies and page size support`
- `feat(grid): multi-entity metadata & column display name fixes`
- `feat(builder): Sprint 1 - Foundation & Layout improvements`
- `feat(save): implement Save View functionality`
- `feat(sort): Add multi-column sort with FetchXML integration`
- `feat(results): Add Column from Results Grid`
- `feat(views): Load view LayoutXML for column configuration`
- `feat(layout): Add column resize and reorder`
- `feat(layout): Add LayoutXML foundation`
- `feat(filter): Add filter link-entity support for any/not any/all scenarios`
- `feat(condition): Add valueof cross-entity comparison`
- `feat(condition): Add valueof same-row column comparison`
- `feat(condition): Add entityname support for outer join filtering`
- `feat(views): Add Load View picker and optimized view execution`
- `feat(parser): Add FetchXML parser with Monaco editor dialog`
- `feat: Implement Phases 1-4 - DataGrid UI polish`
- `feat: add filter consistency, cascading validation, and global metadata caching`
- `feat(entity-selector): implement advanced filtering with publisher/solution multiselect`
- `feat(ui): add resizable panes with visual grip indicators and command bar`
- `Initial commit - setup main branch`
