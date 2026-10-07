# Changelog

## 1.1 (07-10-2026)

Paste-only release: the tool no longer opens, reads or links to any other file.

- **Removed: Excel/CSV file import.** Import Client List, Import FIU Requests, the file picker, opening of source workbooks, CSV parsing, header mapping, preview/commit, the temporary file copy, and the Import, Import Log and Staging sheets are gone. Evidence-folder selection and PDF export are unchanged.
- **Paste Special > Values** into the Client List and FIU Requests tables, with instructions and zone labels on each sheet.
- **New Validate FIU Requests button**; Validate Client List now also sets the client-list revision. Both checks also run at the start of every manual search and every batch.
- **Formulas are refused** before any search or batch, with the cell addresses and how to replace them with values. Text that only looks like a formula, and IDs with leading zeros, stay text.
- **Rows pasted directly under a table are added to it.** Data anywhere else outside the table (after a blank row, to the right, or in the instruction rows) stops the check with its cell addresses instead of being ignored.
- **FIU Requests columns reordered**: the eleven input columns come first, then the columns filled by the tool (grey), then the Response columns (green). A Request ID or Screening Status the tool did not write (for example a paste that was too wide) stops the batch.
- **Repeated reference + name** rows are marked Error and not searched (previously skipped at import). The earliest searchable row is kept.
- **Client-list revision**: starts only when the content changes (sorting does not), is recorded with its record count on every run, search and PDF, and two changes in the same second get different revisions.
- Generated Record IDs and Request IDs are never reused after rows are deleted.
- Client List: the Import ID and Loaded At columns were removed.
- **Client List and FIU Requests are now protected** (no password). Input columns, Response columns and 20,000 empty rows under each table stay editable; headers, instructions, zone labels and the tool's columns are locked. Macros still update the locked columns. New **Delete Selected Rows** button on both sheets, because Excel blocks deleting rows on protected sheets. Excel limits sorting and filtering on protected sheets.

## 1.0 (07-10-2026)

First public release.

- **Standalone and offline.** Each user loads their own client list as a local snapshot (paste, or import from Excel/CSV with header mapping). No live links to other workbooks.
- **VBA matching engine.** Runs on Excel 2021 and Microsoft 365, 32-bit and 64-bit.
- **Whole-word matching (rules MR-1.0).** Names are used in full, names containing digits are searched and flagged, and a short, editable list of trailing suffixes is removed for entities only.
- **Errors are never reported as "No candidate found".** Formula errors, broken input, an empty or invalid client list, and names made only of suffix words become an Error with the reason.
- **No invented decisions.** Zero-candidate searches are recorded as an engine result, without a reviewer or review decision.
- **History is kept.** Rerunning a request keeps the earlier search, decision and PDF in Search History.
- **Evidence PDFs.** One for every search, never overwritten, multipage landscape layout, and each footer shows its own Evidence ID.
- **Stable IDs** for requests, searches, candidates, runs and evidence. Evidence and record IDs are never reused.
- Candidate review with selected candidate IDs; pause, cancel and resume; evidence retry; client-list versioning; import log; Search History.
- English and Arabic names, customer type, request type, FIU status, publication and due dates, reference number, review details, evidence path, response status and date, and the OVERDUE / ESCALATE / REPEAT NAME / CHECK NAME flags.
- Importing from a workbook that is already open reads its current contents, including unsaved edits, and leaves the file open and untouched. The read time is logged.
- Dates: year-only values, plain numbers, two-digit years, invalid dates and day/month-ambiguous dates are kept as typed, flagged, and excluded from date comparisons. Ambiguous dates are converted only when the user sets Text date order.
- Buttons work after the file is renamed or moved, and reapply sheet protection when the file was opened without its startup routine.
