# Changelog

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
