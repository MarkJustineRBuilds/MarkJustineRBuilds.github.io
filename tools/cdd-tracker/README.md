# CDD Tracker

An Excel workbook for tracking customer due diligence (CDD) requirements per client case, chasing relationship managers (RMs) and handing documents over to compliance.

Version 1.0 · By Mark Justine

## What it does

- **Case Desk:** one screen per case. Overdue and due-today follow-ups come first. Tick requirements and move them along: Mark received, Notify compliance, Complete, Reject.
- **Requirements by status:** Outstanding → Requested from RM → Received - awaiting compliance → Sent to compliance → Completed (plus Rejected, Withdrawn and Internal action).
- **RM chasers:** puts a case's open items into the email you have open in Outlook, or builds an RM Report with one Outlook draft and one Excel file per RM.
- **Compliance handover:** drafts an email to your compliance team listing the documents that are ready, with an optional folder link.
- **Your own tasks:** internal work such as form updates is tracked as Tasks with due dates (View: My tasks).
- **Quick entry:** paste a requirement list from any email (Email Paste) or a four-column table (Paste Import). Optional Outlook import for emails in a fixed format.
- **Audit trail:** every status change, note and follow-up is logged in Change History and Follow-up History. Every button click is recorded on the Log sheet.
- **Dashboard:** open requirements, overdue cases, items awaiting compliance, active cases and open tasks.
- **Automatic backups** before imports and bulk updates.

## Requirements

- Windows with Excel desktop (Microsoft 365 or Excel 2016+)
- Classic Outlook for Windows (optional, only for email drafts and Outlook import). The new Outlook and Outlook on the web cannot be used by Excel macros. Everything else works without Outlook.
- Mac: not supported. Buttons show text only and the email features will not work.

## Get started

1. Download the zip and unblock it before extracting. See HOW-TO-UNBLOCK.md.
2. Extract the folder and open `CDD-Tracker.xlsm`. Click **Enable Content** if asked.
3. Read the **Start Here** sheet.
4. Click **Settings** on Start Here and fill in at least `MyName`, plus `ComplianceTo` if you will use Notify compliance. Click **Check settings**.
5. Explore the six sample cases on the Case Desk, then click **Clear sample data** when you're ready to use your own.

## Buttons

| Sheet | Button | What it does |
|---|---|---|
| Start Here | Go to Case Desk | Opens the Case Desk |
| Start Here | Check settings | Lists invalid settings and the features that are switched off because a setting is blank |
| Start Here | Repair layout | Restores buttons, dropdowns, protection and missing settings. Never deletes data |
| Start Here | Clear sample data | Deletes only the sample rows (IDs starting `SAMPLE-`, plus the sample clients and RMs), after a confirmation |
| Start Here | Settings / Lists / Templates / RM Directory / Full guide | Opens that sheet (these sheets are hidden otherwise) |
| Settings, Lists, Templates, RM Directory, Instructions | Back to Start Here | Returns to Start Here and hides the sheet again |
| Dashboard | Import Outlook email | Imports the requirement table from the selected Outlook email (needs `ImportSubjectPrefix`, see below) |
| Dashboard | Apply update email | Applies an "... - Update" email with explicit Added / Amended / Fulfilled / Withdrawn rows |
| Dashboard | Go to Case Desk | Opens the Case Desk |
| Case Desk | Search cases (Ctrl+Shift+F) | Finds an active case by client, person/company or RM |
| Case Desk | Refresh | Rebuilds the desk from the tables |
| Case Desk | RM Report | Outstanding RM items on active cases: one Outlook draft per RM with an Excel file attached, or the Excel files only |
| Case Desk | Open source email | Opens the Outlook email a case was imported from (Outlook import only) |
| Case Desk | Draft to open email | Inserts the case's open RM items at the cursor of the email open in Outlook |
| Case Desk | Mark received | Ticked items → Received - awaiting compliance |
| Case Desk | Notify compliance | Ticked received items and open Tasks → Sent to compliance, and drafts an email to `ComplianceTo` |
| Case Desk | Complete | Ticked items → Completed. Offers to close the case when nothing is open |
| Case Desk | Reject | Ticked items → Rejected, with an optional reason in the note |
| Case Desk | Add requirement | Adds a document, clarification, screening or Task, with optional due date and note |
| Case Desk | History | Imports, status changes and follow-ups for the case |
| Case Desk | Clear ticks | Removes all ticks |
| Case Desk | Client docs / Compliance | Opens the client's folder link from RM Directory (optional) |
| Email Paste | Convert to Paste Import | Turns a pasted email list into Paste Import rows |
| Paste Import | Import pasted rows | Validates and imports the rows as a new case or into an open case |
| Cases | Record follow-up | Logs a follow-up for the selected case row |
| Cases | Draft RM email | New Outlook draft with the selected case's open RM items |

## Settings

Open the Settings sheet with the **Settings** button on Start Here. Only the Value column is editable. Blank optional settings switch that feature off.

| Setting | Default | Purpose |
|---|---|---|
| MyName | (blank) | Optional. Your name, for the `{MyName}` placeholder |
| FollowUpWorkdays | 3 | Next follow-up = today + this many working days (also `{DueDate}`) |
| WeekendPattern | 0000011 | Weekend days, Monday first, 1 = weekend (0000110 = Fri-Sat) |
| AgeAmberDays / AgeRedDays | 7 / 14 | Requirement age colours on the Case Desk |
| ComplianceTo / ComplianceCC | (blank) | Recipients for Notify compliance |
| ClientDocsRootLink / ComplianceRootLink | (blank) | Optional main folders (web link or folder path) |
| ImportSubjectPrefix | (blank) | Turns on Outlook import (see below) |
| BackupFolder / BackupKeep | Backups / 10 | Automatic backups, next to the workbook |
| ReportsFolder | Reports | Where RM Report Excel files are saved |
| RMReportDetailMax | 5 | Above this many items, the RM email shows a per-client summary |
| DefaultReviewType | Onboarding | Default review type on Paste Import |
| ShowIcons | TRUE | Icons on buttons |
| ColorPrimary / ColorAccent / ColorDanger | 23,50,77 / 14,124,134 / 192,0,0 | Button and heading colours (R,G,B) |

Email wording is on the **Templates** sheet (Templates button on Start Here). Placeholders: `{RMName}`, `{ClientName}`, `{MissingDocs}` (the item list), `{DueDate}`, `{Count}`, `{ClientCount}`, `{Date}`, `{MyName}`, `{ComplianceFolderLink}`. Write `**text**` for bold. Your Outlook signature is added automatically.

## Outlook import (optional)

Outlook import is **off** until you fill in `ImportSubjectPrefix`. Until then, its buttons show a message and change nothing. Email Paste and Paste Import work for any email.

It reads the email selected in classic Outlook when that email follows this format:

- **Subject:** `<ImportSubjectPrefix> - <Client name>`, optionally followed by ` - <Review type>` (one of the review types on the Lists sheet; otherwise `DefaultReviewType` is used). `RE:` and `FW:` are ignored. A subject ending in ` - Update` is an update email.
- **Body:** an HTML table whose header row is exactly `No.` | `Company / Individual` | `Request type` | `Required document / action`. Each row is one requirement. `No.` is a whole number, unique within the email, and `Request type` must be on the Request types list.
- **Update emails** add a fifth column, `Change`, with `Added`, `Amended`, `Fulfilled` or `Withdrawn` for each row. A row left out never counts as completed.

Example: with `ImportSubjectPrefix` = `CDD Request`, the subject `CDD Request - Sample Client 01 LLC - Periodic Review` creates or updates a Periodic Review case for Sample Client 01 LLC.

Each email is imported once; a preview is shown before anything is written.

## Sheets you see

Seven sheets are always visible: Start Here, Dashboard, Case Desk, Email Paste, Paste Import, Cases and Requirements.

- **Setup sheets** (Settings, Lists, Templates, RM Directory and the Instructions guide) are hidden and open from the buttons on Start Here. **Back to Start Here** closes them again.
- **History sheets** (Change History, Follow-up History, Import Log and the Log) are hidden. The tracker writes to them in the background; the Case Desk **History** button shows a case's history.

Right-click Unhide is blocked by the workbook protection. To see every sheet, use maintenance mode (below).

## Your data stays with you

The workbook does not send data anywhere. Email features create Outlook drafts for you to review. Nothing is sent automatically.

## If something looks broken

Click **Repair layout** on the Start Here sheet. It restores buttons, formatting and protection without touching your data. If RM Directory or one of the editable lists runs out of blank rows, Repair layout adds more.

Sheets are protected to prevent accidental changes. Password: none (the sheets are protected without a password). It is a safety catch, not security. Unprotect only if you know what you're changing.

## Advanced: maintenance mode

Maintenance mode is for inspecting or correcting the hidden sheets (for example a history row).

1. Press **Alt+F8**, choose **ToggleMaintenanceMode** and click **Run**, then confirm.
2. Every sheet is shown and a red **MAINTENANCE MODE** banner appears on Start Here. Sheets stay protected against accidental edits (password: none, so Review > Unprotect Sheet works if you really need it).
3. Run **ToggleMaintenanceMode** again to return to the normal view.

Closing the file always returns to the normal view: the next time it opens, the hidden sheets are hidden again. **Repair layout** also restores the normal view.

## Updating to a new version

Download the new version and copy your rows from the old tables into the same tables in the new file (paste values only). Keep a copy of your old file until you've checked everything moved across.

## Customising

Request types, task types and document types are on the **Lists** sheet (Lists button on Start Here). Add, rename or remove them there; no code changes are needed. Keep the request types `Screening` and `Task`: the tracker treats them as internal work that is never sent to RMs. Statuses, case statuses and review types are fixed in this version. The VBA source is in the `src` folder if you want to adapt the code.

## Disclaimer

This is a free, unofficial tool provided as is, without warranty. It is not legal or compliance advice and does not replace your firm's procedures or screening systems. See DISCLAIMER.md.

## License

Licensed under Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0). See LICENSE.
