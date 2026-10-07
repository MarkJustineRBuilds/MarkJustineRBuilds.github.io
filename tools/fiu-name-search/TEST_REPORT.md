# Test report: FIU Name Search 1.1

Date: 07-10-2026. Environment: Windows 11, Microsoft 365 Excel 64-bit (16.0), desktop.

All results below were produced on copies of the **final** release files, after the last code change. Earlier passing runs were not counted.

| File | SHA-256 |
|---|---|
| FIU_Name_Search_Public_v1.1.xlsm | F64887F9295B4365A263666CA23405A618594D9149C7984B35B9359DCA5F06E9 |
| FIU_Name_Search_Demo_v1.1.xlsm | F58B4F645E3AD205F1A6B19F285E16BFC7E6861B983400D37F7DB42BC8D2DA8E |

## What changed in 1.1

External Excel/CSV import was **removed**, not hardened. Client List and FIU Requests are now **protected**: only input and response cells (and empty rows under each table) are editable. Data is entered by Paste Special > Values into the Client List and FIU Requests tables and validated by the tool. The 1.0 import tests (CSV, xlsm with an external link and a Workbook_Open macro, open-workbook import) were removed with the feature and were **not** re-run. No test source files were created for 1.1: all test data is fictional and is pasted from an unsaved scratch workbook in the same Excel session.

## 1. Automated checks (Excel automation, real workbook code)

### 1.1 Build and compile
- Both workbooks were built from the `src` modules, round-tripped through `.xlsx` to drop any earlier compiled code, re-imported, and compiled in Excel. **Compile: OK** for both. They contain only the 7 modules in `src` plus the sheet and workbook modules (modImport is gone).

### 1.2 Runtime suite: 232 checks, run on a copy of each file

| File tested | Result |
|---|---|
| Copy of the Public release | **PASS 232 / FAIL 0** |
| Copy of the Demo release | **PASS 232 / FAIL 0** |

The whole suite runs with Client List and FIU Requests protected, so every paste, validation, batch, review and PDF above was done with the protection in place. Prompts were answered by a scripted test harness. Any prompt the harness did not expect fails the test instead of waiting.

| Area | Checks | Covered |
|---|---|---|
| Normalisation | 14 | Unchanged from 1.0. |
| Matching | 17 | Unchanged from 1.0. |
| Client list | 8 | Empty list and invalid rows block the batch; requests stay New; missing Record IDs generated. |
| Batch | 19 | Unchanged from 1.0 (pending vs no candidate, no invented review, Error rows, leading zeros, repeat names, settings restored). |
| Review | 16 | Unchanged from 1.0. |
| Stable IDs / sorting | 2 | Requests and history sorted; links and Review Pending still correct. |
| Rerun, pause / resume / cancel, error handling, PDF failure / folders, paths, multipage, manual search | 44 | Unchanged from 1.0. |
| **Pasted client list** | 16 | 7 rows pasted as values grow the table; leading-zero Record ID and reference kept; IDs generated; same-name people separate; `=SUM(A1:A2) Literal Co`, `=1+1`, `+971 Trading`, `-0042` stay literal text in the table, the history, the review panel and the PDF; Natural / Juridical Person accepted; Arabic kept; unrecognised date noted, not blocking; revision and count recorded. |
| **Paste problems** | 17 | A formula from a normal paste is refused with its cell address by Validate, by manual search and by Run Batch (nothing searched); after Paste Values it validates. Blank Party Type and repeated Record ID block with row numbers. Rows written directly under the table are taken in. Data after a blank row, to the right of the table, and in the instruction area are each rejected with the address. Zero-length text left by Paste Values is not treated as data. Sorting keeps IDs with rows and keeps the revision. |
| **Pasted FIU requests** | 17 | A paste one column too wide (value in Request ID) is refused. A formula in Reference Number refuses the batch with its address; nothing is marked searched. Validate notes the repeated reference + name and a text due date, and assigns Request IDs. Batch: the copy is Error, the first row is searched; leading zeros, same name under other references, Juridical Person, digits warning, formula-looking literal name, real dates. The run records the validated revision and record count. A PDF for every searched request. A corrected row pasted after a broken one is searched; the broken row is Error. |
| **Snapshots after list change** | 7 | A decision is saved, then the list is edited and a row deleted. The new search uses the new revision; the earlier search keeps its revision, decision, candidate names as searched and PDF; the two runs record different revisions. |
| Recovery, versioning | 5 | Interrupted run; half-searched request becomes Error; an edit creates a new revision, an unchanged list keeps it, two changes in the same second get different revisions. |
| Sample data / repair | 4 | Unchanged from 1.0. |
| Dates | 26 | Date rules unchanged; the import part was replaced by pasted values: real date cell, number 1980, text 1980, two-digit year, ambiguous, invalid, blank and unambiguous text. |
| Protection fallback | 4 | Unchanged from 1.0. |
| **Protection** | 15 | Both paste sheets protected; input, Source Label and Response columns unlocked; Validation, Request ID, Screening Status and Flag locked; headers, instructions and zone labels locked; spare rows under the table unlocked for input columns only; cells right of the table locked. Rows pasted under the protected table are taken in by validation, the macro writes the Validation column and the new rows are relocked. A batch writes status, IDs and evidence into the locked columns and restores the Flag formula. Delete Selected Rows removes the selected row and keeps the sheet protected; outside the tables it does nothing. |
| Global | 1 | Calculation mode unchanged after the whole suite. |

After the run, only the test copy was open in that Excel session, and it had no external links and no connections.

### 1.3 Reopen, protection and portability (Verify, on copies of the final files)
- Demo opened with startup macros skipped: a write to a locked history cell was blocked (error 1004); after the first button macro, writes worked; no sheet is left unprotected; Excel state restored.
- Demo batch: all 10 sample requests behaved as in 1.0 (9 Pending Review, 1 No Candidate Found).
- Normal reopen: Start active; five visible tabs; no sheet left unprotected after the batch.
- Public release copied to another folder under another name: Start first; 7 supporting sheets hidden, none very hidden; structure protected; **no external links, connections or queries**; **all 34 buttons point to existing public macros and none is tied to a file name**; navigation and repair macros ran; the release opens empty (one Flag formula placeholder in FIU Requests, as in 1.0).

### 1.4 Protection as a person meets it
A copy of the Public file was opened with startup macros off, so the macros' UI-only exemption was not active and Excel applied the saved protection exactly as it does to typing:

| Action | Result |
|---|---|
| Type in Client List input cells (Full Legal Name, Party Type) | allowed |
| Type in Client List Validation, header, instructions, zone label, or right of the table | blocked |
| Type in FIU Reference Number and Response Status | allowed |
| Type in FIU Request ID, Screening Status or a header | blocked |
| Paste Special > Values of 2 rows under the client table | allowed |
| Paste Special > Values 11 columns wide (reaching the Validation column) | blocked |
| Delete a sheet row with Excel | blocked (use Delete Selected Rows) |
| Filter the client table through the object model | blocked |
| Sort the client table through the object model | allowed |

Then the Validate buttons ran (0.5 s): the 2 pasted rows were taken into the table, the macro wrote "OK" into the locked Validation column, assigned a Request ID in the locked FIU column, and both sheets stayed protected. Sorting and filtering from Excel's own header buttons were not tried by hand (section 4).

### 1.5 No external file access (static check of `src`)
The source was searched for `Workbooks.Open/Add/OpenText`, `GetOpenFilename`, `FileDialog`, `QueryTables`, `Connections`, ODBC/OLEDB/ADO, `FileCopy`, file reads (`Open ... For Input/Binary`, `Line Input`, FileSystemObject), `Dir(`, `Kill`, `Shell`, link updates and web functions. Remaining hits, all retained by design: the evidence **folder** picker (`FileDialog(4)`), `Shell explorer.exe` for Open Evidence Folder, the evidence-folder write test (writes and deletes its own `.tmp` file), `Scripting.Dictionary`, and the Esc-key read (`GetAsyncKeyState`). **No routine opens, reads or links an external Excel or CSV file.**

### 1.6 Sanitisation
- Release-skill leak scan of the distribution folder: **HIGH 0**. MEDIUM 14: the 7 approved hidden supporting sheets in each workbook (accepted; their tables are empty in the Public file apart from the Flag placeholder). INFO: compiled VBA, scanned separately below.
- Compiled VBA (oletools decompression of every stream, both files): 14 "email" HIGH hits are in still-compressed raw streams, where compression markers split words such as `Scripting.Dictionary` around an at-sign (same false positive as 1.0). The decompressed module source has **no** email address. No user paths, network paths, OneDrive/SharePoint links or organisation terms. The Windows user name appears only as the English word "Mark" ("Mark at least one candidate") and in procedure names such as MarkError.
- Private-data cross-check: 137 private values from the original workbook (names, references, notes, reviewer names, paths, organisation terms) were searched for in every XML part, document, `src` file and decompressed VBA stream. Held in memory, not printed. **0 hits.**

## 2. Visual checks (by reading the generated PDFs)
- Manual-search PDF for the client `=SUM(A1:A2) Literal Co`: the name, the government ID `=1+1` and the search name print as literal text.
- Reviewed batch PDF (pasted request 0012345): REVIEWED – NAME MATCH banner, Arabic name left-aligned, selected candidate, reviewer and note, and the client-list line "CLV-… (revision from …; validated …; 10 records; pasted Client List) – records searched: 10".
- Demo reviewed PDF (No Match, fictional reviewer) rendered to PNG for the download page.

## 3. Code review
A structured review of the 1.0 → 1.1 source diff found 8 items. Fixed and covered by tests: a corrected re-paste was being treated as a copy of a broken row; zero-length text from Paste Values counted as data outside the table; help text overstated which grey columns are checked; an over-long Request ID could overflow; two revisions in the same second could share an ID. Not changed (known limitations, section 4): pasting onto the instruction or band-label cells is accepted; validation of very large tables is row-by-row.
A second review of the protection change found 7 items. Fixed: a validation message told users to filter the Validation column, which may not work on the protected sheet. Not changed (section 4): relocking cost on each protect call, sort/filter limits, a formatted paste carrying a locked format, the 20,000-row paste area, deletion of already-searched requests (warned), and a style point.
The built-in /security-review was **not run**: it needs a git branch, and this release folder is not a git repository.

## 4. Not tested, or known limitations (please check manually)
- **Real clicking and real pasting by a person.** Pasting was done by Excel's own Copy and PasteSpecial through automation; every button macro ran, but MsgBox, InputBox and the folder picker were answered by the harness. Please paste a few rows and click through once yourself.
- **Sorting and filtering on the protected tables** from Excel's header buttons were not tried by hand. Through the object model, filtering was refused and sorting was allowed; Excel's interface normally refuses to sort ranges that contain locked cells.
- **A normal paste (Ctrl+V) with formatting** may carry the source's locked format into input cells until the next validation relocks them. Not tested by hand.
- **More than 20,000 new rows at once** reach locked cells and Excel refuses the paste; paste in parts and validate between them.
- **Run time:** the test suite took about 3.5 minutes per file after protection, against 1.5 minutes before, because the input cells are relocked each time a macro re-protects a paste sheet. A single Validate took 0.5 s on a small table; large tables were not timed.
- **Delete Selected Rows** can delete a request that was already searched, after a warning; its searches, decisions and PDFs stay in Search History.
- **Only Request ID and Screening Status are checked for values the tool did not write.** Values pasted into the other grey columns are not detected. Paste only into the white input columns.
- **Values pasted onto the instruction cells** (A1, A2 or the zone labels in row 4) are not reported.
- **Numbers that only display leading zeros** (number 123 formatted as 000123) arrive as 123 with Paste Values; there is nothing for the tool to detect. Format such IDs as text in the source.
- **Esc to pause**: tested through a hook, not a physical key press.
- **Excel 2021 and 32-bit Excel**: not available; static evidence only (`PtrSafe` with `#If VBA7`, late binding, no dynamic-array functions).
- **Mark-of-the-Web / Protected View** on another PC: Zone.Identifier simulated only (section 5).
- **Large volumes**: tested with up to 130 client records and 120 candidates per search. Validation reads some cells one at a time, so very large FIU tables (10,000+ rows) may validate slowly; not measured.
- Write-protected folders, PCs without a PDF driver, non-English Excel or other regional settings: not tested.
- Workbook and sheet protection has no password. It is a safety catch against accidents, not security: anyone can unprotect a sheet.

## 5. Package
- Zip: `FIU_Name_Search_v1.1.zip`. Its SHA-256 is published alongside it (the zip cannot contain its own checksum).
- Contents: the two workbooks, `src/` (7 modules + ThisWorkbook), README, TEST_REPORT, CHANGELOG, SCOPE_NOTE, DISCLAIMER, HOW-TO-UNBLOCK and LICENSE. No private build tools, test files, logs or config.
- Leak scan of the zip and download simulation: see `release-notes.md` for the zip hash; results are recorded in the private leak report.
