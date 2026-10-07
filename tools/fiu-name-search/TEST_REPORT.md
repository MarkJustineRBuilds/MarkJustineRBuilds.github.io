# Test report: FIU Name Search 1.1.1

Date: 07-10-2026. Environment: Windows 11, Microsoft 365 Excel 64-bit (16.0), desktop.

All results below were produced on copies of the **final** release files, after the last code change. Earlier passing runs were not counted. All test data is fictional.

| File | SHA-256 |
|---|---|
| FIU_Name_Search_Public_v1.1.1.xlsm | 106F1B42D3D59A8CD1FFCF13F076801D154329CA24AC004494285F584F2B23F7 |
| FIU_Name_Search_Demo_v1.1.1.xlsm | 3D271C477BB956EAC786B823CB80224362A8CC46EDAE95BDEB3FDF56B6C7E717 |

## What changed in 1.1.1

- **Problem reported:** in 1.1, a multi-row Paste Values into FIU Requests failed with "The cell or chart you're trying to change is on a protected sheet", while typing worked.
- **Cause, reproduced on the published 1.1 file with Excel's own ribbon Paste Values command:** a 3-row paste of the 11 input columns raised the message. The Flag column was an Excel formula (calculated) column, and when a paste grew the table Excel tried to fill that formula into the new rows' locked Flag cells. A 14-column paste of the FIU export was refused outright, because the 12th column was the locked Request ID. The same pastes on Client List (no formula column) worked.
- **Fix:** FIU Requests now starts with the 14 FIU export columns (SNO to CHECKERDATE), all unlocked, and Flag is written by the tool instead of by a formula column.

## 1. Clipboard paste with Excel's own command (not array assignment)

A copy of the final Public file was opened normally, so it carried its real protection. Rows were copied from a separate workbook and pasted with the ribbon **Paste Values** command (`CommandBars.ExecuteMso "PasteValues"`); one case used the ribbon **Paste** (Ctrl+V equivalent). Any Excel message was recorded. Afterwards the tool's Validate buttons were run.

| Paste | Result |
|---|---|
| Client List, 3 rows x 9 columns at A6 (first data cell) | all 27 cells landed; the table grew from 1 to 3 rows; no message |
| Client List, 2 rows x 9 columns directly under the table | landed; no message |
| Client List, 2 rows x 3 columns with normal Paste | landed; no message |
| FIU Requests, 3 rows x **10** columns (SNO to REFNUMBER) at A6 | all 30 cells landed; the table grew from 1 to 3 rows; no message |
| FIU Requests, 3 rows x **14** columns (SNO to CHECKERDATE) directly under the table | all 42 cells landed; no message |
| FIU Requests, 2 rows x 14 columns inside the table (A7) | all 28 cells landed; no message |
| FIU Requests, 1 row x **15** columns (one past CHECKERDATE) | refused by Excel with the protected-sheet message, as intended |

Validate then reported "Client list OK: 5 record(s)" and "FIU requests OK: 6 request(s) waiting", including the rows pasted under each table. Request IDs were assigned. PUBDATE, DUEDATE and CHECKERDATE pasted as `2026-10-01` text were shown as `01-10-2026`, and both sheets were still protected.

The same paste on the published 1.1 file, for comparison: the 11-column FIU paste gave the protected-sheet message and the 14-column paste landed nothing.

## 2. Automated checks (Excel automation, real workbook code)

### 2.1 Build and compile
Both workbooks were built from the `src` modules, round-tripped through `.xlsx` to drop any earlier compiled code, re-imported, and compiled in Excel. **Compile: OK** for both: 7 modules plus the sheet and workbook modules.

### 2.2 Runtime suite: 241 checks, run on a copy of each file

| File tested | Result |
|---|---|
| Copy of the Public release | **PASS 241 / FAIL 0** |
| Copy of the Demo release | **PASS 241 / FAIL 0** |

The suite runs with both paste sheets protected. Pastes inside the suite use Excel's Copy and PasteSpecial from a scratch workbook (VBA), which carries the macros' protection exemption. That is why section 1 tests the user's path separately. Prompts were answered by a scripted harness.

| Area | Checks | Covered |
|---|---|---|
| Normalisation, matching | 31 | Unchanged from 1.1. |
| Client list, batch, review | 43 | Unchanged from 1.1, run against the new FIU column layout. |
| Stable IDs / sorting | 2 | Requests and history sorted; links and Review Pending still correct. |
| Rerun, pause / resume / cancel, error handling, PDF failure / folders, paths, multipage, manual search | 44 | Unchanged from 1.1. |
| Pasted client list | 17 | As 1.1, plus a `1985-03-12` text date stored as a date and shown `12-03-1985`. |
| Paste problems | 17 | Formula cells refused with their address (validate, manual search, batch); blank Party Type and repeated Record ID; rows under the table taken in; data after a blank row, to the right and in the instructions rejected; zero-length text ignored; sorting keeps IDs and revision. |
| **Pasted FIU requests (14-column layout)** | 22 | 3 rows pasted with all 14 export columns at A6 and 3 rows with 10 columns under the table: all 6 taken in. A 15-column paste (value in Request ID) refused. A formula in REFNUMBER refuses the batch with its address. Repeated REFNUMBER + name marked Error; the first row searched. YYYY-MM-DD DUEDATE stored as a date shown DD-MM-YYYY; MAKERDATE with a time keeps the time; SNO, MAKER, CHECKER kept as pasted; source STATUS kept separate from Screening Status; leading zeros; Juridical Person; digits warning; formula-looking literal name; run records the client-list revision and count; a PDF for every searched request; corrected re-paste searched. |
| Snapshots after list change | 7 | Unchanged from 1.1. |
| Recovery, versioning | 5 | Unchanged from 1.1. |
| Sample data / repair | 4 | Clear Sample Data and Repair Layout keep the other rows. |
| Dates | 26 | Unchanged rules; pasted date values. |
| Protection fallback | 4 | Unchanged from 1.1. |
| **Protection** | 18 | Both sheets protected; all 14 FIU export columns and the operator columns unlocked; Request ID, Screening Status and Flag locked; headers, instructions and zone labels locked; spare rows under the table unlocked for input columns only. Rows pasted under the table taken in and relocked. A batch writes status, IDs and evidence into locked columns; Flag holds no formula. A past DUEDATE shows OVERDUE as soon as it is entered, and entering a Response Status clears it at once. Delete Selected Rows works and keeps the sheet protected. |
| Global | 1 | Calculation mode unchanged after the whole suite. |

### 2.3 Reopen, protection and portability (Verify, on copies of the final files)
- Demo batch: all 10 sample requests as in 1.1 (9 Pending Review, 1 No Candidate Found). Flags: REPEAT NAME, CHECK NAME, and OVERDUE on the past-due sample.
- Opened without startup macros: a locked history cell is blocked until the first button macro, then macros write again; no sheet left unprotected.
- Public release moved and renamed: Start first; five visible sheets; 7 hidden, none very hidden; structure protected; **no external links, connections or queries**; **all 34 buttons point to existing public macros**; the release opens empty.
- Typing as a person (opened without the macros' exemption):
  - **allowed:** client input cells; FIU SNO, REFNUMBER, CHECKERDATE and Response Status;
  - **blocked:** client Validation, headers, instructions, zone labels and cells right of the table; FIU Request ID and Screening Status;
  - **blocked:** deleting a sheet row with Excel;
  - **object model only:** filtering was refused and sorting was allowed.

### 2.4 No external file access (static check of `src`)
No routine opens, reads or links an external Excel or CSV file. The only file-related calls are the retained evidence-folder ones: the folder picker (`FileDialog(4)`), `Shell explorer.exe` to open the folder, and the folder write test that deletes its own `.tmp` file.

### 2.5 Sanitisation
- Leak scan of the distribution folder: **HIGH 0**; MEDIUM 14 = the 7 approved hidden supporting sheets in each workbook.
- Compiled VBA (oletools decompression of every stream): 4 "email" HIGH hits inside still-compressed raw streams (compression markers splitting words around an at-sign, the same false positive as 1.0 and 1.1). The decompressed module source has no email address and no local path.
- Private-data cross-check: 137 private values from the original workbook searched in every XML part, document, `src` file and decompressed VBA stream, held in memory and not printed. **0 hits.**

## 3. Visual checks
- Reviewed demo PDF (No Match, "Sample Reviewer"): CUSTOMERNAME ARB, REQUESTTYPE and PUBDATE / DUEDATE read from the new FIU columns; client-list revision line present.
- Screenshots of Start, Client List, FIU Requests, Settings & Help and Search & Review (fictional data). FIU headers are readable at 70% zoom, and Settings & Help shows Quick Start first, then Editable Settings.

## 4. Code review
A review of the 1.1 -> 1.1.1 diff found 6 items:
- **Fixed:** Flag stayed stale until the next button action (it now refreshes when FIU Requests is edited or pasted); the duplicate message said "Reference Number" (now REFNUMBER).
- **Not changed:** errors inside the Flag refresh are silent within the Start refresh; Settings & Help row heights are estimated for merged cells; recognised text dates are converted in place, by design; a build-only diagnostic cell.

The built-in /security-review was **not run**: the release folder is not a git repository.

## 5. Not tested, or known limitations (please check manually)
- **Pasting by a person.** Section 1 used Excel's own Paste Values and Paste commands through automation, not a person's keyboard or mouse; Ctrl+Alt+V (the Paste Special dialog) was not driven. Please paste once yourself (`test-checklist.md`).
- **Sorting and filtering** the protected tables from Excel's header buttons were not tried by hand.
- **Settings & Help paragraph heights** are estimated (merged cells do not auto-fit); check they read cleanly at your zoom and font.
- **Paste wider than the input columns** (Client List past I, FIU past N) is refused by Excel by design; more than 20,000 new rows at once must be pasted in parts.
- Only Request ID and Screening Status are checked for values the tool did not write; values typed onto the instruction cells are not reported.
- Numbers that only display leading zeros arrive as plain numbers with Paste Values.
- Excel 2021 and 32-bit Excel: not available (static evidence only). Mark-of-the-Web / Protected View on another PC: simulated only. Large volumes (10,000+ rows), write-protected folders, PCs without a PDF driver, non-English Excel: not tested.
- Sheet protection has no password: a safety catch, not security.

## 6. Package
- Zip: `FIU_Name_Search_v1.1.1.zip`. Its SHA-256 is published alongside it.
- Contents: the two workbooks, `src/` (7 modules + ThisWorkbook), README, TEST_REPORT, CHANGELOG, SCOPE_NOTE, DISCLAIMER, HOW-TO-UNBLOCK and LICENSE.
