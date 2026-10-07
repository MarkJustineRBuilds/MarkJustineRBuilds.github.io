# Test report: FIU Name Search 1.0

Date: 07-10-2026. Environment: Windows 11, Microsoft 365 Excel 64-bit (16.0, build 20430), desktop.

All results below were produced on copies of the **final** release files. They were re-run after the last code change and after the final package clean-up. Earlier passing runs were not counted.

| File | SHA-256 |
|---|---|
| FIU_Name_Search_Public_v1.0.xlsm | 28323DC80BD2F29593058CDA9892D5670D5953836780D87BE601554208C57207 |
| FIU_Name_Search_Demo_v1.0.xlsm | 6929717A41B5FC10FDFD495F61DF859C4E6BFDFAB77E3632B7261CB1C3CDCB00 |

## 1. Automated checks (Excel automation, real workbook code)

### 1.1 Build and compile
- Both workbooks were built from the `src` modules, round-tripped through `.xlsx` to drop any earlier compiled code, re-imported, and compiled in Excel. **Compile: OK** for both. They contain only the 8 modules in `src` plus the sheet and workbook modules.

### 1.2 Runtime suite: 212 checks, run on a copy of each file

| File tested | Result |
|---|---|
| Copy of the Public release | **PASS 212 / FAIL 0** |
| Copy of the Demo release | **PASS 212 / FAIL 0** |

Prompts were answered by a scripted test harness. Any prompt the harness did not expect fails the test instead of waiting.

| Area | Checks | Covered |
|---|---|---|
| Normalisation | 14 | Case, spaces, punctuation, hyphens, apostrophes, accents, repeated words, 9-word names, digits, Arabic diacritics. Entity suffixes removed only at the end, only for entities. GENERAL/TRADING kept. A name made only of suffixes is kept and flagged. A punctuation-only name gives zero words. |
| Matching | 17 | Exact; reordered; record words in search; one-word search; whole words only (KHAN ≠ KHANDWALA); digits; long names; identical names belonging to different records; optional fields present, absent and conflicting (noted, never excluded); type difference; minimum-words setting. |
| Client list | 8 | Empty list and invalid rows block the batch. Requests stay New and are never marked "No candidate found". Missing Record IDs are generated. |
| Batch | 19 | Pending vs No Candidate Found; no invented review for zero candidates; Error rows not searched; leading-zero references; repeat names under different references; Excel settings restored; completed work not overwritten. |
| Review | 16 | Blank decision, unexplained dismissal and Match without a selection are all rejected. Cancelling the confirmation leaves the review pending. Selected candidate IDs, reviewer and time are saved. A new PDF is made and the earlier one kept. A second decision is blocked. |
| Stable IDs | 2 | After sorting requests and history, requests still link to their searches, and Review Pending opens the right request. |
| Rerun | 8 | Rerun via typed ID and via a selected row. The earlier search, decision and PDF are preserved. A superseded search cannot be decided. |
| Pause / resume / cancel | 9 | Pause, then resume the same run with correct counts. Cancel leaves the remaining requests New. |
| Error handling | 10 | An error injected mid-batch: run marked Failed, the unprocessed request is not marked searched, and ScreenUpdating, EnableEvents, Calculation, Cursor, StatusBar and EnableCancelKey are all restored. |
| PDF failure / folders | 8 | A PDF failure keeps the search with evidence marked Failed, and a retry succeeds. A missing evidence folder means the batch is not started; choosing a new folder lets it run. |
| Paths | 4 | Unsafe characters replaced, no overwrite on a name collision, long names trimmed, an over-long folder refused. |
| Multipage | 2 | 120 candidates with a ~300-character name. |
| Manual search | 3 | Search recorded, PDF saved, blank name refused. |
| Import, clients (CSV) | 16 | Auto-mapping; preview counts; error rows not imported; same-name people kept; leading zeros; Arabic text; import log; client-list version. |
| Import, cancel / mapping | 6 | Cancel at file choice and after mapping change nothing. Unmapped required field, duplicate mapping, commit without preview and mapping changed after preview are all refused. |
| Import, xlsm / replace | 10 | The source's Workbook_Open macro did not run. The cached external-link value is imported as a value only. Source file unchanged. No links in the tool. Source closed and active workbook restored. Append, replace by label, replace all. |
| Import, requests | 9 | Original FIU headers auto-mapped. Duplicate reference + name skipped. Leading zeros, digits-in-name warning, real dates. Repeat import adds nothing. |
| Recovery | 4 | An interrupted run becomes Interrupted and a half-searched request becomes Error. A manual edit creates a new client-list version. |
| Sample data / repair | 4 | Clear Sample Data removes only SAMPLE rows; IDs are not reused afterwards; Repair Layout keeps data. |
| **Dates (review fix)** | 26 | Real Excel date cell → date. Numeric cell 1980, text "1980" and a year typed into a date column → kept as "1980" with a warning. "03/04/45" (two-digit year), "31/02/2020" (invalid) and "03/04/2020" (ambiguous, default setting) → kept as typed with a warning. Blank → blank. "31/12/1980" and ISO → dates. Ambiguous dates convert only when the user sets DMY, with a note. Unresolved values are excluded from date comparisons, and full dates are still compared. |
| **Open workbook import (review fix)** | 12 | A source open in the same Excel with unsaved edits: cancel at sheet choice, successful import, and cancel after mapping. The source stays open and unsaved, the edit is kept, and the disk file is never written. The in-memory read and its time are logged. The active workbook is unchanged. |
| **Protection fallback (review fix)** | 4 | Batch completes, sheets stay protected, Excel settings unchanged, and a pending VBA error is preserved across the protection routine. |
| Global | 1 | Calculation mode unchanged after the whole suite. |

Test source files (CSV, xlsm, xlsx) were hashed before and after the suite and were **unchanged**. The macro marker file written by the test source's Workbook_Open **was never created**.

### 1.3 Reopen, protection and portability (Verify run on copies of the final files)
- **Demo opened with startup macros skipped** (events off, so Workbook_Open did not run):
  - a macro write to a locked history cell was blocked (error 1004), as expected;
  - after the first button macro (Run Batch), writes worked;
  - Client List and FIU Requests are the only unprotected sheets, by design;
  - Excel state afterwards: ScreenUpdating on, EnableEvents on, Calculation automatic, StatusBar cleared.
- **Demo batch results:** all 10 sample requests behaved as designed. 9 Pending Review and 1 No Candidate Found; identical names returned as separate records; LLC/FZE variants matched; long name matched; digits and one-word warnings shown; date and country comparisons shown.
- **Normal reopen of the saved demo copy:** Start is active; five tabs are visible (Start, Client List, FIU Requests, Search & Review, Settings & Help); macro writes allowed; Start counts and review queue correct.
- **Public release copied to a different folder under a different file name:**
  - Start opens first, with the same five visible tabs;
  - 10 supporting sheets hidden, none very hidden;
  - workbook structure protected;
  - no external links, connections or queries;
  - **all 37 buttons point to existing public macros and none is tied to a file name**;
  - navigation and repair macros ran through the renamed file;
  - the release opens empty (0 clients, 0 requests).

### 1.4 Sanitisation
- **Release-skill leak scan, distribution folder:** HIGH 0 after the fix. The scan first found the build folder path saved in `workbook.xml` (`x15ac:absPath`); this was removed and the workbooks re-tested.
  - MEDIUM 20, accepted: the hidden Setup and Backend sheets, which are part of the approved design. In the Public file their tables are empty, apart from the Flag formula placeholder.
  - INFO: compiled VBA, which is scanned separately below.
- **Compiled VBA** (oletools decompression of every stream in `vbaProject.bin`, both files): no user paths, user names, network paths, SharePoint/OneDrive links or organisation terms.
  - 10 "email" pattern hits are false positives: compression markers inside still-compressed module streams happen to split words such as `Scripting.Dictionary` around an at-sign. The decompressed source has none.
  - No UserForms, so no MSForms temporary path.
- **Private-data cross-check:** 137 private values (names, Arabic names, reference numbers, notes, reviewer names, paths and organisation terms) were searched for, case-insensitively, in every XML part, document, `src` file and decompressed VBA stream of the distribution. They were held in memory and not printed. **0 hits.**
- **Document properties:** author, last-modified-by, company and manager are blank. No thumbnail, printer settings, comments, customXml or external-link parts.
- **Zip contents and Zone.Identifier simulation:** see section 4.

## 2. Visual checks (performed by reading the generated PDFs)
- **Pending-review PDF** (demo SAMPLE-REF-0001): amber PENDING REVIEW banner; Arabic name shown left-aligned; dates DD-MM-YYYY; candidate table with match reasons and "Date matches / Date differs"; footer shows its own Evidence ID and reference.
- **No-candidate PDF:** green "ENGINE RESULT: NO CANDIDATE FOUND – no human review decision is recorded"; table states that no candidates were returned.
- **Reviewed PDF:** "REVIEWED – DECISION: NAME MATCH"; selected candidate ID, reviewer, time and note; Selected = Yes on the chosen row.
- **Multipage PDF:** 10 landscape pages, header row repeated, all 120 candidates, the full ~300-character name wrapped, "Page n of 10" footers.
- **Defects found and fixed during these checks:**
  - the second and later PDFs reused the first PDF's footer (a `PrintCommunication` issue);
  - Arabic text was right-aligned.
  Both were re-checked on the final build.

## 3. Not tested, or tested only indirectly (please check manually)
- **Real clicking of buttons and dialogs.** Every button's macro was run, but MsgBox, InputBox, file picker and folder picker were answered by the harness. Please click through the flow once yourself.
- **Esc to pause.** The pause/cancel/resume logic is tested through a test hook; physically pressing Esc during a batch was not tested.
- **Excel 2021 (perpetual licence) and 32-bit Excel.** These were not available. The code uses `PtrSafe` with `#If VBA7`, late binding only and no dynamic-array or Microsoft 365-only functions, but this is static evidence, not a runtime test.
- **Mark-of-the-Web / Protected View** on another PC: the Zone.Identifier was simulated only (section 4).
- **Large volumes:** tested with 130 client records and up to 120 candidates per search. Speed with 10,000+ clients is not measured. PDF export takes about 1 second per request on this PC.
- **A folder that exists but is write-protected** (Windows permissions): only missing folders were tested. The writability check uses a real test write.
- **PCs without a default printer/PDF driver**, and non-English Excel UI or regional settings (dates are parsed by the tool, but this was not run under other locales).

## 4. Package
- Zip: `FIU_Name_Search_v1.0.zip`. Its SHA-256 is published alongside it (the zip cannot contain its own checksum).
- Contents: the two workbooks, `src/` (8 modules + ThisWorkbook), README, TEST_REPORT, CHANGELOG, SCOPE_NOTE, DISCLAIMER and HOW-TO-UNBLOCK. No private build tools, test files, logs or config.
- **Leak scan of the zip:** HIGH 0. MEDIUM 20 are the accepted hidden sheets described in 1.4.
- **Download simulation:** a copy of the zip was marked `Zone.Identifier ZoneId=3`. After `Unblock-File` (the Properties > Unblock equivalent) the mark was gone. All 17 files of the tested package extracted without a mark, and both workbooks were byte-identical to the tested files. Excel's own Protected View / Enable Content prompt on a downloaded file was not exercised (see section 3).

- **After testing:** `LICENSE` was added, and README, CHANGELOG and this report were edited (text only). The workbooks and `src` files are byte-identical to the tested files. The zip SHA-256 published on the download page is for this final package.
