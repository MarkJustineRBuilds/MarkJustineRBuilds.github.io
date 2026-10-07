# FIU Name Search

An offline Excel tool for searching names from FIU requests against **your own** client list (clients and associated parties), recording a human review decision for every possible match and saving a PDF evidence record for every search.

Version 1.0

## What is in this folder

| File | Purpose |
|---|---|
| `FIU_Name_Search_Public_v1.0.xlsm` | The tool. Opens empty and ready for your data. |
| `FIU_Name_Search_Demo_v1.0.xlsm` | The same tool with fictional sample clients and requests, for practice. |
| `src/` | Every line of VBA as plain text, for review by you or your IT team. |
| `SCOPE_NOTE.md` | What a result does and does not mean. Please read. |
| `LICENSE`, `DISCLAIMER.md`, `HOW-TO-UNBLOCK.md`, `CHANGELOG.md`, `TEST_REPORT.md` | Disclaimer, unblocking steps, version history, test results. |

## Requirements

- Windows desktop Excel 2021 or Microsoft 365, 32-bit or 64-bit.
- Macros enabled for this file (see below).
- Not supported: Excel for Mac, Excel for the web, Excel 2019 or older (untested).

The tool sends nothing anywhere. It does not connect to any FIU portal, email system, cloud service or API. Your data stays in the workbook and in the evidence folder you choose.

## Enabling macros safely

1. Before extracting the zip, right-click it > **Properties** > tick **Unblock** > **OK** (details in `HOW-TO-UNBLOCK.md`).
2. Open the workbook and click **Enable Content** when asked.
3. If your organisation blocks macros from the internet, ask IT. They can review the code in `src/` first. Do not lower your Trust Center settings.

## First use (5 steps)

1. **Load your client list.** Start > **Import Client List**, or paste values into the Client List sheet.
2. **Load FIU requests.** Start > **Import FIU Requests**, or paste into the FIU Requests sheet.
3. **Run Batch.** Choose an evidence folder when asked. Every request is searched once and gets a PDF.
4. **Review candidates.** Start > **Review Pending**. Mark matching candidates with **Select = Yes**, choose a **Decision**, write a **Review note**, then click **Save Decision**.
5. **Open evidence.** Start > **Open Evidence Folder**. PDFs are filed in a dated sub-folder.

Then record your response to the FIU in the green **Response Status / Response Date** columns. Producing evidence does not mean a response was submitted.

Try the demo workbook first: open it, click **Run Batch** and review the sample results. In your own copy, **Clear Sample Data** removes rows whose IDs start with `SAMPLE-`.

## The sheets

Five sheets are visible for daily use:

| Sheet | Use |
|---|---|
| **Start** | Setup status, request totals, review queue and all main buttons. |
| **Client List** | Your clients and associated parties (one row per person or entity). |
| **FIU Requests** | Requests to search, screening progress and your response tracking. |
| **Search & Review** | Manual search, plus the candidate review panel. |
| **Settings & Help** | Settings, entity suffix list, response list and operating notes. |

Supporting sheets are hidden. They are kept as history, not removed:

- **Import**: header mapping while an import is in progress.
- **Search History**: every search, including superseded ones (Start > Search History).
- **Candidates, Runs, Evidence Log, Import Log, Log**: the audit trail behind reviews and PDFs.
- **System, Staging, Evidence Layout**: internal working areas.

Sheets are protected without a password, as a safety catch against accidental edits. Client List and FIU Requests are not protected, so you can paste, sort and filter freely. Every request, search, candidate and decision has a stable ID, so sorting never breaks the links between them.

## Loading the client list

### Option A: paste

Paste **values** into the Client List table:

| Standard field | Required | Notes |
|---|---|---|
| Record ID | No | Generated (CL-000001...) when blank. Must be unique. |
| Full Legal Name | Yes | Kept exactly as entered. Digits allowed. |
| Party Type | Yes | `Individual` or `Entity` (Natural Person / Juridical Person are accepted). |
| Client/Parent Reference | No | e.g. the parent client of a UBO. |
| Relationship/Role | No | e.g. UBO, Director, Signatory. |
| Date of Birth/Incorporation | No | A real Excel date, or text such as 31-12-1980. |
| Country | No | |
| Government ID | No | Stored as text, so leading zeros are kept. |
| Source Label | No | e.g. `Clients`, `UBOs`. |

Click **Validate Client List** to check the list. Problems are written to the **Validation** column. Filter that column, then fix or delete the listed rows. A batch will not run while any row is invalid; no row is ever dropped silently.

### Option B: import with header mapping

1. **Import Client List** > choose an Excel (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) or CSV file > choose the sheet or table > confirm the header row.
2. On the Import sheet, check the **Source Column** for each standard field. Common header names are matched automatically.
3. Set the **Source label**, the **Import mode** and, if your file has no type column, a **Default type**.
4. Click **Preview & Validate**. Every row is listed as OK, Warning, Error (will not be imported) or Skip (blank, or already imported).
5. Click **Commit Import**, or **Cancel Import**. Cancelling changes nothing.

Import modes for the client list:

- **Append**: adds the rows.
- **Replace rows with this source label**: for example, reload only the UBOs.
- **Replace entire client list**.

Import clients and UBOs separately, each with its own label. People with the same name are always kept as separate records.

How the source file is treated:

- It opens read-only, links are not updated, its macros are disabled, and it is closed without saving.
- Only values are copied. The result is a snapshot, not a live link.
- CSV files are read as UTF-8 text, so leading zeros and Arabic text are kept.

**If the file is already open in Excel**, the tool reads what is currently in Excel, including unsaved edits. It never saves, closes or switches to that workbook. The Import sheet and the Import Log record that the open copy was read, and when.

The time of each import, its source label and its record count appear on Start and in the Import Log. Every search records the client-list version it used.

### Dates

These are converted to real dates:

- real Excel date cells;
- ISO dates (1980-12-31);
- dates that cannot be misread (31/12/1980).

These are **kept exactly as typed**, flagged as a Warning, and **not used in date comparisons**:

- year-only values (1980);
- plain numbers;
- two-digit years (03/04/45);
- invalid dates (31/02/2020);
- dates where day and month could swap (03/04/2020).

If all your files use day-first (or month-first) dates, set **Text date order** to `DMY` (or `MDY`) on Settings & Help, and such dates will then be converted. A year typed into a date column, which Excel displays as a 1905 date, is detected and treated as the year text.

## Loading FIU requests

Paste into FIU Requests, or use **Import FIU Requests** (always Append). The original FIU column names map automatically: REFNUMBER, CUSTOMERNAME ENG, CUSTOMERNAME ARB, CUSTOMERTYPE, REQUESTTYPE, STATUS, PUBDATE and DUEDATE.

- **Required:** Reference Number, Name (English), Customer Type.
- **Optional:** Name (Arabic), Request Type, FIU Status, Publication Date, Due Date, and a date of birth, country or ID for comparison.
- Reference numbers and names are kept as text, digits included.
- The same name under different references stays as separate requests.
- A row with the same reference **and** the same name as an existing request is skipped as already imported.

## Searching and reviewing

**Run Batch** searches every request whose Screening Status is New, blank or Error.

- Pending Review, No Candidate Found and Reviewed requests are never overwritten.
- To search a request again, select its row and click **Rerun Selected Request**. The earlier search, decision and PDF stay in Search History and in the evidence folder.
- **Pause:** hold **Esc** during a batch. You can pause (continue later with **Resume Batch**), cancel the run (remaining requests stay New), or keep going.
- The client list is checked first. An empty or invalid list, or a name with no letters or digits, never produces "No candidate found". The request stays New, or shows Error with the reason in Processing Message.

| Screening Status | Meaning |
|---|---|
| New | Waiting to be searched. |
| Pending Review | Candidates were returned; a person must decide. A PDF marked PENDING REVIEW is already saved. |
| No Candidate Found | The engine found no candidate. No human decision is recorded or invented. A PDF is saved. |
| Reviewed | A decision was saved and an updated PDF was produced. |
| Error | **Not searched.** The reason is in Processing Message. It is retried by the next batch. |

Review decisions:

- **Name Match** and **Partial Name Match**: mark at least one candidate as Select = Yes.
- **No Match**: needs a written reason of at least 15 characters.
- **Pending Review**: saves your note and leaves the request pending.
- Before saving, a confirmation appears. Answering No, or clicking **Cancel / Leave Pending**, leaves the review pending.

Response Status and Response Date are entirely separate from screening, and you fill them in yourself.

**Manual search:** on Search & Review, type a name and customer type (plus an optional date, country and ID), then click **Search Name**. Click **Save Evidence PDF** to keep a PDF of the result.

## Matching rules (MR-1.0)

- Names are upper-cased and Latin accents removed (É → E).
- Apostrophes and full stops are removed (O'NEIL = ONEIL, L.L.C = LLC). Hyphens and other punctuation become spaces.
- Repeated words count once. Digits are kept and matched like letters. Long names are used in full, with no word limit.
- **Entity names only:** suffix words at the end of the name, from the editable list on Settings & Help (LLC, FZE, LTD...), are ignored. GENERAL, TRADING, GROUP and HOLDING are kept. A name made only of suffix words is kept as is, and flagged.
- A client record becomes a candidate when it shares at least **N whole words** with the search name. N is the smallest of: **Minimum matching words** (default 2), the number of words in the search name, and the number of words in the record.
- Each candidate shows its match type and the reason:
  - Exact (normalised);
  - Same words, different order;
  - All search words in record;
  - All record words in search;
  - Partial words.
- **Not supported:** substring, fuzzy, phonetic or transliteration matching. MOHAMMED does not match MUHAMMAD, and KHAN does not match KHANDWALA.
- Arabic text matches only identical Arabic words (diacritics and tatweel are ignored). English and Arabic names are not cross-matched.
- Party type, date, country and government ID are compared and shown in Comparison Notes. **They never exclude a candidate.**

## Evidence PDFs

A PDF is produced for every successfully searched request: zero-candidate searches, pending reviews (clearly labelled PENDING REVIEW) and each saved decision. Each PDF contains:

- the request reference, the name as received, the customer type and the search time;
- the client-list snapshot (version, load time, record count);
- the matching rules and settings;
- the candidate count and every candidate with its match reason;
- the review status, reviewer and note.

Long names and many candidates wrap over several landscape pages, and the table header repeats on each page.

Files are named `EV-000123_<reference>_<name>.pdf` and placed in a dated sub-folder of your evidence folder. Unsafe characters are replaced, long names are shortened, and existing files are never overwritten.

If a PDF fails (for example, the folder is missing or not writable), the search and decision are kept and Evidence Status shows Failed. Use **Retry Failed Evidence**.

## Buttons (Start sheet)

| Button | Macro | What it does |
|---|---|---|
| Import Client List / Import FIU Requests | ImportClientList / ImportRequests | Import with header mapping. |
| Validate Client List | ValidateClientListUI | Checks every client row. |
| Run Batch / Resume Batch | RunBatch / ResumeBatch | Searches New and Error requests. |
| Review Pending | ReviewPending | Opens the next request waiting for review. |
| Retry Failed Evidence | RetryEvidence | Produces PDFs that failed earlier. |
| Client List, FIU Requests, Search & Review, Settings & Help, Search History | navigation | |
| Open Evidence Folder | OpenEvidenceFolder | |
| Repair Layout | RepairLayout | Restores buttons, lists, formats, protection and sheet visibility. Data is not changed. |
| Clear Sample Data | ClearSampleData | Removes only `SAMPLE-` rows, after confirmation. |

## Settings (Settings & Help sheet)

| Setting | Default | Notes |
|---|---|---|
| Evidence folder | blank | You are asked on first use. Keep the path to 150 characters or fewer. |
| Reviewer name | blank | Blank uses your Office user name. |
| Minimum matching words | 2 | Whole number, 1 to 5. |
| Text date order | Keep ambiguous as text | Or DMY / MDY. |
| Organisation name (evidence header) | blank | Printed on PDFs. |
| Show button icons | TRUE | Run Repair Layout after changing. |

## Advanced: maintenance mode

Alt+F8 > **ToggleMaintenanceMode** shows every hidden sheet with a red banner. Run it again, or close the file, to return to the normal view. Protection has no password.

## Updating to a new version

Copy your rows (paste values) from the old Client List and FIU Requests sheets into the new file. Keep the old file, with its history sheets and PDFs, as your record.

## Licence and disclaimer

Licensed under Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0). See `LICENSE`, `DISCLAIMER.md` and `SCOPE_NOTE.md`.
