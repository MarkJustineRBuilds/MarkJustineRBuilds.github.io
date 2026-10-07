# FIU Name Search

An offline Excel tool for searching names from FIU requests against **your own** client list (clients and associated parties), recording a human review decision for every possible match and saving a PDF evidence record for every search.

Version 1.1

## What is in this folder

| File | Purpose |
|---|---|
| `FIU_Name_Search_Public_v1.1.xlsm` | The tool. Opens empty and ready for your data. |
| `FIU_Name_Search_Demo_v1.1.xlsm` | The same tool with fictional sample clients and requests, for practice. |
| `src/` | Every line of VBA as plain text, for review by you or your IT team. |
| `SCOPE_NOTE.md` | What a result does and does not mean. Please read. |
| `LICENSE`, `DISCLAIMER.md`, `HOW-TO-UNBLOCK.md`, `CHANGELOG.md`, `TEST_REPORT.md` | Disclaimer, unblocking steps, version history, test results. |

## Requirements

- Windows desktop Excel 2021 or Microsoft 365, 32-bit or 64-bit.
- Macros enabled for this file (see below).
- Not supported: Excel for Mac, Excel for the web, Excel 2019 or older (untested).

The tool sends nothing anywhere. It does not connect to any FIU portal, email system, cloud service or API, and it does not open, import or link to any other Excel or CSV file: you paste your data in. Your data stays in the workbook and in the evidence folder you choose.

## Enabling macros safely

1. Before extracting the zip, right-click it > **Properties** > tick **Unblock** > **OK** (details in `HOW-TO-UNBLOCK.md`).
2. Open the workbook and click **Enable Content** when asked.
3. If your organisation blocks macros from the internet, ask IT. They can review the code in `src/` first. Do not lower your Trust Center settings.

## First use (5 steps)

1. **Paste your client list.** Copy the rows from your own register, then **Paste Special > Values** into the Client List table. Click **Validate Client List**.
2. **Paste FIU requests.** **Paste Special > Values** into the white input columns of FIU Requests. Click **Validate FIU Requests**.
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

- **Search History**: every search, including superseded ones (Start > Search History).
- **Candidates, Runs, Evidence Log, Log**: the audit trail behind reviews and PDFs.
- **System, Evidence Layout**: internal working areas.

Every sheet is protected without a password, as a safety catch against accidental edits (not security). On **Client List** and **FIU Requests** you can edit only the input columns (and, on FIU Requests, the green Response columns), in the table and in the empty rows under it. Headers, the instruction rows, the zone labels and the columns the tool fills are locked.

The tool's own macros still update the locked columns. To remove rows, select a cell in each row and click **Delete Selected Rows** on that sheet (Excel's own row delete is blocked on protected sheets). Excel limits sorting and filtering on protected sheets, so the table header buttons may not sort or filter these two tables. Every request, search, candidate and decision has a stable ID, so the order of rows never breaks the links between them.

## Pasting your data

The tool has no file import. You copy rows from your own workbook and paste them as **values**:

1. In your own register, put the columns in the same order as the tool's table (insert or move columns in a copy if needed).
2. Select and copy the data rows (not the headers).
3. In the tool, right-click the first empty row of the table > **Paste Special > Values** (keyboard: Ctrl+Alt+V, then V, Enter).
4. Click the **Validate** button for that sheet.

Paste inside the table, or directly under its last row with no blank row in between. The table grows to include those rows when you validate. Up to 20,000 empty rows under each table are open for pasting; paste a larger list in parts and validate between them. A paste that is too wide for the input columns is refused by Excel, because it would reach a locked column. Anything else on the sheet outside the table (after a blank row, to the right of the table, or in the instruction rows above it) stops validation, and the message gives the cell addresses.

**Why values only:** a normal paste (Ctrl+V) also brings formulas and formatting. Formula cells are refused before any search or batch, with their addresses. To fix them, select the cells, Copy, then **Home > Paste > Paste Values** over the same cells. Text that only looks like a formula (for example `=ABC`, `+971 ...` or `-0042`) is fine when the source cell holds it as text, and it stays text in every table, review and PDF.

**Leading zeros:** an ID such as `000123` is kept when the source cell holds it as text. If your source stores the number 123 and only *displays* `000123`, Paste Values brings in 123. Format such columns as Text in your source first.

**Validation runs every time.** Validate Client List and Validate FIU Requests show problems straight away, but the same checks run again at the start of every manual search and every batch. A normal paste after validating cannot slip through, and no earlier result is trusted.

## Client List

| Column | Required | Notes |
|---|---|---|
| Record ID | No | Your own client ID, or blank: CL-000001... is generated. Must be unique. Generated IDs are never reused. |
| Full Legal Name | Yes | Kept exactly as pasted. Digits allowed. |
| Party Type | Yes | `Individual` or `Entity` (Natural Person / Juridical Person and similar are accepted). |
| Client/Parent Reference | No | e.g. the parent client of a UBO. |
| Relationship/Role | No | e.g. UBO, Director, Signatory. |
| Date of Birth/Incorporation | No | A real Excel date, or text such as 31-12-1980. |
| Country | No | |
| Government ID | No | Kept as text. |
| Source Label | No | e.g. `Clients`, `UBOs`. |
| Validation | Filled by the tool | OK, OK with a date note, or the problem. |

**Validate Client List** checks every row: required fields, party type, dates (notes only), repeated Record IDs, formulas and data outside the table. Problems are written to the **Validation** column and the message names the sheet rows. A search or batch will not run while any row has a problem; no row is ever dropped silently. People with the same name are kept as separate records.

Each validated list gets a **revision ID** (`CLV-yyyymmdd-hhmmss`). A new revision starts only when the content changes; sorting the list does not. Every run, search and PDF records the revision and the record count it used. Earlier searches, candidates, decisions and PDFs keep the values they had, so changing the list later never alters past results.

### Dates

These are converted to real dates:

- real Excel date cells;
- ISO dates (1980-12-31);
- dates that cannot be misread (31/12/1980).

These are **kept exactly as typed**, noted, and **not used in date comparisons**:

- year-only values (1980);
- plain numbers;
- two-digit years (03/04/45);
- invalid dates (31/02/2020);
- dates where day and month could swap (03/04/2020).

If all the dates you paste are day-first (or month-first), set **Text date order** to `DMY` (or `MDY`) on Settings & Help, and such dates will then be converted. A year typed into a date column, which Excel displays as a 1905 date, is detected and treated as the year text.

## FIU Requests

The table has three zones, labelled in the row above the headers:

- **White input columns** (paste here): Reference Number, Name (English), Name (Arabic), Customer Type, Request Type, FIU Status, Publication Date, Due Date, Date of Birth/Incorporation, Country, Government ID. The usual FIU export columns are REFNUMBER, CUSTOMERNAME ENG, CUSTOMERNAME ARB, CUSTOMERTYPE, REQUESTTYPE, STATUS, PUBDATE and DUEDATE: arrange them in this order before copying.
- **Grey columns** are filled by the tool (Request ID, screening status, results, review, evidence, Processing Message, Flag). Never paste into them.
- **Green Response columns** are yours to complete after review.

Rules:

- **Required:** Reference Number, Name (English), Customer Type.
- Reference numbers and names are kept as text, digits included.
- The same name under different references stays as separate requests.
- A row with the same reference **and** the same name as an earlier searchable row is marked **Error** ("Same Reference Number and name as REQ-...") and is not searched. Delete the copy with **Delete Selected Rows**.

**Validate FIU Requests** assigns Request IDs and writes any problem in a New row to Processing Message (prefixed `Check:`), including dates that will not be used (for example a text due date). It stops with a message, and nothing is searched, if it finds formulas, data outside the table, or a Request ID or Screening Status the tool did not write (usually a paste that was too wide). Values pasted into the other grey columns are not detected, so paste only into the white columns.

## Searching and reviewing

**Run Batch** searches every request whose Screening Status is New, blank or Error.

- Pending Review, No Candidate Found and Reviewed requests are never overwritten.
- To search a request again, select its row and click **Rerun Selected Request**. The earlier search, decision and PDF stay in Search History and in the evidence folder.
- **Pause:** hold **Esc** during a batch. You can pause (continue later with **Resume Batch**), cancel the run (remaining requests stay New), or keep going.
- Both tables are validated first. Formulas, data outside a table, or an empty or invalid client list stop the batch before anything is searched. A request row that cannot be searched (blank required field, invalid type, no letters or digits, repeated reference + name) shows Error with the reason in Processing Message. None of these ever produces "No candidate found".

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
- the client-list snapshot (revision, when it started and was last validated, record count);
- the matching rules and settings;
- the candidate count and every candidate with its match reason;
- the review status, reviewer and note.

Long names and many candidates wrap over several landscape pages, and the table header repeats on each page.

Files are named `EV-000123_<reference>_<name>.pdf` and placed in a dated sub-folder of your evidence folder. Unsafe characters are replaced, long names are shortened, and existing files are never overwritten.

If a PDF fails (for example, the folder is missing or not writable), the search and decision are kept and Evidence Status shows Failed. Use **Retry Failed Evidence**.

## Buttons (Start sheet)

| Button | Macro | What it does |
|---|---|---|
| Validate Client List | ValidateClientListUI | Checks the pasted client list, generates missing Record IDs, sets the revision. |
| Validate FIU Requests | ValidateRequestsUI | Checks the pasted requests and assigns Request IDs. Searches nothing. |
| Run Batch / Resume Batch | RunBatch / ResumeBatch | Searches New and Error requests. |
| Review Pending | ReviewPending | Opens the next request waiting for review. |
| Retry Failed Evidence | RetryEvidence | Produces PDFs that failed earlier. |
| Client List, FIU Requests, Search & Review, Settings & Help, Search History | navigation | |
| Open Evidence Folder | OpenEvidenceFolder | |
| Repair Layout | RepairLayout | Restores buttons, lists, formats, protection and sheet visibility. Data is not changed. |
| Clear Sample Data | ClearSampleData | Removes only `SAMPLE-` rows, after confirmation. |
| Delete Selected Rows (on Client List and FIU Requests) | DeleteSelectedRows | Deletes the table rows of the selected cells, after confirmation. Searches, decisions and PDFs of a deleted request stay in Search History. |

## Settings (Settings & Help sheet)

| Setting | Default | Notes |
|---|---|---|
| Evidence folder | blank | You are asked on first use. Keep the path to 150 characters or fewer. |
| Reviewer name | blank | Blank uses your Office user name. |
| Minimum matching words | 2 | Whole number, 1 to 5. |
| Text date order | Keep ambiguous as text | Or DMY / MDY, if all pasted dates use that order. |
| Organisation name (evidence header) | blank | Printed on PDFs. |
| Show button icons | TRUE | Run Repair Layout after changing. |

## Advanced: maintenance mode

Alt+F8 > **ToggleMaintenanceMode** shows every hidden sheet with a red banner. Run it again, or close the file, to return to the normal view. Protection has no password.

## Updating to a new version

Copy your rows from the old Client List and FIU Requests sheets and Paste Special > Values them into the new file. In 1.1 the FIU Requests input columns come first: copy Reference Number to Government ID as one block, without the old Request ID column. Keep the old file, with its history sheets and PDFs, as your record.

## Licence and disclaimer

Licensed under Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0). See `LICENSE`, `DISCLAIMER.md` and `SCOPE_NOTE.md`.
