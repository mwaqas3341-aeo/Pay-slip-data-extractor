---
name: data-extraction-from-payslips
description: Extract structured, one-row-per-employee data from Punjab Government AG/District payroll exports (files often named like PAY_JUNE_2026.XLS, which are actually UTF-16 tab-delimited text dumps of a two-column-per-page payslip printout, NOT real binary Excel files). Use this skill whenever the user uploads a payroll/pay-bill/payslip file from a Pakistani (esp. Punjab) government department and asks to extract, tabulate, list, or analyze employee pay data, personal numbers, designations, bank details, dates of birth, or allowance/deduction breakdowns — even if they don't use the word "extract" (e.g. "pull employee data from this pay file", "list all staff and their salaries", "who's in this payroll"). Also use it to merge such extracted payroll data against an HR/EMIS master list (e.g. AEO/school staff rosters) to build budget proformas, find employees missing from a roster, or verify pay figures.
---

# Data Extraction from Payslips (Punjab Govt payroll .XLS dumps)

## What this file format actually is

Despite the `.XLS` extension, these files are **not binary Excel** — they are
**UTF-16 encoded, tab-delimited plain text**, a raw dump of a print layout
that shows **two employee payslips side by side per printed page** (hence
every line in the file has two "columns" of content separated by a tab).
Always confirm with `file <path>` — it will report `UTF-16 ... text`, not
`Microsoft Excel`.

A single file can contain **all departments in a district** (Police, Health,
Education, Agriculture, Judiciary, Revenue/Land Records, Colleges, etc.), not
just one — check office codes before assuming scope. Office codes seen so far
follow patterns like `LL####` (School Education), `LY####` (a wide mix of
other line departments), `LH####` (a small residual group) — **prefixes are
district-specific**, re-derive them per file rather than hardcoding.

## Earnings vs Deductions (auto-classified, color-coded)

The script splits each employee's pay codes into two groups **by which
section of the payslip they physically came from**, not by numeric prefix
guesswork (prefixes like `6xxx` are used for *both* allowance adjustments
and deduction adjustments, so prefix alone is not reliable):

- Everything between `PAYS AND ALLOWANCES:` and `Gross Pay and Allowances`
  → **earnings**, green header fill.
- Everything between `DEDUCTIONS:` and `Total Deductions` → **deductions**,
  red header fill.

A `Total Deductions` column is appended at the end, summing every red
column for that row.

**Column sets are fully dynamic** — the script does not hardcode a code
list. Every run re-derives the earnings and deduction code sets from
whatever is actually present in that file, so a new allowance/deduction
code that didn't exist in a previous district's or month's file is
automatically picked up and added as its own correctly-colored column, with
no script changes needed. This also means **column count and order can
differ between runs/files** — always re-check the header row rather than
assuming column positions match a previous extract.

## Quick start

```bash
python3 scripts/extract_payslips.py <input.XLS> <output.xlsx> [--office-prefix LL]
```

- Omit `--office-prefix` to extract **every department** in the file.
- Pass e.g. `--office-prefix LL` to restrict to one department/office-code family.
- Run `file <input>` first if unsure of encoding; the script assumes UTF-16.
- For very large files (100MB+ is normal), this can take 1-2 minutes — that's expected.

The script prints row/column counts as it runs so you can sanity-check scale
before opening the output.

## What one employee's raw record looks like

```
                         Layyah	                         Layyah
    S#:1                                      P Sec:001  Month:June 2026	    S#:2 ...
                                              LL6098 -HEAD MISTRESS GIRLS HIGH S	...
    Pers #: 30010209      Buckle: 0                Home	...
    Name:   KISHWAR BANO                      NTN:	...
           ENGLISH TEACHER                    GPF #:	...
    CNIC No.3220104814810                     Old #:	...
    GPF Interest Applied	...
           11  Active Permanent                                LL6098    -	...
   PAYS AND ALLOWANCES:	...
    0001-Basic Pay                                                44,850.00	...
    1000-House Rent Allowance                                      1,853.00	...
    ...
      Gross Pay and Allowances                                   166,200.00	...
   DEDUCTIONS:	...
    ...
                    D.O.B      LFP Quota:      4	...
                    16.11.1966   NATIONAL BANK OF PAKChoubara	...
      30 Years 10 Months 000 Days       3106098838	...
```

Key parsing facts (all handled by the script, documented here so you can
debug or extend it):

1. **Two-column print layout**: split every line on the first `\t`. Left half
   = one employee's stream, right half = the next. No line ever starts with
   a tab (i.e. the left side is never empty) — confirmed empirically.
2. **Continuation slots**: when an employee has too many pay/allowance lines
   to fit one print slot, the payroll system spills into a **second (or
   third) consecutive `S#:` slot**, repeating the full header (name, CNIC,
   office, etc.) and continuing the `PAYS AND ALLOWANCES` list. **You must
   merge these back into one row per employee** — group by `(Personal No.,
   Office Code)` and union their pay-component dicts. About 80% of employees
   use 2 slots, most of the rest use 1, a small tail uses 3.
3. **Pay component lines**: format is `CODE-Description....amount`, e.g.
   `2378-Adhoc Relief All 2023 35%    13,744.00`. The description can contain
   **hyphens** (e.g. `Social Security Ben - 30%`) — the regex must allow `-`
   in the description character class or those lines silently fail to match
   and get dropped. This bit once before — double-check any regex changes
   against `grep -c` counts before/after.
4. **D.O.B / Bank / Service length block** (near the end of each slip):
   - Line A (label only): `D.O.B      LFP Quota:      4`
   - Line B (the actual data): `<DD.MM.YYYY>   <BANK NAME>  <BRANCH/ADDRESS>`
   - Line C: `<NN Years NN Months NNN Days>   <bank account number>`
   - The bank account number is a **raw account number, not a formatted
     IBAN** (no `PK` prefix / check digits). Don't call it an IBAN in output.
   - There is **no stored "date of entry in service" field** — only length
     of service as of the payslip month. An entry date can be *estimated* by
     subtracting that duration from the payslip's reference date, but flag
     it clearly as an approximation, not a source fact.
5. **`Buckle` field anomaly**: on rare records (seen in judiciary dept), this
   field literally contains the text `IBAN` instead of a real buckle number
   — a source-data artifact, not a real value. Don't over-interpret it.
6. **A different, richer single-employee format also exists**: "Monthly
   Salary Statement" slips (e.g. Education Authority Punjab, District
   Accounts Office letterhead) contain several fields the bulk register
   never has: **exact "Entry into Govt. Service" date** (not just a
   duration), full untruncated DDO/office name, BPS, Pay Stage, GPF
   Balance/Account No., and full bank account number. If the user has this
   format instead of (or alongside) the bulk register, prefer its literal
   fields over anything estimated from the bulk file — e.g. always use a
   real "Entry into Govt. Service" date when available rather than the
   `Est. Date of Entry in Service` estimate below (checked against 2 real
   examples: the estimate landed 2-4 days off). This format has not yet had
   a batch parser built for it — if the user provides many of these (e.g. a
   multi-page PDF/export of many individual statements), that's a new
   extraction path to build, not a variant of the bulk-register parser.

## Office codes are cost-centers / disbursing offices, NOT postings

This is the single biggest trap. An office code like `LL6009` can be a
**DDO (disbursing office) that pays staff across 200+ different physical
schools** — it is not a 1:1 stand-in for "this employee's school." Do not
assume office code ⇒ specific school/EMIS/Markaz without checking.

**If you need to map employees to a specific school/EMIS/Markaz**, you need
a separate HR/EMIS master list (the department usually has one). To merge:

1. Match master-list rows to extracted payroll rows by **Personal No.**
   (reliable primary key on both sides).
2. From the matches, build an `OfficeCode -> (EMIS, School, Markaz)` map,
   but only trust it where **exactly one** school combination is seen for
   that office code across all matches. If an office code maps to many
   different schools in your matched data, it's a shared DDO — do NOT
   auto-assign a school to unmatched employees under that code. Flag them
   for manual verification instead (list Name/PersNo/Designation/pay figures
   and let the human confirm).
3. Employees present in payroll but absent from the master list, under an
   **unambiguous** office code, can safely be added with the inferred
   School/EMIS/Markaz. Under an **ambiguous** (multi-school) office code,
   add them but leave School/EMIS blank or use their Personal No. as a
   placeholder in the EMIS column (per user preference — confirm with the
   user which convention they want) and clearly flag for review.
4. Master-list employees with no matching payroll record aren't necessarily
   errors — could be on leave, vacant post, retired mid-month, etc. Keep
   them in output with blank pay columns and a note, don't silently drop.

## Building budget/HR proformas from a template

If the user has a target spreadsheet format (a "proforma") they want the
data arranged into:

1. **Read their exact header row first** (`openpyxl`, check every cell for
   exact text, spacing, and casing — e.g. `' House Rent A01202'` has a
   leading space, `'DESIGNATION '` has a trailing space). Reproduce headers
   **verbatim**, character-for-character, including whitespace. Getting this
   wrong is the #1 thing users will bounce back for correction.
2. Ask before combining or splitting pay codes into template columns when
   the mapping isn't obvious (e.g. one template column might need 2+ raw
   pay codes summed — like "Qualification/M.Phil" = Qualification Allowance
   + Ph.D/M.Phil Allowance). Present your proposed mapping and confirm
   before generating the final file, especially for anything money-related.
3. Only add extra columns beyond the user's template if they explicitly ask
   for "all allowances" or similar — otherwise stick strictly to their
   headers. Users may push back hard if extra invented columns appear
   without being asked, since this is often a real government budget
   document.
4. Add a `TOTAL` row at the bottom using **live SUM formulas** (not
   hardcoded sums) referencing each pay column's range, and recalc with
   `/mnt/skills/public/xlsx/scripts/recalc.py <file> <timeout>` before
   presenting — verify `total_errors: 0`.
5. Put explanatory notes (pay-code mapping assumptions, row-count
   breakdowns, which rows were added/flagged and why) in a **separate
   "Legend & Notes" sheet**, never in the main data sheet's headers or as
   extra unrequested columns.
6. Color-code flagged rows (e.g. light blue = added-with-confidence, light
   red = needs manual verification) so the user can scan visually — but
   don't over-decorate a document they described as strictly needing to
   match their template.

## Full unique pay-code catalog (Layyah district, June 2026 snapshot)

185 distinct pay/allowance/deduction codes exist across all departments in
one such file; ~60 of those apply to School Education staff specifically.
Codes and amounts **will differ by district, department mix, and pay
period** — always re-derive the live catalog from the current file rather
than assuming last time's codes still apply (`scripts/extract_payslips.py`
does this automatically now, per-run, per-earnings/deductions group — see
above).

In the **generic full extraction** (`extract_payslips.py`), every code gets
its own column, including adjustment codes (prefix `5xxx`/`6xxx`, e.g.
`5002-Adjustment House Rent`) — they are kept **separate from their base
code**, not summed together, so the user can see exactly what was posted.
Only when building a **specific budget proforma with a fixed template**
(see next section) should adjustment codes be summed into their
corresponding base-allowance template column — and only after confirming
that mapping with the user, since it's a deliberate simplification for that
one deliverable, not the default extraction behavior.

## Two source formats — bulk register vs individual PIFRA slip

This skill now handles **two distinct payslip formats**, which look
completely different and need different scripts:

### 1. Bulk district-wide register (`scripts/extract_payslips.py`)
The `.XLS` file that's actually UTF-16 tab-delimited text, two employees
printed side-by-side per line, covering an entire district/all departments
in one file. See the sections above for its quirks. Use for large-scale
extraction across many employees at once.

### 2. Individual PIFRA "Monthly Salary Statement" PDF (`scripts/extract_individual_slip.py`)
A **per-employee** PDF (or occasionally a phone screenshot of the same
layout), generated by the Education Authority Punjab / District Accounts
Office system, one employee per file. Usage:

```bash
python3 scripts/extract_individual_slip.py <file_or_folder.pdf> -o <output.xlsx>
```

Accepts either a single PDF or a folder of many (one row per PDF found).
**Always check the PDF has a real text layer first** — try `pdfplumber`
(`pdf.pages[0].extract_text()`); if it returns nothing or garbage, it's a
scanned/flattened image PDF and this regex-based script won't work — fall
back to the image/screenshot approach below.

This format is **richer than the bulk register** — it has fields the bulk
file never has:
- **Exact "Entry into Govt. Service" date** (no estimation needed — always
  prefer this over the bulk-register estimate when both are available for
  the same employee)
- Full untruncated DDO/office name, BPS, Pay Stage, Pay Scale description
- GPF A/C No., GPF Interest status (Free/Applied), GPF Balance
- Full (not truncated) bank name/branch and full account number
- Guardian's name (d/w/s of ... / s/o of ...)
- Income Tax computation detail (Payable, Recovered-to-date, Exempted %, Recoverable)
- City, Domicile, Housing Status, Email

Verified field-for-field against two real examples (different DDOs, one
Permanent/one Contract employee, one with GPF balance/one "GPF Interest
Free", differing allowance sets) — every field matched exactly, including
the earnings/deductions split and totals. Known regex trickiness worth
preserving if you touch this script:
- The `Designation:` line runs straight into the department code with only
  a single space in `pdfplumber`'s extracted text (not the double-space gap
  you'd expect from the visual layout) — split on the `\d{8}-` department
  code pattern, not on whitespace width.
- `GPF A/C No:` is often **blank** followed immediately by `GPF Interest
  Free` or `GPF Interest applied` — don't let a naive "next token" regex
  swallow that status text as if it were the account number.
- Same earnings/deductions section-boundary approach as the bulk parser
  (`Wage type Amount Wage type Amount` marks each table; split on `Deductions
  - General` / `Deductions - Loans and Advances` as section boundaries), and
  same dynamic green/red column-per-code behavior.

### Recognizing which format you've been given
- File literally named `.PDF`/`.pdf` with "Salary Statement" in the visible
  text → individual PIFRA slip → `extract_individual_slip.py`.
- File named `PAY_<month>_<year>.XLS` (or similar), very large (tens/hundreds
  of MB), `file` reports UTF-16 text → bulk register → `extract_payslips.py`.
- A phone screenshot (`.jpg`/`.png`) of a Monthly Salary Statement → same
  individual-slip field set, but read via vision and fed in as JSON (see
  below) since there's no text layer to regex against in an image.

### Screenshots / images of the individual-slip format — now supported

`extract_individual_slip.py` accepts a mix of PDFs, JSON files, and folders
of either, all merged into one colored output workbook:

```bash
python3 scripts/extract_individual_slip.py <pdf_or_json_or_folder> [more...] -o <output.xlsx>
python3 scripts/extract_individual_slip.py --write-template record.json
```

For an image/screenshot (no text layer to regex against), the workflow is:

1. **Read the image directly with vision** (the `view` tool) — this format's
   fields are consistently laid out enough that vision transcription is
   reliable; don't reach for OCR (tesseract) first.
2. Generate an empty JSON record with `--write-template <path>`, or just
   write one matching `templates/manual_record_template.json` in this skill
   — same keys `parse_pdf()` produces: `PersNo`, `Name`, `GuardianName`,
   `CNIC`, `PayPeriod`, `DOB`, `EntryIntoService`, `LengthOfService`,
   `EmploymentCategory`, `Designation`, `DepartmentCode`, `DDOCode`,
   `DDOName`, `BPS`, `PayStage`, `PayScale`, `GPFAccountNo`,
   `GPFInterestStatus`, `GPFBalance`, `BankAccountNo`, `BankDetails`,
   `GrossPay`, `NetPay`, `IncomeTaxPayable`, `IncomeTaxRecoverable`, `City`,
   `Email`, `SourceFile`, plus `Earnings` and `Deductions` dicts keyed by
   4-digit code (`{"0001": {"desc": "Basic Pay", "amt": "46150.00"}}`).
3. Fill in every value visible in the image (leave a field `""`/omit it if
   genuinely not shown — don't guess).
4. Pass the JSON file(s) to the script alongside/instead of PDFs — it
   auto-detects `.json` vs `.pdf` and merges everything into one workbook
   with the same green/red column coloring and dynamic code-column
   discovery as PDF-parsed records.
5. **Sanity-check before trusting the output**: Gross Pay should equal the
   sum of all Earnings codes; Net Pay should equal Gross Pay minus the sum
   of Deductions codes (deductions are stored as negative amounts). If
   these don't reconcile, re-check the transcription against the image.

Validated end-to-end on the two real screenshots from this session (Zaib Ul
Nissa, June 2026 and Muhammad Waqas, May 2026) — both transcribed via
vision into JSON, run through the script together with the PDF version of
one of them, and produced identical field values across both extraction
paths (confirming the JSON/vision route is just as reliable as the native
PDF text-layer route for this format).

For a genuinely large batch of scanned images where per-image vision
transcription would be impractical, `tesseract` is available in the
sandbox as a fallback — but it hasn't been validated on this table-heavy
layout, so spot-check its output against a few images before trusting it
at scale, and prefer vision transcription whenever the volume is
manageable.

## Common follow-up requests and how to handle them

- **"One row per employee, show me a sample"** — pull one full record
  (identity fields + every non-zero pay component) and present as a table;
  don't dump raw regex output.
- **"Is this data department-wise?"** — be honest: the file has office
  codes/names, not an explicit "Department" field. Offer to derive a
  Department column via office-name pattern classification if wanted, but
  flag that this involves inference, not a literal source field.
- **"Add columns for X field you mentioned but didn't include"** — check
  whether X is a literal source field (extract it properly, verify with
  grep against raw text first) or something that must be derived/estimated
  (compute it, but label the column and document the caveat, e.g. "Est. ...*"
  with a footnote).
- **Large output files** (150+ columns, 10,000+ rows) can be slow to open on
  mobile — mention this if relevant, and consider offering a trimmed/split
  version.
