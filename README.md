# Payroll Extract

A browser-only tool that turns Punjab Government payroll exports into a
clean, one-row-per-employee spreadsheet — no server, no upload, no backend.
It's a JavaScript port of the `data-extraction-from-payslips` skill, so it
handles the same two source formats:

1. **District-wide payroll register** (`PAY_<month>_<year>.XLS`) — despite
   the `.XLS` extension, this is UTF-16, tab-delimited text: a raw print
   dump with two employees' slips side by side per line.
2. **Individual PIFRA "Monthly Salary Statement"** — a per-employee PDF
   (or a screenshot, transcribed by hand into JSON using the built-in
   template).

Everything — decoding, regex parsing, and building the colored `.xlsx`
output — runs client-side in the visitor's browser using plain JavaScript,
[pdf.js](https://mozilla.github.io/pdf.js/) (for PDF text extraction) and
[ExcelJS](https://github.com/exceljs/exceljs) (for the styled workbook).
No file is ever sent anywhere.

## Hosting on GitHub Pages

1. Push this repo's contents to a GitHub repository.
2. Repo **Settings → Pages → Source**: deploy from the branch containing
   these files (root folder).
3. That's it — no build step, no server, no environment variables. Open
   the published URL and start uploading files.

You can also just open `index.html` directly in a browser for local
testing (double-click it, or `python3 -m http.server` from this folder).

## File layout

```
index.html              Page shell, both extraction modes
css/style.css            Styling
js/bulkParser.js         Parses the bulk two-column payroll register
js/individualParser.js   Parses individual PIFRA slip text (from PDF or JSON)
js/xlsxWriter.js         Builds the colored .xlsx workbook (ExcelJS) + triggers download
js/main.js               UI wiring: drag-drop, pdf.js calls, preview table, download buttons
SKILL_NOTES.md            Field reference and parsing quirks carried over from the original skill
```

## What the output looks like

One row per employee, with pay/allowance codes split into two groups:

- **Green columns** — earnings, taken from the payslip's
  `PAYS AND ALLOWANCES` section.
- **Red columns** — deductions, taken from the `DEDUCTIONS` section.
- Column sets are **fully dynamic** — whatever codes are actually present
  in the uploaded file become columns; nothing is hardcoded.

The bulk-register mode also merges multi-slot "continuation" records (an
employee whose allowance list overflowed one print slot spills into a
second or third `S#:` slot) back into a single row, keyed on
`(Personal No., Office Code)`.

## Known limitations (carried over from the original skill)

- **Office codes are disbursing offices, not postings.** One office code
  can pay staff across 200+ physical schools — don't treat it as a 1:1
  school identifier.
- **"Est. Date of Entry in Service"** (bulk mode only) is calculated by
  subtracting the printed "Length of Service" from the payslip's reference
  date — an approximation, typically accurate to within a few days.
  Individual slips carry the exact "Entry into Govt. Service" date instead
  and should be preferred when available.
- **Bank Account No.** is the raw account number as stored in payroll, not
  a formatted IBAN.
- **Scanned image PDFs** (no text layer) aren't parsed automatically —
  the tool warns you and points to the manual JSON template route for
  screenshots.
- A **large tesseract-OCR fallback** for batches of scanned images was
  noted as unvalidated in the original skill and hasn't been ported here;
  the manual JSON template is the supported path for images.

## Editing / extending

The parsing logic in `js/bulkParser.js` and `js/individualParser.js` is a
line-for-line port of the Python regexes in the original skill — if the
source payroll system changes its print layout, start by comparing against
`SKILL_NOTES.md`, which documents every regex's reasoning and the specific
edge cases (multi-hyphen descriptions, GPF A/C No. sometimes blank, the
`Buckle` field occasionally containing literal text `IBAN`, etc.).
