/**
 * xlsxWriter.js
 * -----------------------------------------------------------------------
 * Builds the colored .xlsx workbook (green = earnings, red = deductions)
 * for both the bulk register and individual-slip extractors, mirroring
 * write_workbook() in the two Python scripts. Requires ExcelJS
 * (loaded globally as `ExcelJS` via CDN in index.html).
 * -----------------------------------------------------------------------
 */

const XlsxWriter = (() => {

  const EARN_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
  const DED_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
  const EARN_FONT = { bold: true, size: 10, color: { argb: 'FF006100' } };
  const DED_FONT = { bold: true, size: 10, color: { argb: 'FF9C0006' } };
  const FIXED_FONT = { bold: true, size: 10 };

  function styleHeaderRow(row, nFixed, nEarn, nDed) {
    row.eachCell((cell, colNumber) => {
      const i = colNumber - 1;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      if (i < nFixed) {
        cell.font = FIXED_FONT;
      } else if (i < nFixed + nEarn) {
        cell.font = EARN_FONT;
        cell.fill = EARN_FILL;
      } else {
        cell.font = DED_FONT;
        cell.fill = DED_FILL;
      }
    });
  }

  /** Builds the workbook for the BULK register extractor.
   *  merged: array of records as produced by BulkParser.mergeRecords
   *  refDate: JS Date used for the "Est. Date of Entry in Service" estimate
   *  returns: { workbook, nRows, nCols } */
  async function buildBulkWorkbook(merged, refDate) {
    const earnDesc = {};
    const dedDesc = {};
    for (const m of merged) {
      for (const [code, info] of Object.entries(m.PayComponents)) {
        if (!(code in earnDesc)) earnDesc[code] = info.desc;
      }
      for (const [code, info] of Object.entries(m.DeductionComponents || {})) {
        if (!(code in dedDesc)) dedDesc[code] = info.desc;
      }
    }
    const earnCodes = Object.keys(earnDesc).sort();
    const dedCodes = Object.keys(dedDesc).sort();

    const FIXED_HEADERS = ["S.No", "Personal No.", "Name", "CNIC No.", "Buckle No.",
      "Office Code", "Office Name (as on payslip)", "Designation", "Gross Pay",
      "Date of Birth", "Bank Name / Branch", "Bank Account No.",
      "Length of Service", "Est. Date of Entry in Service*"];
    const EARN_HEADERS = earnCodes.map(c => `${c}-${earnDesc[c]}`);
    const DED_HEADERS = dedCodes.map(c => `${c}-${dedDesc[c]}`);
    const ALL_HEADERS = [...FIXED_HEADERS, ...EARN_HEADERS, ...DED_HEADERS, "Total Deductions"];
    const nCols = ALL_HEADERS.length;

    const sorted = [...merged].sort((a, b) => {
      const oc = (a.OfficeCode || '').localeCompare(b.OfficeCode || '');
      if (oc !== 0) return oc;
      return (a.Name || '').localeCompare(b.Name || '');
    });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Payroll Extract');

    ws.addRow(["Extracted Employee Pay Data (Green = Earnings, Red = Deductions)"]);
    ws.mergeCells(1, 1, 1, nCols);

    const headerRow = ws.addRow(ALL_HEADERS);
    styleHeaderRow(headerRow, FIXED_HEADERS.length, EARN_HEADERS.length, DED_HEADERS.length);

    sorted.forEach((m, i) => {
      const row = [i + 1, m.PersNo, m.Name, m.CNIC, m.Buckle,
        m.OfficeCode, m.OfficeNameTrunc, m.Designation,
        m.GrossPay ? parseFloat(m.GrossPay) : 0,
        m.DOB, m.BankNameBranch, m.BankAccountNo,
        m.LengthOfService,
        BulkParser.estJoiningDate(m.LengthOfService, refDate)];
      let dedTotal = 0;
      for (const code of earnCodes) {
        const info = m.PayComponents[code];
        row.push(info ? parseFloat(info.amt) : 0);
      }
      for (const code of dedCodes) {
        const info = (m.DeductionComponents || {})[code];
        const amt = info ? parseFloat(info.amt) : 0;
        dedTotal += amt;
        row.push(amt);
      }
      row.push(Math.round(dedTotal * 100) / 100);
      ws.addRow(row);
    });

    ws.columns.forEach(col => { col.width = 16; });

    const notes = wb.addWorksheet('Notes');
    const noteLines = [
      "Notes on this extract",
      "",
      "Columns are split into two groups, colored for quick scanning:",
      "  GREEN headers = Earnings (pay / allowances) - from the payslip's",
      "                  'PAYS AND ALLOWANCES' section",
      "  RED headers   = Deductions - from the payslip's 'DEDUCTIONS' section",
      "  'Total Deductions' (last column) sums every red column for that row.",
      "",
      "Column sets are built automatically from whatever codes are found in the source",
      "file each run - if a new allowance or deduction code appears, it is automatically",
      "added as its own new column, correctly colored green or red. No code list is hardcoded.",
      "",
      "Office Name is truncated to ~26 characters by the original print layout — it is NOT the",
      "full school/office name. EMIS code, Markaz, and BPS are NOT present in payslip data at all;",
      "they must come from a separate HR/EMIS master list if you need them.",
      "",
      "* Est. Date of Entry in Service is NOT a stored field in this bulk-register format",
      "  (it IS a stored, exact field in individual 'Monthly Salary Statement' slips — use",
      "  that directly, it's exact, this estimate is not).",
      "  Calculated by subtracting the printed 'Length of Service' from the payslip's reference",
      "  date. Cross-checked against real slips: off by 2-4 days. Treat as approximate.",
      "",
      "Bank Account No. is the raw account number as stored in payroll — NOT a formatted IBAN",
      "(no PK country code / check digits).",
      "",
      "Two consecutive S# print slots often belong to the SAME employee (their allowance list",
      "was too long for one slot) — this tool automatically merges those into a single row.",
    ];
    noteLines.forEach(l => notes.addRow([l]));
    notes.getColumn(1).width = 100;

    return { workbook: wb, nRows: sorted.length, nCols };
  }

  /** Builds the workbook for the INDIVIDUAL PIFRA slip extractor.
   *  records: array of records as produced by IndividualParser.parseText / normalizeJsonRecords
   *  returns: { workbook, nRows, nCols } */
  async function buildIndividualWorkbook(records) {
    const earnDesc = {}, dedDesc = {};
    for (const r of records) {
      for (const [c, info] of Object.entries(r.Earnings || {})) {
        if (!(c in earnDesc)) earnDesc[c] = info.desc;
      }
      for (const [c, info] of Object.entries(r.Deductions || {})) {
        if (!(c in dedDesc)) dedDesc[c] = info.desc;
      }
    }
    const earnCodes = Object.keys(earnDesc).sort();
    const dedCodes = Object.keys(dedDesc).sort();

    const FIXED = ["S.No", "Personal No.", "Name", "Guardian Name", "CNIC", "Pay Period",
      "Date of Birth", "Entry into Govt. Service", "Length of Service",
      "Employment Category", "Designation", "Department", "DDO Code", "DDO Name",
      "BPS", "Pay Stage", "Pay Scale", "GPF A/C No", "GPF Interest Status", "GPF Balance",
      "Bank Account No.", "Bank Details", "Gross Pay", "Net Pay",
      "Income Tax Payable", "Income Tax Recoverable", "City", "Email", "Source File"];
    const EARN_H = earnCodes.map(c => `${c}-${earnDesc[c]}`);
    const DED_H = dedCodes.map(c => `${c}-${dedDesc[c]}`);
    const HEADERS = [...FIXED, ...EARN_H, ...DED_H, "Total Deductions (calc)"];
    const nCols = HEADERS.length;

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Individual Payslips');

    ws.addRow(["Individual PIFRA Salary Statements (Green=Earnings, Red=Deductions)"]);
    ws.mergeCells(1, 1, 1, nCols);

    const headerRow = ws.addRow(HEADERS);
    styleHeaderRow(headerRow, FIXED.length, EARN_H.length, DED_H.length);

    records.forEach((r, i) => {
      const row = [i + 1, r.PersNo, r.Name, r.GuardianName, r.CNIC,
        r.PayPeriod, r.DOB, r.EntryIntoService, r.LengthOfService,
        r.EmploymentCategory, r.Designation, r.DepartmentCode,
        r.DDOCode, r.DDOName, r.BPS, r.PayStage,
        r.PayScale, r.GPFAccountNo, r.GPFInterestStatus, r.GPFBalance,
        r.BankAccountNo, r.BankDetails,
        r.GrossPay ? parseFloat(r.GrossPay) : 0,
        r.NetPay ? parseFloat(r.NetPay) : 0,
        r.IncomeTaxPayable ? parseFloat(r.IncomeTaxPayable) : 0,
        r.IncomeTaxRecoverable ? parseFloat(r.IncomeTaxRecoverable) : 0,
        r.City, r.Email, r.SourceFile];
      let dedTotal = 0;
      for (const c of earnCodes) {
        const info = (r.Earnings || {})[c];
        row.push(info ? parseFloat(info.amt) : 0);
      }
      for (const c of dedCodes) {
        const info = (r.Deductions || {})[c];
        const amt = info ? parseFloat(info.amt) : 0;
        dedTotal += amt;
        row.push(amt);
      }
      row.push(Math.round(dedTotal * 100) / 100);
      ws.addRow(row);
    });

    ws.columns.forEach(col => { col.width = 16; });

    return { workbook: wb, nRows: records.length, nCols };
  }

  /** Trigger a browser download of the given ExcelJS workbook. */
  async function downloadWorkbook(workbook, filename) {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  return { buildBulkWorkbook, buildIndividualWorkbook, downloadWorkbook };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = XlsxWriter;
}
