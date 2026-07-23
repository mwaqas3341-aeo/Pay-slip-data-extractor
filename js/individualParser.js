/**
 * individualParser.js
 * -----------------------------------------------------------------------
 * Parses text extracted from a PIFRA / District Accounts Office "Monthly
 * Salary Statement" PDF (the single-employee format) into one structured
 * record. Also accepts manually-transcribed JSON records for
 * screenshots/scanned images of the same layout (see RECORD_TEMPLATE).
 *
 * JS port of scripts/extract_individual_slip.py from the
 * data-extraction-from-payslips skill.
 * -----------------------------------------------------------------------
 */

const IndividualParser = (() => {

  const CODE_RE = /(\d{4})\s+([A-Za-z][A-Za-z0-9 .()%/&'-]*?)\s+(-?[\d,]+\.\d{2})/g;

  const RECORD_TEMPLATE = {
    PersNo: "", Name: "", GuardianName: "", CNIC: "", PayPeriod: "",
    DOB: "", EntryIntoService: "", LengthOfService: "",
    EmploymentCategory: "", Designation: "", DepartmentCode: "",
    DDOCode: "", DDOName: "", BPS: "", PayStage: "", PayScale: "",
    GPFAccountNo: "", GPFInterestStatus: "", GPFBalance: "",
    BankAccountNo: "", BankDetails: "", GrossPay: "", NetPay: "",
    IncomeTaxPayable: "", IncomeTaxRecoverable: "", City: "", Email: "",
    SourceFile: "",
    Earnings: { "0001": { desc: "Basic Pay", amt: "0.00" } },
    Deductions: { "3609": { desc: "Income Tax", amt: "0.00" } },
  };

  /** Parse the full text of one PDF (already extracted, e.g. via pdf.js
   *  joining all pages' text with newlines) into a structured record. */
  function parseText(text, sourceFileName) {
    const d = { SourceFile: sourceFileName || '' };
    let m;

    m = text.match(/Monthly Salary Statement\s*\(([A-Za-z]+-\d{4})\)/);
    d.PayPeriod = m ? m[1] : null;

    m = text.match(/Personal Information of\s+(?:Mr|Ms|Mrs)\s+(.+?)\s+(?:d\/w\/s|s\/o|d\/o|w\/o)\s+of\s+(.+)/);
    d.Name = m ? m[1].trim() : null;
    d.GuardianName = m ? m[2].trim() : null;

    m = text.match(/Personnel Number:\s*(\S+)\s+CNIC:\s*(\S*)/);
    d.PersNo = m ? m[1] : null;
    d.CNIC = m ? m[2] : null;

    m = text.match(/Date of Birth:\s*(\S+)\s+Entry into Govt\. Service:\s*(\S+)\s+Length of Service:\s*(.+)/);
    d.DOB = m ? m[1] : null;
    d.EntryIntoService = m ? m[2] : null; // exact date - no estimation needed here
    d.LengthOfService = m ? m[3].trim() : null;

    m = text.match(/Employment Category:\s*(.+)/);
    d.EmploymentCategory = m ? m[1].trim() : null;

    m = text.match(/Designation:\s*(.+?)\s+(\d{8}-\S.*)/);
    d.Designation = m ? m[1].trim() : null;
    d.DepartmentCode = m ? m[2].trim() : null;

    m = text.match(/DDO Code:\s*(\S+)-(.+)/);
    d.DDOCode = m ? m[1] : null;
    d.DDOName = m ? m[2].trim() : null;

    m = text.match(/GPF A\/C No:\s*(\d*)\s*(GPF Interest \S+)?\s*GPF Balance:\s*([\d,.]+)/);
    d.GPFAccountNo = (m && m[1]) ? m[1] : null;
    d.GPFInterestStatus = (m && m[2]) ? m[2] : null;
    d.GPFBalance = m ? m[3].replace(/,/g, '') : null;

    m = text.match(/Pay scale:\s*(.+?)\s+Pay Scale Type:\s*(\S+)\s+BPS:\s*(\d+)\s+Pay Stage:\s*(\d+)/);
    d.PayScale = m ? m[1].trim() : null;
    d.PayScaleType = m ? m[2] : null;
    d.BPS = m ? m[3] : null;
    d.PayStage = m ? m[4] : null;

    // Earnings: between the table header and "Deductions - General"
    const em = text.match(/Wage type\s+Amount\s+Wage type\s+Amount\s*\n([\s\S]*?)Deductions - General/);
    const earnText = em ? em[1] : '';
    d.Earnings = {};
    for (const cm of earnText.matchAll(CODE_RE)) {
      d.Earnings[cm[1]] = { desc: cm[2].trim(), amt: cm[3].replace(/,/g, '') };
    }

    // Deductions - General: between the 2nd table header and "Deductions - Loans"
    const dm = text.match(/Deductions - General\s*\n\s*Wage type\s+Amount\s+Wage type\s+Amount\s*\n([\s\S]*?)Deductions - Loans and Advances/);
    const dedText = dm ? dm[1] : '';
    d.Deductions = {};
    for (const cm of dedText.matchAll(CODE_RE)) {
      d.Deductions[cm[1]] = { desc: cm[2].trim(), amt: cm[3].replace(/,/g, '') };
    }

    m = text.match(/Payable:\s*([\d,.]+)\s+Recovered till \S+:\s*([\d,.]+)\s+Exempted:\s*([\d.,\-]+)\s+Recoverable:\s*([\d,.]+)/);
    if (m) {
      d.IncomeTaxPayable = m[1].replace(/,/g, '');
      d.IncomeTaxRecoveredTillDate = m[2].replace(/,/g, '');
      d.IncomeTaxExemptedPct = m[3];
      d.IncomeTaxRecoverable = m[4].replace(/,/g, '');
    }

    m = text.match(/Gross Pay \(Rs\.\):\s*([\d,.]+)\s+Deductions: \(Rs\.\):\s*(-?[\d,.]+)\s+Net Pay: \(Rs\.\):\s*([\d,.]+)/);
    d.GrossPay = m ? m[1].replace(/,/g, '') : null;
    d.TotalDeductions = m ? m[2].replace(/,/g, '') : null;
    d.NetPay = m ? m[3].replace(/,/g, '') : null;

    m = text.match(/Payee Name:\s*(.+)/);
    d.PayeeName = m ? m[1].trim() : null;
    m = text.match(/Account Number:\s*(\S+)/);
    d.BankAccountNo = m ? m[1] : null;
    m = text.match(/Bank Details:\s*(.+)/);
    d.BankDetails = m ? m[1].trim() : null;

    m = text.match(/City:\s*(\S*)\s+Domicile:\s*(.+?)\s+Housing Status:\s*(.+)/);
    d.City = m ? m[1] : null;
    d.Domicile = m ? m[2].trim() : null;
    d.HousingStatus = m ? m[3].trim() : null;

    m = text.match(/Email:\s*(\S+)/);
    d.Email = m ? m[1] : null;

    return d;
  }

  /** Load one or more manually-transcribed records from parsed JSON
   *  (object or array of objects), filling in defaults. */
  function normalizeJsonRecords(data, sourceFileName) {
    const arr = Array.isArray(data) ? data : [data];
    return arr.map(r => ({
      Earnings: {}, Deductions: {}, SourceFile: sourceFileName || '', ...r,
    }));
  }

  return { parseText, normalizeJsonRecords, RECORD_TEMPLATE, CODE_RE };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = IndividualParser;
}
