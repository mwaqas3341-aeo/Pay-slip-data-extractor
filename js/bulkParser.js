/**
 * bulkParser.js
 * -----------------------------------------------------------------------
 * Parses a Punjab Govt AG/District payroll ".XLS" export. Despite the
 * extension, these files are UTF-16 encoded, tab-delimited plain text: a
 * raw print dump showing TWO employee payslips side by side per line.
 *
 * This is a JS port of scripts/extract_payslips.py from the
 * data-extraction-from-payslips skill. See README.md / SKILL_NOTES.md for
 * the full field reference and known quirks.
 * -----------------------------------------------------------------------
 */

const BulkParser = (() => {

  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june',
    'july', 'august', 'september', 'october', 'november', 'december'];

  // Matches a pay/allowance/deduction line, e.g.:
  //   "2378-Adhoc Relief All 2023 35%    13,744.00"
  // Description may contain hyphens, so the char class must allow '-'.
  const CODE_RE = /(\d{4})-([A-Za-z0-9 .()%/&'-]+?)\s{2,}([\d,]+\.\d{2})/g;

  /** Detect UTF-16 BOM (LE or BE) and decode; falls back to LE if no BOM found
   *  (observed default for this payroll system's export). */
  function decodeUtf16(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    let little = true;
    let offset = 0;
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
      little = true; offset = 2;
    } else if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
      little = false; offset = 2;
    }
    const decoder = new TextDecoder(little ? 'utf-16le' : 'utf-16be');
    return decoder.decode(arrayBuffer.slice(offset));
  }

  /** Real .xlsx/.xlsm files are zip archives — they start with the "PK"
   *  signature. If someone opens the raw UTF-16 dump in Excel and does a
   *  genuine "Save As .xlsx", the file stops being raw text and becomes one
   *  of these; it needs to be read cell-by-cell instead of decoded as text. */
  function isRealXlsx(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer.slice(0, 4));
    return bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4B; // "PK"
  }

  /** Reconstructs the original tab-delimited print-dump text from a real
   *  .xlsx workbook: each worksheet row becomes one text line, each column
   *  in that row is rejoined with a tab — reversing whatever Excel did when
   *  it split the original tab-delimited line into columns on Save As. */
  async function reconstructTextFromXlsx(arrayBuffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    const lines = [];
    wb.eachSheet(sheet => {
      sheet.eachRow({ includeEmpty: true }, row => {
        const maxCol = row.cellCount || (row.values ? row.values.length - 1 : 0);
        const cells = [];
        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c);
          let v = cell.value;
          if (v === null || v === undefined) v = '';
          else if (typeof v === 'object' && v.richText) v = v.richText.map(t => t.text).join('');
          else if (typeof v === 'object' && v.text) v = v.text;
          else v = String(v);
          cells.push(v);
        }
        lines.push(cells.join('\t'));
      });
    });
    return lines.join('\n');
  }

  /** The source prints two employee slips side by side per line, separated
   *  by the first tab on that line. Split into two independent streams. */
  function buildStreams(text) {
    const lines = text.split(/\r\n|\r|\n/);
    const left = [];
    const right = [];
    for (const line of lines) {
      const idx = line.indexOf('\t');
      if (idx >= 0) {
        left.push(line.slice(0, idx));
        right.push(line.slice(idx + 1));
      } else {
        left.push(line);
      }
    }
    return [left.join('\n'), right.join('\n')];
  }

  function splitRecords(text) {
    const parts = text.split(/(?=S#:\d+)/);
    return parts.filter(p => p.trim().startsWith('S#:'));
  }

  function parseRecord(chunk) {
    const d = {};
    let m;

    m = chunk.match(/S#:(\d+)/);
    d.SNo = m ? m[1] : null;

    m = chunk.match(/P Sec:(\S+)/);
    d.PSec = m ? m[1] : null;

    m = chunk.match(/\b([A-Z]{2}\d{4})\s*-([^\r\n]*)/);
    d.OfficeCode = m ? m[1] : null;
    d.OfficeNameTrunc = m ? m[2].trim() : null;

    m = chunk.match(/Pers #:\s*(\S+)\s+Buckle:\s*(\S*)/);
    d.PersNo = m ? m[1] : null;
    d.Buckle = m ? m[2] : null;

    m = chunk.match(/Name:\s+(.*?)\s{2,}NTN/);
    d.Name = m ? m[1].trim() : null;

    m = chunk.match(/NTN:[^\n]*\n\s*(.*?)\s{2,}GPF #:/s);
    d.Designation = m ? m[1].trim() : null;

    m = chunk.match(/CNIC No\.?\s*(\S*)/);
    d.CNIC = m ? m[1] : null;

    // Split into EARNINGS ("PAYS AND ALLOWANCES:" .. "Gross Pay and Allowances")
    // and DEDUCTIONS ("DEDUCTIONS:" .. "Total Deductions") sections, so codes
    // are classified by *section*, not by prefix guesswork (a 6xxx code can
    // be either an earning or a deduction adjustment).
    const em = chunk.match(/PAYS AND ALLOWANCES:([\s\S]*?)Gross Pay and Allowances/);
    const earningsText = em ? em[1] : chunk;
    const dm = chunk.match(/DEDUCTIONS:([\s\S]*?)Total Deductions/);
    const deductionsText = dm ? dm[1] : '';

    d.PayComponents = {};
    for (const cm of earningsText.matchAll(CODE_RE)) {
      d.PayComponents[cm[1]] = { desc: cm[2].trim(), amt: cm[3].replace(/,/g, '') };
    }
    d.DeductionComponents = {};
    for (const cm of deductionsText.matchAll(CODE_RE)) {
      d.DeductionComponents[cm[1]] = { desc: cm[2].trim(), amt: cm[3].replace(/,/g, '') };
    }

    m = chunk.match(/Gross Pay and Allowances\s+([\d,]+\.\d{2})/);
    d.GrossPay = m ? m[1].replace(/,/g, '') : null;

    // D.O.B + Bank Name/Branch line (the line right after "D.O.B ... LFP Quota")
    m = chunk.match(/D\.O\.B\b[\s\S]*?\n\s*(\d{2}\.\d{2}\.\d{4})\s+(\S.*?)\s*\n/);
    d.DOB = m ? m[1] : null;
    d.BankNameBranch = m ? m[2].trim() : null;

    // Length of service + bank account number line
    m = chunk.match(/(\d{2} Years \d{2} Months \d{3} Days)\s+(\S+)/);
    d.LengthOfService = m ? m[1] : null;
    d.BankAccountNo = m ? m[2] : null;

    return d;
  }

  /** Combine multi-slot continuation records (same employee spread across
   *  2-3 consecutive S# slots) into one row per employee. */
  function mergeRecords(allRecords, officePrefix) {
    let records = allRecords;
    if (officePrefix) {
      const prefix = officePrefix.toUpperCase();
      records = records.filter(r => r.OfficeCode && r.OfficeCode.startsWith(prefix));
    }
    const byKey = new Map();
    for (const r of records) {
      const key = `${r.PersNo}||${r.OfficeCode}`;
      if (!byKey.has(key)) byKey.set(key, []);
      byKey.get(key).push(r);
    }

    const merged = [];
    for (const group of byKey.values()) {
      group.sort((a, b) => (parseInt(a.SNo, 10) || 0) - (parseInt(b.SNo, 10) || 0));
      const base = group[0];
      const comps = {};
      const dcomps = {};
      for (const g of group) {
        Object.assign(comps, g.PayComponents);
        Object.assign(dcomps, g.DeductionComponents || {});
      }
      merged.push({
        PersNo: base.PersNo, OfficeCode: base.OfficeCode,
        OfficeNameTrunc: base.OfficeNameTrunc, Name: base.Name,
        Designation: base.Designation, CNIC: base.CNIC,
        Buckle: base.Buckle, GrossPay: base.GrossPay,
        DOB: base.DOB, BankNameBranch: base.BankNameBranch,
        BankAccountNo: base.BankAccountNo, LengthOfService: base.LengthOfService,
        PayComponents: comps, DeductionComponents: dcomps,
      });
    }
    return merged;
  }

  /** Estimate joining date by subtracting "Length of Service" from a
   *  reference date. Approximate only (checked 2-4 days off against real
   *  slips) — never a substitute for a real "Entry into Govt. Service" date
   *  when one is available (individual slip format). */
  function estJoiningDate(lengthStr, refDate) {
    if (!lengthStr) return null;
    const m = lengthStr.match(/(\d{2}) Years (\d{2}) Months (\d{3}) Days/);
    if (!m) return null;
    const y = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
    const totalMonths = y * 12 + mo;
    let yy = refDate.getFullYear() - Math.floor(totalMonths / 12);
    let mm = (refDate.getMonth() + 1) - (totalMonths % 12);
    if (mm <= 0) { mm += 12; yy -= 1; }
    let dd = refDate.getDate() - d;
    let guard = 0;
    while (dd <= 0 && guard < 60) {
      mm -= 1;
      if (mm <= 0) { mm += 12; yy -= 1; }
      dd += 28;
      guard += 1;
    }
    try {
      const day = Math.min(dd, 28);
      const dt = new Date(yy, mm - 1, day);
      const pad = n => String(n).padStart(2, '0');
      return `${pad(dt.getDate())}.${pad(dt.getMonth() + 1)}.${dt.getFullYear()}`;
    } catch (e) {
      return null;
    }
  }

  function detectRefDate(leftText) {
    const m = leftText.match(/Month:(\w+)\s+(\d{4})/);
    if (m) {
      const idx = MONTHS.indexOf(m[1].toLowerCase());
      if (idx >= 0) return new Date(parseInt(m[2], 10), idx, 28);
    }
    return new Date();
  }

  /**
   * Full pipeline: raw ArrayBuffer -> { merged, refDate, rawCount }
   * officePrefix: optional string, e.g. "LL"
   * onProgress: optional callback(stage:string) for UI feedback
   */
  async function parseBulkFile(arrayBuffer, officePrefix, onProgress) {
    const notify = (s) => { if (onProgress) onProgress(s); };

    let text;
    if (isRealXlsx(arrayBuffer)) {
      notify('Detected a real Excel workbook (not raw text) — reading cells and reconstructing the original print layout…');
      text = await reconstructTextFromXlsx(arrayBuffer);
    } else {
      notify('Decoding UTF-16 text stream…');
      text = decodeUtf16(arrayBuffer);
    }

    notify('Splitting two-column print layout…');
    const [leftText, rightText] = buildStreams(text);

    notify('Parsing individual payslip records…');
    const leftRecs = splitRecords(leftText).map(parseRecord);
    const rightRecs = splitRecords(rightText).map(parseRecord);
    const allRecords = leftRecs.concat(rightRecs);

    notify('Merging continuation slots into one row per employee…');
    const merged = mergeRecords(allRecords, officePrefix);

    const refDate = detectRefDate(leftText);

    return { merged, refDate, rawCount: allRecords.length };
  }

  return {
    decodeUtf16, buildStreams, splitRecords, parseRecord, mergeRecords,
    estJoiningDate, detectRefDate, parseBulkFile, isRealXlsx, reconstructTextFromXlsx,
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = BulkParser;
}
