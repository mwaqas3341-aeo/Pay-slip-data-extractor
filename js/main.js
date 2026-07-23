/**
 * main.js — wires the UI to bulkParser.js / individualParser.js / xlsxWriter.js.
 * Everything here runs client-side; no file is ever sent to a server.
 */
(() => {
  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }

  // ---------------------------------------------------------------- Tabs
  const tabs = document.querySelectorAll('.mode-btn');
  const panels = { bulk: document.getElementById('panel-bulk'), individual: document.getElementById('panel-individual') };
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => { t.classList.remove('active'); t.setAttribute('aria-selected', 'false'); });
      tab.classList.add('active');
      tab.setAttribute('aria-selected', 'true');
      const mode = tab.dataset.mode;
      Object.entries(panels).forEach(([key, el]) => { el.hidden = key !== mode; });
    });
  });

  // ---------------------------------------------------------------- Helpers
  function log(boxId, message, cls) {
    const box = document.getElementById(boxId);
    const empty = box.querySelector('.log-empty');
    if (empty) empty.remove();
    const p = document.createElement('p');
    if (cls) p.className = cls;
    p.textContent = message;
    box.appendChild(p);
    box.scrollTop = box.scrollHeight;
  }
  function clearLog(boxId) {
    const box = document.getElementById(boxId);
    box.innerHTML = '<p class="log-empty">Extraction log will appear here.</p>';
  }
  function fmtBytes(n) {
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / (1024 * 1024)).toFixed(1)} MB`;
  }

  function wireDropzone(zoneId, inputId, onFiles) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    input.addEventListener('change', () => onFiles(Array.from(input.files)));
    ['dragenter', 'dragover'].forEach(evt =>
      zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add('drag-over'); }));
    ['dragleave', 'drop'].forEach(evt =>
      zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove('drag-over'); }));
    zone.addEventListener('drop', e => {
      const files = Array.from(e.dataTransfer.files);
      if (files.length) onFiles(files);
    });
  }

  // ============================================================ BULK MODE
  let bulkFile = null;
  const chipBulk = document.getElementById('file-chip-bulk');
  const btnRunBulk = document.getElementById('btn-run-bulk');

  wireDropzone('dropzone-bulk', 'file-bulk', files => {
    if (!files.length) return;
    bulkFile = files[0];
    chipBulk.hidden = false;
    chipBulk.innerHTML = `<span>${bulkFile.name}</span><span>${fmtBytes(bulkFile.size)}</span>`;
    btnRunBulk.disabled = false;
    clearLog('log-bulk');
    document.getElementById('result-bulk').hidden = true;
  });

  btnRunBulk.addEventListener('click', async () => {
    if (!bulkFile) return;
    btnRunBulk.disabled = true;
    clearLog('log-bulk');
    document.getElementById('result-bulk').hidden = true;
    const officePrefix = document.getElementById('office-prefix').value.trim() || null;

    try {
      log('log-bulk', `Reading ${bulkFile.name} (${fmtBytes(bulkFile.size)})…`);
      const arrayBuffer = await bulkFile.arrayBuffer();

      const { merged, refDate, rawCount } = BulkParser.parseBulkFile(
        arrayBuffer, officePrefix, stage => log('log-bulk', stage));

      log('log-bulk', `Raw slip records found: ${rawCount}`);
      log('log-bulk', `Employee rows after merge${officePrefix ? ` (filtered to "${officePrefix}")` : ''}: ${merged.length}`, 'log-ok');

      if (merged.length === 0) {
        log('log-bulk', 'No matching records — check the office-code prefix, or that this is really the UTF-16 payroll export.', 'log-error');
        btnRunBulk.disabled = false;
        return;
      }

      log('log-bulk', 'Building colored .xlsx workbook…');
      const { workbook, nRows, nCols } = await XlsxWriter.buildBulkWorkbook(merged, refDate);
      log('log-bulk', `Done — ${nRows} rows × ${nCols} columns.`, 'log-ok');

      renderBulkResult(merged, nRows, nCols, workbook, officePrefix);
    } catch (err) {
      console.error(err);
      log('log-bulk', `Error: ${err.message}`, 'log-error');
    } finally {
      btnRunBulk.disabled = false;
    }
  });

  function renderBulkResult(merged, nRows, nCols, workbook, officePrefix) {
    const summary = document.getElementById('result-summary-bulk');
    summary.textContent = `${nRows} employee${nRows === 1 ? '' : 's'} · ${nCols} columns` +
      (officePrefix ? ` · office prefix "${officePrefix}"` : ' · all departments');

    const earnDesc = {}, dedDesc = {};
    for (const m of merged) {
      for (const [c, i] of Object.entries(m.PayComponents)) if (!(c in earnDesc)) earnDesc[c] = i.desc;
      for (const [c, i] of Object.entries(m.DeductionComponents || {})) if (!(c in dedDesc)) dedDesc[c] = i.desc;
    }
    const earnCodes = Object.keys(earnDesc).sort();
    const dedCodes = Object.keys(dedDesc).sort();
    const fixed = ['S.No', 'Personal No.', 'Name', 'Office Code', 'Designation', 'Gross Pay'];
    const headers = [...fixed, ...earnCodes.map(c => `${c}-${earnDesc[c]}`), ...dedCodes.map(c => `${c}-${dedDesc[c]}`), 'Total Ded.'];

    const table = document.getElementById('preview-table-bulk');
    table.innerHTML = '';
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    headers.forEach((h, i) => {
      const th = document.createElement('th');
      th.textContent = h;
      if (i >= fixed.length && i < fixed.length + earnCodes.length) th.className = 'col-earn';
      else if (i >= fixed.length + earnCodes.length && i < headers.length - 1) th.className = 'col-ded';
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    const previewRows = merged.slice(0, 25);
    previewRows.forEach((m, idx) => {
      const tr = document.createElement('tr');
      const vals = [idx + 1, m.PersNo, m.Name, m.OfficeCode, m.Designation, m.GrossPay];
      vals.forEach(v => { const td = document.createElement('td'); td.textContent = v ?? ''; tr.appendChild(td); });
      let dedTotal = 0;
      earnCodes.forEach(c => {
        const info = m.PayComponents[c];
        const td = document.createElement('td'); td.className = 'col-earn';
        td.textContent = info ? Number(info.amt).toLocaleString() : '0';
        tr.appendChild(td);
      });
      dedCodes.forEach(c => {
        const info = (m.DeductionComponents || {})[c];
        const amt = info ? parseFloat(info.amt) : 0;
        dedTotal += amt;
        const td = document.createElement('td'); td.className = 'col-ded';
        td.textContent = amt.toLocaleString();
        tr.appendChild(td);
      });
      const tdTotal = document.createElement('td'); tdTotal.className = 'col-ded';
      tdTotal.textContent = dedTotal.toLocaleString();
      tr.appendChild(tdTotal);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    if (merged.length > previewRows.length) {
      log('log-bulk', `Preview shows first ${previewRows.length} of ${merged.length} rows — the download has all of them.`);
    }

    document.getElementById('result-bulk').hidden = false;
    const btnDownload = document.getElementById('btn-download-bulk');
    btnDownload.onclick = () => {
      const stamp = new Date().toISOString().slice(0, 10);
      XlsxWriter.downloadWorkbook(workbook, `payroll_extract_${stamp}.xlsx`);
    };
  }

  // ======================================================= INDIVIDUAL MODE
  let individualFiles = [];
  const listIndividual = document.getElementById('file-list-individual');
  const btnRunIndividual = document.getElementById('btn-run-individual');

  wireDropzone('dropzone-individual', 'file-individual', files => {
    individualFiles = individualFiles.concat(files);
    renderFileList();
    btnRunIndividual.disabled = individualFiles.length === 0;
    clearLog('log-individual');
    document.getElementById('result-individual').hidden = true;
  });

  function renderFileList() {
    listIndividual.innerHTML = '';
    individualFiles.forEach((f, idx) => {
      const li = document.createElement('li');
      const nameSpan = document.createElement('span');
      nameSpan.textContent = f.name;
      const rmBtn = document.createElement('span');
      rmBtn.textContent = '✕ ' + fmtBytes(f.size);
      rmBtn.style.cursor = 'pointer';
      rmBtn.title = 'Remove';
      rmBtn.addEventListener('click', () => {
        individualFiles.splice(idx, 1);
        renderFileList();
        btnRunIndividual.disabled = individualFiles.length === 0;
      });
      li.appendChild(nameSpan);
      li.appendChild(rmBtn);
      listIndividual.appendChild(li);
    });
  }

  document.getElementById('btn-template').addEventListener('click', () => {
    const blob = new Blob([JSON.stringify(IndividualParser.RECORD_TEMPLATE, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'manual_record_template.json';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  });

  /** pdf.js hands back text as disconnected fragments with x/y positions —
   *  NOT pre-arranged into lines. Field regexes below rely on real line
   *  breaks (matching the slip's visual layout), so fragments are grouped
   *  into lines by y-position (rows), then ordered left-to-right by x
   *  within each row, before being joined. Without this, an entire page
   *  becomes one run-on line and every regex over-captures to the end
   *  of the page. */
  function linesFromTextContent(items) {
    const frags = items
      .filter(it => it.str && it.str.trim().length)
      .map(it => ({ x: it.transform[4], y: it.transform[5], str: it.str }));
    frags.sort((a, b) => b.y - a.y || a.x - b.x);

    const rows = [];
    const Y_TOLERANCE = 2.5;
    let current = [];
    let currentY = null;
    for (const f of frags) {
      if (currentY === null || Math.abs(f.y - currentY) <= Y_TOLERANCE) {
        current.push(f);
        if (currentY === null) currentY = f.y;
      } else {
        rows.push(current);
        current = [f];
        currentY = f.y;
      }
    }
    if (current.length) rows.push(current);

    return rows.map(row => row.sort((a, b) => a.x - b.x).map(f => f.str).join(' '));
  }

  /** Returns an array of page texts (one string per PDF page), each with
   *  reconstructed line breaks. */
  async function extractPdfPages(file) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      pages.push(linesFromTextContent(content.items).join('\n'));
    }
    return pages;
  }

  btnRunIndividual.addEventListener('click', async () => {
    if (!individualFiles.length) return;
    btnRunIndividual.disabled = true;
    clearLog('log-individual');
    document.getElementById('result-individual').hidden = true;

    const records = [];
    try {
      for (const file of individualFiles) {
        const lower = file.name.toLowerCase();
        try {
          if (lower.endsWith('.json')) {
            log('log-individual', `Reading ${file.name} as JSON…`);
            const text = await file.text();
            const data = JSON.parse(text);
            const recs = IndividualParser.normalizeJsonRecords(data, file.name);
            records.push(...recs);
            log('log-individual', `  loaded ${recs.length} record(s) from ${file.name}`, 'log-ok');
          } else if (lower.endsWith('.pdf')) {
            log('log-individual', `Extracting text layer from ${file.name}…`);
            const pages = await extractPdfPages(file);
            const multiPage = pages.length > 1;
            let pageOk = 0;
            pages.forEach((pageText, idx) => {
              if (!pageText.trim()) return;
              const sourceLabel = multiPage ? `${file.name} (p.${idx + 1})` : file.name;
              const rec = IndividualParser.parseText(pageText, sourceLabel);
              if (!rec.PersNo && !rec.Name) {
                log('log-individual', `  WARNING: no recognizable fields on page ${idx + 1} of ${file.name} — may be a scanned image, a cover/blank page, or a different layout.`, 'log-error');
              } else {
                pageOk += 1;
                records.push(rec);
              }
            });
            if (multiPage) {
              log('log-individual', `  ${file.name}: ${pageOk} of ${pages.length} pages parsed as employee records`, 'log-ok');
            } else if (pageOk) {
              log('log-individual', `  parsed ${records[records.length - 1].Name || records[records.length - 1].PersNo || file.name}`, 'log-ok');
            }
          } else {
            log('log-individual', `Skipping ${file.name} — unsupported file type.`, 'log-error');
          }
        } catch (innerErr) {
          console.error(innerErr);
          log('log-individual', `  WARNING: failed to parse ${file.name}: ${innerErr.message}`, 'log-error');
        }
      }

      if (!records.length) {
        log('log-individual', 'No records extracted.', 'log-error');
        return;
      }

      log('log-individual', 'Building colored .xlsx workbook…');
      const { workbook, nRows, nCols } = await XlsxWriter.buildIndividualWorkbook(records);
      log('log-individual', `Done — ${nRows} rows × ${nCols} columns.`, 'log-ok');

      renderIndividualResult(records, nRows, nCols, workbook);
    } catch (err) {
      console.error(err);
      log('log-individual', `Error: ${err.message}`, 'log-error');
    } finally {
      btnRunIndividual.disabled = false;
    }
  });

  function renderIndividualResult(records, nRows, nCols, workbook) {
    const summary = document.getElementById('result-summary-individual');
    summary.textContent = `${nRows} employee${nRows === 1 ? '' : 's'} · ${nCols} columns`;

    const earnDesc = {}, dedDesc = {};
    for (const r of records) {
      for (const [c, i] of Object.entries(r.Earnings || {})) if (!(c in earnDesc)) earnDesc[c] = i.desc;
      for (const [c, i] of Object.entries(r.Deductions || {})) if (!(c in dedDesc)) dedDesc[c] = i.desc;
    }
    const earnCodes = Object.keys(earnDesc).sort();
    const dedCodes = Object.keys(dedDesc).sort();
    const fixed = ['S.No', 'Personal No.', 'Name', 'Designation', 'DDO Name', 'Gross Pay', 'Net Pay'];
    const headers = [...fixed, ...earnCodes.map(c => `${c}-${earnDesc[c]}`), ...dedCodes.map(c => `${c}-${dedDesc[c]}`)];

    const table = document.getElementById('preview-table-individual');
    table.innerHTML = '';
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    headers.forEach((h, i) => {
      const th = document.createElement('th');
      th.textContent = h;
      if (i >= fixed.length && i < fixed.length + earnCodes.length) th.className = 'col-earn';
      else if (i >= fixed.length + earnCodes.length) th.className = 'col-ded';
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    records.forEach((r, idx) => {
      const tr = document.createElement('tr');
      const vals = [idx + 1, r.PersNo, r.Name, r.Designation, r.DDOName, r.GrossPay, r.NetPay];
      vals.forEach(v => { const td = document.createElement('td'); td.textContent = v ?? ''; tr.appendChild(td); });
      earnCodes.forEach(c => {
        const info = (r.Earnings || {})[c];
        const td = document.createElement('td'); td.className = 'col-earn';
        td.textContent = info ? Number(info.amt).toLocaleString() : '0';
        tr.appendChild(td);
      });
      dedCodes.forEach(c => {
        const info = (r.Deductions || {})[c];
        const td = document.createElement('td'); td.className = 'col-ded';
        td.textContent = info ? Number(info.amt).toLocaleString() : '0';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    document.getElementById('result-individual').hidden = false;
    const btnDownload = document.getElementById('btn-download-individual');
    btnDownload.onclick = () => {
      const stamp = new Date().toISOString().slice(0, 10);
      XlsxWriter.downloadWorkbook(workbook, `individual_slips_${stamp}.xlsx`);
    };
  }
})();
