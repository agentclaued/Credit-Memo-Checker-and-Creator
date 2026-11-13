(function () {
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const drop = $('rcDrop'), input = $('rcFile'), status = $('rcStatus');

  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') input.click(); });
  // only the two boxes take files; a drop anywhere else does nothing (and the browser must not open the file)
  const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  document.addEventListener('dragover', (e) => { if (hasFiles(e)) { e.preventDefault(); e.dataTransfer.dropEffect = 'none'; } });
  document.addEventListener('drop', (e) => { if (hasFiles(e)) e.preventDefault(); });
  input.addEventListener('change', () => { runAll(Array.from(input.files), 'batch'); input.value = ''; });
  // each box takes what is dropped on it
  [[drop, 'batch'], [$('rcSale'), 'sale']].forEach(([box, target]) => {
    box.addEventListener('dragover', (e) => { if (hasFiles(e)) { e.preventDefault(); e.stopPropagation(); e.dataTransfer.dropEffect = 'copy'; box.classList.add('over'); } });
    box.addEventListener('dragleave', (e) => { if (!box.contains(e.relatedTarget)) box.classList.remove('over'); });
    box.addEventListener('drop', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault(); e.stopPropagation();
      box.classList.remove('over');
      runAll(Array.from(e.dataTransfer.files), target);
    });
  });

  // one tab per checked file; only the selected file's result is shown
  // results are kept for this browser session, so leaving the page (e.g. to the DataBase Room) and coming back does not lose them
  const KEEP = 'rc-results', KEEP_TAB = 'rc-tab';
  let saved = [], restoring = false;
  const persist = () => { try { sessionStorage.setItem(KEEP, JSON.stringify(saved)); } catch (e) { /* storage full or blocked: skip */ } };
  const keepTab = (i) => { try { sessionStorage.setItem(KEEP_TAB, String(i)); } catch (e) { /* ignore */ } };
  // the message with the required value ('expected') marked green
  const msgHtml = (e) => {
    const m = esc(e.msg);
    if (e.expected == null || e.expected === '') return m;
    const x = esc(e.expected), at = m.indexOf(x);
    return at < 0 ? m : m.slice(0, at) + '<b class="rci-exp">' + x + '</b>' + m.slice(at + x.length);
  };
  function reset() { $('rcResults').innerHTML = ''; $('rcTabs').innerHTML = ''; saved = []; persist(); keepTab(0); }
  function addResult(name, card, errors) {
    const tab = document.createElement('button');
    tab.type = 'button'; tab.className = 'rc-tab'; tab.title = name;
    tab.innerHTML = '<span class="rc-tab-name">' + esc(name) + '</span><span class="rcn-count ' + (errors ? 'bad' : 'ok') + '">' + (errors ? errors : '✓') + '</span>';
    card.hidden = true;
    const pick = () => {
      $('rcTabs').querySelectorAll('.rc-tab').forEach((t) => t.classList.toggle('on', t === tab));
      $('rcResults').querySelectorAll(':scope > .card').forEach((c) => { c.hidden = c !== card; });
    };
    tab.addEventListener('click', () => { pick(); keepTab(Array.prototype.indexOf.call($('rcTabs').children, tab)); });
    $('rcTabs').appendChild(tab);
    $('rcResults').appendChild(card);
    if ($('rcTabs').children.length === 1) pick();
  }
  function errorCard(name, msg) {
    const c = document.createElement('section');
    c.className = 'card mt';
    c.innerHTML = '<b>' + esc(name) + '</b><p class="muted">' + esc(msg) + '</p>';
    addResult(name, c, 1);
  }

  let batchFiles = [];      // kept so the batches can be rechecked after a Sale Report is imported
  let needed = new Set();   // accounts whose sale report is missing (from the P4 errors)

  // `target` = the box the files were dropped on ('batch' | 'sale'); a drop elsewhere on the page is sorted by file type
  async function runAll(files, target) {
    const sale = target === 'sale' ? files.filter((f) => /\.(csv|xls|xlsx)$/i.test(f.name)) : target === 'batch' ? [] : files.filter((f) => /\.(csv|xls)$/i.test(f.name));
    const batches = target === 'sale' ? [] : files.filter((f) => /\.xlsx$/i.test(f.name));
    if (!sale.length && !batches.length) { status.textContent = target === 'sale' ? 'Drop a Sale Report .csv or .xlsx here.' : target === 'batch' ? 'Drop Credit Memo Batch .xlsx files here.' : 'Drop Credit Memo Batch .xlsx files, or a Sale Report .csv.'; return; }
    if (sale.length && !(await importSale(sale))) return;
    if (batches.length) batchFiles = batches;
    if (batchFiles.length) await checkBatches();
  }

  async function checkBatches() {
    reset(); needed = new Set();
    for (let i = 0; i < batchFiles.length; i++) await run(batchFiles[i], i + 1, batchFiles.length);
    showNeeded();
  }

  // ---- Sale Report import box: before or after the check
  async function importSale(files) {
    for (const f of files) {
      status.textContent = 'Importing Sale Report ' + f.name + '…';
      const fd = new FormData();
      fd.append('file', f);
      try {
        const res = await fetch('/upload/orders', { method: 'POST', body: fd });
        const data = await res.json().catch(() => ({}));
        if (res.status === 409) { status.textContent = f.name + ' is already imported, no changes.'; continue; }   // identical duplicate: just recheck
        if (!res.ok) throw new Error(data.detail || 'The import failed.');
        for (let t = 0; t < 600; t++) {                       // the import runs in the background; wait until it is done
          await new Promise((r) => setTimeout(r, 500));
          const list = await (await fetch('/api/files')).json();
          const row = list.find((x) => x.id === data.id);
          if (row && row.status === 'imported') break;
          if (row && row.status === 'failed') throw new Error(row.note || 'The import failed.');
        }
      } catch (e) {
        status.textContent = f.name + ': ' + e.message;
        return false;
      }
    }
    status.textContent = 'Sale Report imported.' + (batchFiles.length ? ' Rechecking…' : ' Now drop the batch files.');
    return true;
  }
  const sale = $('rcSale'), saleIn = $('rcSaleFile');
  sale.addEventListener('click', (e) => { if (!e.target.closest('button')) saleIn.click(); });
  saleIn.addEventListener('change', async () => {
    const files = Array.from(saleIn.files); saleIn.value = '';
    if (files.length && (await importSale(files)) && batchFiles.length) { await checkBatches(); status.textContent = ''; }
  });
  function showNeeded() {
    const list = Array.from(needed);
    $('rcSaleNeed').innerHTML = list.length ? 'Needed for: ' + list.map((a) => '<span class="chip">' + esc(a) + '</span>').join('') : '';
    sale.classList.toggle('need', list.length > 0);
  }

  async function run(file, k, n) {
    status.textContent = 'Checking ' + file.name + (n > 1 ? ' (' + k + ' of ' + n + ')' : '') + '…';
    const fd = new FormData();
    fd.append('file', file);
    try {
      const res = await fetch('/api/recheck', { method: 'POST', body: fd });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.detail || 'The check failed.');
      status.textContent = '';
      show(data);
    } catch (e) {
      status.textContent = '';
      saved.push({ err: file.name, msg: e.message }); persist();
      errorCard(file.name, e.message);
    }
  }

  function show(d) {
    if (!restoring) { saved.push({ d }); persist(); }
    d.results.forEach((r) => r.issues.forEach((i) => { const m = r.n === 4 && /sale report of (\S+)/.exec(i.msg); if (m) needed.add(m[1]); }));
    const errors = [];
    d.results.forEach((r) => r.issues.forEach((i) => errors.push({ ...i, n: r.n, title: r.title, rule: r.info, stage: r.stage || 0 })));
    const order = (d.sheets || []).slice();
    const rank = (sh) => { const k = order.indexOf(sh); return k < 0 ? 999 : k; };
    const rowOf = (w) => +((w.split('!')[1] || '').replace(/\D/g, '')) || 0;
    errors.sort((x, y) => rank(x.sheet) - rank(y.sheet) || x.n - y.n || rowOf(x.where) - rowOf(y.where));
    const total = d.count_total != null ? d.count_total : d.results.reduce((a, r) => a + r.count, 0);
    const sub = [d.memos + ' memo sheets', d.total != null ? 'total $' + Number(d.total).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '', d.date ? 'dated ' + d.date : ''].filter(Boolean).join(', ');
    const title = '<div class="rc-head"><span class="rc-stamp ' + (total ? 'bad' : 'ok') + '">' + (total ? total + ' ERROR' + (total > 1 ? 'S' : '') : 'ALL CLEAR') +
      '</span><div><div class="muted">' + esc(sub) + ' · ' + d.passed + ' of ' + d.results.length + ' checks passed' +
      (d.stopped ? '</div><div class="muted"><b>Stopped after "' + esc(d.stopped) + '"</b>: fix these first. ' + d.skipped + ' later checks run when this is clean.' : '') + '</div></div></div>';
    // reference sources: one chip each, green when found, amber when missing (with what to do)
    const chipOf = (ok, label, text, tip) => '<span class="chip ' + (ok ? 'ok' : 'warn') + '" title="' + esc(tip) + '"><b>' + (ok ? '✓' : '!') + '</b> ' + esc(label) + ': <span>' + esc(text) + '</span></span>';
    let chips = chipOf(!!d.ageing_file, 'Ageing', d.ageing_file || 'not imported', d.ageing_file ? 'The Ageing file used for this check' : 'Import the Ageing in the DataBase Room');
    (d.sale_files.length ? d.sale_files : ['not imported']).forEach((f) => {
      const m = /^(\S+): not imported$/.exec(f), ok = f !== 'not imported' && !m;
      chips += chipOf(ok, 'Sale Report' + (m ? ' ' + m[1] : ''), ok ? f : 'missing', ok ? 'The Sale Report used for this check' : 'Drop the Sale Report in the box above to finish the check');
    });
    // left-style tree: batch file > sheet (workbook order) > each problem
    const bySheet = {};
    errors.forEach((e) => { (bySheet[e.sheet] = bySheet[e.sheet] || []).push(e); });
    const names = order.slice();
    Object.keys(bySheet).forEach((x) => { if (!names.includes(x)) names.push(x); });
    // sheets with errors first (workbook order); the passed sheets follow inline in one row
    const label = (sh) => (/^\d+$/.test(sh) ? 'Memo ' : '') + sh;
    const bad = names.filter((sh) => (bySheet[sh] || []).length), good = names.filter((sh) => !(bySheet[sh] || []).length);
    // one flat table: Sheet | Cell | Error | Fix (format first, then Batch values / Ageing, then the rest)
    const rowHtml = (e, isNote) =>
      '<tr><td>' + esc(label(e.sheet)) + '</td><td><b class="rcn-cell">' + esc(e.where.split('!')[1] || e.where) + '</b></td><td><span class="rc-rule' + (isNote ? ' note' : '') + '" title="' + esc(isNote ? (e.title || '') : (e.title || '') + (e.rule ? ': ' + e.rule : '')) + '">' + (isNote ? (e.title || 'NOTE') : (e.label || 'P' + e.n)) + '</span>'  + '</td><td>' +
      (e.value != null ? '<span class="rci-val' + (isNote ? ' note' : '') + '" title="Value found">' + esc(e.value) + '</span>' : '') + '</td><td>' +
      (e.expected != null && e.expected !== '' ? '<b class="rci-exp">' + esc(e.expected) + '</b> ' : '') +
      (e.fix ? '<span class="info" tabindex="0" aria-label="' + esc(e.fix) + '">i<span class="info-tip">' + esc(e.fix) + '</span></span>' : '') + '</td><td>' + esc(e.msg) + '</td></tr>';
    const rowsHtml = [].concat(...bad.map((sh) => bySheet[sh].slice().sort((x, y) => (x.stage || 0) - (y.stage || 0) || x.n - y.n || rowOf(x.where) - rowOf(y.where)))).map((e) => rowHtml(e, false)).join('');
    const tableHtml = rowsHtml ? '<table class="rc-table"><thead><tr><th>Sheet</th><th>Cell</th><th>Error Type</th><th>Error</th><th>Fix</th><th>Ref</th></tr></thead><tbody>' + rowsHtml + '</tbody></table>' : '';
    const goodHtml = good.length ? '<div class="rc-passed"><span class="muted">Passed:</span>' + good.map((sh) => '<span class="rcn ok">' + esc(label(sh)) + ' ✓</span>').join('') + '</div>' : '';
    const notes = d.notes || [];
    const head = '<thead><tr><th>Sheet</th><th>Cell</th><th>Error Type</th><th>Error</th><th>Fix</th><th>Ref</th></tr></thead>';
    const notesHtml = notes.length ? '<hr class="rc-hr"><table class="rc-table">' + head + '<tbody>' + notes.map((e) => rowHtml(e, true)).join('') + '</tbody></table>' : '';
    const tree = '<div class="rcn root">' + esc(d.filename) + '</div>' + tableHtml + '<div class="rc-extra">' + goodHtml + '</div>' + notesHtml;
    const card = document.createElement('section');
    card.className = 'card mt';
    card.innerHTML = '<div class="rc-top"><h3>' + title + '</h3><div class="chips">' + chips + '</div></div><div class="rc-tree">' + tree + '</div>';
    card.addEventListener('click', (e) => {
      const b = e.target.closest('[data-all]');
      if (b) card.querySelectorAll('details').forEach((x) => { x.open = b.dataset.all === '1'; });
    });
    addResult(d.filename, card, total);
  }

  // ---- demo: example errors
  $('rcDemo').addEventListener('click', () => {
    status.textContent = '';
    reset();
    const mk = (n, title, issues) => ({ n, title, pass: !issues.length, count: issues.length, issues: issues.map(([where, msg]) => { const sh = where.split('!')[0]; return { where, msg, fix: 'Fix it in the cell', sheet: sh.startsWith('Memo ') ? sh.slice(5) : sh }; }) });
    show({
      filename: 'DEMO batch #10-00000.xlsx', memos: 3, total: 1206.83, date: '10/08/2026', sheets: ['REASON', 'PRICE LIST', 'Work1', 'Work2', 'Work3', 'Batch', '1', '2', '3'], ageing_file: 'Demo ageing', sale_files: ['Demo sale report'],
      passed: 6,
      results: [
        mk(1, 'No spaces in cells', [['Batch!B9', 'Space before or after the text: "LOWE\'S OD SALEM. NH "']]),
        mk(2, 'DB# with hashtag', [['Batch!E10', '"DB 536852" is not like DB#123456'], ['Memo 3!A36', 'Batch has "DB#536852" but the memo has "DB 536852"']]),
        mk(3, 'DB amount = Ageing', [['Memo 2', 'Memo total 940.11 but Ageing has 942.11 for DB#536851']]),
        mk(4, 'Item in sale report', []),
        mk(5, 'Salesman, Terms, Reason', [['Memo 1!F9', 'Terms is empty']]),
        mk(6, 'No extra formula', []),
        mk(7, 'Template formulas kept', []),
        mk(8, 'Account name = Ageing', []),
        mk(9, 'No duplicate item', [['Memo 1!A16', '7782LQP-24 is also on row 15']]),
        mk(10, 'Qty and price within sale report', []),
      ],
    });
  });

  // bring back the results of this session (after visiting another page); a refresh (F5) starts clean
  try {
    const nav = performance.getEntriesByType('navigation')[0];
    if (nav && nav.type === 'reload') { sessionStorage.removeItem(KEEP); sessionStorage.removeItem(KEEP_TAB); }
    const old = JSON.parse(sessionStorage.getItem(KEEP) || '[]');
    if (old.length) {
      restoring = true; needed = new Set();
      old.forEach((r) => { saved.push(r); if (r.d) show(r.d); else errorCard(r.err, r.msg); });
      restoring = false;
      showNeeded();
      const k = +(sessionStorage.getItem(KEEP_TAB) || 0), tabs = $('rcTabs').children;
      if (tabs[k]) tabs[k].click();
    }
  } catch (e) { restoring = false; }
})();
