/* Demo of the template form. Math mirrors the template formulas only:
   amount = price*qty, subtotal = SUM(lines), total = SUM(subtotal..adjustment), batch = SUM(amounts). */
(function () {
  const R = window.MEMO_RULES, $ = s => document.querySelector(s);
  const state = { memos: [], cur: -1, prices: {}, reasons: [] };
  const money = n => (Math.round((Number(n) || 0) * 100) / 100);
  const fmt = n => money(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const newMemo = () => ({ date: $('#bDate').value, soldTo: '', soldAcct: '', shipTo: '', shipAcct: '', code: '', salesman: '', terms: '', via: '',
    claim: '', gst: '', freight: '', adj: '', db: '', lines: Array.from({ length: R.lines }, () => ({ sku: '', qty: '', price: '' })) });
  const lineAmt = l => money((Number(l.price) || 0) * (Number(l.qty) || 0));            // =G*D
  const subtotal = m => money(m.lines.reduce((s, l) => s + lineAmt(l), 0));              // =SUM(I13:I33)
  const total = m => money(subtotal(m) + (Number(m.gst) || 0) + (Number(m.freight) || 0) + (Number(m.adj) || 0)); // =SUM(I34:I37)
  const reasonOf = m => (state.reasons.find(r => r.code === m.code) || {}).description || '';

  // ---- Batch sheet: 20 editable rows (template rows 8-27); columns after the fixed # column
  const BCOLS = ['company', 'date', 'amount', 'db', 'claim', 'rt', 'rtAmt', 'disc', 'over'], MONEY_COLS = ['amount', 'rtAmt', 'disc', 'over'];
  const BROWS = R.maxMemos;                                  // the template has 20 rows; more are added when needed
  state.batch = Array.from({ length: BROWS }, () => Object.fromEntries(BCOLS.map(c => [c, ''])));
  const toNum = v => { const n = parseFloat(String(v ?? '').replace(/[^0-9.\-]/g, '')); return isNaN(n) ? null : n; };
  const showMoney = v => { const n = toNum(v); return n === null || v === '' ? '' : '$ ' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  const mdy = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? `${+m[2]}/${+m[3]}/${m[1]}` : (iso || ''); };

  function buildBatch() {
    $('#batchTable tbody').innerHTML = state.batch.map((_, r) => `<tr><td class="rn">${r + 1}</td>` + BCOLS.map((c, k) =>
      `<td><input class="cell ${MONEY_COLS.includes(c) ? 'num' : ''}" data-r="${r}" data-c="${k}" autocomplete="off"></td>`).join('') + '</tr>').join('');
    renderBatch();
  }
  function renderBatch(force) {
    document.querySelectorAll('#batchTable input.cell').forEach(el => {
      if (!force && document.activeElement === el) return;
      const c = BCOLS[+el.dataset.c], v = state.batch[+el.dataset.r][c];
      el.value = MONEY_COLS.includes(c) ? showMoney(v) : v;
    });
    const t = state.batch.reduce((s, row) => s + (toNum(row.amount) || 0), 0);       // =SUM(D8:D27)
    $('#batchTotal').textContent = fmt(t);
  }
  function syncRow(i) {                                      // a memo fills its Batch row; the row can still be edited by hand
    const m = state.memos[i], row = state.batch[i];
    if (!m || !row) return;
    row.company = m.soldTo; row.date = mdy(m.date); row.db = m.db; row.claim = m.claim;
    row.amount = total(m) ? String(total(m)) : row.amount;
  }
  const isoOf = mdyText => { const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec((mdyText || '').trim()); return m ? `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}` : ''; };
  // Pasted Batch rows -> one memo per row, filled from the Ageing / Account list / Report 2 / Sale Report
  async function makeMemos(only) {                        // only = one Batch row typed by hand (index): just that row is fetched again
    const last = only !== undefined ? only : state.batch.reduce((n, row, i) => (Object.values(row).some(v => String(v).trim()) ? i : n), -1);
    if (last < 0) return;
    const msg = $('#pasteMsg'); msg.textContent = 'Creating the memos…';
    for (let i = 0; i <= last; i++) if (!state.memos[i]) state.memos.push(newMemo());
    if (only === undefined) buildBatch();                   // a typed row keeps the cursor: the table is not rebuilt
    const first = only !== undefined ? only : 0, force = only !== undefined;
    const refs = state.batch.slice(first, last + 1).map(r => r.db.trim()).filter(Boolean);
    let info = {};
    try { info = await (await fetch('/api/autofill', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ refs, division: $('#bDiv').value }) })).json(); } catch (e) { info = {}; }
    const missing = [];
    for (let i = first; i <= last; i++) {
      const row = state.batch[i], m = state.memos[i], a = info[row.db.trim()];
      if (a && a.found) {                                    // the Ageing fills what the row leaves empty
        if (force || !row.claim.trim()) row.claim = a.claim || '';
        if ((force || !String(row.amount).trim()) && a.total != null) row.amount = String(a.total);
        if (force || !row.date.trim()) row.date = mdy($('#bDate').value);
      }
      m.db = row.db; m.claim = row.claim; m.date = isoOf(row.date) || m.date;
      if (a && a.found) {
        m.soldTo = a.name; m.soldAcct = a.acct; m.shipTo = a.name; m.shipAcct = a.shipAcct || m.shipAcct;
        m.salesman = a.salesman; m.terms = a.terms;
        row.company = a.name;
      } else { m.soldTo = m.soldTo || row.company; if (row.db.trim()) missing.push(row.db.trim()); }
    }
    refresh();
    msg.textContent = `${last - first + 1} ${last > first ? 'memos' : 'memo'} created.` + (missing.length ? ` Not found in the Ageing: ${missing.join(', ')}.` : ' All DB# found in the Ageing.') + ' Choose the reason code and add the item lines in each memo.';
  }
  function pasteInto(e) {
    const t = e.target;
    if (t.matches && t.matches('#priceTable input.cell')) return pastePrice(e);
    if (!t.matches('#batchTable input.cell')) return;
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!/[\t\n]/.test(text.trim())) return;                 // a single value: normal paste into the cell
    e.preventDefault();
    let rows = text.replace(/\r/g, '').split('\n').map(l => l.split('\t'));
    while (rows.length && rows[rows.length - 1].every(c => !c.trim())) rows.pop();
    if ($('#dbOnly').checked && rows.length && rows.every(r => r.length === 1 && /^\s*(DB|IN)#?\s*\d+\s*$/i.test(r[0]))) {   // only DB# pasted: they go in the DB# column, the Ageing fills the rest
      const r0 = Math.max(0, +t.dataset.r);
      rows.forEach((r, i) => {
        while (state.batch.length <= r0 + i) state.batch.push(Object.fromEntries(BCOLS.map(k => [k, ''])));
        state.batch[r0 + i].db = r[0].trim().replace(/^(DB|IN)#?\s*/i, (m, p) => p.toUpperCase() + '#');
      });
      buildBatch(); makeMemos(); return;
    }
    if (rows.length && /company/i.test(rows[0].join(' '))) rows.shift();          // copied header row
    const isNo = v => /^\d+$/.test((v || '').trim());
    const hasNo = rows.length > 0 && rows.some(r => isNo(r[0])) && rows.every(r => isNo(r[0]) || !(r[0] || '').trim()) && rows.some(r => r.length > 1);
    let r0 = +t.dataset.r, c0 = +t.dataset.c;
    if (hasNo) { rows = rows.filter(r => isNo(r[0])).map(r => r.slice(1)); c0 = 0; }   // the # column is not pasted and the total row (no #) is skipped; columns start at COMPANY
    rows.forEach((cells, i) => cells.forEach((v, j) => {
      while (state.batch.length <= r0 + i) state.batch.push(Object.fromEntries(BCOLS.map(k => [k, ''])));
      const row = state.batch[r0 + i], c = BCOLS[c0 + j];
      if (row && c) row[c] = v.trim();
    }));
    buildBatch();
    makeMemos();
  }
  // ---- PRICE LIST sheet: editable, paste Item # and Description from Excel
  const PRICE_ROWS = 131;                                    // the template's PRICE LIST rows 2-132
  state.priceRows = [];
  function buildPrice() {
    const used = state.priceRows.reduce((n, r, i) => (r.item.trim() || r.desc.trim() ? i + 1 : n), 0), want = Math.max(PRICE_ROWS, used + 10);
    while (state.priceRows.length < want) state.priceRows.push({ item: '', desc: '' });
    $('#priceTable tbody').innerHTML = state.priceRows.map((_, r) => `<tr>` + ['item', 'desc'].map((c, k) =>
      `<td><input class="cell" data-r="${r}" data-c="${k}" autocomplete="off"></td>`).join('') + '</tr>').join('');
    renderPrice(true);
  }
  function renderPrice(force) {
    document.querySelectorAll('#priceTable input.cell').forEach(el => {
      if (!force && document.activeElement === el) return;
      el.value = state.priceRows[+el.dataset.r][+el.dataset.c ? 'desc' : 'item'];
    });
  }
  function syncPrices() {                                    // the memos' description lookup (VLOOKUP) and the item list follow the sheet
    state.prices = {};
    state.priceRows.forEach(r => { if (r.item.trim()) state.prices[r.item.trim()] = r.desc.trim(); });
    $('#skuList').innerHTML = Object.entries(state.prices).map(([i, d]) => `<option value="${esc(i)}">${esc(d)}</option>`).join('');
    $('#priceMsg').textContent = `${Object.keys(state.prices).length} items.`;
    renderMemo();
  }
  function pastePrice(e) {
    const t = e.target, text = (e.clipboardData || window.clipboardData).getData('text');
    if (!/[\t\n]/.test(text.trim())) return;
    e.preventDefault();
    let rows = text.replace(/\r/g, '').split('\n').map(l => l.split('\t'));
    while (rows.length && rows[rows.length - 1].every(c => !c.trim())) rows.pop();
    if (rows.length && /item/i.test(rows[0][0] || '') && /desc/i.test(rows[0][1] || '')) rows.shift();   // copied header row
    const r0 = +t.dataset.r, c0 = +t.dataset.c;
    while (state.priceRows.length < r0 + rows.length) state.priceRows.push({ item: '', desc: '' });
    rows.forEach((cells, i) => cells.forEach((v, j) => { const c = ['item', 'desc'][c0 + j]; if (c) state.priceRows[r0 + i][c] = v.trim(); }));
    buildPrice(); syncPrices();
  }
  document.addEventListener('paste', pasteInto);
  document.addEventListener('focusout', e => {
    if (!e.target.matches('#batchTable input.cell')) return;
    const r = +e.target.dataset.r, db = (state.batch[r].db || '').trim();
    if (BCOLS[+e.target.dataset.c] === 'db' && db && db !== e.target.dataset.fetched) { e.target.dataset.fetched = db; makeMemos(r); }   // DB# typed by hand and left (Tab): fetch it from the Ageing like a paste
    renderBatch();
  });

  function renderTabs() {
    $('#memoTabs').innerHTML = state.memos.map((_, i) => `<button class="mtab ${i === state.cur ? 'on' : ''}" data-i="${i}">${i + 1}</button>`).join('')
      + '<button class="mtab add" id="addMemo" title="Add memo">+</button>';
  }

  function renderMemo() {
    const m = state.memos[state.cur], form = $('#memoForm');
    form.hidden = !m; if (!m) return;
    form.querySelectorAll('[data-f]').forEach(el => { if (document.activeElement !== el) el.value = m[el.dataset.f] ?? ''; });
    $('#reasonText').textContent = reasonOf(m) || '—';
    $('#claimText').textContent = m.claim ? `${reasonOf(m)} : ${m.claim}` : '—';
    const tb = $('#linesTable tbody');
    if (!tb.children.length) tb.innerHTML = m.lines.map((_, i) => `<tr data-l="${i}">
      <td><input list="skuList" data-k="sku"></td><td><input class="num" type="number" data-k="qty" step="1"></td>
      <td class="desc"></td><td><input class="num" type="number" step="0.01" data-k="price"></td><td class="num amt"></td></tr>`).join('');
    [...tb.children].forEach((tr, i) => {
      const l = m.lines[i];
      tr.querySelectorAll('input').forEach(el => { if (document.activeElement !== el) el.value = l[el.dataset.k]; });
      tr.querySelector('.desc').textContent = l.sku ? (state.prices[l.sku.trim()] || '#N/A') : '';   // VLOOKUP on PRICE LIST
      tr.querySelector('.desc').classList.toggle('bad', !!l.sku && !state.prices[l.sku.trim()]);
      tr.querySelector('.amt').textContent = lineAmt(l) ? fmt(lineAmt(l)) : '';
    });
    $('#mSub').textContent = fmt(subtotal(m)); $('#mTotal').textContent = fmt(total(m));
  }

  async function exportXlsx() {
    const msg = $('#exportMsg'), btn = $('#exportBtn');
    msg.textContent = 'Creating the file…'; btn.disabled = true;
    try {
      const body = { batchRows: state.batch, priceList: state.priceRows.filter(r => r.item.trim() || r.desc.trim()), batch: { no: $('#bNo').value, date: $('#bDate').value, return: $('#bReturn').value, division: $('#bDiv').value }, memos: state.memos };
      const res = await fetch('/api/export', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).detail || 'The export failed.');
      const url = URL.createObjectURL(await res.blob()), a = document.createElement('a');
      a.href = url; a.download = `Credit Memo Batch #${$('#bNo').value || 'draft'}.xlsx`; a.click(); URL.revokeObjectURL(url);
      msg.textContent = 'Downloaded.';
    } catch (err) { msg.textContent = err.message; } finally { btn.disabled = false; }
  }

  function showSheet(name) {
    document.querySelectorAll('.stab').forEach(b => b.classList.toggle('on', b.dataset.sheet === name && (name !== 'Batch' || state.cur < 0)));
    if (name !== 'Batch') document.querySelectorAll('.mtab.on').forEach(b => b.classList.remove('on'));
    const ref = name !== 'Batch', price = name === 'PRICE LIST';
    $('#batchSheets').hidden = ref; $('#priceCard').hidden = !price; $('#refCard').hidden = !ref || price;
    if (!ref || price) return;
    const head = $('#refTable thead'), body = $('#refTable tbody'), note = $('#refNote');
    if (name === 'REASON' || name === 'PRICE LIST') {
      const rows = name === 'REASON' ? state.reasons.map(r => [r.code, r.description]) : Object.entries(state.prices);
      head.innerHTML = name === 'REASON' ? '<tr><th>Code</th><th>Description</th></tr>' : '<tr><th>Item #</th><th>Description</th></tr>';
      body.innerHTML = rows.map(r => `<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td></tr>`).join('') || '<tr><td colspan="2" class="muted">Empty.</td></tr>';
      note.textContent = `${rows.length} rows. They are written to the ${name} sheet when you export.`;
    } else {
      head.innerHTML = ''; body.innerHTML = '<tr><td class="muted">Empty, like in the template.</td></tr>'; note.textContent = 'This sheet stays empty.';
    }
  }

  const refresh = () => { renderTabs(); renderBatch(); renderMemo(); document.querySelector('.stab[data-sheet="Batch"]').classList.toggle('on', state.cur < 0); $('#batchCard').hidden = state.cur >= 0; $('#memoCard').hidden = state.cur < 0; };   // a memo shows in its own sheet tab
  const select = i => { state.cur = i; refresh(); };

  document.addEventListener('click', e => {
    if (e.target.dataset.sheet) { if (e.target.dataset.sheet === 'Batch') { state.cur = -1; refresh(); } showSheet(e.target.dataset.sheet); }
    else if (e.target.id === 'exportBtn') exportXlsx();
    else if (e.target.id === 'clearPrice' && confirm('Clear all values of the PRICE LIST?')) { state.priceRows = []; buildPrice(); syncPrices(); }
    else if (e.target.id === 'clearBatch' && confirm('Clear the Batch sheet and delete all memos?')) {
      state.memos = []; state.cur = -1;
      state.batch = Array.from({ length: BROWS }, () => Object.fromEntries(BCOLS.map(c => [c, ''])));
      $('#pasteMsg').textContent = ''; $('#exportMsg').textContent = ''; buildBatch(); refresh();
    }
    else if (e.target.id === 'addMemo') { state.memos.push(newMemo()); while (state.batch.length < state.memos.length) state.batch.push(Object.fromEntries(BCOLS.map(k => [k, '']))); buildBatch(); select(state.memos.length - 1); showSheet('Batch'); }
    else if (e.target.classList.contains('mtab')) { select(+e.target.dataset.i); showSheet('Batch'); }
    else if (e.target.id === 'delMemo' && confirm('Delete this memo?')) { state.memos.splice(state.cur, 1); select(Math.min(state.cur, state.memos.length - 1)); }
  });
  document.addEventListener('change', e => {                 // reason chosen on the Batch sheet: every memo gets it, each memo can be changed by hand
    if (e.target.id !== 'reasonAll') return;
    state.memos.forEach((x, i) => { x.code = e.target.value; syncRow(i); });
    refresh();
  });
  document.addEventListener('input', e => {
    const t = e.target, m = state.memos[state.cur];
    if (t.matches('#priceTable input.cell')) { state.priceRows[+t.dataset.r][+t.dataset.c ? 'desc' : 'item'] = t.value; syncPrices(); return; }
    if (t.matches('#batchTable input.cell')) { state.batch[+t.dataset.r][BCOLS[+t.dataset.c]] = t.value; renderBatch(); return; }
    if (!m) return;
    if (t.dataset.f === 'code') { state.memos.forEach((x, i) => { x.code = t.value; syncRow(i); }); }   // a new reason code goes to every memo; change a memo by hand if it differs
    else if (t.dataset.f) m[t.dataset.f] = t.value;
    else if (t.dataset.k) m.lines[+t.closest('tr').dataset.l][t.dataset.k] = t.value;
    else return;
    syncRow(state.cur); renderBatch(); renderMemo();
  });

  Promise.all([fetch('/api/price-list').then(r => r.json()), fetch('/api/reasons').then(r => r.json())]).then(([p, r]) => {
    state.priceRows = [];                                    // the PRICE LIST starts empty: the user pastes it
    buildPrice(); syncPrices();
    state.reasons = r;
    $('#reasonSel').innerHTML = $('#reasonAll').innerHTML = '<option value="">—</option>' + r.map(x => `<option value="${esc(x.code)}"${x.blocked ? ' disabled' : ''}>${esc(x.code)} · ${esc(x.description)}${x.blocked ? ' (not allowed)' : ''}</option>`).join('');
    $('#bDate').value = new Date().toISOString().slice(0, 10);
    buildBatch();
    refresh();
  });
})();
