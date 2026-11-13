/* DataBase Room: one main tab per file type; sub tabs Data / More; per-file Re-import and Delete. */
(function () {
  const KINDS = window.IMPORT_KINDS, $q = s => document.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const size = n => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
  const kindOf = k => KINDS.find(x => x.key === k);

  let files = [], kind = KINDS[0].key, sub = 'data', timer = null;
  const activeId = {};       // kind -> file id shown in the Data tab
  const local = {};          // file id -> progress text for work done in this browser (reading a PDF)
  let shown = { id: null, key: '' };   // which file/state the viewer currently displays

  const byKind = k => files.filter(f => f.kind === k);                       // newest first (API order)
  const keyOf = f => f ? f.id + '|' + f.status + '|' + f.rows_imported : '';
  const current = () => {
    const list = byKind(kind);
    return list.find(f => f.id === activeId[kind]) || list.find(f => f.status === 'imported') || list[0] || null;
  };

  /* ------------------------------------------------------------------ data */
  async function loadFiles() {
    files = await (await fetch('/api/files')).json();
    render();
    clearTimeout(timer);
    if (files.some(f => f.status === 'importing' || local[f.id])) timer = setTimeout(loadFiles, 1500);
  }

  /* ------------------------------------------------------------------ rendering */
  function render() {
    renderTabs(); renderSub();
    if (sub === 'data') renderData(); else renderHistory();
  }

  function renderTabs() {
    document.querySelectorAll('.btab').forEach(b => {
      const k = b.dataset.kind, list = byKind(k);
      b.classList.toggle('on', k === kind); b.setAttribute('aria-selected', k === kind);
      b.querySelector('.bt-n').textContent = list.length || '';
    });
  }

  function renderSub() {
    document.querySelectorAll('.subtabs button').forEach(b => b.classList.toggle('on', b.dataset.sub === sub));
    $q('#viewData').hidden = sub !== 'data'; $q('#viewMore').hidden = sub !== 'more';
    $q('#moreCount').textContent = byKind(kind).length ? '(' + byKind(kind).length + ')' : '';
    const k = kindOf(kind);
    $q('#kindInfo').textContent = k.info;
    $q('#fileInput').accept = k.accept;
    $q('#importBtn').textContent = 'Import ' + k.label.toLowerCase();
  }

  function actions(f) {
    const label = f.kind === 'ar_pdf' ? 'Re-read PDF' : 'Re-import';
    return (f.kind === 'reasons' && f.status === 'imported' ? `<button class="btn sm" data-act="editcodes" data-id="${f.id}">Edit codes</button>` : '') +
           `<button class="btn ghost sm" data-act="reimport" data-id="${f.id}">${label}</button>` +
           `<button class="btn danger sm" data-act="delete" data-id="${f.id}">Delete</button>`;
  }

  function renderData() {
    const f = current(), bar = $q('#fileBar'), empty = $q('#emptyMsg'), mount = $q('#dvmount');
    if (!f) {
      shown = { id: null, key: '' }; window.DV.clear();
      bar.hidden = true; mount.hidden = true; empty.hidden = false;
      empty.innerHTML = `No ${esc(kindOf(kind).label)} file yet.<br><button class="btn" data-act="pick">Import file</button>`;
      return;
    }
    activeId[kind] = f.id;
    bar.hidden = false; empty.hidden = true;
    $q('#viewerTitle').textContent = f.filename;
    const note = local[f.id] || f.note || '';
    $q('#viewerStatus').textContent = [f.status, f.rows_imported ? f.rows_imported.toLocaleString() + ' rows' : '', f.uploaded_at].filter(Boolean).join(' · ') +
      (note && f.status !== 'imported' ? ' — ' + note : '');
    $q('#fbActions').innerHTML = actions(f);
    if (f.status === 'imported') {
      mount.hidden = false;
      if (shown.id !== f.id || shown.key !== keyOf(f)) {          // rebuild only when the file or its data changed
        shown = { id: f.id, key: keyOf(f) };
        window.DV.clear(); mount.innerHTML = '<div class="empty">Loading data…</div>';
        window.DV.open(f.id);
      }
    } else {
      shown = { id: null, key: '' }; window.DV.clear(); mount.hidden = true; empty.hidden = false;
      const wait = f.status === 'importing' || local[f.id];
      empty.innerHTML = wait ? esc(note || 'Importing…')
        : f.status === 'failed' ? 'Import failed: ' + esc(note)
        : f.kind === 'ar_pdf' ? 'The PDF is saved but not read yet.<br><button class="btn" data-act="reimport" data-id="' + f.id + '">Read PDF</button>'
        : 'Waiting to import.<br><button class="btn" data-act="reimport" data-id="' + f.id + '">Import now</button>';
    }
  }

  let lastHist = '';
  function renderHistory() {
    const list = byKind(kind), cur = current();
    const html = list.map(r => {
      const note = local[r.id] || r.note || '';
      return `<tr data-id="${r.id}" class="${cur && r.id === cur.id ? 'cur' : ''}"><td>${esc(r.filename)}</td><td class="num">${size(r.size_bytes)}</td>
        <td><span class="pill ${esc(r.status)}">${esc(r.status)}</span></td><td class="num">${r.rows_imported ? r.rows_imported.toLocaleString() : ''}</td>
        <td class="note">${esc(note)}</td><td>${esc(r.uploaded_at)}</td>
        <td><div class="rowacts"><button class="btn ghost sm" data-act="view" data-id="${r.id}">View</button>${actions(r)}</div></td></tr>`;
    }).join('') || '<tr><td colspan="7" class="muted">No uploads yet for this file type.</td></tr>';
    if (html === lastHist) return;                           // unchanged: leave the DOM alone
    lastHist = html; $q('#histTable tbody').innerHTML = html;
  }

  /* ------------------------------------------------------------------ navigation */
  function go(k, s) {
    kind = k; sub = s;
    history.replaceState(null, '', '#' + k + (s === 'more' ? '/more' : ''));
    lastHist = ''; render();
  }
  $q('#mainTabs').addEventListener('click', e => { const b = e.target.closest('.btab'); if (b) go(b.dataset.kind, 'data'); });
  $q('.subtabs').addEventListener('click', e => { const b = e.target.closest('button'); if (b) go(kind, b.dataset.sub); });
  function fromHash() {
    const [k, s] = location.hash.slice(1).split('/');
    if (kindOf(k)) { kind = k; sub = s === 'more' ? 'more' : 'data'; }
  }

  /* ------------------------------------------------------------------ actions */
  document.addEventListener('click', async e => {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const id = +b.dataset.id, f = files.find(x => x.id === id);
    if (b.dataset.act === 'pick') $q('#fileInput').click();
    else if (b.dataset.act === 'editcodes') openEditor(id);
    else if (b.dataset.act === 'retry') { shown = { id: null, key: '' }; render(); }
    else if (b.dataset.act === 'view') { activeId[kind] = id; go(kind, 'data'); }
    else if (b.dataset.act === 'reimport' && f) {
      if (f.kind === 'ar_pdf') readPdf(id);
      else { await fetch('/api/files/' + id + '/import', { method: 'POST' }); await loadFiles(); }
    } else if (b.dataset.act === 'delete' && f) {
      if (!confirm('Delete "' + f.filename + '"?\n\nIts ' + (f.rows_imported || 0).toLocaleString() +
        ' rows are removed from the database. The file itself is kept in data/uploads/_deleted.')) return;
      const r = await fetch('/api/files/' + id, { method: 'DELETE' });
      if (!r.ok) { alert((await r.json().catch(() => ({}))).detail || 'Delete failed.'); return; }
      if (activeId[f.kind] === id) delete activeId[f.kind];
      shown = { id: null, key: '' }; lastHist = '';
      await loadFiles();
    }
  });

  /* ------------------------------------------------------------------ reason codes: edit by hand */
  async function openEditor(id) {
    const rows = await (await fetch('/api/files/' + id + '/reasons')).json();
    const dlg = document.createElement('dialog'); dlg.className = 'rc-editor';
    dlg.innerHTML = `<h3>Edit reason codes</h3><input type="search" class="rc-find" placeholder="Find a code or text">
      <div class="rc-scroll"><table class="grid"><thead><tr><th>Reason Code</th><th>Description</th><th></th></tr></thead><tbody></tbody></table></div>
      <p class="muted rc-msg" role="status"></p>
      <div class="actions"><button class="btn ghost sm" data-e="add">Add code</button><span style="flex:1"></span>
      <button class="btn ghost sm" data-e="cancel">Cancel</button><button class="btn sm" data-e="save">Save</button></div>`;
    const body = dlg.querySelector('tbody'), msg = dlg.querySelector('.rc-msg');
    const addRow = (r, top) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><input class="c" value="${esc(r.code)}" size="6"></td><td><input class="d" value="${esc(r.description)}" style="width:100%"></td>
        <td><button class="btn danger sm" data-e="del">Remove</button></td>`;
      top ? body.prepend(tr) : body.append(tr); return tr;
    };
    rows.forEach(r => addRow(r));
    dlg.addEventListener('input', e => {
      if (!e.target.classList.contains('rc-find')) return;
      const q = e.target.value.toLowerCase();
      body.querySelectorAll('tr').forEach(tr => { tr.hidden = !!q && !(tr.querySelector('.c').value + ' ' + tr.querySelector('.d').value).toLowerCase().includes(q); });
    });
    dlg.addEventListener('click', async e => {
      const a = e.target.closest('[data-e]'); if (!a) return;
      if (a.dataset.e === 'del') a.closest('tr').remove();
      else if (a.dataset.e === 'add') addRow({ code: '', description: '' }, true).querySelector('.c').focus();
      else if (a.dataset.e === 'cancel') { dlg.close(); dlg.remove(); }
      else if (a.dataset.e === 'save') {
        const data = [...body.querySelectorAll('tr')].map(tr => ({ code: tr.querySelector('.c').value, description: tr.querySelector('.d').value }));
        msg.textContent = 'Saving…';
        const r = await fetch('/api/files/' + id + '/reasons', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
        if (!r.ok) { msg.textContent = (await r.json().catch(() => ({}))).detail || 'Save failed.'; return; }
        dlg.close(); dlg.remove(); shown = { id: null, key: '' }; lastHist = ''; await loadFiles();
      }
    });
    document.body.append(dlg); dlg.showModal();
  }

  /* ------------------------------------------------------------------ A/R PDF: read with pdf.js in the browser, post the parsed rows */
  async function readPdf(id) {
    const f = files.find(x => x.id === id);
    local[id] = 'Reading PDF…'; render();
    const setNote = t => { local[id] = t; const c = current(); if (c && c.id === id) { const el = $q('#viewerStatus'); if (el) el.textContent = t; const em = $q('#emptyMsg'); if (em && !em.hidden) em.textContent = t; } };
    try {
      pdfjsLib.GlobalWorkerOptions.workerSrc = '/static/vendor/pdf.worker.js';
      const buf = await (await fetch('/api/files/' + id + '/download')).arrayBuffer();
      const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf), verbosity: 0 }).promise;
      const pages = [];
      for (let p = 1; p <= doc.numPages; p++) {
        if (p % 10 === 0 || p === doc.numPages) { setNote(`Reading page ${p} of ${doc.numPages}`); await new Promise(r => setTimeout(r)); }
        const tc = await (await doc.getPage(p)).getTextContent();
        pages.push(tc.items.filter(i => i.str.trim()).map(i => ({ str: i.str.trim(), x: i.transform[4], y: i.transform[5], w: i.width })));
      }
      setNote('Saving rows…');
      const res = window.parseReport(pages, f.filename);
      const r = await fetch('/api/files/' + id + '/rows', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ detail: res.detail, aging: res.aging, customers: res.customers, warnings: res.warnings }) });
      if (!r.ok) throw new Error((await r.json().catch(() => ({}))).detail || 'Save failed');
    } catch (e) { alert('Could not read the PDF: ' + e.message); }
    delete local[id];
    shown = { id: null, key: '' };
    await loadFiles();
  }

  /* ------------------------------------------------------------------ upload (Import button, or drop a file on the card) */
  async function send(file) {
    const k = kindOf(kind), ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!k.accept.split(',').includes(ext)) { alert(k.label + ' accepts ' + k.accept + ' files only.'); return; }
    const btn = $q('#importBtn'), old = btn.textContent; btn.disabled = true; btn.textContent = 'Uploading…';
    const fd = new FormData(); fd.append('file', file);
    const r = await fetch('/upload/' + kind, { method: 'POST', body: fd });
    btn.disabled = false; btn.textContent = old;
    if (!r.ok) { alert((await r.json().catch(() => ({}))).detail || 'Upload failed.'); return; }
    const j = await r.json();
    activeId[kind] = j.id; lastHist = '';
    await loadFiles(); go(kind, 'data');
    if (j.kind === 'ar_pdf' && ext === '.pdf') readPdf(j.id);
  }
  $q('#importBtn').addEventListener('click', () => $q('#fileInput').click());
  $q('#fileInput').addEventListener('change', e => { if (e.target.files[0]) send(e.target.files[0]); e.target.value = ''; });
  const card = $q('#impCard'), veil = $q('#dropVeil');
  ['dragenter', 'dragover'].forEach(ev => card.addEventListener(ev, e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { e.preventDefault(); veil.hidden = false; } }));
  card.addEventListener('dragleave', e => { if (!card.contains(e.relatedTarget)) veil.hidden = true; });
  card.addEventListener('drop', e => { e.preventDefault(); veil.hidden = true; if (e.dataTransfer.files[0]) send(e.dataTransfer.files[0]); });

  fromHash();
  window.addEventListener('hashchange', () => { fromHash(); lastHist = ''; render(); });
  loadFiles();
})();
