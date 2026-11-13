/* Data viewer: builds the Report2Excel-style tables (DataTables + column filters + advanced search) for one uploaded file.
   Engine code (search bar, filters, pills, highlight) is viewer_engine.js, extracted from the v7 tool. */
(function () {
  const $ = window.jQuery;
  const MARKUP = '<div id="dvload" hidden><span class="dvbar"><i></i></span><span class="dvtxt"></span></div><div id="dockslot"><div id="dock"></div><div id="btnslot"></div></div><div class="dvtabs" id="tabs"></div><div class="dvpanes" id="panes"></div><div id="fp"></div>';
  let TABLES = [];

  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const fmtDate = v => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || ''); return m ? m[2] + '/' + m[3] + '/' + m[1] : (v || ''); };
  const fmtNum = (v, d) => {
    const x = Number(v); if (d === undefined) return x.toLocaleString('en-US');
    const p = Math.abs(x).toFixed(d).split('.');
    return (x < 0 ? '-' : '') + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (p[1] ? '.' + p[1] : '');
  };
  const showVal = (c, v) => {
    if (c.type === 'd') return fmtDate(v);
    if (c.type === 'n') return v === null || v === undefined || v === '' ? '' : c.plain ? String(v) : (c.cur ? '$' : '') + fmtNum(v, c.dec);
    return v === null || v === undefined ? '' : String(v);
  };
  const d2iso = s => { const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s); return m ? m[3] + m[1] + m[2] : s; };
  const WRAP = ['Name', 'Customer', 'Claim#', 'Order#', 'Comment', 'Description', 'Item Name', 'Claim #', 'Customer Claim #'];


  /* ---- downloads: DataTables Buttons (copy, CSV, XLSX, PDF, print, column visibility) + an XLS (Excel 2003 XML) button.
          Every export uses the rows as shown: filtered and sorted. ---- */
  const stamp = () => new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' }).replace(/\//g, '-');
  const safeName = s => String(s).replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]+/g, '_').trim();
  const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const BIG = 3000;     // PDF / print of more rows than this is slow: ask first
  function saveBlob(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function xlsOf(t, rows) {   // Excel 2003 XML spreadsheet; Excel and LibreOffice open it as .xls
    const num = c => c.type === 'n' && !c.plain ? (c.dec ? ' ss:StyleID="m"' : ' ss:StyleID="i"') : '';
    const head = '<Row>' + t.cols.map(c => '<Cell ss:StyleID="h"><Data ss:Type="String">' + xmlEsc(c.label) + '</Data></Cell>').join('') + '</Row>';
    const body = rows.map(r => '<Row>' + t.cols.map((c, i) => {
      const v = r[i];
      if (v === null || v === undefined || v === '') return '<Cell/>';
      if (c.type === 'n') return '<Cell' + num(c) + '><Data ss:Type="Number">' + v + '</Data></Cell>';
      if (c.type === 'd' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return '<Cell ss:StyleID="d"><Data ss:Type="DateTime">' + v + 'T00:00:00.000</Data></Cell>';
      return '<Cell><Data ss:Type="String">' + xmlEsc(v) + '</Data></Cell>';
    }).join('') + '</Row>').join('');
    return '<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
      '<Styles><Style ss:ID="h"><Font ss:Bold="1"/><Interior ss:Color="#E6F4EA" ss:Pattern="Solid"/></Style><Style ss:ID="d"><NumberFormat ss:Format="mm/dd/yyyy"/></Style>' +
      '<Style ss:ID="m"><NumberFormat ss:Format="#,##0.00"/></Style><Style ss:ID="i"><NumberFormat ss:Format="#,##0"/></Style></Styles>' +
      '<Worksheet ss:Name="' + xmlEsc(t.title.slice(0, 31)) + '"><Table>' + head + body + '</Table></Worksheet></Workbook>';
  }
  function buttonSet(t) {
    const heading = () => document.getElementById('viewerTitle').textContent + ' - ' + t.title;
    const base = () => safeName(document.getElementById('viewerTitle').textContent) + ' - ' + t.title + ' ' + stamp();
    const exportOptions = { orthogonal: 'export', modifier: { search: 'applied', order: 'applied' }, columns: ':visible' };
    const guarded = name => function (e, dt, node, config, cb) {
      const n = dt.rows({ search: 'applied' }).count();
      if (n > BIG && !confirm(n.toLocaleString() + ' rows will be processed. This can take a while. Continue?')) return;
      $.fn.dataTable.ext.buttons[name].action.call(this, e, dt, node, config, cb);
    };
    return [
      { extend: 'copyHtml5', text: 'Copy', footer: true, exportOptions },
      { extend: 'csvHtml5', text: 'CSV', filename: base, footer: true, bom: true, exportOptions },
      { extend: 'excelHtml5', text: 'XLSX', filename: base, title: null, sheetName: t.title.slice(0, 31), footer: true, exportOptions },
      { text: 'XLS', className: 'buttons-xls', titleAttr: 'Excel 2003 .xls (rows as shown)', action: function (e, dt) {
        const rows = dt.rows({ search: 'applied', order: 'applied' }).data().toArray();
        if (!rows.length) { alert('No rows to download.'); return; }
        saveBlob(new Blob([xlsOf(t, rows)], { type: 'application/vnd.ms-excel' }), base() + '.xls');
      } },
      { extend: 'pdfHtml5', text: 'PDF', filename: base, title: heading, orientation: 'landscape', pageSize: 'A3', footer: true, exportOptions,
        action: guarded('pdfHtml5'), customize: doc => { doc.defaultStyle.fontSize = 6; doc.styles.tableHeader.fontSize = 7; doc.pageMargins = [14, 14, 14, 14]; } },
      { extend: 'print', text: 'Print', title: heading, footer: true, exportOptions, action: guarded('print') },
      { extend: 'colvis', text: 'Columns' }
    ];
  }

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) { navigator.clipboard.writeText(text).catch(() => legacyCopy(text)); return; }
    legacyCopy(text);
  }
  function legacyCopy(text) {
    const ta = document.createElement('textarea'); ta.value = text; ta.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); } catch (e) { /* nothing more to try */ } ta.remove();
  }
  let toastTimer;
  function toast(msg) {
    let el = document.getElementById('dvToast');
    if (!el) { el = document.createElement('div'); el.id = 'dvToast'; document.body.appendChild(el); }
    el.textContent = msg; el.classList.add('show');
    clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 1800);
  }

  function destroy() {
    TABLES.forEach(t => { try { t.dt.destroy(); } catch (e) { /* already gone */ } });
    TABLES = [];
    $.fn.dataTable.ext.search.length = 0;      // row filters of the previous file must not apply to the new tables
  }

  function build(data) {
    destroy();
    const mount = $('#dvmount').html(MARKUP).addClass('hasdock');
    const filters = {};
    TABLES = data.tables.map(t => ({ id: t.id, title: t.title, name: t.title, rows: t.rows, total: t.total,
      cols: t.cols.map((c, i) => Object.assign({}, c, { key: i })) }));
    TABLES.forEach((t, ti) => {
      filters[t.id] = {};
      $('#tabs').append('<button type="button" class="tab' + (ti ? '' : ' on') + '" data-t="' + t.id + '">' + esc(t.title) + ' (' + (t.total != null ? t.total : t.rows.length).toLocaleString() + ')</button>');
      $('#panes').append('<div class="pane' + (ti ? '' : ' on') + '" id="pane-' + t.id + '"><table id="tbl-' + t.id + '" class="display stripe hover order-column" style="width:100%">' +
        '<thead><tr>' + t.cols.map(c => '<th class="' + (c.type === 'n' ? 'n' : '') + '">' + esc(c.label) + '</th>').join('') + '</tr></thead>' +
        '<tfoot><tr>' + t.cols.map((c, i) => '<th class="n">' + (i === 0 ? '<span style="float:left">Total (filtered)</span>' : '') + '</th>').join('') + '</tr></tfoot></table></div>');
      t.dt = $('#tbl-' + t.id).DataTable({
        data: t.rows, pageLength: 100, lengthMenu: [[10, 25, 50, 100, 500, -1], [10, 25, 50, 100, 500, 'All']],
        autoWidth: false, ordering: { handler: false, indicators: false }, order: [],
        columns: t.cols.map((c, i) => ({
          data: i, className: (c.type === 'n' ? 'n' : '') + (WRAP.includes(c.label) ? ' wr' : ''),
          // speed: our own search engine does the matching, so DataTables must not build its text index or scan every row to guess the type
          searchable: false, type: c.type === 'n' ? 'num' : 'string',
          render: (d, type) => type === 'export' ? (c.type === 'd' ? fmtDate(d) : (d === null || d === undefined ? '' : d))
            : (type === 'display' || type === 'filter') ? showVal(c, [].concat(d)[0]) : (d === null || d === undefined ? '' : d)
        })),
        footerCallback: function () {
          const api = this.api(), rows = api.rows({ search: 'applied' }).data().toArray();
          t.cols.forEach((c, i) => { if (c.sum) $(api.column(i).footer()).html((c.cur ? '$' : '') + fmtNum(rows.reduce((s, r) => s + (r[i] || 0), 0), c.dec)); });
        }
      });
    });
    TABLES.forEach(t => {
      new $.fn.dataTable.Buttons(t.dt, { buttons: buttonSet(t) });
      $('#btnslot').append(t.dt.buttons().container().toggle(t === TABLES[0]));   // one button group per table, only the active table's is shown
    });
    window.R2Engine.setupUI($, TABLES, (c, r) => showVal(c, r[c.key]), esc, filters, d2iso);
    TABLES.forEach(t => t.dt.on('draw.dt', () => window.R2Engine.hlTable(t.dt)));
    $('#tabs').on('click', '.tab', function () {
      $('#tabs .tab').removeClass('on'); $(this).addClass('on');
      $('#panes .pane').removeClass('on'); $('#pane-' + this.dataset.t).addClass('on');
      TABLES.find(t => t.id === this.dataset.t).dt.columns.adjust();
      TABLES.forEach(t => t.dt.buttons().container().toggle(t.id === this.dataset.t));
      $('#dock .sbar:visible .sq').trigger('focus');
    });
    // double-click any cell to copy its value
    $('#panes').on('dblclick', 'tbody td', function () {
      const text = this.textContent.trim();
      if (!text) return;
      copyText(text);
      const sel = window.getSelection(); if (sel) sel.removeAllRanges();
      $(this).addClass('copied'); setTimeout(() => $(this).removeClass('copied'), 600);
      toast('Copied: ' + (text.length > 60 ? text.slice(0, 60) + '…' : text));
    });
    // Reserve the search box's space: when it floats to the bottom-right (#dock.low) the table must not jump up.
    const slot = document.getElementById('dockslot'), dock = document.getElementById('dock');
    setTimeout(() => { if (!dock.classList.contains('low')) slot.style.minHeight = slot.offsetHeight + 'px'; }, 0);
    TABLES.forEach(t => { t.dt.columns.adjust(); t.fitSoon && t.fitSoon(); });
  }

  /* ---- loading pipeline: first rows at once, the rest streams in behind a progress bar ---- */
  const FIRST = 100;
  let loadSeq = 0;       // only the most recent open() may touch the screen; older, slower responses are dropped

  function setLoading(on, text, frac) {
    $('#dvload').prop('hidden', !on);
    $('#btnslot').toggleClass('wait', on);
    if (on) { $('#dvload .dvtxt').text(text || ''); $('#dvload .dvbar i').css('width', frac == null ? '30%' : Math.round(frac * 100) + '%').toggleClass('indet', frac == null); }
  }

  async function fetchFull(url, seq, onProgress) {
    const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), 180000);
    try {
      const r = await fetch(url, { signal: ctl.signal });
      if (!r.ok) throw new Error('The server answered ' + r.status + '.');
      const raw = +r.headers.get('X-Raw-Size') || 0, reader = r.body.getReader(), chunks = [];
      let got = 0;
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        if (seq !== loadSeq) { ctl.abort(); return null; }
        chunks.push(value); got += value.length; onProgress(raw ? Math.min(got / raw, 0.99) : null);
      }
      const buf = new Uint8Array(got); let o = 0;
      for (const c of chunks) { buf.set(c, o); o += c.length; }
      return JSON.parse(new TextDecoder().decode(buf));
    } catch (e) {
      if (e.name === 'AbortError') throw new Error('Loading took longer than 3 minutes.');
      throw e;
    } finally { clearTimeout(timer); }
  }

  async function fillRows(full) {
    await new Promise(r => setTimeout(r, 30));
    full.tables.forEach(ft => {
      const t = TABLES.find(x => x.id === ft.id); if (!t) return;
      t.rows = ft.rows;
      t.dt.clear().rows.add(ft.rows).draw(false);
    });
    TABLES.forEach(t => t.fitSoon && t.fitSoon());
  }

  function showError(msg) {
    destroy();
    $('#dvmount').html('<div class="empty">Could not load the data.<br><span class="muted">' + esc(msg) + '</span><br><button class="btn" data-act="retry">Try again</button></div>');
  }

  window.DV = {
    /** Shows one stored file. Resolves with the row count, or null when it failed or a newer open() replaced this one. */
    async open(fileId) {
      const seq = ++loadSeq;
      try {
        const r = await fetch('/api/files/' + fileId + '/data?limit=' + FIRST);
        if (!r.ok) throw new Error('The server answered ' + r.status + '.');
        const first = await r.json();
        if (seq !== loadSeq) return null;
        build(first);                                         // instant first screen
        const total = first.tables.reduce((n, t) => n + t.total, 0);
        if (first.tables.every(t => t.total <= t.rows.length)) return total;      // small file: already complete
        setLoading(true, 'Loading all ' + total.toLocaleString() + ' rows…', null);
        const full = await fetchFull('/api/files/' + fileId + '/data', seq, f => { if (seq === loadSeq) setLoading(true, 'Loading all ' + total.toLocaleString() + ' rows… ' + (f == null ? '' : Math.round(f * 100) + '%'), f); });
        if (!full || seq !== loadSeq) return null;
        await fillRows(full);
        setLoading(false);
        return total;
      } catch (e) {
        if (seq === loadSeq) showError(e.message || String(e));
        return null;
      }
    },
    clear() { loadSeq++; destroy(); $('#dvmount').empty(); }
  };
})();
