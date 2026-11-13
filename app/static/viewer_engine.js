/* Extracted from the Report2Excel v7 tool: advanced search bar, column filters, pills, highlight. */
(function () {
  /* Shared UI: advanced search bar, header filter funnel + pills. Self-contained so the HTML export can reuse it. */
  function setupUI($, T, show, esc, filters, d2iso) {
    const escRx = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // cache each row's display text so searching never re-formats 35k rows x 19 columns
    const CACHE = new WeakMap();
    const txt = (t, row, i) => { let a = CACHE.get(row); if (!a) { a = new Array(t.cols.length); CACHE.set(row, a); } return a[i] !== undefined ? a[i] : (a[i] = show(t.cols[i], row)); };
    let FS = {};
    try { FS = JSON.parse(localStorage.getItem('r2e-fontsize6') || '{}'); } catch (e) {}
    const fsStyle = $('<style></style>').appendTo(document.head);
    const BASE = 11;
    function applyFS() {
      let css = '';
      T.forEach(t => { const m = FS[t.id] || {}; for (const i in m) css += '#tbl-' + t.id + '>tbody>tr>td:nth-child(' + (+i + 1) + '){font-size:' + m[i] + 'px !important}'; });
      fsStyle.text(css);
      try { localStorage.setItem('r2e-fontsize6', JSON.stringify(FS)); } catch (e) {}
    }
    applyFS();
    // columns that can be collapsed to a dot (click the dot to show/hide); Cust Code starts collapsed
    let CL = null;
    try { CL = JSON.parse(localStorage.getItem('r2e-collapsed8') || 'null'); } catch (e) {}
    if (!CL) { CL = {}; T.forEach(t => t.cols.forEach((c, i) => { if (['Pg#', 'Cust Code'].includes(c.label)) (CL[t.id] = CL[t.id] || {})[i] = true; })); }
    const fsLess = $('<style></style>').appendTo(document.head);
    function applyCL() {
      let css = '';
      T.forEach(t => t.cols.forEach((c, i) => {
        const on = !!(CL[t.id] || {})[i], th = $(t.dt.column(i).header());
        th.toggleClass('col-c', on);
        th.find('.dotbtn').html(on ? '&#9679;' : '&#9675;').attr('title', on ? 'Show ' + c.label : 'Collapse ' + c.label + ' to a dot');
        if (on) css += '#tbl-' + t.id + '>tbody>tr>td:nth-child(' + (i + 1) + '),#tbl-' + t.id + '>tfoot>tr>th:nth-child(' + (i + 1) + '){font-size:0 !important;width:14px;max-width:14px;min-width:0;padding:0 !important;overflow:hidden}';
      }));
      fsLess.text(css);
      try { localStorage.setItem('r2e-collapsed8', JSON.stringify(CL)); } catch (e) {}
    }
    const FUNNEL = '<svg viewBox="0 0 16 16"><path d="M1 2h14l-5.5 6.5V14l-3-1.5V8.5z"/></svg>';
    function tokenize(q, mode) {
      if (mode === 'phrase') return q.trim() ? [{ text: q.trim(), neg: false }] : [];
      const out = [], re = /(-?)"([^"]*)"|(\S+)/g; let m;
      while ((m = re.exec(q))) {
        if (m[3] !== undefined) { const neg = m[3].length > 1 && m[3][0] === '-'; out.push({ text: neg ? m[3].slice(1) : m[3], neg }); }
        else if (m[2]) out.push({ text: m[2], neg: m[1] === '-', quoted: true });
      }
      return out;
    }
    function compileSearch(t) {
      const s = t.sq; t.match = null; t.hl = null; t.bad = false;
      if (!s.q.trim()) return;
      const toks = tokenize(s.q, s.mode), pos = [], neg = [], hlSrc = [];
      for (const k of toks) {
        const o = {};
        const plain = !s.rx && !k.quoted && s.mode !== 'phrase';
        const nm = plain && /^(>=|<=|<>|!=|>|<|=)\s*(-?[\d,]*\.?\d+)$/.exec(k.text);
        const rg = plain && /^(-?[\d,]*\.?\d+)\.\.(-?[\d,]*\.?\d+)$/.exec(k.text);
        if (nm) { const v = parseFloat(nm[2].replace(/,/g, '')), op = nm[1]; o.num = x => op === '>' ? x > v : op === '<' ? x < v : op === '>=' ? x >= v : op === '<=' ? x <= v : op === '=' ? x === v : x !== v; }
        else if (rg) { const a = parseFloat(rg[1].replace(/,/g, '')), b = parseFloat(rg[2].replace(/,/g, '')); o.num = x => x >= Math.min(a, b) && x <= Math.max(a, b); }
        else {
          let src = s.rx ? k.text : escRx(k.text).replace(/\\\*/g, '.*').replace(/\\\?/g, '.');
          if (s.whole) src = '^(?:' + src + ')$';
          try { o.re = new RegExp(src, s.cs ? '' : 'i'); } catch (e) { t.bad = true; return; }
          if (!k.neg) hlSrc.push('(?:' + src + ')');
        }
        (k.neg ? neg : pos).push(o);
      }
      const cols = s.col >= 0 ? [s.col] : t.cols.map((c, i) => i);
      const test = (o, row, texts) => o.num ? cols.some(i => t.cols[i].type === 'n' && typeof row[t.cols[i].key] === 'number' && o.num(row[t.cols[i].key])) : texts.some(x => o.re.test(x));
      t.match = row => {
        const texts = cols.map(i => txt(t, row, i));
        if (neg.some(o => test(o, row, texts))) return false;
        if (!pos.length) return true;
        return s.mode === 'any' ? pos.some(o => test(o, row, texts)) : pos.every(o => test(o, row, texts));
      };
      t.hl = hlSrc.length ? { src: hlSrc.join('|'), cs: s.cs, col: s.col } : null;
    }
    // row filter: header checklists + advanced search (column text filters use DataTables' own column search)
    $.fn.dataTable.ext.search.push(function (settings, data, idx, row) {
      const t = T.find(x => 'tbl-' + x.id === settings.nTable.id);
      if (!t) return true;
      const f = filters[t.id];
      for (const ci in f) if (!f[ci].has(txt(t, row, +ci))) return false;
      return !(t.match && !t.match(row));
    });

    const fp = $('#fp');
    let fpCtx = null;
    function renderPills(t, ci) {
      const th = $(t.dt.column(ci).header()), txt = (t.dt.column(ci).search() || '').trim(), sel = filters[t.id][ci];
      const so = t.dt.order()[0], dir = so && so[0] === ci ? so[1] : '';
      th.find('.pills').remove();
      th.find('.fbtn').toggleClass('act', !!(txt || sel || dir));
      if (!txt && !sel && !dir) return;
      const box = $('<div class="pills"></div>').appendTo(th).on('click mousedown keydown keyup', e => e.stopPropagation());
      const add = (label, kind) => $('<span class="pill" title="Click to edit"></span>').append($('<span></span>').text(label), '<b title="Remove filter">&times;</b>').data('k', kind).appendTo(box);
      if (dir) add(dir === 'asc' ? '\u2191 Ascending' : '\u2193 Descending', 'sort');
      if (txt) add(txt, 'text');
      if (sel) { if (sel.size <= 4) [...sel].forEach(v => add(v === '' ? '(Blanks)' : v, 'val').data('v', v)); else add(sel.size + ' selected', 'list'); }
      box.on('click', '.pill', function (e) {
        const k = $(this).data('k');
        if ($(e.target).is('b')) {
          if (k === 'text') t.dt.column(ci).search(''); else if (k === 'sort') t.dt.order([]); else if (k === 'val') { const sl = filters[t.id][ci]; sl.delete($(this).data('v')); if (!sl.size) delete filters[t.id][ci]; } else delete filters[t.id][ci];
          renderPills(t, ci); t.dt.draw();
        } else openFilter(t, ci, th.find('.fbtn')[0]);
      });
    }
    function openFilter(t, ci, btn) {
      const c = t.cols[ci], rows = t.dt.rows().data().toArray();
      const counts = new Map();
      rows.forEach(r => { const v = show(c, r); counts.set(v, (counts.get(v) || 0) + 1); });
      const vals = [...counts.keys()];
      vals.sort(c.type === 'n' ? (a, b) => (parseFloat(a.replace(/[$,]/g, '')) || 0) - (parseFloat(b.replace(/[$,]/g, '')) || 0)
        : c.type === 'd' ? (a, b) => d2iso(a).localeCompare(d2iso(b)) : (a, b) => a.localeCompare(b, undefined, { numeric: true }));
      const cur = filters[t.id][ci];
      fpCtx = { t, ci, vals, counts, open: new Set(), groups: {}, sel: new Set(cur ? vals.filter(v => cur.has(v)) : vals) };
      const isDate = true; // font size / text box / 'Pick values' heading are hidden in every column panel
      const so = t.dt.order()[0], dir = so && so[0] === ci ? so[1] : '', nu = c.type === 'n' || c.type === 'd';
      fp.html('<div class="sortrow"><button type="button" data-sort="asc" class="' + (dir === 'asc' ? 'on' : '') + '">&#8593; ' + (nu ? 'Low to high' : 'A to Z') + '</button><button type="button" data-sort="desc" class="' + (dir === 'desc' ? 'on' : '') + '">&#8595; ' + (nu ? 'High to low' : 'Z to A') + '</button></div>' +
        (isDate ? '' : '<div class="lbl">Column font size</div><div class="sortrow fsrow"><button type="button" data-fs="-1">A&minus;</button><button type="button" class="fsv" disabled>' + ((FS[t.id] || {})[ci] || BASE) + ' px</button><button type="button" data-fs="1">A+</button><button type="button" data-fs="0">Reset</button></div>' +
        '<div class="lbl">Text contains</div><input type="text" class="ftext" placeholder="Type to filter ' + esc(c.label) + '">' +
        '<div class="lbl">Pick values</div>') + '<input type="search" placeholder="Search values..."><div class="list"></div>' +
        '<div class="row"><a id="fclr">Clear filter</a><button id="fcancel">Cancel</button><button class="ok" id="fok">OK</button></div>');
      fp.find('.ftext').val(t.dt.column(ci).search() || '');
      drawList('');
      const rc = btn.getBoundingClientRect();   // visual position (correct even when the table is zoomed to fit)
      fp.css({ display: 'block', top: rc.bottom + window.scrollY + 6, left: Math.max(8, Math.min(rc.left + window.scrollX - 20, document.documentElement.clientWidth - 222)) });
      (fp.find('.ftext').length ? fp.find('.ftext') : fp.find('input[type=search]')).focus();
    }
    function drawList(q) {
      const x = fpCtx, shown = x.vals.filter(v => v.toLowerCase().includes(q.toLowerCase()));
      const all = shown.length && shown.every(v => x.sel.has(v));
      const list = fp.find('.list'), top = list.scrollTop();
      const selAll = '<label><input type="checkbox" id="fall"' + (all ? ' checked' : '') + '> <b>(Select All)</b></label>';
      const leaf = (v, pad, txt) => '<label style="padding-left:' + pad + 'px"><input type="checkbox" data-i="' + x.vals.indexOf(v) + '"' + (x.sel.has(v) ? ' checked' : '') + '> ' + (txt === undefined ? (v === '' ? '(Blanks)' : esc(v)) : txt) + '<span class="c">' + x.counts.get(v) + '</span></label>';
      if (x.t.cols[x.ci].type !== 'd') {
        list.html(selAll + shown.map(v => leaf(v, 0)).join(''));
      } else {
        // Excel-style date tree: Year > Month > Day
        const MN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        const tree = {}, flat = []; x.groups = {};
        shown.forEach(v => { const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(v); if (!m) { flat.push(v); return; } ((tree[m[3]] = tree[m[3]] || {})[m[1]] = (tree[m[3]] || {})[m[1]] || []).push(v); });
        const cnt = arr => arr.reduce((n, v) => n + x.counts.get(v), 0);
        const node = (key, label, arr, pad, body) => {
          x.groups[key] = arr;
          const n = arr.filter(v => x.sel.has(v)).length, open = q || x.open.has(key);
          return '<div class="tn" style="padding-left:' + pad + 'px"><span class="tg" data-t="' + key + '">' + (open ? '&minus;' : '+') + '</span><label><input type="checkbox" data-g="' + key + '" data-s="' + (n === arr.length ? 'all' : n ? 'some' : 'none') + '"' + (n === arr.length ? ' checked' : '') + '> ' + label + '<span class="c">' + cnt(arr) + '</span></label></div>' + (open ? body() : '');
        };
        let html = selAll;
        Object.keys(tree).sort((a, b) => b - a).forEach(y => {
          const months = tree[y], yall = [].concat(...Object.values(months));
          html += node(y, y, yall, 0, () => Object.keys(months).sort().map(m =>
            node(y + '-' + m, MN[+m - 1], months[m], 18, () => months[m].map(v => leaf(v, 54, v.slice(3, 5))).join(''))).join(''));
        });
        html += flat.map(v => leaf(v, 0)).join('');
        list.html(html);
        list.find('input[data-s=some]').prop('indeterminate', true);
      }
      list.scrollTop(top);
      fp.data('shown', shown);
    }
    fp.on('input', 'input[type=search]', e => drawList(e.target.value));
    fp.on('keydown', '.ftext', e => { if (e.key === 'Enter') applyFilter(false); else if (e.key === 'Escape') fp.hide(); });
    fp.on('change', 'input[type=checkbox]', function () {
      const x = fpCtx;
      if (this.id === 'fall') fp.data('shown').forEach(v => this.checked ? x.sel.add(v) : x.sel.delete(v));
      else if (this.dataset.g) (x.groups[this.dataset.g] || []).forEach(v => this.checked ? x.sel.add(v) : x.sel.delete(v));
      else { const v = x.vals[+this.dataset.i]; this.checked ? x.sel.add(v) : x.sel.delete(v); }
      drawList(fp.find('input[type=search]').val());
    });
    function applyFilter(clear) {
      const x = fpCtx, f = filters[x.t.id];
      if (clear || x.sel.size === x.vals.length) delete f[x.ci]; else f[x.ci] = new Set(x.sel);
      if (clear || fp.find('.ftext').length) x.t.dt.column(x.ci).search(clear ? '' : fp.find('.ftext').val().trim());
      renderPills(x.t, x.ci);
      x.t.dt.draw();
      fp.hide();
    }
    fp.on('click', '.tg', function () {
      const x = fpCtx, k = this.dataset.t;
      x.open.has(k) ? x.open.delete(k) : x.open.add(k);
      drawList(fp.find('input[type=search]').val());
    });
    fp.on('click', '[data-fs]', function () {
      const x = fpCtx, m = FS[x.t.id] = FS[x.t.id] || {}, d = +this.dataset.fs;
      if (d === 0) delete m[x.ci]; else m[x.ci] = Math.max(7, Math.min(28, (m[x.ci] || BASE) + d));
      applyFS();
      fp.find('.fsv').text((m[x.ci] || BASE) + ' px');
    });
    fp.on('click', '[data-sort]', function () {
      const x = fpCtx, was = $(this).hasClass('on');
      x.t.dt.order(was ? [] : [[x.ci, this.dataset.sort]]).draw();
      x.t.cols.forEach((c, i) => renderPills(x.t, i));
      fp.hide();
    });
    fp.on('click', '#fok', () => applyFilter(false));
    fp.on('click', '#fclr', () => applyFilter(true));
    fp.on('click', '#fcancel', () => fp.hide());
    $(document).on('mousedown', e => { if (!$(e.target).closest('#fp, .fbtn, .pill').length) fp.hide(); });

    // once the table header is frozen at the top, drop the search card to the bottom-right so it never covers the header
    $(window).on('scroll resize', () => { const p = $('#panes')[0]; $('#dock').toggleClass('low', !!p && p.getBoundingClientRect().top < -4); });
    T.forEach(t => {
      t.sq = { q: '', col: -1, cs: false, whole: false, rx: false, mode: 'all' };
      $('#dock').append(
        '<div class="sbar" data-t="' + t.id + '"><div class="r1"><div class="sbox"><svg class="sico" viewBox="0 0 24 24"><path d="M15.5 14h-.8l-.3-.3A6.5 6.5 0 1 0 14 15.5l.3.3v.8l5 5 1.5-1.5-5-5zm-6 0a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9z"/></svg>' +
        '<input type="search" class="sq" placeholder="Search this table" autocomplete="off"><span class="mt"></span><span class="nc"></span>' +
        '<button type="button" class="nprev" title="Previous match (Shift+Enter)">&#9650;</button><button type="button" class="nnext" title="Next match (Enter)">&#9660;</button>' +
        '<button type="button" class="sclr" title="Clear search and options">&times;</button></div><span class="info" data-tip="Words are combined per the mode. Use quotes for an exact phrase. A leading minus (-word) excludes rows containing it. * and ? are wildcards like Excel. Number columns accept &gt;1000, &lt;=50, =0, &lt;&gt;0 and ranges like 100..500. With Regex on, the text is used as a regular expression. Enter = next match, Shift+Enter = previous.">i</span></div>' +
        '<div class="r2"><select class="scol" title="Search in"><option value="-1">All columns</option>' + t.cols.map((c, i) => '<option value="' + i + '">' + esc(c.label) + '</option>').join('') + '</select>' +
        '<span class="seg"><label class="opt" title="Every word must match"><input type="radio" class="smode" name="m-' + t.id + '" value="all" checked>All</label>' +
        '<label class="opt" title="Any word may match"><input type="radio" class="smode" name="m-' + t.id + '" value="any">Any</label>' +
        '<label class="opt" title="Match the text as one phrase"><input type="radio" class="smode" name="m-' + t.id + '" value="phrase">Phrase</label></span>' +
        '<label class="opt" title="Match case"><input type="checkbox" class="scs">Aa</label>' +
        '<label class="opt" title="Match entire cell contents"><input type="checkbox" class="swhole">Cell</label>' +
        '<label class="opt" title="Regular expression"><input type="checkbox" class="srx">.*</label><span class="cnt"></span></div></div>');
      const bar = $('#dock .sbar[data-t="' + t.id + '"]'); bar.toggle(T.indexOf(t) === 0);
      let timer;
      const apply = () => {
        const s = t.sq;
        s.q = bar.find('.sq').val(); s.col = +bar.find('.scol').val(); s.mode = bar.find('.smode:checked').val();
        s.cs = bar.find('.scs').prop('checked'); s.whole = bar.find('.swhole').prop('checked'); s.rx = bar.find('.srx').prop('checked');
        // show a "Searching..." spinner first, then do the (synchronous) filtering on the next tick so the user sees it
        bar.find('.sbox').addClass('busy'); bar.find('.mt').removeClass('zero').text('Searching…');
        setTimeout(() => {
          compileSearch(t); t.dt.settings()[0].hl = t.hl; t.jump = true;
          bar.find('.sq').toggleClass('bad', t.bad);
          t.dt.draw();
          bar.find('.sbox').removeClass('busy');
        }, 30);
      };
      bar.on('input', '.sq', () => { clearTimeout(timer); timer = setTimeout(apply, 120); });
      bar.on('change', 'select, input[type=checkbox], input[type=radio]', apply);
      bar.on('click', '.sclr', () => { bar.find('.sq').val(''); bar.find('.scol').val('-1'); bar.find('input[type=checkbox]').prop('checked', false); bar.find('.smode[value=all]').prop('checked', true); apply(); });
      bar.on('click', '.shelp', () => bar.toggleClass('showhelp'));
      // browser-style find: step through the highlighted matches (Enter / Shift+Enter, or the arrows)
      let cur = -1;
      const marks = () => $('#pane-' + t.id + ' tbody mark');
      const go = d => {
        const m = marks(); if (!m.length) { bar.find('.nc').text(''); return; }
        cur = d === 0 ? 0 : (cur + d + m.length) % m.length;
        m.removeClass('cur'); const el = m.eq(cur).addClass('cur')[0];
        el.scrollIntoView({ block: 'center', inline: 'center' });
        bar.find('.nc').text((cur + 1) + '/' + m.length);
      };
      bar.on('click', '.nnext', () => go(1));
      bar.on('click', '.nprev', () => go(-1));
      bar.on('keydown', '.sq', e => { if (e.key === 'Enter') { e.preventDefault(); go(e.shiftKey ? -1 : 1); } });
      t.dt.on('draw.dt', () => setTimeout(() => {
        cur = -1; const n = marks().length;
        bar.find('.nc').text(n ? n + ' found' : '');
        if (t.jump && n) go(0);
        t.jump = false;
      }, 0));
      t.dt.on('draw.dt', () => {
        const n = t.dt.rows({ search: 'applied' }).count(), all = t.dt.rows().count();
        bar.find('.cnt').text(n.toLocaleString() + ' of ' + all.toLocaleString() + ' rows');
        bar.find('.mt').text(n < all ? n.toLocaleString() + (n === 1 ? ' entry matched' : ' entries matched') : '').toggleClass('zero', n === 0 && n < all);
      });
      // funnel icon in every header cell
      t.cols.forEach((c, i) => {
        const th = $(t.dt.column(i).header());
        if (['Cust Code', 'DV'].includes(c.label)) $('<button type="button" class="dotbtn"></button>').prependTo(th.find('.dt-column-header').length ? th.find('.dt-column-header') : th)
          .on('click keydown keyup mousedown', e => e.stopPropagation())
          .on('click', function () { const m = CL[t.id] = CL[t.id] || {}; m[i] = !m[i]; applyCL(); t.dt.columns.adjust(); setTimeout(() => t.fit && t.fit(), 0); });
        $('<button type="button" class="fbtn" title="Filter">' + FUNNEL + '</button>').appendTo(th.find('.dt-column-header').length ? th.find('.dt-column-header') : th)
          .on('click keydown keyup mousedown', e => e.stopPropagation())
          .on('click', function () { openFilter(t, i, this); });
      });
      applyCL();
      // fit every column in the viewport: shrink the whole table (zoom) only when it is wider than the window
      t.fit = () => {
        const tb = $('#tbl-' + t.id), pane = $('#pane-' + t.id);
        if (!pane.is(':visible')) return;
        tb.css('zoom', 1);
        const avail = pane.width() - 4, need = Math.max(tb[0].scrollWidth, tb[0].offsetWidth);
        if (need > avail) tb.css('zoom', Math.max(0.4, avail / need));
      };
      let ft; const fitSoon = () => { clearTimeout(ft); ft = setTimeout(t.fit, 30); };
      t.dt.on('draw.dt', fitSoon);
      $(window).on('resize', fitSoon);
      $('#tabs').on('click', '.tab', function () { bar.toggle(this.dataset.t === t.id); setTimeout(fitSoon, 30); });
      t.fitSoon = fitSoon;
      // clear every filter (used when the data is reloaded)
      t.resetUI = () => {
        t.dt.columns().search('');
        t.cols.forEach((c, i) => renderPills(t, i));
        bar.find('.sclr').trigger('click');
      };
    });
  }

  function hlTable(dt) {
    const $ = jQuery, H = dt.settings()[0].hl;
    const g = H ? [] : (dt.search() || '').trim().split(/\s+/).filter(Boolean);
    const cols = dt.columns().indexes().toArray().map(i => (dt.column(i).search() || '').trim());
    let brk = false;
    const esc = s => s.replace(/[&<>\/]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '/': brk ? '/<br>' : '/' }[c]));
    const q = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    $(dt.table().body()).find('tr').each(function () {
      $(this).children('td').each(function (i) {
        brk = this.classList.contains('slm');
        const cq = cols[i] ? [q(cols[i])] : [];
        const terms = g.map(q).concat(cq);
        const hasH = H && (H.col < 0 || H.col === i);
        const had = this.querySelector('mark');
        if (!terms.length && !hasH) { if (had) this.innerHTML = esc(this.textContent); return; }
        const txt = this.textContent;
        const out = [];
        try {
          if (terms.length) out.push(new RegExp(terms.join('|'), 'gi'));
          if (hasH) out.push(new RegExp(H.src, H.cs ? 'g' : 'gi'));
        } catch (e) { return; }
        const hit = new Array(txt.length).fill(false);
        let any = false, m;
        out.forEach(re => { while ((m = re.exec(txt))) { if (!m[0]) { re.lastIndex++; continue; } any = true; for (let k = m.index; k < m.index + m[0].length; k++) hit[k] = true; } });
        if (!any) { if (had) this.innerHTML = esc(txt); return; }
        let html = '', open = false;
        for (let k = 0; k < txt.length; k++) {
          if (hit[k] !== open) { html += open ? '</mark>' : '<mark>'; open = hit[k]; }
          html += esc(txt[k]);
        }
        this.innerHTML = html + (open ? '</mark>' : '');
      });
    });
  }

  window.R2Engine = { setupUI: setupUI, hlTable: hlTable };
})();
