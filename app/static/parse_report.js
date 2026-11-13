/* A/R Inquiry PDF parser, copied unchanged from the Report2Excel v7 tool. */
(function () {
/* ---------- PDF parser: A/R Inquiry Report (detail lines, claims, aging totals) ---------- */
function groupLines(items) {
  const sorted = items.slice().sort((a, b) => b.y - a.y || a.x - b.x);
  const lines = [];
  for (const it of sorted) {
    const ln = lines[lines.length - 1];
    if (ln && Math.abs(ln.y - it.y) <= 2) ln.items.push(it);
    else lines.push({ y: it.y, items: [it] });
  }
  for (const ln of lines) {
    ln.items.sort((a, b) => a.x - b.x);
    ln.text = ln.items.map(i => i.str).join(' ');
  }
  return lines;
}
const MONEY = /^[\d,]*\.\d\d-?$/, DATE = /^\d\d\/\d\d\/\d\d$/;
const money = s => { const neg = /-$/.test(s); const n = parseFloat(s.replace(/[,-]/g, '')); return isNaN(n) ? null : (neg ? -n : n); };
const iso = s => { const m = /^(\d\d)\/(\d\d)\/(\d\d)$/.exec(s || ''); return m ? '20' + m[3] + '-' + m[1] + '-' + m[2] : ''; };
const round2 = n => Math.round(n * 100) / 100;
const SHIP = {"VAAM9285":"AM9285","EMQV9020":"QV1281","EMAA9014":"AA6902","EMAH9046":"AH0009","EMAC9012":"AL2010","EMAD9302":"AD1302","EMAM9988":"AM1988","EMAM9878":"AM1998","EMAM9286":"AXYEG2","EMAM9285":"AZIND9","WQBJ9009":"BJB800","EMBJ9999":"BJB800","EMBP9099":"BP902S","EMBK9000":"BK0412","EMBI9040":"BI4004","EMCT9001":"CT9007","EMKN9100":"KN1100","EMCE9955":"CE1273","EMCY9855":"CYDAY1","EMCB9075":"SE1536","EMCC9145":"CC9999","EMDB9465":"DB0002","VADG9250":"DG6900","EMDO9500":"DO1650","EMEB9025":"EB1025","EMFD9080":"SH1514","EMFE9500":"FE1500","EMFL9104":"FL1104","EMFL9101":"FL1101","EMFL9102":"FL1102","EMFM9058":"FM4461","EMGO9201":"BG0827","EMGR9004":"GR1132","EMHB9003":"HE4040","EMHD9290":"HD5034","EMHC9000":"HC7275","EMHO9097":"HH0001","EMHA9025":"HN1025","EMHA9225":"HN1225","EMHS9029":"HS0094","EMIN9370":"IN3371","EMJC9051":"JC0010","EMKA9000":"KA1001","EMKE9950":"KE1550","EMKO9036":"KO9036","EMKO9035":"KO0826","EMPE9084":"PE8004","EMKV9000":"KV0000","EMLR9000":"LR1111","EMLE9900":"EN1275","EMLW9301":"LW0955","EMLW9000":"LW1100","EMMD9102":"MD1102","EMMA9359":"MD7000","EMMX9050":"M11100","EMMD9009":"MD1001","EMME9040":"ME0093","EMME9145":"ME1715","EMMI9091":"MI8030","EMMR994450":"MR114450","EMNA9010":"NA0984","EMPM9620":"PM1620","EMQV9200":"QV1017","EMQV9201":"QV9201","EMQV9090":"QV1290","EMRO9900":"RO0000","EMRU9900":"RU1002","EMSA9100":"SA1100","EMSA9200":"SA6464","EMSD9165":"SD1165","EMSD9167":"SD1167","EMSD9166":"SD1166","EMBM9501":"EN1660","EMST9252":"ST1259","EMST9990":"ST0274","EMSU9189":"SU1003","EMTA9098":"TA3857","EMTA9199":"TA1199","EMTW9001":"TW2222","EMNW9000":"NW1101","EMTS9620":"TS1602","EMTD9120":"TD1120","EMTD9110":"TD1110","EMTD9100":"TD1100","EMTI9099":"TC1230","EMTD9091":"TD1090","EMTR9012":"TR0009","EMTU9000":"TU1150","EMVA9040":"VA1960","EMWC9085":"WC7102","EMWC9087":"WC1187","EMWA9085":"WA6561","EMWD9709":"WD1772","EMWF9100":"WF1100","EMCS9100":"CS1100","EMWE9349":"WE0111","EMHI9980":"HI1900","EMTE9005":"TE1611","EMCH9610":"CH1610","EMCA9118":"CA3070","EMWC9086":"WC1100","EMSC9999":"SC9999","EMGI9025":"GI1241","EMVE9630":"VE0372","EMBN9001":"BN0001","EMHH9800":"HE1280","EMOR9009":"OG0001","EMTY9090":"TY0398","EMGN9009":"GN4042","EMAF9039":"AF1039","EMFD9010":"FD9515","EMRD9877":"RD3000"};
const shipFor = c => SHIP[String(c||'').replace(/\s+/g,'').toUpperCase()] || '';
const HEAD = /^([A-Z]{2})\s+([A-Z0-9]+)\s+([A-Z]{2}\w*)\s+(\d\d\/\d\d\/\d\d)\s+(\d\d\/\d\d\/\d\d)/;
const AGE_ROWS = { 'INVOICES': 'Invoices', 'RETURNS': 'Returns', 'DEBIT MEMOS': 'Debit Memos', 'CREDIT MEMOS': 'Credit Memos' };
const AGE_KEYS = ['total', 'current', 'd1_30', 'd31_60', 'd61_90', 'over90'];

function parseReport(pages, fileName) {
  const res = { file: fileName, detail: [], aging: [], customers: [], warnings: [] };
  let cust = null, last = null;
  for (let pageNo = 1; pageNo <= pages.length; pageNo++) {
    const items = pages[pageNo - 1];
    for (const ln of groupLines(items)) {
      const t = ln.text, f = ln.items[0];
      if (/E\. MISHAN|Inquiry - Report|DB1 - A\/R|Powered by YCS/.test(t) || (f.str === 'DV' && /Ref#/.test(t))) continue;

      if (/CREDIT LINE:/.test(t)) {
        cust = { file: fileName, page: pageNo, code: f.str, name: ln.items.filter(i => i.x > 85 && i.x < 300).map(i => i.str).join(' '), terms: ln.items.filter(i => i.x > 500 && i.x < 675 && !/^TERMS:?$/.test(i.str)).map(i => i.str).join(' '), slm: '', vendor: '', balance: null, mtd: null, ytd: null, ytdRet: null };
        res.customers.push(cust); last = null; continue;
      }
      let x;
      if (cust) {
        const tail = () => ln.items.filter(i => i.x > 500 && !/^(SLM :?|Vendor#:?)$/.test(i.str)).map(i => i.str).join(' ');
        if (ln.items.some(i => /^SLM :?$/.test(i.str))) cust.slm = tail();
        if (ln.items.some(i => /^Vendor#:?$/.test(i.str))) cust.vendor = tail();
        if (cust.balance === null && (x = /Balance:\s*(\S+)/.exec(t)) && MONEY.test(x[1])) { cust.balance = money(x[1]); continue; }
        if (cust.mtd === null && (x = /M-T-D Sales:\s*(\S+)/.exec(t)) && MONEY.test(x[1])) cust.mtd = money(x[1]);
        if (cust.ytd === null && (x = /Y-T-D Sales:\s*(\S+)/.exec(t)) && MONEY.test(x[1])) cust.ytd = money(x[1]);
        if (cust.ytdRet === null && (x = /Y-T-D Retns:\s*(\S+)/.exec(t)) && MONEY.test(x[1])) cust.ytdRet = money(x[1]);
      }

      const h = HEAD.exec(t);
      if (h) {
        const rest = ln.items.filter(i => i.x >= 225);
        const row = { file: fileName, custCode: cust ? cust.code : '', customer: cust ? cust.name : '', _c: cust, dv: h[1], acct: h[2], ship: shipFor(cust && cust.code), ref: h[3], type: h[3].slice(0, 2),
          page: pageNo, date: iso(h[4]), due: iso(h[5]), amount: null, pdDate: '', past: '', order: '', check: '', checkAmt: null, checkDate: '', claim: '' };
        const order = [], chk = [];
        for (const i of rest) {
          if (MONEY.test(i.str) && i.x < 312 && row.amount === null) row.amount = money(i.str);
          else if (DATE.test(i.str) && i.x < 365) row.pdDate = iso(i.str);
          else if (i.x >= 360 && i.x < 396) row.past += (row.past ? ' ' : '') + i.str;
          else if (i.x >= 396 && i.x < 505) order.push(i.str);
          else if (MONEY.test(i.str) && i.x >= 600) row.checkAmt = money(i.str);
          else if (DATE.test(i.str) && i.x >= 690) row.checkDate = iso(i.str);
          else if (i.x >= 505 && i.x < 600) chk.push(i.str);
        }
        row.order = order.join(' '); row.check = chk.join(' ');
        row._c = cust; res.detail.push(row); last = row;
        continue;
      }
      if (f.str.startsWith('Claim#:')) {
        if (last) last.claim = t.replace(/^Claim#:\s*/, '').trim();
        continue;
      }

      const vals = ln.items.filter(i => MONEY.test(i.str) && i.x > 150).map(i => money(i.str));
      const label = ln.items.filter(i => !MONEY.test(i.str) && !/^_+$/.test(i.str)).map(i => i.str).join(' ').replace(/^_+\s*/, '');
      if (vals.length === 6 && (AGE_ROWS[label] || /Totals$/.test(label))) {
        const type = AGE_ROWS[label] || 'Totals';
        const rec = { file: fileName, page: pageNo, scope: '', custCode: '', name: '', type };
        AGE_KEYS.forEach((k, i) => rec[k] = vals[i]);
        if (AGE_ROWS[label]) { rec._pending = true; res.aging.push(rec); }
        else {
          const scope = label.replace(/\s*Totals$/, '') || 'Totals';
          const isCust = /\d/.test(scope) && cust;
          let j = res.aging.length; while (j > 0 && res.aging[j - 1]._pending) j--;
          for (let k = j; k < res.aging.length; k++) { Object.assign(res.aging[k], { scope, _c: isCust ? cust : null }); delete res.aging[k]._pending; }
          rec.scope = scope; rec._c = isCust ? cust : null; res.aging.push(rec);
        }
      }
    }
  }
  const sums = new Map(), dvSums = new Map();
  for (const d of res.detail) {
    sums.set(d._c, round2((sums.get(d._c) || 0) + (d.amount || 0)));
    dvSums.set(d.dv, round2((dvSums.get(d.dv) || 0) + (d.amount || 0)));
  }
  for (const r of res.aging) {
    r.custCode = r._c ? r._c.code : '';
    r.name = r._c ? r._c.name : (/^Final/.test(r.scope) ? 'All customers' : 'Division total');
    if (r.type === 'Totals') {
      const calc = r._c ? sums.get(r._c) : dvSums.get(r.scope);
      if (calc !== undefined) {
        r.calc = calc; r.diff = round2(r.total - calc);
        if (r.diff) res.warnings.push((r._c ? r._c.code : r.scope) + ': printed total ' + r.total + ' differs from line sum ' + calc);
      }
    }
    delete r._c;
  }
  for (const d of res.detail) { if (d._c) { d.terms = d._c.terms; d.slm = d._c.slm; d.vendor = d._c.vendor; } delete d._c; }
  if (!res.detail.length) res.warnings.push('No A/R detail lines found in this PDF.');
  return res;
}


  window.parseReport = parseReport;
})();
