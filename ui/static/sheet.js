'use strict';

// Spreadsheet preview: draws the workbook model of /files/<name>/sheet (ui/xlsx_view.py)
// as an Excel-like grid with a formula bar, sheet tabs and a "show formulas" mode.
const Sheet = (() => {
  const BORDER = {
    thin: '1px solid', hair: '1px solid', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed',
    dashDotDot: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed',
    mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', thick: '3px solid', double: '3px double',
  };
  const GRID = '1px solid #dadcdf';
  const NONE = '1px solid transparent';
  const FONTS = "Carlito, Calibri, 'Helvetica Neue', Arial, sans-serif";  // single quotes: goes into style="..."
  const fontName = n => `'${String(n).replace(/['"\\<>;{}]/g, '')}'`;
  const H_ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'justify', fill: 'left',
    centerContinuous: 'center', distributed: 'center' };
  const V_ALIGN = { top: 'top', center: 'middle', justify: 'top', distributed: 'middle' };
  const ROW_HEAD = 44;

  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const colLetter = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; } return s; };
  const key = (r, c) => r * 4096 + c;
  const inRange = (rg, r, c) => r >= rg[0] && r <= rg[2] && c >= rg[1] && c <= rg[3];

  function mount(root, wb, opts = {}) {
    const st = { sheet: wb.active || 0, formulas: false, sel: null };
    root.innerHTML = `<div class="xl">
      <div class="xl-bar"><div class="xl-ref"></div><div class="xl-fx">fx</div><div class="xl-formula"></div>
        <div class="xl-val"></div><label class="xl-toggle" title="Formülleri göster (Ctrl+\`)"><input type="checkbox"><span>Formüller</span></label></div>
      <div class="xl-info" hidden></div>
      <div class="xl-scroll" tabindex="0"></div>
      <div class="xl-notes" hidden></div>
      <div class="xl-tabs"></div></div>`;
    const $ = sel => root.querySelector(sel);
    const scroll = $('.xl-scroll');
    let S, cells, merged, covered, colW, colHidden, colStyle, rows, cols, memo, fWidth;

    function prepare() {
      S = wb.sheets[st.sheet];
      cells = new Map(S.cells.map(c => [key(c.r, c.c), c]));
      merged = new Map();
      covered = new Map();
      for (const [r1, c1, r2, c2] of S.merges) {
        merged.set(key(r1, c1), [r2, c2]);
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered.set(key(r, c), [r1, c1]);
      }
      const R = Math.max(S.rows + (S.rows < 1990 ? 8 : 0), 30);
      const C = Math.max(S.cols + 3, 10);
      colW = new Array(C + 2).fill(S.defaultW);
      colHidden = new Array(C + 2).fill(false);
      colStyle = new Array(C + 2).fill(null);
      for (const ci of S.colInfo) {
        for (let c = ci.min; c <= Math.min(ci.max, C + 1); c++) {
          colW[c] = ci.w;
          colHidden[c] = ci.hidden || ci.w === 0;
          if (ci.s) colStyle[c] = ci.s;
        }
      }
      rows = [];
      for (let r = 1; r <= R; r++) if (!S.rowInfo[r]?.hidden) rows.push(r);
      cols = [];
      for (let c = 1; c <= C; c++) if (!colHidden[c]) cols.push(c);
      memo = new Map();
      fWidth = new Map();  // formula view widens formula columns to their longest formula
      for (const cell of S.cells) {
        if (cell.f && !merged.has(key(cell.r, cell.c))) {
          fWidth.set(cell.c, Math.max(fWidth.get(cell.c) || 0, Math.min(420, Math.round(cell.f.length * 7.3 + 12))));
        }
      }
    }

    const width = c => (st.formulas ? Math.max(colW[c], fWidth.get(c) || 0) : colW[c]);

    // ---- styles: table style < cell format < number-format color < conditional format ----
    function tableStyle(r, c) {
      for (const t of S.tables) {
        if (!inRange(t.range, r, c)) continue;
        const look = t.look;
        let s = { ...(look.body || {}) };
        if (r < t.range[0] + t.header) s = { ...s, ...look.hdr };
        else if (t.totals && r > t.range[2] - t.totals) s = { ...s, b: 1, bd: { t: ['double', look.hdr.bg || '#000000'] } };
        else if (t.stripes && (r - t.range[0] - t.header) % 2 === 0) s = { ...s, ...look.band };
        if (t.first && c === t.range[1]) s.b = 1;
        return s;
      }
      return null;
    }

    function styleAt(r, c) {
      const k = key(r, c);
      if (memo.has(k)) return memo.get(k);
      const cell = cells.get(k);
      const idx = cell ? (cell.s || 0) : (S.rowInfo[r]?.s ?? colStyle[c] ?? 0);
      let s = { ...(wb.styles[idx] || {}) };
      const t = tableStyle(r, c);
      if (t) {
        for (const p of ['fn', 'sz', 'c']) if (s[p] === wb.font[p] || (p === 'c' && s.c === '#000000')) delete s[p];
        s = { ...t, ...s };
      }
      if (cell?.fc) s.c = cell.fc;
      if (cell?.cf) s = { ...s, ...cell.cf };
      memo.set(k, s);
      return s;
    }

    function css(s, cell, shown) {
      const out = [];
      if (s.fn && s.fn !== wb.font.fn) out.push(`font-family:${fontName(s.fn)},${FONTS}`);
      if (s.sz && s.sz !== wb.font.sz) out.push(`font-size:${s.sz}pt`);
      if (s.b) out.push('font-weight:700');
      if (s.i) out.push('font-style:italic');
      const deco = [s.u ? 'underline' : '', s.s ? 'line-through' : ''].filter(Boolean).join(' ');
      if (deco) out.push(`text-decoration:${deco}${s.u === 'double' || s.u === 'doubleAccounting' ? ' double' : ''}`);
      if (s.c) out.push(`color:${s.c}`);
      if (s.bg) out.push(`background-color:${s.bg}`);
      const kind = cell?.k;
      let h = H_ALIGN[s.h];
      if (!h) h = shown === 'formula' ? 'left' : kind === 'n' || kind === 'd' ? 'right' : kind === 'b' || kind === 'e' ? 'center' : 'left';
      if (h !== 'left') out.push(`text-align:${h}`);
      if (V_ALIGN[s.v]) out.push(`vertical-align:${V_ALIGN[s.v]}`);
      if (s.w) out.push('white-space:pre-wrap;overflow-wrap:anywhere');
      if (s.ind) out.push(`padding-${h === 'right' ? 'right' : 'left'}:${3 + s.ind * 9}px`);
      return out.join(';');
    }

    function side(s, name) {
      const b = s.bd && s.bd[name];
      return b ? `${BORDER[b[0]] || '1px solid'} ${b[1]}` : null;
    }

    function edge(own, other, ownStyle, otherStyle) {
      if (own || other) return own || other;
      if (!S.grid || ownStyle.bg || (otherStyle && otherStyle.bg)) return NONE;
      return GRID;
    }

    const isEmpty = (r, c) => {
      const cell = cells.get(key(r, c));
      return !(cell && (cell.v || (st.formulas && cell.f))) && !covered.has(key(r, c)) && !merged.has(key(r, c));
    };

    // ---- grid ----
    function render() {
      const out = ['<table class="xl-grid"><colgroup>', `<col style="width:${ROW_HEAD}px">`];
      for (const c of cols) out.push(`<col style="width:${width(c)}px">`);
      out.push('</colgroup><thead><tr><th class="xl-corner"></th>');
      for (const c of cols) out.push(`<th data-c="${c}">${colLetter(c)}</th>`);
      out.push('</tr></thead><tbody>');
      const visRow = new Set(rows);
      const visCol = new Set(cols);
      for (const r of rows) {
        const h = S.rowInfo[r]?.h ?? S.defaultH;
        const first = styleAt(r, cols[0] || 1);
        out.push(`<tr style="height:${h}pt"><th data-r="${r}" style="border-right:${side(first, 'l') || '1px solid #c8cbd0'}">${r}</th>`);
        for (const c of cols) {
          if (covered.has(key(r, c))) continue;
          const cell = cells.get(key(r, c));
          const s = styleAt(r, c);
          let span = '';
          let rEnd = r;
          let cEnd = c;
          const m = merged.get(key(r, c));
          if (m) {
            [rEnd, cEnd] = m;
            let rs = 0;
            let cs = 0;
            for (let i = r; i <= rEnd; i++) if (visRow.has(i)) rs++;
            for (let j = c; j <= cEnd; j++) if (visCol.has(j)) cs++;
            if (rs > 1) span += ` rowspan="${rs}"`;
            if (cs > 1) span += ` colspan="${cs}"`;
          }
          const right = styleAt(r, cEnd + 1);
          const below = styleAt(rEnd + 1, c);
          const er = edge(side(styleAt(r, cEnd), 'r'), side(right, 'l'), s, right);
          const eb = edge(side(styleAt(rEnd, c), 'b'), side(below, 't'), s, below);
          const showF = st.formulas && cell?.f;
          const text = showF ? cell.f : (cell?.v ?? '');
          const cls = [];
          const leftish = showF || ((cell?.k === 's' || !cell?.k) && (!s.h || s.h === 'left' || s.h === 'general'));
          if (text && !s.w && !m && leftish && isEmpty(r, c + 1)) cls.push('ov');
          if (cell?.note) cls.push('note');
          if (showF) cls.push('fx-on');
          else if ((cell?.k === 'n' || cell?.k === 'd') && !s.w) cls.push('n');
          const title = cell?.f && !showF ? ` title="${esc(cell.f)}"` : showF ? ` title="${esc(cell.v)}"` : '';
          out.push(`<td data-r="${r}" data-c="${c}"${span}${cls.length ? ` class="${cls.join(' ')}"` : ''}${title} style="${css(s, cell, showF ? 'formula' : '')};border-right:${er};border-bottom:${eb}">${esc(text)}</td>`);
        }
        out.push('</tr>');
      }
      out.push('</tbody></table>');
      const top = scroll.scrollTop;
      const left = scroll.scrollLeft;
      scroll.innerHTML = out.join('');
      scroll.scrollTop = top;
      scroll.scrollLeft = left;
      const table = scroll.querySelector('table');
      table.style.fontFamily = `${fontName(wb.font.fn || 'Calibri')},${FONTS}`;
      table.style.fontSize = `${wb.font.sz || 11}pt`;
      table.style.width = `${cols.reduce((w, c) => w + width(c), ROW_HEAD)}px`;  // columns keep their width, like Excel
      // a number that does not fit shows ####, like Excel (measure everything first: one layout, not one per cell)
      const tooWide = [...table.querySelectorAll('td.n')]
        .filter(td => td.scrollWidth > td.clientWidth + 1).map(td => [td, td.clientWidth]);
      for (const [td, w] of tooWide) {
        td.title = td.textContent;
        td.textContent = '#'.repeat(Math.max(1, Math.floor((w - 6) / 7)));
      }
      freeze(table);
      if (st.sel) select(st.sel.r, st.sel.c, false);
    }

    function freeze(table) {
      const fz = S.freeze;
      if (!fz || (!fz.rows && !fz.cols)) return;
      const head = table.tHead.rows[0];
      const fr = rows.filter(r => r <= fz.rows);
      const fc = cols.filter(c => c <= fz.cols);
      const tops = new Map();
      for (const tr of table.tBodies[0].rows) {
        const r = Number(tr.cells[0].dataset.r);
        if (fr.includes(r)) tops.set(r, tr.offsetTop);
      }
      const lefts = new Map();
      for (const th of head.cells) if (th.dataset.c && fc.includes(Number(th.dataset.c))) lefts.set(Number(th.dataset.c), th.offsetLeft);
      const lastR = fr[fr.length - 1];
      const lastC = fc[fc.length - 1];
      for (const th of table.tBodies[0].querySelectorAll('th')) {
        const r = Number(th.dataset.r);
        if (tops.has(r)) { th.style.top = `${tops.get(r)}px`; th.classList.add('fz-r'); }
      }
      for (const td of table.tBodies[0].querySelectorAll('td')) {
        const r = Number(td.dataset.r);
        const c = Number(td.dataset.c);
        const inR = tops.has(r);
        const inC = lefts.has(c);
        if (!inR && !inC) continue;
        td.classList.add('fz');
        if (inR) td.style.top = `${tops.get(r)}px`;
        if (inC) td.style.left = `${lefts.get(c)}px`;
        if (inR && inC) td.style.zIndex = 3;
        const shadow = [];
        if (r === lastR) shadow.push('inset 0 -1px 0 #8f9399');
        if (c === lastC) shadow.push('inset -1px 0 0 #8f9399');
        if (shadow.length) td.style.boxShadow = shadow.join(',');
      }
    }

    // ---- selection + formula bar ----
    function select(r, c, focus = true) {
      const a = covered.get(key(r, c));
      if (a) [r, c] = a;
      st.sel = { r, c };
      scroll.querySelectorAll('.sel, .hl').forEach(e => e.classList.remove('sel', 'hl'));
      const td = scroll.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
      if (td) {
        td.classList.add('sel');
        if (focus) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
      const m = merged.get(key(r, c)) || [r, c];
      for (let j = c; j <= m[1]; j++) scroll.querySelector(`thead th[data-c="${j}"]`)?.classList.add('hl');
      for (let i = r; i <= m[0]; i++) scroll.querySelector(`tbody th[data-r="${i}"]`)?.classList.add('hl');
      const cell = cells.get(key(r, c));
      $('.xl-ref').textContent = colLetter(c) + r + (m[0] !== r || m[1] !== c ? `:${colLetter(m[1])}${m[0]}` : '');
      const content = cell?.f || cell?.raw || cell?.v || '';
      $('.xl-formula').textContent = content;
      $('.xl-formula').title = content;
      $('.xl-formula').classList.toggle('is-f', !!cell?.f);
      $('.xl-val').textContent = cell?.f ? `= ${cell.v || '(hesaplanmamış)'}` : '';
      const info = [];
      if (cell?.note) info.push(`💬 ${cell.note}`);
      for (const v of S.validations) if (v.ranges.some(rg => inRange(rg, r, c))) info.push(`✓ ${v.text}`);
      $('.xl-info').hidden = !info.length;
      $('.xl-info').textContent = info.join('   ');
    }

    function move(dr, dc) {
      if (!st.sel) return select(rows[0], cols[0]);
      let { r, c } = st.sel;
      const m = merged.get(key(r, c));
      if (m && dr > 0) r = m[0];
      if (m && dc > 0) c = m[1];
      const ri = rows.indexOf(r) + dr;
      const ci = cols.indexOf(c) + dc;
      if (ri < 0 || ri >= rows.length || ci < 0 || ci >= cols.length) return;
      select(rows[ri], cols[ci]);
    }

    // ---- tabs, notes ----
    function renderTabs() {
      const box = $('.xl-tabs');
      box.innerHTML = '';
      wb.sheets.forEach((sh, i) => {
        if (sh.state === 'veryHidden') return;
        const b = document.createElement('button');
        b.type = 'button';
        b.className = `xl-tab${i === st.sheet ? ' on' : ''}${sh.state === 'hidden' ? ' hidden-sheet' : ''}`;
        b.innerHTML = `${sh.tab ? `<span class="sw" style="background:${sh.tab}"></span>` : ''}${esc(sh.name)}${sh.state === 'hidden' ? ' <small>(gizli)</small>' : ''}`;
        b.onclick = () => { if (i !== st.sheet) { st.sheet = i; st.sel = null; show(); } };
        box.append(b);
      });
    }

    function renderNotes() {
      const n = $('.xl-notes');
      n.innerHTML = S.notes.map(esc).join('<br>');
      if (opts.onPdf && S.notes.some(x => x.includes('PDF'))) {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = 'PDF görünümüne geç';
        b.onclick = opts.onPdf;
        n.append(' ', b);
      }
      n.hidden = !S.notes.length;
    }

    function show() {
      prepare();
      renderTabs();
      renderNotes();
      scroll.scrollTop = 0;
      scroll.scrollLeft = 0;
      if (S.kind !== 'sheet') {
        scroll.innerHTML = `<div class="xl-empty">${S.kind === 'chart' ? 'Bu bir grafik sayfası; PDF görünümünde görebilirsiniz.' : 'Bu sayfa türü gösterilemiyor.'}</div>`;
        return;
      }
      render();
      select(rows[0], cols[0], false);
    }

    const toggle = $('.xl-toggle input');
    toggle.onchange = () => { st.formulas = toggle.checked; memo = new Map(); render(); };
    scroll.addEventListener('click', e => {
      const td = e.target.closest('td');
      if (td) select(Number(td.dataset.r), Number(td.dataset.c), false);
    });
    scroll.addEventListener('keydown', e => {
      const k = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, e.shiftKey ? -1 : 1] }[e.key];
      if (k) { e.preventDefault(); move(...k); }
      if (e.key === '`' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); toggle.checked = !toggle.checked; toggle.onchange(); }
    });
    show();
    return { focus: () => scroll.focus() };
  }

  return { mount };
})();
