'use strict';

// ---- state + helpers --------------------------------------------------------
const S = {
  sid: null, sess: null, sessions: [], health: { checking: true },
  pending: [], uploading: [], preview: null, blobs: new Map(), sheets: new Map(), pvMode: 'sheet',
  pollToken: 0, previewToken: 0,
};
const SHEETS = ['.xlsx', '.xlsm', '.xltx', '.csv'];
const UPLOADABLE = ['.docx', '.dotx', '.docm', '.pptx', '.potx', '.pptm', '.xlsx', '.xltx', '.xlsm', '.pdf',
  '.doc', '.ppt', '.xls', '.rtf', '.odt', '.odp', '.ods', '.csv', '.txt', '.md', '.json', '.png', '.jpg', '.jpeg'];
const EXAMPLES = [
  ['Sunum', 'Satış ekibi için 2026 Q3 özetini anlatan 5 slaytlık temalı bir sunum hazırla'],
  ['Excel', 'Aylık ev bütçesi için gelir, gider ve kalan bakiyeyi hesaplayan formüllü bir Excel tablosu oluştur'],
  ['PDF', 'Yeni başlayan çalışanlar için 2 sayfalık bir oryantasyon rehberini PDF olarak hazırla'],
  ['Word', 'Bir yazılım projesi için haftalık durum raporu hazırla (Word)'],
];

const $ = (sel, root = document) => root.querySelector(sel);
const el = (tag, cls, html) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html != null) e.innerHTML = html;
  return e;
};
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const enc = encodeURIComponent;
const extOf = n => (String(n).match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
const FAMILY = {
  docx: 'docx', dotx: 'docx', docm: 'docx', doc: 'docx', rtf: 'docx', odt: 'docx',
  pptx: 'pptx', potx: 'pptx', pptm: 'pptx', ppt: 'pptx', odp: 'pptx',
  xlsx: 'xlsx', xltx: 'xlsx', xlsm: 'xlsx', xls: 'xlsx', ods: 'xlsx', csv: 'xlsx', pdf: 'pdf',
};
const family = n => FAMILY[extOf(n)] || 'other';
const fmtSize = b => b < 1024 ? `${b} B` : b < 1048576 ? `${Math.round(b / 1024)} KB` : `${(b / 1048576).toFixed(1)} MB`;
const fmtDur = s => { s = Math.max(0, Math.round(s)); return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`; };
const now = () => Date.now() / 1000;
const ago = t => {
  const d = now() - t;
  if (d < 60) return 'şimdi';
  if (d < 3600) return `${Math.floor(d / 60)} dk`;
  if (d < 86400) return `${Math.floor(d / 3600)} sa`;
  if (d < 7 * 86400) return `${Math.floor(d / 86400)} g`;
  return new Date(t * 1000).toLocaleDateString('tr-TR', { day: 'numeric', month: 'short' });
};
const fileUrl = (name, extra = '') => `/api/sessions/${S.sid}/files/${enc(name)}${extra}`;
const fileRec = name => (S.sess?.files || []).find(f => f.name === name);
const isSheet = name => (S.health.sheets || SHEETS).includes('.' + extOf(name));
const runningTurn = () => (S.sess?.turns || []).find(t => t.status === 'running');

const ICONS = {
  clip: '<path d="m21.4 11.1-9.2 9.2a6 6 0 0 1-8.5-8.5l9.2-9.2a4 4 0 0 1 5.7 5.7l-9.2 9.2a2 2 0 0 1-2.8-2.8l8.5-8.5"/>',
  send: '<path d="M12 19V5M5 12l7-7 7 7"/>',
  stop: '<rect x="7" y="7" width="10" height="10" rx="1.5" fill="currentColor" stroke="none"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>',
  close: '<path d="M18 6 6 18M6 6l12 12"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  trash: '<path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/>',
  menu: '<path d="M3 6h18M3 12h18M3 18h18"/>',
  external: '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6M15 3h6v6M10 14 21 3"/>',
  eye: '<path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
  terminal: '<path d="m4 17 6-6-6-6M12 19h8"/>',
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>',
  pencil: '<path d="M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',
  book: '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5z"/><path d="M4 19.5A2.5 2.5 0 0 0 6.5 22H20v-5"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
  tool: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.4-.6-.6-2.4z"/>',
  msg: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
  files: '<path d="M15.5 2H8a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V6.5z"/><path d="M3 7.5V20a2 2 0 0 0 2 2h9.5"/>',
};
const svg = (name, size = 18) =>
  `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name]}</svg>`;

async function api(path, opts = {}) {
  const r = await fetch(path, opts);
  const data = (r.headers.get('content-type') || '').includes('application/json') ? await r.json() : null;
  if (!r.ok) throw new Error((data && data.error) || `HTTP ${r.status}`);
  return data;
}

let toastTimer;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 5000);
}

// ---- markdown (agent answers) ----------------------------------------------
function inline(src) {
  const codes = [];
  let s = String(src).replace(/`([^`]+)`/g, (_, c) => { codes.push(c); return `\u0000${codes.length - 1}\u0000`; });
  s = esc(s)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*\w])\*(?!\s)([^*]+?)\*(?!\*)/g, '$1<em>$2</em>')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');
  return s.replace(/\u0000(\d+)\u0000/g, (_, i) => `<code>${esc(codes[i])}</code>`);
}

const LIST = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/;
const isTableSep = l => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(l);
const cells = l => l.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());

function renderList(items) {
  let html = '';
  const stack = [];
  for (const it of items) {
    while (stack.length && it.ind < stack[stack.length - 1].ind) html += `</li></${stack.pop().tag}>`;
    const top = stack[stack.length - 1];
    if (!top || it.ind > top.ind) {
      const tag = it.ord ? 'ol' : 'ul';
      stack.push({ ind: it.ind, tag });
      html += `<${tag}${it.ord && it.start > 1 ? ` start="${it.start}"` : ''}><li>`;
    } else {
      html += '</li><li>';
    }
    html += inline(it.text);
  }
  while (stack.length) html += `</li></${stack.pop().tag}>`;
  return html;
}

function md(src) {
  const lines = String(src || '').replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  let i = 0;
  const startsBlock = j => /^\s*(```|#{1,6}\s|>)/.test(lines[j]) || LIST.test(lines[j]) ||
    (lines[j].includes('|') && j + 1 < lines.length && isTableSep(lines[j + 1]));
  while (i < lines.length) {
    const l = lines[i];
    let m;
    if (!l.trim()) { i++; continue; }
    if (/^\s*```/.test(l)) {
      const buf = [];
      for (i++; i < lines.length && !/^\s*```/.test(lines[i]); i++) buf.push(lines[i]);
      i++;
      out.push(`<pre><code>${esc(buf.join('\n'))}</code></pre>`);
    } else if ((m = l.match(/^(#{1,6})\s+(.*)$/))) {
      const lv = Math.min(m[1].length + 2, 6);
      out.push(`<h${lv}>${inline(m[2])}</h${lv}>`);
      i++;
    } else if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(l)) {
      out.push('<hr>');
      i++;
    } else if (l.includes('|') && i + 1 < lines.length && isTableSep(lines[i + 1])) {
      const head = cells(l);
      const rows = [];
      for (i += 2; i < lines.length && lines[i].includes('|') && lines[i].trim(); i++) rows.push(cells(lines[i]));
      out.push(`<div class="table-wrap"><table><thead><tr>${head.map(c => `<th>${inline(c)}</th>`).join('')}</tr></thead><tbody>${
        rows.map(r => `<tr>${r.map(c => `<td>${inline(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`);
    } else if (/^\s*>/.test(l)) {
      const buf = [];
      for (; i < lines.length && /^\s*>/.test(lines[i]); i++) buf.push(lines[i].replace(/^\s*>\s?/, ''));
      out.push(`<blockquote>${md(buf.join('\n'))}</blockquote>`);
    } else if (LIST.test(l)) {
      const items = [];
      for (; i < lines.length; i++) {
        const mm = lines[i].match(LIST);
        if (mm) items.push({ ind: mm[1].replace(/\t/g, '  ').length, ord: /\d/.test(mm[2]), start: parseInt(mm[2], 10), text: mm[3] });
        else if (/^\s{2,}\S/.test(lines[i])) items[items.length - 1].text += ' ' + lines[i].trim();
        else break;
      }
      out.push(renderList(items));
    } else {
      const buf = [];
      for (; i < lines.length && lines[i].trim() && (!buf.length || !startsBlock(i)); i++) buf.push(lines[i]);
      out.push(`<p>${buf.map(inline).join('<br>')}</p>`);
    }
  }
  return out.join('');
}

// file names in an answer (`outputs/rapor.docx`) open the preview
function linkFiles(root) {
  for (const c of root.querySelectorAll('code')) {
    const name = c.textContent.trim().split('/').pop();
    if (fileRec(name)) {
      c.classList.add('file-link');
      c.title = 'Önizle';
      c.onclick = () => openPreview(name);
    }
  }
}

// ---- health -------------------------------------------------------------------
async function loadHealth(refresh) {
  try { S.health = await api('/api/health' + (refresh ? '?refresh=1' : '')); } catch { return; }
  renderHealth();
  if (S.health.checking) setTimeout(loadHealth, 1500);
}

function renderHealth() {
  const h = S.health;
  const box = $('#health');
  if (h.checking) {
    box.innerHTML = '<div class="hrow"><span class="dot"></span>Sistem kontrol ediliyor…</div>';
    return;
  }
  const row = (ok, label) => `<div class="hrow"><span class="dot ${ok ? 'ok' : 'bad'}"></span>${label}</div>`;
  box.innerHTML = row(h.docker, 'Docker') + row(h.image, 'Agent imajı') + row(h.gotenberg, 'Gotenberg') + row(h.env, '.env') +
    (h.model ? `<div class="model" title="Model">${esc(h.model)}</div>` : '');
  $('#fileInput').accept = (h.uploadable || UPLOADABLE).join(',');

  const problems = [];
  if (!h.docker) problems.push('Docker çalışmıyor — Docker Desktop\'u başlatın.');
  else if (!h.image) problems.push(`<code>${esc(h.image_name)}</code> imajı yok — <code>docker compose build</code> çalıştırın.`);
  if (!h.env) problems.push('<code>.env</code> yok — <code>cp .env.example .env</code> ile model ayarlarını yazın.');
  if (h.image && !h.gotenberg) problems.push('Gotenberg\'e ulaşılamıyor — önizleme ve PDF üretimi çalışmaz (<code>docker compose up -d gotenberg</code>).');
  const b = $('#banner');
  b.hidden = !problems.length;
  if (problems.length) {
    b.innerHTML = `<strong>Sistem hazır değil</strong><ul>${problems.map(p => `<li>${p}</li>`).join('')}</ul><button type="button">Yeniden kontrol et</button>`;
    b.querySelector('button').onclick = () => { S.health = { checking: true }; renderHealth(); loadHealth(true); };
  }
}

// ---- sessions -----------------------------------------------------------------
async function loadSessions() {
  try { S.sessions = await api('/api/sessions'); } catch { return; }
  renderSessions();
}

function renderSessions() {
  const nav = $('#sessions');
  nav.innerHTML = '';
  if (!S.sessions.length) {
    nav.append(el('div', 'empty-list', 'Henüz sohbet yok.'));
    return;
  }
  for (const s of S.sessions) {
    const a = el('a', `sess${s.id === S.sid ? ' active' : ''}`);
    a.href = `#/${s.id}`;
    a.innerHTML = `<span class="sess-title">${s.running ? '<span class="live"></span>' : ''}${esc(s.title)}</span>` +
      `<span class="sess-time">${ago(s.updated)}</span><button type="button" class="sess-del" title="Sohbeti sil">${svg('trash', 14)}</button>`;
    a.onclick = e => {
      e.preventDefault();
      if (e.target.closest('.sess-del')) return deleteSession(s);
      document.body.classList.remove('sidebar-open');
      if (s.id !== S.sid) openSession(s.id);
    };
    nav.append(a);
  }
}

async function deleteSession(s) {
  if (!confirm(`“${s.title}” sohbeti ve tüm dosyaları silinsin mi?`)) return;
  try { await api(`/api/sessions/${s.id}`, { method: 'DELETE' }); } catch (e) { return toast(e.message); }
  if (s.id === S.sid) newChat();
  loadSessions();
}

async function openSession(sid, opts = {}) {
  let sess;
  try { sess = await api(`/api/sessions/${sid}`); } catch (e) { toast(e.message); return newChat(); }
  if (sid !== S.sid) {
    closePreview();
    revokeBlobs();
  }
  S.sid = sid;
  S.sess = sess;
  S.pending = sess.files.filter(f => f.origin === 'upload' && f.turn == null).map(f => f.name);
  history.replaceState(null, '', `#/${sid}`);
  renderAll();
  if (runningTurn()) poll();
  const outs = opts.finished?.outputs || [];
  if (outs.length) openPreview(outs[outs.length - 1]);
  else if (S.preview) renderPreviewHead();
}

function newChat() {
  S.pollToken++;
  S.sid = null;
  S.sess = null;
  S.pending = [];
  closePreview();
  revokeBlobs();
  history.replaceState(null, '', location.pathname);
  renderAll();
  $('#input').focus();
}

async function ensureSession() {
  if (S.sid) return;
  const sess = await api('/api/sessions', { method: 'POST' });
  S.sid = sess.id;
  S.sess = sess;
  S.pending = [];
  history.replaceState(null, '', `#/${sess.id}`);
}

// ---- messages -----------------------------------------------------------------
function renderAll() {
  $('#title').textContent = S.sess?.title || 'Yeni sohbet';
  const box = $('#messages');
  box.innerHTML = '';
  const turns = S.sess?.turns || [];
  if (!turns.length) box.append(emptyState());
  for (const t of turns) {
    box.append(userMsg(t));
    const m = agentMsg(t);
    box.append(m);
    updateAgent(m, t);
  }
  renderContext();
  renderComposer();
  renderFilesBtn();
  renderSessions();
  box.scrollTop = box.scrollHeight;
}

function emptyState() {
  const d = el('div', 'empty');
  d.innerHTML = `<div class="empty-mark">D</div><h2>Hangi belgeyi hazırlayalım?</h2>
    <p>İsteğinizi yazın; Word, PowerPoint, Excel ya da PDF olarak hazırlanıp doğrulanır ve yanda önizlenir.
    Bir dosya yükleyip değişiklik de isteyebilirsiniz: dosya baştan üretilmez, yalnızca istediğiniz yer değişir.</p>
    <div class="formats">${['DOCX', 'PPTX', 'XLSX', 'PDF'].map(f => `<span class="badge f-${f.toLowerCase()}">${f}</span>`).join('')}
    <span class="small">.doc · .xls · .ppt dönüştürülür · şablon: .dotx · .potx · .xltx</span></div>
    <div class="examples"></div>`;
  for (const [kind, text] of EXAMPLES) {
    const b = el('button', 'example', `<small>${esc(kind)}</small>${esc(text)}`);
    b.type = 'button';
    b.onclick = () => { $('#input').value = text; autoGrow(); renderComposer(); $('#input').focus(); };
    $('.examples', d).append(b);
  }
  const up = el('button', 'example', '<small>Dosya düzenleme</small>Bir dosya yükleyin, sonra örneğin “Sonuç başlığını kırmızı yap”, “B4 hücresini 20000 yap” ya da “3. slaytı sil” yazın');
  up.type = 'button';
  up.onclick = () => $('#fileInput').click();
  $('.examples', d).append(up);
  return d;
}

function fileChip(name, onRemove) {
  const c = el('span', 'chip', `<span class="ext-dot f-${family(name)}"></span><span class="chip-name">${esc(name)}</span>`);
  c.title = 'Önizle';
  c.tabIndex = 0;
  c.onclick = () => openPreview(name);
  c.onkeydown = e => { if (e.key === 'Enter') openPreview(name); };
  if (onRemove) {
    const x = el('span', 'chip-x', svg('close', 13));
    x.title = 'Kaldır';
    x.onclick = e => { e.stopPropagation(); onRemove(); };
    c.append(x);
  }
  return c;
}

function userMsg(t) {
  const m = el('div', 'msg user');
  const b = el('div', 'bubble');
  b.append(el('div', 'text', esc(t.text)));
  if (t.attachments.length || t.active || t.confirm) {
    const meta = el('div', 'meta');
    for (const a of t.attachments) meta.append(fileChip(a));
    if (t.active && !t.attachments.includes(t.active)) meta.append(el('span', 'tag', `${svg('eye', 13)}${esc(t.active)}`));
    if (t.confirm) meta.append(el('span', 'tag warn', '--onay'));
    b.append(meta);
  }
  m.append(b);
  return m;
}

function agentMsg(t) {
  const m = el('div', 'msg agent');
  m.dataset.n = t.n;
  m.innerHTML = `<div class="avatar">D</div><div class="body"><div class="status"></div><div class="note" hidden></div>
    <details class="steps"><summary><span class="count"></span><span class="last"></span></summary><ol></ol></details>
    <div class="reply md"></div><div class="error"></div><div class="outputs"></div></div>`;
  $('.steps', m).open = t.status === 'running';
  return m;
}

const STATUS = { running: 'Çalışıyor', done: 'Tamamlandı', error: 'Hata', cancelled: 'Durduruldu', interrupted: 'Yarıda kaldı' };
const KIND = {
  shell: ['terminal', ''], read: ['file', 'Okudu'], write: ['pencil', 'Yazdı'], edit: ['pencil', 'Düzenledi'],
  skill: ['book', 'Skill'], search: ['search', 'Aradı'], tool: ['tool', ''], text: ['msg', ''],
};

function stepEl(s, running) {
  const li = el('li', `step k-${s.k}`);
  const state = s.ok === false ? 'err' : s.ok === true ? 'ok' : running && s.k !== 'text' ? 'run' : '';
  const [icon, verb] = KIND[s.k] || KIND.tool;
  let head;
  if (s.k === 'text') head = `<div class="step-text">${inline(s.label)}</div>`;
  else if (s.k === 'shell') head = `${s.note ? `<div class="step-note">${esc(s.note)}</div>` : ''}<code class="cmd">${esc(s.label)}</code>`;
  else head = `${verb ? `<span class="verb">${verb}</span> ` : ''}<code>${esc(s.label)}</code>`;
  let more = '';
  if (s.detail) more += `<details><summary>içerik</summary><pre>${esc(s.detail)}</pre></details>`;
  if (s.out) more += `<details><summary>${s.ok === false ? 'hata çıktısı' : 'çıktı'}</summary><pre>${esc(s.out)}</pre></details>`;
  li.innerHTML = `<span class="step-ico ${state}">${svg(icon, 13)}</span><div class="step-main">${head}${more}</div>`;
  const cmd = $('.cmd', li);
  if (cmd) cmd.onclick = () => cmd.classList.toggle('full');
  return li;
}

function stepSummary(s) {
  if (!s) return '';
  if (s.k === 'shell') return '$ ' + s.label.split('\n')[0];
  return `${(KIND[s.k] || KIND.tool)[1]} ${s.label}`.trim().split('\n')[0];
}

function updateAgent(m, t) {
  const running = t.status === 'running';
  const st = $('.status', m);
  st.className = `status s-${t.status}`;
  const elapsed = (t.finished || now()) - t.started;
  st.innerHTML = `${running ? '<span class="spinner"></span>' : '<span class="sdot"></span>'}<span class="label">${STATUS[t.status] || t.status}</span>` +
    `<span class="sep">·</span><span class="elapsed" data-started="${t.started}">${fmtDur(elapsed)}</span>` +
    (t.steps.length ? `<span class="sep">·</span><span>${t.steps.length} adım</span>` : '');

  const note = $('.note', m);
  note.hidden = !(running && S.health.streaming === false);
  note.innerHTML = 'Bu imaj adımları canlı yazmıyor; adımlar iş bitince görünecek. Canlı izlemek için imajı yeniden derleyin: <code>docker compose build</code>';

  const det = $('.steps', m);
  det.hidden = !t.steps.length;
  $('.count', det).textContent = `Adımlar (${t.steps.length})`;
  $('.last', det).textContent = running ? stepSummary(t.steps[t.steps.length - 1]) : '';
  const ol = $('ol', det);
  t.steps.forEach((s, i) => {
    const sig = JSON.stringify([s, running && i === t.steps.length - 1]);
    const old = ol.children[i];
    if (old && old._sig === sig) return;
    const li = stepEl(s, running && i === t.steps.length - 1);
    li._sig = sig;
    if (old) {
      if ($('details', old)?.open) $('details', li)?.setAttribute('open', '');
      ol.replaceChild(li, old);
    } else {
      ol.append(li);
    }
  });
  while (ol.children.length > t.steps.length) ol.lastChild.remove();
  if (m.dataset.status === 'running' && !running) det.open = false;
  m.dataset.status = t.status;

  const r = $('.reply', m);
  if (r._src !== t.reply) {
    r.innerHTML = md(t.reply);
    r._src = t.reply;
    linkFiles(r);
  }
  r.hidden = !t.reply;
  const e = $('.error', m);
  e.textContent = t.error || '';
  e.hidden = !t.error;

  const o = $('.outputs', m);
  const sig = (t.outputs || []).join('|');
  if (o._sig !== sig) {
    o.innerHTML = '';
    for (const n of t.outputs || []) o.append(fileCard(n));
    o._sig = sig;
  }
  o.hidden = !(t.outputs || []).length;
}

function fileCard(name) {
  const rec = fileRec(name);
  const c = el('div', `file-card${S.preview === name ? ' active' : ''}`);
  c.dataset.name = name;
  c.innerHTML = `<span class="badge f-${family(name)}">${esc(extOf(name).toUpperCase())}</span>
    <div class="fc-main"><div class="fc-name" title="${esc(name)}">${esc(name)}</div><div class="fc-meta">${rec ? fmtSize(rec.size) : ''}</div></div>
    <button type="button" class="btn">${svg('eye', 15)}<span>Önizle</span></button>
    <a class="btn" href="${fileUrl(name, '?download=1')}" download title="İndir">${svg('download', 15)}<span>İndir</span></a>`;
  c.onclick = e => { if (!e.target.closest('a')) openPreview(name); };
  return c;
}

function markActiveCards() {
  for (const c of document.querySelectorAll('.file-card')) c.classList.toggle('active', c.dataset.name === S.preview);
}

function nearBottom() {
  const b = $('#messages');
  return b.scrollHeight - b.scrollTop - b.clientHeight < 120;
}

// running message: poll until it ends, then reload the chat (new files)
function poll() {
  const sid = S.sid;
  const token = ++S.pollToken;
  const tick = async () => {
    if (token !== S.pollToken) return;
    const t = runningTurn();
    if (!t) return;
    try {
      const nt = await api(`/api/sessions/${sid}/turns/${t.n}`);
      if (token !== S.pollToken) return;
      Object.assign(t, nt);
      if (nt.status !== 'running') {
        await openSession(sid, { finished: nt });
        loadSessions();
        return;
      }
      const m = document.querySelector(`.msg.agent[data-n="${t.n}"]`);
      if (m) {
        const stick = nearBottom();
        updateAgent(m, t);
        if (stick) $('#messages').scrollTop = $('#messages').scrollHeight;
      }
    } catch { /* server restarting: keep trying */ }
    setTimeout(tick, 1000);
  };
  setTimeout(tick, 500);
}

setInterval(() => {
  for (const e of document.querySelectorAll('.s-running .elapsed')) e.textContent = fmtDur(now() - Number(e.dataset.started));
}, 1000);
setInterval(loadSessions, 10000);

// ---- composer ------------------------------------------------------------------
function autoGrow() {
  const i = $('#input');
  i.style.height = 'auto';
  i.style.height = Math.min(i.scrollHeight, 240) + 'px';
}

function renderComposer() {
  const btn = $('#sendBtn');
  const running = !!runningTurn();
  btn.classList.toggle('stop', running);
  btn.innerHTML = svg(running ? 'stop' : 'send', 18);
  btn.title = running ? 'Durdur' : 'Gönder';
  btn.disabled = !running && (!$('#input').value.trim() || S.uploading.length > 0);
}

function renderContext() {
  const box = $('#context');
  box.innerHTML = '';
  for (const u of S.uploading) {
    box.append(el('span', 'chip uploading', `<span class="spinner"></span><span class="chip-name">${esc(u.name)}</span><span class="pct">%${u.pct}</span>`));
  }
  for (const n of S.pending) box.append(fileChip(n, () => removePending(n)));
  if (S.preview && !S.pending.includes(S.preview)) {
    const a = el('span', 'active-file', `${svg('eye', 14)}<span>Açık dosya: <b>${esc(S.preview)}</b></span>`);
    a.title = 'Değişiklik isterken dosya adı vermezseniz bu dosya düzenlenir.';
    box.append(a);
  }
  box.hidden = !box.children.length;
}

function renderFilesBtn() {
  const files = S.sess?.files || [];
  const b = $('#filesBtn');
  b.hidden = !files.length;
  b.innerHTML = `${svg('files', 15)}<span>Dosyalar (${files.length})</span>`;
}

async function send() {
  const input = $('#input');
  const text = input.value.trim();
  if (!text || S.uploading.length || runningTurn()) return;
  try {
    await ensureSession();
    await api(`/api/sessions/${S.sid}/turns`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text, active: S.preview, confirm: $('#confirm').checked }),
    });
  } catch (e) {
    return toast(e.message);
  }
  input.value = '';
  autoGrow();
  $('#confirm').checked = false;
  await openSession(S.sid);
  loadSessions();
}

async function cancel() {
  const t = runningTurn();
  if (!t) return;
  try { await api(`/api/sessions/${S.sid}/turns/${t.n}/cancel`, { method: 'POST' }); } catch (e) { toast(e.message); }
}

function xhrUpload(url, file, onProgress) {
  return new Promise((resolve, reject) => {
    const x = new XMLHttpRequest();
    x.open('PUT', url);
    x.upload.onprogress = e => { if (e.lengthComputable) onProgress(Math.round(e.loaded * 100 / e.total)); };
    x.onload = () => {
      let d = null;
      try { d = JSON.parse(x.responseText); } catch { /* not JSON */ }
      if (x.status < 300) resolve(d);
      else reject(new Error(d?.error || `HTTP ${x.status}`));
    };
    x.onerror = () => reject(new Error('Bağlantı hatası'));
    x.send(file);
  });
}

async function uploadFiles(list) {
  const allowed = new Set(S.health.uploadable || UPLOADABLE);
  for (const f of list) {
    if (!allowed.has('.' + extOf(f.name))) { toast(`${f.name}: desteklenmeyen dosya türü`); continue; }
    if (f.size > 200 * 1024 * 1024) { toast(`${f.name}: en fazla 200 MB`); continue; }
    try { await ensureSession(); } catch (e) { return toast(e.message); }
    const u = { name: f.name, pct: 0 };
    S.uploading.push(u);
    renderContext();
    renderComposer();
    try {
      const rec = await xhrUpload(`/api/sessions/${S.sid}/files?name=${enc(f.name)}`, f, p => { u.pct = p; renderContext(); });
      S.sess.files.push(rec);
      S.pending.push(rec.name);
      openPreview(rec.name);
    } catch (e) {
      toast(`${f.name}: ${e.message}`);
    } finally {
      S.uploading = S.uploading.filter(x => x !== u);
      renderContext();
      renderComposer();
      renderFilesBtn();
    }
  }
  loadSessions();
}

async function removePending(name) {
  try { await api(fileUrl(name), { method: 'DELETE' }); } catch (e) { return toast(e.message); }
  S.pending = S.pending.filter(n => n !== name);
  S.sess.files = S.sess.files.filter(f => f.name !== name);
  const key = `${S.sid}/${name}`;  // the next upload may reuse the name
  if (S.blobs.has(key)) URL.revokeObjectURL(S.blobs.get(key).url);
  S.blobs.delete(key);
  S.sheets.delete(key);
  if (S.preview === name) closePreview();
  renderContext();
  renderFilesBtn();
}

// ---- preview -------------------------------------------------------------------
function revokeBlobs() {
  for (const b of S.blobs.values()) URL.revokeObjectURL(b.url);
  S.blobs.clear();
  S.sheets.clear();
}

function renderPreviewHead() {
  const sel = $('#fileSelect');
  const files = [...(S.sess?.files || [])].reverse();
  sel.innerHTML = '';
  for (const [label, list] of [['Çıktılar', files.filter(f => f.origin === 'output')], ['Yüklenenler', files.filter(f => f.origin === 'upload')]]) {
    if (!list.length) continue;
    const g = el('optgroup');
    g.label = label;
    for (const f of list) {
      const o = el('option');
      o.value = f.name;
      o.textContent = f.origin === 'output' ? `${f.name}  ·  ${f.turn}. yanıt` : f.name;
      g.append(o);
    }
    sel.append(g);
  }
  sel.value = S.preview;
  $('#downloadBtn').href = fileUrl(S.preview, '?download=1');
  $('#openBtn').href = fileUrl(S.preview, '/preview');
  const vt = $('#viewToggle');
  vt.hidden = !isSheet(S.preview);
  $('[data-mode=pdf]', vt).textContent = extOf(S.preview) === 'csv' ? 'Metin' : 'PDF';
  for (const b of vt.querySelectorAll('button')) b.classList.toggle('on', b.dataset.mode === S.pvMode);
  $('#previewFoot').textContent = isSheet(S.preview) && S.pvMode === 'sheet'
    ? 'Tablo görünümü: değerler dosyadaki son hesaplanmış sonuçlardır. Bir hücreye tıklayın; formülü üstte görünür. Grafikler için PDF görünümüne geçin.'
    : 'Ofis dosyaları Gotenberg (LibreOffice) ile PDF\'e çevrilerek gösterilir; Word / PowerPoint / Excel görünümünden küçük farklar olabilir. Asıl dosya için İndir\'i kullanın.';
}

function setPreviewMode(mode) {
  S.pvMode = mode;
  if (S.preview) openPreview(S.preview);
}

async function showSheet(name, key, token) {
  const body = $('#previewBody');
  if (!S.sheets.has(key)) {
    body.innerHTML = '<div class="pv-state"><span class="spinner big"></span><div>Tablo hazırlanıyor…</div></div>';
    try {
      const model = await api(fileUrl(name, '/sheet'));
      S.sheets.set(key, model);
    } catch (e) {
      if (token !== S.previewToken) return;
      body.innerHTML = `<div class="pv-state err"><div>${esc(e.message)}</div><div class="row">
        <button type="button" class="btn" id="pvPdf">PDF görünümüne geç</button></div></div>`;
      $('#pvPdf').onclick = () => setPreviewMode('pdf');
      return;
    }
    if (token !== S.previewToken) return;
  }
  body.innerHTML = '';
  Sheet.mount(body, S.sheets.get(key), { onPdf: () => setPreviewMode('pdf') });
}

async function openPreview(name) {
  if (!S.sid || !fileRec(name)) return;
  S.preview = name;
  $('#preview').hidden = false;
  $('#splitter').hidden = false;
  renderPreviewHead();
  renderContext();
  markActiveCards();
  const body = $('#previewBody');
  const key = `${S.sid}/${name}`;
  const token = ++S.previewToken;
  if (isSheet(name) && S.pvMode === 'sheet') return showSheet(name, key, token);
  if (S.blobs.has(key)) return showBlob(S.blobs.get(key), token);
  const office = (S.health.office || []).includes('.' + extOf(name));
  body.innerHTML = `<div class="pv-state"><span class="spinner big"></span><div>${office ? 'LibreOffice ile PDF\'e çevriliyor…' : 'Yükleniyor…'}</div></div>`;
  try {
    const r = await fetch(fileUrl(name, '/preview'));
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { msg = (await r.json()).error || msg; } catch { /* not JSON */ }
      throw new Error(msg);
    }
    const blob = await r.blob();
    const item = { url: URL.createObjectURL(blob), type: blob.type, blob };
    S.blobs.set(key, item);
    if (token === S.previewToken) showBlob(item, token);
  } catch (e) {
    if (token !== S.previewToken) return;
    body.innerHTML = `<div class="pv-state err"><div>${esc(e.message)}</div><div class="row">
      <button type="button" class="btn" id="pvRetry">Tekrar dene</button>
      <a class="btn" href="${fileUrl(name, '?download=1')}" download>${svg('download', 15)}<span>İndir</span></a></div></div>`;
    $('#pvRetry').onclick = () => openPreview(name);
  }
}

function showBlob(item, token) {
  const body = $('#previewBody');
  body.innerHTML = '';
  if (item.type.includes('pdf')) {
    const f = el('iframe');
    f.title = 'Önizleme';
    f.src = `${item.url}#view=FitH&navpanes=0`;
    body.append(f);
  } else if (item.type.startsWith('image/')) {
    const wrap = el('div', 'pv-img');
    const img = el('img');
    img.src = item.url;
    img.alt = S.preview;
    wrap.append(img);
    body.append(wrap);
  } else if (item.type.startsWith('text/') || item.type.includes('json')) {
    item.blob.text().then(t => {
      if (token !== S.previewToken) return;
      const pre = el('pre', 'pv-text');
      pre.textContent = t.length > 400000 ? t.slice(0, 400000) + '\n…' : t;
      body.append(pre);
    });
  } else {
    body.innerHTML = '<div class="pv-state">Bu dosya türü önizlenemiyor; İndir ile açın.</div>';
  }
}

function closePreview() {
  S.preview = null;
  S.previewToken++;
  $('#preview').hidden = true;
  $('#splitter').hidden = true;
  $('#previewBody').innerHTML = '';
  renderContext();
  markActiveCards();
}

// drag the splitter to resize the preview
(() => {
  const root = document.documentElement;
  try {
    const w = localStorage.getItem('docgen.previewWidth');
    if (w) root.style.setProperty('--preview-w', w);
  } catch { /* storage blocked */ }
  $('#splitter').addEventListener('pointerdown', e => {
    e.preventDefault();
    document.body.classList.add('dragging');
    const move = ev => {
      const w = Math.max(320, Math.min(window.innerWidth * 0.75, window.innerWidth - ev.clientX));
      root.style.setProperty('--preview-w', `${Math.round(w)}px`);
    };
    const up = () => {
      document.body.classList.remove('dragging');
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      try { localStorage.setItem('docgen.previewWidth', root.style.getPropertyValue('--preview-w')); } catch { /* storage blocked */ }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  });
})();

// ---- wiring --------------------------------------------------------------------
$('#newChat').innerHTML = `${svg('plus', 16)}<span>Yeni sohbet</span>`;
$('#menuBtn').innerHTML = svg('menu', 18);
$('#attachBtn').innerHTML = svg('clip', 18);
$('#openBtn').innerHTML = svg('external', 17);
$('#downloadBtn').innerHTML = svg('download', 17);
$('#closePreview').innerHTML = svg('close', 18);

$('#newChat').onclick = () => { document.body.classList.remove('sidebar-open'); newChat(); };
$('#menuBtn').onclick = () => document.body.classList.toggle('sidebar-open');
$('#scrim').onclick = () => document.body.classList.remove('sidebar-open');
$('#closePreview').onclick = closePreview;
$('#fileSelect').onchange = e => openPreview(e.target.value);
$('#viewToggle').onclick = e => {
  const b = e.target.closest('button');
  if (b && b.dataset.mode !== S.pvMode) setPreviewMode(b.dataset.mode);
};
$('#filesBtn').onclick = () => {
  if (S.preview) return closePreview();
  const files = S.sess?.files || [];
  const last = [...files].reverse().find(f => f.origin === 'output') || files[files.length - 1];
  if (last) openPreview(last.name);
};
$('#attachBtn').onclick = () => $('#fileInput').click();
$('#fileInput').onchange = e => { uploadFiles([...e.target.files]); e.target.value = ''; };
$('#composer').onsubmit = e => { e.preventDefault(); runningTurn() ? cancel() : send(); };
$('#input').addEventListener('input', () => { autoGrow(); renderComposer(); });
$('#input').addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    if (!runningTurn()) send();
  }
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && S.preview && document.activeElement !== $('#input')) closePreview();
});

let dragDepth = 0;
const hasFiles = e => [...(e.dataTransfer?.types || [])].includes('Files');
window.addEventListener('dragenter', e => { if (hasFiles(e)) { dragDepth++; $('#drop').hidden = false; } });
window.addEventListener('dragleave', e => { if (hasFiles(e) && --dragDepth <= 0) { dragDepth = 0; $('#drop').hidden = true; } });
window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
window.addEventListener('drop', e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  $('#drop').hidden = true;
  uploadFiles([...e.dataTransfer.files]);
});

const m = location.hash.match(/^#\/([0-9a-f]{12})$/);
if (m) openSession(m[1]); else renderAll();
loadSessions();
loadHealth();
