/* Platform365 EasyWrite 2.0.0 — tarayıcı tabanlı kelime işlemci */
'use strict';

const APP_VERSION = '2.0.0';
const FILE_FORMAT = 'easywrite';
const P = window.P365;
const C = window.Convert;

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
const trLower = (s) => s.toLocaleLowerCase('tr');
const trUpper = (s) => s.toLocaleUpperCase('tr');

/* ------------------------------------------------------------------ */
/* Ayarlar                                                             */
/* ------------------------------------------------------------------ */

const SETTINGS_KEY = 'ew-settings';
const settings = Object.assign({ smart: true, markdown: true, typewriter: false }, (() => {
  try { return JSON.parse(localStorage.getItem(SETTINGS_KEY)) || {}; } catch (e) { return {}; }
})());
const saveSettings = () => { try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch (e) { /* özel pencere */ } };

const PAGE_SIZES = { A4: [794, 1123], Letter: [816, 1056], A5: [559, 794] };
const MARGINS = { narrow: 48, normal: 96, wide: 144 };
const PAPERS = { white: 'Beyaz', sepia: 'Sepya', mint: 'Nane', dark: 'Gece' };
const FONTS = ['Calibri', 'Arial', 'Georgia', 'Times New Roman', 'Verdana', 'Trebuchet MS', 'Garamond', 'Courier New', 'Segoe UI'];

/* ------------------------------------------------------------------ */
/* Şablonlar                                                           */
/* ------------------------------------------------------------------ */

const today = () => new Date().toLocaleDateString('tr-TR', { dateStyle: 'long' });
const TEMPLATES = [
  { id: 'blank', name: 'Boş belge', icon: '📄', desc: 'Sıfırdan başlayın', html: () => '<p><br></p>' },
  { id: 'letter', name: 'Mektup', icon: '✉️', desc: 'Resmî yazı / dilekçe', html: () => `<p style="text-align: right">${today()}</p><p><b>Sayın [Alıcı Adı],</b></p><p>[Mektubunuzun amacını belirten giriş paragrafı.]</p><p>[Ayrıntılar ve gerekçeler.]</p><p>Gereğini bilgilerinize arz ederim.</p><p><br></p><p style="text-align: right">Saygılarımla,<br><b>[Adınız Soyadınız]</b><br>[İletişim bilgileri]</p>` },
  { id: 'cv', name: 'Özgeçmiş', icon: '🧑‍💼', desc: 'Sade ve modern CV', html: () => '<h1>Ad Soyad</h1><p>📧 eposta@ornek.com · 📱 +90 5xx xxx xx xx · 📍 Şehir · 🔗 linkedin.com/in/...</p><hr><h2>Özet</h2><p>Deneyiminizi ve hedeflerinizi anlatan 2–3 cümle.</p><h2>Deneyim</h2><h3>Pozisyon — Şirket</h3><p><i>2022 – Günümüz</i></p><ul><li>Somut bir başarı (rakamlarla)</li><li>Sorumluluk alanı</li></ul><h2>Eğitim</h2><h3>Bölüm — Üniversite</h3><p><i>2016 – 2020</i></p><h2>Yetenekler</h2><table><tr><th>Teknik</th><th>Diller</th></tr><tr><td>Excel, Python, Figma</td><td>Türkçe (ana dil), İngilizce (C1)</td></tr></table>' },
  { id: 'meeting', name: 'Toplantı notu', icon: '🗒️', desc: 'Gündem, kararlar, görevler', html: () => `<h1>Toplantı notu</h1><p><b>Tarih:</b> ${today()} · <b>Katılımcılar:</b> @...</p><h2>Gündem</h2><ol><li>Konu 1</li><li>Konu 2</li></ol><h2>Kararlar</h2><div data-callout="success"><p>Alınan karar…</p></div><h2>Görevler</h2><ul class="checklist"><li data-check="0">Görev — Sorumlu — Tarih</li><li data-check="0">Görev — Sorumlu — Tarih</li></ul><h2>Notlar</h2><p><br></p>` },
  { id: 'report', name: 'Rapor', icon: '📊', desc: 'İçindekiler ve bölümler', html: () => `<h1>Rapor başlığı</h1><p><i>Hazırlayan: [Ad] · ${today()}</i></p><nav data-toc contenteditable="false"></nav><h2>Yönetici özeti</h2><div data-callout="info"><p>Raporun en önemli üç bulgusu.</p></div><h2>Giriş</h2><p>Amaç ve kapsam.</p><h2>Bulgular</h2><h3>Bulgu 1</h3><p>Açıklama.</p><table><tr><th>Ölçüt</th><th>Değer</th><th>Değişim</th></tr><tr><td>Örnek</td><td>120</td><td>+%8</td></tr></table><h2>Sonuç ve öneriler</h2><ul><li>Öneri 1</li><li>Öneri 2</li></ul>` },
  { id: 'essay', name: 'Ödev / Makale', icon: '🎓', desc: 'Akademik düzen', html: () => `<h1 style="text-align: center">Makale Başlığı</h1><p style="text-align: center">Ad Soyad · Kurum · ${today()}</p><h2>Özet</h2><p style="text-align: justify">Çalışmanın kısa özeti (150–250 kelime).</p><p><b>Anahtar kelimeler:</b> …</p><h2>1. Giriş</h2><p style="text-align: justify"><br></p><h2>2. Yöntem</h2><p style="text-align: justify"><br></p><h2>3. Bulgular</h2><p style="text-align: justify"><br></p><h2>4. Tartışma</h2><p style="text-align: justify"><br></p><h2>Kaynakça</h2><ol><li>Yazar, A. (Yıl). <i>Eser adı</i>. Yayınevi.</li></ol>` },
  { id: 'plan', name: 'Proje planı', icon: '🚀', desc: 'Hedefler, takvim, riskler', html: () => '<h1>Proje planı</h1><h2>Hedefler</h2><ul><li>Ölçülebilir hedef 1</li><li>Ölçülebilir hedef 2</li></ul><h2>Takvim</h2><table><tr><th>Aşama</th><th>Başlangıç</th><th>Bitiş</th><th>Sorumlu</th></tr><tr><td>Analiz</td><td></td><td></td><td></td></tr><tr><td>Tasarım</td><td></td><td></td><td></td></tr><tr><td>Uygulama</td><td></td><td></td><td></td></tr></table><h2>Yapılacaklar</h2><ul class="checklist"><li data-check="1">Proje başlatıldı</li><li data-check="0">Ekip toplantısı</li></ul><h2>Riskler</h2><div data-callout="warning"><p>Olası risk ve önlemi.</p></div><details open><summary>Bütçe ayrıntıları</summary><p>Kalemler…</p></details>' },
  { id: 'journal', name: 'Günlük', icon: '📔', desc: 'Bugün ne oldu?', html: () => `<h1>${today()}</h1><h3>Bugün için minnettar olduklarım</h3><ol><li><br></li></ol><h3>Bugün neler oldu?</h3><p><br></p><h3>Yarın için niyetim</h3><ul class="checklist"><li data-check="0"><br></li></ul>` },
];

/* ------------------------------------------------------------------ */
/* Depolama                                                            */
/* ------------------------------------------------------------------ */

const Store = {
  db: null,
  mem: new Map(),
  async open() {
    if (!('indexedDB' in window)) return;
    try {
      this.db = await new Promise((res, rej) => {
        const r = indexedDB.open('easywrite-db', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('docs', { keyPath: 'id' });
        r.onsuccess = () => res(r.result);
        r.onerror = () => rej(r.error);
      });
    } catch (e) {
      console.warn('IndexedDB kullanılamıyor, bellek kullanılacak', e);
      this.db = null;
    }
  },
  req(mode, fn) {
    return new Promise((res, rej) => {
      const r = fn(this.db.transaction('docs', mode).objectStore('docs'));
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  },
  async all() { return this.db ? this.req('readonly', (s) => s.getAll()) : [...this.mem.values()].map((d) => ({ ...d })); },
  async get(id) { return this.db ? this.req('readonly', (s) => s.get(id)) : (this.mem.get(id) ? { ...this.mem.get(id) } : null); },
  async put(d) { return this.db ? this.req('readwrite', (s) => s.put(d)) : this.mem.set(d.id, { ...d }); },
  async del(id) { return this.db ? this.req('readwrite', (s) => s.delete(id)) : this.mem.delete(id); },
};

function newDoc(title = 'Adsız belge', html = '<p><br></p>') {
  const now = Date.now();
  return { id: P.uid(), title, html, created: now, updated: now, version: APP_VERSION, page: { size: 'A4', margin: 'normal', paper: 'white', font: 'Calibri', fontSize: 16 }, goal: 0 };
}

/* ------------------------------------------------------------------ */
/* Durum                                                               */
/* ------------------------------------------------------------------ */

let doc = null;
let saveTimer = null;
let statsTimer = null;
let lastRange = null;
let sessionStartWords = 0;
const page = () => $('#page');

function scheduleSave() {
  if (!doc) return;
  $('#saveState').textContent = 'Kaydediliyor…';
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveNow, 500);
}
function serialize() {
  const c = page().cloneNode(true);
  $$('.is-active', c).forEach((n) => n.classList.remove('is-active'));
  $$('img.sel', c).forEach((n) => n.classList.remove('sel'));
  $$('[class=""]', c).forEach((n) => n.removeAttribute('class'));
  return c.innerHTML;
}
async function saveNow() {
  clearTimeout(saveTimer);
  if (!doc) return;
  doc.html = serialize();
  doc.updated = Date.now();
  doc.version = APP_VERSION;
  try {
    await Store.put(doc);
    $('#saveState').textContent = 'Kaydedildi';
  } catch (e) {
    console.error(e);
    $('#saveState').textContent = 'Kaydedilemedi!';
    P.toast('Kaydedilemedi: depolama alanı dolu olabilir.');
  }
}
function changed() { scheduleSave(); scheduleStats(); }

/* ------------------------------------------------------------------ */
/* Ana sayfa                                                           */
/* ------------------------------------------------------------------ */

function renderTemplates() {
  $('#tplGrid').innerHTML = TEMPLATES.map((t) => `<button class="tpl-card" data-tpl="${t.id}"><div class="tp"><span class="ico">${t.icon}</span>${t.id === 'blank' ? '' : C.sanitize(t.html()).slice(0, 600)}</div><b>${esc(t.name)}</b><small>${esc(t.desc)}</small></button>`).join('');
}

async function renderHome() {
  const docs = (await Store.all()).sort((a, b) => b.updated - a.updated);
  const grid = $('#docGrid');
  grid.innerHTML = '';
  $('#emptyHint').hidden = docs.length > 0;
  for (const d of docs) {
    const card = document.createElement('div');
    card.className = 'doc-card';
    card.tabIndex = 0;
    const words = (d.html.replace(/<[^>]+>/g, ' ').match(/[\p{L}\p{N}]+/gu) || []).length;
    card.innerHTML = `
      <div class="thumb"></div>
      <div class="info">
        <div class="name">${esc(d.title || 'Adsız belge')}</div>
        <div class="meta">${words} kelime · ${P.relTime(d.updated)}</div>
      </div>
      <div class="card-actions">
        <button data-act="dup" title="Kopyasını oluştur">⧉</button>
        <button data-act="dl" title="İndir (.docx)">⬇</button>
        <button data-act="del" title="Sil">🗑</button>
      </div>`;
    $('.thumb', card).innerHTML = d.html.length > 60000 ? C.sanitize(d.html.slice(0, 60000)) : d.html;
    card.addEventListener('click', async (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (!act) return openDoc(d.id);
      e.stopPropagation();
      if (act === 'del') {
        if (confirm(`“${d.title}” silinsin mi? Bu işlem geri alınamaz.`)) { await Store.del(d.id); renderHome(); }
      } else if (act === 'dup') {
        const copy = { ...d, id: P.uid(), title: `${d.title} (kopya)`, created: Date.now(), updated: Date.now() };
        await Store.put(copy);
        renderHome();
      } else if (act === 'dl') exportAs('docx', d);
    });
    card.addEventListener('keydown', (e) => { if (e.key === 'Enter') openDoc(d.id); });
    grid.appendChild(card);
  }
}

function showHome() {
  document.body.classList.remove('in-editor', 'focus-mode');
  $('#editorView').hidden = true;
  $('#homeView').hidden = false;
  doc = null;
  closeFind();
  stopDictation();
  renderHome();
}

async function openDoc(id) {
  if (location.hash !== `#/d/${id}`) { location.hash = `#/d/${id}`; return; }
  const d = await Store.get(id);
  if (!d) { P.toast('Belge bulunamadı.'); location.hash = ''; return; }
  doc = d;
  doc.page = Object.assign({ size: 'A4', margin: 'normal', paper: 'white', font: 'Calibri', fontSize: 16 }, doc.page || {});
  document.body.classList.add('in-editor');
  $('#homeView').hidden = true;
  $('#editorView').hidden = false;
  $('#docTitle').value = doc.title;
  $('#saveState').textContent = 'Kaydedildi';
  const pg = page();
  pg.innerHTML = doc.html || '<p><br></p>';
  if (!pg.firstElementChild) pg.innerHTML = '<p><br></p>';
  applyPage();
  autoZoom();
  refreshStructure(true);
  updateStats();
  sessionStartWords = lastStats ? lastStats.words : 0;
  pg.focus({ preventScroll: true });
  const r = document.createRange();
  r.selectNodeContents(pg.firstElementChild || pg);
  r.collapse(true);
  getSelection().removeAllRanges();
  getSelection().addRange(r);
  $('#pageWrap').scrollTop = 0;
  try {
    const list = await P.snapshots.list('easywrite', doc.id);
    if (!list.length || Date.now() - list[0].at > 30 * 60 * 1000) await P.snapshots.save('easywrite', doc.id, doc.title, { title: doc.title, html: doc.html, page: doc.page });
  } catch (e) { /* sürüm yok */ }
}

async function createDoc(tplId = 'blank') {
  const t = TEMPLATES.find((x) => x.id === tplId) || TEMPLATES[0];
  const d = newDoc(tplId === 'blank' ? 'Adsız belge' : t.name, t.html());
  await Store.put(d);
  openDoc(d.id);
}

async function route() {
  if (doc) await saveNow();
  const h = location.hash;
  let m;
  if ((m = /^#\/new(?:\/(\w+))?$/.exec(h))) {
    const t = TEMPLATES.find((x) => x.id === m[1]) || TEMPLATES[0];
    const d = newDoc(t.id === 'blank' ? 'Adsız belge' : t.name, t.html());
    await Store.put(d);
    location.replace(`#/d/${d.id}`);
    return;
  }
  if ((m = /^#\/inbox\/(.+)$/.exec(h))) { handleInbox(decodeURIComponent(m[1])); return; }
  if ((m = /^#\/d\/(.+)$/.exec(h))) openDoc(decodeURIComponent(m[1]));
  else showHome();
}

async function handleInbox(id) {
  let item = null;
  try { item = await P.inbox.take(id); } catch (e) { /* yok */ }
  if (!item) { location.replace('#'); return; }
  if (item.kind === 'html') {
    const d = newDoc(item.title || 'İçe aktarılan belge', C.sanitize(item.html || '<p><br></p>'));
    await Store.put(d);
    location.replace(`#/d/${d.id}`);
    P.toast('Belge oluşturuldu.');
  } else if (item.kind === 'file') {
    window.history.replaceState(null, '', location.pathname);
    showHome();
    importFile(new File([item.blob], item.name, { type: item.type || '' }));
  } else if (item.kind === 'new') {
    location.replace(`#/new/${item.tpl || 'blank'}`);
  } else location.replace('#');
}

/* ------------------------------------------------------------------ */
/* Sayfa düzeni ve yakınlaştırma                                       */
/* ------------------------------------------------------------------ */

function applyPage() {
  const pg = page();
  const [w, h] = PAGE_SIZES[doc.page.size] || PAGE_SIZES.A4;
  pg.style.setProperty('--page-w', `${w}px`);
  pg.style.setProperty('--page-h', `${h}px`);
  pg.style.setProperty('--margin', `${Math.min(MARGINS[doc.page.margin] || 96, w / 4)}px`);
  pg.style.setProperty('--doc-font', `"${doc.page.font || 'Calibri'}"`);
  pg.style.setProperty('--doc-size', `${doc.page.fontSize || 16}px`);
  pg.className = `page${doc.page.paper && doc.page.paper !== 'white' ? ` paper-${doc.page.paper}` : ''}${pg.classList.contains('page-fluid') ? ' page-fluid' : ''}`;
}
function setZoom(v) {
  v = clamp(Math.round(v), 40, 200);
  $('#pageZoom').style.zoom = v / 100;
  $('#zoom').value = v;
  $('#zoomVal').textContent = `${v}%`;
}
function autoZoom() {
  const wrap = $('#pageWrap');
  const [w] = PAGE_SIZES[doc.page.size] || PAGE_SIZES.A4;
  const avail = wrap.clientWidth - 20;
  // Telefon gibi dar ekranlarda sayfa küçültülmez, ekran genişliğine akar
  const fluid = avail > 0 && avail < w * 0.75;
  page().classList.toggle('page-fluid', fluid);
  setZoom(fluid || avail <= 0 || avail >= w ? 100 : Math.floor((avail / w) * 20) * 5);
}

async function pageSettings() {
  const p = doc.page;
  const opt = (obj, cur) => Object.entries(obj).map(([k, v]) => `<option value="${k}"${k === cur ? ' selected' : ''}>${esc(v)}</option>`).join('');
  const r = await P.dialog({
    title: 'Sayfa düzeni ve tercihler',
    html: `<div class="p365-keys" style="grid-template-columns: 140px 1fr; gap: 10px 14px">
      <span>Kağıt boyutu</span><select id="psSize">${opt({ A4: 'A4 (21 × 29,7 cm)', Letter: 'Letter (ABD)', A5: 'A5 (14,8 × 21 cm)' }, p.size)}</select>
      <span>Kenar boşluğu</span><select id="psMargin">${opt({ narrow: 'Dar (1,27 cm)', normal: 'Normal (2,54 cm)', wide: 'Geniş (3,81 cm)' }, p.margin)}</select>
      <span>Kağıt rengi</span><select id="psPaper">${opt(PAPERS, p.paper)}</select>
      <span>Yazı tipi</span><select id="psFont">${FONTS.map((f) => `<option${f === p.font ? ' selected' : ''}>${f}</option>`).join('')}</select>
      <span>Yazı boyutu</span><select id="psFs">${[10, 11, 12, 13, 14, 16].map((pt) => `<option value="${Math.round((pt * 4) / 3)}"${Math.round((pt * 4) / 3) === p.fontSize ? ' selected' : ''}>${pt} pt</option>`).join('')}</select>
      <span>Akıllı yazım</span><label class="chk"><input type="checkbox" id="psSmart"${settings.smart ? ' checked' : ''}> “tırnak”, — tire, …, ©, →, satır içi hesap (12*3=)</label>
      <span>Markdown</span><label class="chk"><input type="checkbox" id="psMd"${settings.markdown ? ' checked' : ''}> # başlık, - madde, **kalın** gibi kısayollar</label>
    </div><p class="muted">Kağıt rengi yalnızca ekranda görünür; yazdırırken beyaz kullanılır.</p>`,
    actions: [{ label: 'Vazgeç', value: null }, {
      label: 'Uygula', primary: true,
      value: (b) => ({ size: $('#psSize', b).value, margin: $('#psMargin', b).value, paper: $('#psPaper', b).value, font: $('#psFont', b).value, fontSize: +$('#psFs', b).value, smart: $('#psSmart', b).checked, md: $('#psMd', b).checked }),
    }],
  });
  if (!r) return;
  Object.assign(doc.page, { size: r.size, margin: r.margin, paper: r.paper, font: r.font, fontSize: r.fontSize });
  settings.smart = r.smart;
  settings.markdown = r.md;
  saveSettings();
  applyPage();
  autoZoom();
  changed();
}

/* ------------------------------------------------------------------ */
/* Seçim yardımcıları                                                  */
/* ------------------------------------------------------------------ */

const BLOCK_RE = /^(P|H[1-6]|LI|PRE|BLOCKQUOTE|DIV|TD|TH|SUMMARY|DETAILS|FIGCAPTION)$/;
function closestBlock(n) {
  const pg = page();
  while (n && n !== pg) {
    if (n.nodeType === 1 && BLOCK_RE.test(n.tagName)) return n;
    n = n.parentNode;
  }
  return null;
}
const inPage = (n) => n && page().contains(n);
function saveSel() {
  const s = getSelection();
  if (s.rangeCount && inPage(s.anchorNode)) lastRange = s.getRangeAt(0).cloneRange();
}
function restoreSel() {
  const pg = page();
  if (document.activeElement !== pg) pg.focus({ preventScroll: true });
  if (lastRange && inPage(lastRange.startContainer)) {
    const s = getSelection();
    s.removeAllRanges();
    s.addRange(lastRange);
  }
}
function exec(cmd, val = null) {
  restoreSel();
  document.execCommand(cmd, false, val);
  saveSel();
  changed();
  updateToolbar();
}
function insertHtml(html) { exec('insertHTML', html); }
/** HTML ekler ve imleci {c} işaretinin yerine koyar; replaceNode verilirse o düğümün yerine geçer */
function insertHtmlCaret(html, replaceNode = null) {
  restoreSel();
  const s = getSelection();
  if (replaceNode) {
    const r = document.createRange();
    r.selectNode(replaceNode);
    s.removeAllRanges();
    s.addRange(r);
  }
  document.execCommand('insertHTML', false, html.replace('{c}', '\u2063'));
  const w = document.createTreeWalker(page(), NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) {
    const i = n.data.indexOf('\u2063');
    if (i < 0) continue;
    n.deleteData(i, 1);
    const r = document.createRange();
    if (!n.data && n.parentNode && !n.parentNode.textContent) {
      const br = document.createElement('br');
      n.replaceWith(br);
      r.setStartBefore(br);
    } else r.setStart(n, i);
    r.collapse(true);
    s.removeAllRanges();
    s.addRange(r);
    break;
  }
  saveSel();
  changed();
  updateToolbar();
}
/** Liste oluşturur: ul, ol veya check (yapılacaklar). Chrome'un boş paragrafta iç içe liste hatasını önler. */
function listify(kind) {
  restoreSel();
  const ctx = caretCtx();
  const b = ctx && ctx.block;
  const isP = b && b.tagName === 'P' && b.parentNode === page();
  const empty = isP && !b.textContent.replace(/[\u200b\u2063]/g, '').trim() && !b.querySelector('img');
  if (kind === 'check' && isP) {
    insertHtmlCaret(`<ul class="checklist"><li data-check="0">${empty ? '' : b.innerHTML.replace(/<br>$/, '')}{c}</li></ul>`, b);
    return;
  }
  if (empty) {
    insertHtmlCaret(kind === 'ol' ? '<ol><li>{c}</li></ol>' : '<ul><li>{c}</li></ul>', b);
    return;
  }
  document.execCommand(kind === 'ol' ? 'insertOrderedList' : 'insertUnorderedList');
  if (kind === 'check') {
    const c2 = caretCtx();
    const ul = c2 && c2.block.closest('ul');
    if (ul) { ul.classList.add('checklist'); $$(':scope > li', ul).forEach((li) => { if (!li.dataset.check) li.dataset.check = '0'; }); }
  }
  saveSel();
  changed();
  updateToolbar();
}
function caretCtx() {
  const s = getSelection();
  if (!s.rangeCount || !s.isCollapsed) return null;
  const r = s.getRangeAt(0);
  if (!inPage(r.startContainer)) return null;
  const block = closestBlock(r.startContainer);
  if (!block) return null;
  const pre = document.createRange();
  pre.setStart(block, 0);
  pre.setEnd(r.startContainer, r.startOffset);
  return { s, r, block, pre, before: pre.toString() };
}
/** İmlecin hemen önündeki n karakteri text ile değiştirir (geri alınabilir) */
function replaceBefore(n, text) {
  const s = getSelection();
  if (!s.rangeCount) return false;
  const r = s.getRangeAt(0);
  if (r.startContainer.nodeType !== 3 || r.startOffset < n) return false;
  const rr = document.createRange();
  rr.setStart(r.startContainer, r.startOffset - n);
  rr.setEnd(r.startContainer, r.startOffset);
  s.removeAllRanges();
  s.addRange(rr);
  if (text.startsWith('<')) document.execCommand('insertHTML', false, text);
  else document.execCommand('insertText', false, text);
  return true;
}
function selectedBlocks() {
  const s = getSelection();
  if (!s.rangeCount) return [];
  const r = s.getRangeAt(0);
  const all = $$('p, h1, h2, h3, h4, h5, h6, li, pre, td, th, summary', page()).filter((b) => r.intersectsNode(b));
  const leaf = all.filter((b) => !all.some((o) => o !== b && b.contains(o) && o.tagName !== 'LI'));
  if (leaf.length) return leaf;
  const b = closestBlock(r.startContainer);
  return b ? [b] : [];
}
function topBlock(n) {
  const pg = page();
  while (n && n.parentNode !== pg) n = n.parentNode;
  return n && n.parentNode === pg ? n : null;
}

/* ------------------------------------------------------------------ */
/* Biçimlendirme                                                       */
/* ------------------------------------------------------------------ */

function formatBlock(tag) {
  restoreSel();
  const ctx = caretCtx();
  if (ctx && ctx.block.tagName === 'LI' && tag !== 'p') {
    // Liste maddesini önce listeden çıkar
    document.execCommand(ctx.block.parentNode.tagName === 'OL' ? 'insertOrderedList' : 'insertUnorderedList');
  }
  document.execCommand('formatBlock', false, `<${tag}>`);
  saveSel();
  changed();
  updateToolbar();
}
function setFontSize(pt) {
  restoreSel();
  document.execCommand('fontSize', false, '7');
  $$('font[size="7"]', page()).forEach((f) => {
    const span = document.createElement('span');
    span.style.fontSize = `${pt}pt`;
    while (f.firstChild) span.appendChild(f.firstChild);
    f.replaceWith(span);
    $$('span', span).forEach((inner) => { if (inner.style.fontSize) inner.style.fontSize = ''; if (!inner.getAttribute('style')) inner.replaceWith(...inner.childNodes); });
  });
  saveSel();
  changed();
}
function setHilite(color) {
  restoreSel();
  document.execCommand('styleWithCSS', false, true);
  if (!document.execCommand('hiliteColor', false, color)) document.execCommand('backColor', false, color);
  document.execCommand('styleWithCSS', false, false);
  saveSel();
  changed();
}
function setLineHeight(v) {
  restoreSel();
  selectedBlocks().forEach((b) => { b.style.lineHeight = v || ''; if (!b.getAttribute('style')) b.removeAttribute('style'); });
  changed();
}
function toggleChecklist() {
  restoreSel();
  const ctx = caretCtx() || {};
  let ul = ctx.block && ctx.block.closest('ul');
  if (ul && ul.classList.contains('checklist')) {
    ul.classList.remove('checklist');
    $$(':scope > li', ul).forEach((li) => li.removeAttribute('data-check'));
    if (!ul.className) ul.removeAttribute('class');
  } else {
    if (!ul) { listify('check'); return; }
    if (ul) {
      ul.classList.add('checklist');
      $$(':scope > li', ul).forEach((li) => { if (!li.dataset.check) li.dataset.check = '0'; });
    }
  }
  changed();
}

async function insertLink() {
  const sel = lastRange && !lastRange.collapsed ? lastRange.toString() : '';
  const url = await P.prompt('Bağlantı ekle', 'https://', sel ? `“${sel.slice(0, 60)}” metnine bağlantı verilecek.` : 'Adres metin olarak eklenecek.');
  if (!url) return;
  const u = url.trim();
  if (!/^(https?:|mailto:|tel:|#)/i.test(u)) { P.toast('Geçerli bir adres girin (https://…).'); return; }
  if (sel) exec('createLink', u);
  else insertHtml(`<a href="${esc(u)}">${esc(u)}</a>&nbsp;`);
  $$('a', page()).forEach((a) => { a.target = '_blank'; a.rel = 'noopener noreferrer'; });
}

function readImageFile(file) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => {
      const img = new Image();
      img.onload = () => {
        const max = 1600;
        const r = Math.min(1, max / Math.max(img.naturalWidth || 1, img.naturalHeight || 1));
        const c = document.createElement('canvas');
        c.width = Math.max(1, Math.round((img.naturalWidth || 800) * r));
        c.height = Math.max(1, Math.round((img.naturalHeight || 600) * r));
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        res(c.toDataURL(file.type === 'image/jpeg' ? 'image/jpeg' : 'image/png', 0.88));
      };
      img.onerror = () => rej(new Error('Görsel okunamadı'));
      img.src = fr.result;
    };
    fr.onerror = () => rej(fr.error);
    fr.readAsDataURL(file);
  });
}
async function insertImageFile(file) {
  if (!file || !file.type.startsWith('image/')) return;
  try {
    const src = await readImageFile(file);
    insertHtml(`<img src="${src}" alt="" style="width: 60%">`);
  } catch (e) { P.toast(e.message); }
}

async function insertTable() {
  const r = await P.dialog({
    title: 'Tablo ekle',
    html: `<div class="p365-keys" style="grid-template-columns: 120px 1fr"><span>Satır</span><input type="number" id="tRows" value="3" min="1" max="50"><span>Sütun</span><input type="number" id="tCols" value="3" min="1" max="12"><span>Başlık satırı</span><label class="chk"><input type="checkbox" id="tHead" checked> İlk satır başlık</label></div>`,
    actions: [{ label: 'Vazgeç', value: null }, { label: 'Ekle', primary: true, value: (b) => ({ r: clamp(+$('#tRows', b).value || 3, 1, 50), c: clamp(+$('#tCols', b).value || 3, 1, 12), h: $('#tHead', b).checked }) }],
  });
  if (!r) return;
  let html = '<table>';
  for (let i = 0; i < r.r; i++) html += `<tr>${Array.from({ length: r.c }, (_, j) => { const t = i === 0 && r.h ? 'th' : 'td'; return `<${t}>${i === 0 && j === 0 ? '{c}' : '<br>'}</${t}>`; }).join('')}</tr>`;
  insertHtmlCaret(`${html}</table><p><br></p>`);
}

function insertToc() {
  const existing = $('nav[data-toc]', page());
  if (existing) { refreshStructure(true); existing.scrollIntoView({ block: 'center' }); P.toast('İçindekiler güncellendi.'); return; }
  insertHtmlCaret('<nav data-toc contenteditable="false"></nav><p>{c}</p>');
  refreshStructure(true);
}

const CALLOUTS = ['info', 'success', 'warning', 'danger'];
function insertCallout(kind = 'info') {
  const txt = lastRange && !lastRange.collapsed ? esc(lastRange.toString()) : '';
  insertHtmlCaret(`<div data-callout="${kind}"><p>${txt}{c}</p></div><p><br></p>`);
}

/* ------------------------------------------------------------------ */
/* Araç çubuğu durumu                                                  */
/* ------------------------------------------------------------------ */

function updateToolbar() {
  if (!doc) return;
  const s = getSelection();
  if (!s.rangeCount || !inPage(s.anchorNode)) return;
  for (const b of $$('#ribbon [data-cmd]')) {
    let on = false;
    try { on = document.queryCommandState(b.dataset.cmd); } catch (e) { /* desteklenmiyor */ }
    b.classList.toggle('on', on);
  }
  const block = closestBlock(s.anchorNode);
  const tag = block ? block.tagName.toLowerCase() : 'p';
  const bs = $('#blockStyle');
  bs.value = ['h1', 'h2', 'h3', 'pre'].includes(tag) ? tag : block && block.closest('blockquote') ? 'blockquote' : 'p';
  const el = s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentElement;
  if (el) {
    const cs = getComputedStyle(el);
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const fn = $('#fontName');
    if (![...fn.options].some((o) => o.value === fam)) fn.add(new Option(fam, fam));
    fn.value = fam;
    const pt = Math.round((parseFloat(cs.fontSize) * 3) / 4);
    const fs = $('#fontSize');
    if (![...fs.options].some((o) => +o.value === pt)) fs.add(new Option(String(pt), String(pt)));
    fs.value = String(pt);
  }
  $('[data-act="checklist"]').classList.toggle('on', !!(block && block.closest('ul.checklist')));
  // Odak modu: etkin paragraf
  if (document.body.classList.contains('focus-mode')) {
    $$('.is-active', page()).forEach((n) => n.classList.remove('is-active'));
    topBlock(s.anchorNode)?.classList.add('is-active');
  }
  updateCtxBar();
}

/* ------------------------------------------------------------------ */
/* Bağlam araçları (görsel / tablo)                                    */
/* ------------------------------------------------------------------ */

let selImg = null;
function showCtx(anchor, buttons) {
  const bar = $('#ctxBar');
  bar.innerHTML = buttons.map((b) => (b === '|' ? '<span class="sep"></span>' : `<button data-ctx="${b[0]}">${esc(b[1])}</button>`)).join('');
  bar.hidden = false;
  const r = anchor.getBoundingClientRect();
  const bw = bar.offsetWidth;
  bar.style.left = `${clamp(r.left + r.width / 2 - bw / 2, 8, window.innerWidth - bw - 8)}px`;
  bar.style.top = `${Math.max(60, r.top - bar.offsetHeight - 8)}px`;
}
function hideCtx() { $('#ctxBar').hidden = true; }
function updateCtxBar() {
  if (selImg && inPage(selImg)) {
    showCtx(selImg, [['w25', '25%'], ['w50', '50%'], ['w75', '75%'], ['w100', '100%'], '|', ['al', '⯇'], ['ac', '≡'], ['ar', '⯈'], '|', ['alt', 'Alt metin'], ['idel', '🗑']]);
    return;
  }
  const s = getSelection();
  const cell = s.rangeCount && inPage(s.anchorNode) ? (s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentElement).closest('td, th') : null;
  if (cell && inPage(cell)) {
    showCtx(cell.closest('table'), [['rAbove', '+ Satır ↑'], ['rBelow', '+ Satır ↓'], ['cLeft', '+ Sütun ←'], ['cRight', '+ Sütun →'], '|', ['rDel', '− Satır'], ['cDel', '− Sütun'], '|', ['head', 'Başlık satırı'], ['tDel', '🗑 Tablo']]);
    return;
  }
  hideCtx();
}
function onCtx(act) {
  if (selImg && inPage(selImg)) {
    const img = selImg;
    if (act.startsWith('w')) img.style.width = `${act.slice(1)}%`;
    else if (act === 'al') { img.style.display = ''; img.style.margin = ''; img.style.float = 'left'; img.style.marginRight = '12px'; }
    else if (act === 'ac') { img.style.float = ''; img.style.marginRight = ''; img.style.display = 'block'; img.style.margin = '0 auto'; }
    else if (act === 'ar') { img.style.display = ''; img.style.margin = ''; img.style.float = 'right'; img.style.marginLeft = '12px'; }
    else if (act === 'alt') { P.prompt('Alternatif metin', img.alt || '', 'Ekran okuyucular ve erişilebilirlik için görseli kısaca tanımlayın.').then((v) => { if (v !== null) { img.alt = v; changed(); } }); return; }
    else if (act === 'idel') { img.remove(); selImg = null; hideCtx(); }
    changed();
    updateCtxBar();
    return;
  }
  const s = getSelection();
  const cell = s.rangeCount ? (s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentElement).closest('td, th') : (lastRange && closestBlock(lastRange.startContainer)?.closest('td, th'));
  if (!cell) return;
  const tr = cell.parentNode;
  const table = cell.closest('table');
  const ci = Array.from(tr.children).indexOf(cell);
  const rows = $$('tr', table);
  const newCell = (tag = 'td') => { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; };
  if (act === 'rAbove' || act === 'rBelow') {
    const nr = document.createElement('tr');
    for (let i = 0; i < tr.children.length; i++) nr.appendChild(newCell());
    tr.parentNode.insertBefore(nr, act === 'rAbove' ? tr : tr.nextSibling);
  } else if (act === 'cLeft' || act === 'cRight') {
    rows.forEach((r, ri) => {
      const ref = r.children[ci];
      const c = newCell(r.children[0] && r.children[0].tagName === 'TH' && ri === 0 ? 'th' : 'td');
      if (ref) r.insertBefore(c, act === 'cLeft' ? ref : ref.nextSibling); else r.appendChild(c);
    });
  } else if (act === 'rDel') { if (rows.length <= 1) table.remove(); else tr.remove(); }
  else if (act === 'cDel') { if (tr.children.length <= 1) table.remove(); else rows.forEach((r) => r.children[ci] && r.children[ci].remove()); }
  else if (act === 'head') {
    const first = rows[0];
    const toTh = first.children[0].tagName !== 'TH';
    Array.from(first.children).forEach((c) => { const n = document.createElement(toTh ? 'th' : 'td'); n.innerHTML = c.innerHTML; c.replaceWith(n); });
  } else if (act === 'tDel') table.remove();
  changed();
  hideCtx();
}

/* ------------------------------------------------------------------ */
/* Yazarken: Markdown, akıllı yazım, hesap, eğik çizgi menüsü          */
/* ------------------------------------------------------------------ */

function tryBlockMarkdown(ctx) {
  if (ctx.block.tagName !== 'P' || ctx.block.closest('li, td, th, blockquote, [data-callout]')) return false;
  const m = ctx.before;
  const run = (fn) => {
    const s = getSelection();
    s.removeAllRanges();
    s.addRange(ctx.pre);
    document.execCommand('delete');
    saveSel();
    fn();
    changed();
    updateToolbar();
    return true;
  };
  if (/^#{1,4}$/.test(m)) return run(() => document.execCommand('formatBlock', false, `<h${m.length}>`));
  if (/^[-*+•]$/.test(m)) return run(() => listify('ul'));
  if (/^1[.)]$/.test(m)) return run(() => listify('ol'));
  if (/^\[[ xX]?\]$/.test(m)) {
    const checked = /x/i.test(m);
    return run(() => { listify('check'); if (checked) { const c2 = caretCtx(); if (c2 && c2.block.tagName === 'LI') c2.block.dataset.check = '1'; } });
  }
  if (m === '>') return run(() => document.execCommand('formatBlock', false, '<blockquote>'));
  if (/^!(i|bilgi)?$/.test(m) || m === ':::') return run(() => insertHtmlCaret('<div data-callout="info"><p>{c}</p></div>', ctx.block));
  return false;
}
function tryEnterMarkdown(ctx) {
  if (ctx.block.tagName !== 'P') return false;
  const t = ctx.block.textContent;
  if (ctx.before !== t) return false;
  const run = (html) => {
    const r = document.createRange();
    r.selectNodeContents(ctx.block);
    const s = getSelection();
    s.removeAllRanges();
    s.addRange(r);
    document.execCommand('insertHTML', false, html);
    changed();
    return true;
  };
  if (/^(---|—-|___|\*\*\*)$/.test(t)) { saveSel(); insertHtmlCaret('<hr><p>{c}</p>', ctx.block); return true; }
  if (/^```\w*$/.test(t)) { saveSel(); insertHtmlCaret('<pre>{c}</pre><p><br></p>', ctx.block); return true; }
  if (/^(\[toc\]|\[içindekiler\])$/i.test(t)) { saveSel(); insertHtmlCaret('<nav data-toc contenteditable="false"></nav><p>{c}</p>', ctx.block); refreshStructure(true); return true; }
  return false;
}
function tryInlineMarkdown(ctx) {
  const r = ctx.r;
  if (r.startContainer.nodeType !== 3) return false;
  const text = r.startContainer.data.slice(0, r.startOffset);
  const rules = [
    [/\*\*([^*\s](?:[^*]*[^*\s])?)\*\*$/, 'b'], [/__([^_\s](?:[^_]*[^_\s])?)__$/, 'b'],
    [/(?<![*\p{L}\d])\*([^*\s](?:[^*]*[^*\s])?)\*$/u, 'i'], [/(?<![_\p{L}\d])_([^_\s](?:[^_]*[^_\s])?)_$/u, 'i'],
    [/~~([^~]+)~~$/, 's'], [/==([^=]+)==$/, 'mark'], [/`([^`]+)`$/, 'code'],
  ];
  for (const [re, tag] of rules) {
    const m = re.exec(text);
    if (!m) continue;
    if (replaceBefore(m[0].length, `<${tag}>${esc(m[1])}</${tag}>\u00a0`)) { changed(); return true; }
    return false;
  }
  return false;
}
function smartType(e, ctx) {
  const b = ctx.before;
  const k = e.key;
  const prev = b.slice(-1);
  if (k === '"') { e.preventDefault(); document.execCommand('insertText', false, !prev || /[\s([{“‘—–-]/.test(prev) ? '“' : '”'); return true; }
  if (k === "'") { e.preventDefault(); document.execCommand('insertText', false, !prev || /[\s([{“—–-]/.test(prev) ? '‘' : '’'); return true; }
  if (k === '.' && b.endsWith('..')) { e.preventDefault(); return replaceBefore(2, '…'); }
  if (k === '-' && b.endsWith('<')) { e.preventDefault(); return replaceBefore(1, '←'); }
  if (k === '-' && b.endsWith('-') && !/^-+$/.test(b)) { e.preventDefault(); return replaceBefore(1, '—'); }
  if (k === '>' && b.endsWith('-')) { e.preventDefault(); return replaceBefore(1, '→'); }
  if (k === '>' && b.endsWith('=') && !/\d=$/.test(b)) { e.preventDefault(); return replaceBefore(1, '⇒'); }
  if (k === ')') {
    const m = /\((c|r|tm)$/i.exec(b);
    if (m) { e.preventDefault(); return replaceBefore(m[0].length, { c: '©', r: '®', tm: '™' }[m[1].toLowerCase()]); }
  }
  if (k === ' ') {
    const m = /(?:^|\s)(1\/2|1\/4|3\/4)$/.exec(b);
    if (m) { replaceBefore(3, { '1/2': '½', '1/4': '¼', '3/4': '¾' }[m[1]]); return false; }
  }
  return false;
}
function tryCalc(ctx) {
  const m = /([-+*/×÷x^%().,\d\s]+)$/.exec(ctx.before);
  if (!m) return false;
  let expr = m[1].replace(/^[^\d(-]+/, '').trim();
  if (!/\d/.test(expr) || !/\d\s*[-+*/×÷x^%]\s*[\d(]|\d%/.test(expr)) return false;
  if (expr.includes(',')) expr = expr.replace(/\.(?=\d{3}\b)/g, '').replace(/,/g, '.');
  expr = expr.replace(/(\d)\s*[x×]\s*(?=[\d(])/g, '$1*').replace(/÷/g, '/').replace(/\^/g, '**').replace(/(\d+(?:\.\d+)?)%/g, '($1/100)');
  if (!/^[\d+\-*/().\s]+$/.test(expr)) return false;
  let v;
  try { v = Function(`"use strict"; return (${expr});`)(); } catch (e) { return false; }
  if (typeof v !== 'number' || !Number.isFinite(v)) return false;
  document.execCommand('insertText', false, `=${v.toLocaleString('tr-TR', { maximumFractionDigits: 6 })}`);
  changed();
  return true;
}

/* Eğik çizgi (/) menüsü */
const SLASH = [
  { k: 'baslik1 h1 heading', n: 'Başlık 1', d: 'Büyük bölüm başlığı', ic: 'H1', run: () => formatBlock('h1') },
  { k: 'baslik2 h2', n: 'Başlık 2', d: 'Orta boy başlık', ic: 'H2', run: () => formatBlock('h2') },
  { k: 'baslik3 h3', n: 'Başlık 3', d: 'Küçük başlık', ic: 'H3', run: () => formatBlock('h3') },
  { k: 'metin normal paragraf', n: 'Metin', d: 'Normal paragraf', ic: '¶', run: () => formatBlock('p') },
  { k: 'madde liste bullet', n: 'Madde listesi', d: '• ile liste', ic: '•', run: () => listify('ul') },
  { k: 'numarali liste sirali', n: 'Numaralı liste', d: '1. 2. 3.', ic: '1.', run: () => listify('ol') },
  { k: 'yapilacaklar todo gorev checklist', n: 'Yapılacaklar', d: 'İşaretlenebilir liste', ic: '☑', run: toggleChecklist },
  { k: 'alinti quote', n: 'Alıntı', d: 'Vurgulu alıntı bloğu', ic: '❝', run: () => formatBlock('blockquote') },
  { k: 'kod code', n: 'Kod bloğu', d: 'Eş aralıklı yazı', ic: '</>', run: () => formatBlock('pre') },
  { k: 'bilgi not callout', n: 'Bilgi kutusu', d: 'Mavi not', ic: '💡', run: () => insertCallout('info') },
  { k: 'basari ipucu tamam', n: 'Başarı kutusu', d: 'Yeşil not', ic: '✅', run: () => insertCallout('success') },
  { k: 'uyari dikkat warning', n: 'Uyarı kutusu', d: 'Sarı not', ic: '⚠️', run: () => insertCallout('warning') },
  { k: 'tehlike hata danger', n: 'Tehlike kutusu', d: 'Kırmızı not', ic: '⛔', run: () => insertCallout('danger') },
  { k: 'acilir bolum details toggle', n: 'Açılır bölüm', d: 'Tıklayınca açılan içerik', ic: '▸', run: () => insertHtmlCaret('<details open><summary>{c}Başlık</summary><p><br></p></details><p><br></p>') },
  { k: 'tablo table', n: 'Tablo', d: 'Satır ve sütun', ic: '▦', run: insertTable },
  { k: 'gorsel resim image', n: 'Görsel', d: 'Dosyadan resim', ic: '🖼', run: () => $('#fileImage').click() },
  { k: 'baglanti link', n: 'Bağlantı', d: 'Web adresi', ic: '🔗', run: insertLink },
  { k: 'ayrac cizgi hr', n: 'Ayraç çizgisi', d: 'Yatay çizgi', ic: '―', run: () => exec('insertHorizontalRule') },
  { k: 'sayfa sonu page break', n: 'Sayfa sonu', d: 'Yeni sayfaya geç', ic: '⤓', run: () => insertHtmlCaret('<hr data-page-break><p>{c}</p>') },
  { k: 'icindekiler toc', n: 'İçindekiler', d: 'Başlıklardan otomatik dizin', ic: '☰', run: insertToc },
  { k: 'tarih bugun date', n: 'Bugünün tarihi', d: today(), ic: '📅', run: () => exec('insertText', today()) },
  { k: 'saat zaman time', n: 'Saat', d: 'Şimdiki saat', ic: '⏰', run: () => exec('insertText', new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })) },
  { k: 'emoji', n: 'Emoji', d: 'Emoji seçici', ic: '😀', run: insertEmoji },
  { k: 'sunum present slayt', n: 'Sunuma dönüştür', d: 'Bu belgeden Present sunumu', ic: '▶', run: toPresent },
];
let slash = null; // { node, offset, items, on }
const normTr = (s) => trLower(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ı/g, 'i');
function openSlash() {
  const s = getSelection();
  if (!s.rangeCount) return;
  const r = s.getRangeAt(0);
  if (r.startContainer.nodeType !== 3 || r.startOffset < 1) return;
  slash = { node: r.startContainer, offset: r.startOffset - 1, items: SLASH, on: 0 };
  drawSlash();
}
function slashQuery() {
  const s = getSelection();
  if (!slash || !s.rangeCount) return null;
  const r = s.getRangeAt(0);
  if (r.startContainer !== slash.node || r.startOffset <= slash.offset) return null;
  const q = slash.node.data.slice(slash.offset, r.startOffset);
  if (!q.startsWith('/') || q.length > 24 || /\s{2}/.test(q)) return null;
  return q.slice(1);
}
function drawSlash() {
  const q = slashQuery();
  if (q === null) { closeSlash(); return; }
  const nq = normTr(q);
  slash.items = SLASH.filter((x) => !nq || normTr(`${x.n} ${x.k}`).includes(nq));
  slash.on = Math.min(slash.on, Math.max(0, slash.items.length - 1));
  const m = $('#slashMenu');
  m.innerHTML = slash.items.length ? slash.items.map((x, i) => `<div class="si${i === slash.on ? ' on' : ''}" data-i="${i}"><span class="ic">${x.ic}</span><span><b>${esc(x.n)}</b><small>${esc(x.d)}</small></span></div>`).join('') : '<div class="empty">Sonuç yok — Esc ile kapatın</div>';
  m.hidden = false;
  const rr = document.createRange();
  rr.setStart(slash.node, slash.offset);
  rr.setEnd(slash.node, slash.offset + 1);
  const rect = rr.getBoundingClientRect();
  const h = m.offsetHeight;
  const top = rect.bottom + h + 8 > window.innerHeight ? rect.top - h - 6 : rect.bottom + 6;
  m.style.left = `${clamp(rect.left, 8, window.innerWidth - 290)}px`;
  m.style.top = `${Math.max(8, top)}px`;
  m.querySelector('.on')?.scrollIntoView({ block: 'nearest' });
}
function closeSlash() { slash = null; $('#slashMenu').hidden = true; }
function chooseSlash(i) {
  const item = slash && slash.items[i];
  if (!item) { closeSlash(); return; }
  const s = getSelection();
  const r = s.getRangeAt(0);
  const del = document.createRange();
  del.setStart(slash.node, slash.offset);
  del.setEnd(r.startContainer, r.startOffset);
  s.removeAllRanges();
  s.addRange(del);
  document.execCommand('delete');
  closeSlash();
  saveSel();
  item.run();
}

function onEditorKey(e) {
  const mod = e.ctrlKey || e.metaKey;
  if (slash) {
    if (e.key === 'ArrowDown') { e.preventDefault(); slash.on = (slash.on + 1) % Math.max(1, slash.items.length); drawSlash(); return; }
    if (e.key === 'ArrowUp') { e.preventDefault(); slash.on = (slash.on - 1 + slash.items.length) % Math.max(1, slash.items.length); drawSlash(); return; }
    if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); chooseSlash(slash.on); return; }
    if (e.key === 'Escape') { e.preventDefault(); closeSlash(); return; }
  }
  if (mod && e.key === 'Enter') { e.preventDefault(); saveSel(); insertHtmlCaret('<hr data-page-break><p>{c}</p>'); return; }
  if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) { e.preventDefault(); moveBlock(e.key === 'ArrowUp' ? -1 : 1); return; }
  if (e.key === 'Tab' && !mod) {
    const ctx = caretCtx() || {};
    const cell = ctx.block && ctx.block.closest ? (ctx.block.closest('td, th') || (ctx.block.tagName === 'TD' || ctx.block.tagName === 'TH' ? ctx.block : null)) : null;
    e.preventDefault();
    if (cell) {
      const cells = $$('td, th', cell.closest('table'));
      let i = cells.indexOf(cell) + (e.shiftKey ? -1 : 1);
      if (i >= cells.length) { lastRange = null; onCtx('rBelow'); const c2 = $$('td, th', cell.closest('table')); i = Math.min(i, c2.length - 1); placeCaret(c2[i]); return; }
      if (i >= 0) placeCaret(cells[i]);
      return;
    }
    const s = getSelection();
    const inList = s.rangeCount && (s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentElement).closest('li');
    if (inList) document.execCommand(e.shiftKey ? 'outdent' : 'indent');
    else if (!e.shiftKey) document.execCommand('insertText', false, ' ');
    changed();
    return;
  }
  if (mod || e.altKey) return;
  const ctx = caretCtx();
  if (!ctx) return;
  if (e.key === ' ' && settings.markdown) {
    if (tryBlockMarkdown(ctx)) { e.preventDefault(); return; }
    if (tryInlineMarkdown(ctx)) { e.preventDefault(); return; }
  }
  if (e.key === 'Enter' && !e.shiftKey) {
    if (settings.markdown && tryEnterMarkdown(ctx)) { e.preventDefault(); return; }
    setTimeout(fixNewListItem, 0);
  }
  if (settings.smart && e.key === '=' && tryCalc(ctx)) { e.preventDefault(); return; }
  if (settings.smart && smartType(e, ctx)) { changed(); return; }
  if (e.key === '/' && (!ctx.before || /\s$/.test(ctx.before))) setTimeout(openSlash, 0);
}
function fixNewListItem() {
  const ctx = caretCtx();
  if (ctx && ctx.block.tagName === 'LI' && ctx.block.dataset.check === '1' && !ctx.block.textContent.trim()) ctx.block.dataset.check = '0';
}
function placeCaret(node, end = false) {
  if (!node) return;
  const r = document.createRange();
  r.selectNodeContents(node);
  r.collapse(!end);
  const s = getSelection();
  s.removeAllRanges();
  s.addRange(r);
  saveSel();
}
function moveBlock(dir) {
  const s = getSelection();
  if (!s.rangeCount) return;
  const b = topBlock(s.anchorNode);
  if (!b) return;
  const sib = dir < 0 ? b.previousElementSibling : b.nextElementSibling;
  if (!sib) return;
  const r = s.getRangeAt(0).cloneRange();
  if (dir < 0) page().insertBefore(b, sib); else page().insertBefore(sib, b);
  s.removeAllRanges();
  s.addRange(r);
  b.scrollIntoView({ block: 'nearest' });
  changed();
}

/* ------------------------------------------------------------------ */
/* Ana hat, içindekiler, istatistikler                                 */
/* ------------------------------------------------------------------ */

let structSig = '';
function refreshStructure(force = false) {
  const pg = page();
  const hs = $$('h1, h2, h3', pg).filter((h) => !h.closest('nav[data-toc]'));
  hs.forEach((h) => { if (!h.id) h.id = `h-${Math.random().toString(36).slice(2, 8)}`; });
  const sig = hs.map((h) => `${h.tagName}${h.id}${h.textContent}`).join('|');
  if (!force && sig === structSig) return;
  structSig = sig;
  const lvl = (h) => +h.tagName[1];
  $('#outline').innerHTML = hs.length ? hs.map((h) => `<a data-lvl="${lvl(h)}" data-target="${h.id}" title="${esc(h.textContent)}">${esc(h.textContent.trim() || '(boş başlık)')}</a>`).join('')
    : '<p class="empty">Başlık eklediğinizde (# ve boşluk, ya da Başlık 1–3) burada tıklanabilir ana hat oluşur.</p>';
  for (const nav of $$('nav[data-toc]', pg)) {
    nav.setAttribute('contenteditable', 'false');
    nav.innerHTML = `<b>İçindekiler</b>${hs.length ? hs.map((h) => `<a data-lvl="${lvl(h)}" data-target="${h.id}">${esc(h.textContent.trim())}</a>`).join('') : '<small>Henüz başlık yok.</small>'}`;
  }
}
function jumpTo(id) {
  const h = document.getElementById(id);
  if (!h) return;
  h.scrollIntoView({ behavior: 'smooth', block: 'start' });
  placeCaret(h, true);
}

const STOP = new Set('ve ile ama fakat ancak veya ya da de da ki bu şu o bir için gibi daha çok en mi mı mu mü ne her ise olan olarak olan değil kadar sonra önce göre bunu şunu onu bunun onun şöyle böyle hem diye ben sen biz siz onlar benim senin bizim sizin artık yine zaten bile sadece hiç nasıl neden niçin hangi kendi dolayı üzere ise iken bazı tüm bütün birçok çünkü eğer yani'.split(' '));
const FILLERS = ['aslında', 'gerçekten', 'yani', 'sanki', 'oldukça', 'bayağı', 'resmen', 'açıkçası', 'kesinlikle', 'tabii ki', 'şey', 'falan', 'filan', 'işte', 'bir nevi', 'bir şekilde', 'adeta', 'son derece', 'gayet'];
let lastStats = null;
const hl = { long: false, filler: false, word: null };

function textIndex() {
  const pg = page();
  const nodes = [];
  let s = '';
  let lastBlock = null;
  const w = document.createTreeWalker(pg, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement && n.parentElement.closest('nav[data-toc]') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  let n;
  while ((n = w.nextNode())) {
    const b = closestBlock(n);
    if (lastBlock && b !== lastBlock) s += '\n';
    nodes.push({ node: n, start: s.length });
    s += n.data;
    lastBlock = b;
  }
  return { s, nodes };
}
function rangeFor(idx, start, end) {
  const find = (pos, isEnd) => {
    let lo = 0;
    let hi = idx.nodes.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (idx.nodes[mid].start <= pos - (isEnd ? 1 : 0)) lo = mid; else hi = mid - 1;
    }
    const e = idx.nodes[lo];
    return e ? [e.node, clamp(pos - e.start, 0, e.node.data.length)] : null;
  };
  const a = find(start, false);
  const b = find(end, true);
  if (!a || !b) return null;
  const r = document.createRange();
  try { r.setStart(a[0], a[1]); r.setEnd(b[0], b[1]); } catch (e) { return null; }
  return r;
}
const HAS_HL = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';
function setHighlight(name, ranges) {
  if (!HAS_HL) return;
  if (ranges && ranges.length) CSS.highlights.set(name, new Highlight(...ranges)); else CSS.highlights.delete(name);
}

function computeStats() {
  const idx = textIndex();
  const text = idx.s;
  const words = text.match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) || [];
  const sentences = [];
  const re = /[^.!?…\n]+(?:[.!?…]+|(?=\n)|$)/g;
  let m;
  while ((m = re.exec(text))) {
    const t = m[0];
    if (/\p{L}/u.test(t)) sentences.push({ start: m.index + (t.length - t.trimStart().length), end: m.index + t.trimEnd().length, words: (t.match(/[\p{L}\p{N}]+/gu) || []).length });
    if (m[0] === '') re.lastIndex++;
  }
  const syl = words.reduce((n, w) => n + Math.max(1, (w.match(/[aeıioöuüâîû]/gi) || []).length), 0);
  const nw = words.length;
  const ns = Math.max(1, sentences.length);
  const atesman = nw ? clamp(198.825 - 40.175 * (syl / nw) - 2.61 * (nw / ns), 0, 100) : null;
  const freq = new Map();
  for (const w of words) {
    const k = trLower(w);
    if (k.length < 4 || STOP.has(k) || /^\d+$/.test(k)) continue;
    freq.set(k, (freq.get(k) || 0) + 1);
  }
  const top = [...freq.entries()].filter(([, c]) => c >= 3).sort((a, b) => b[1] - a[1]).slice(0, 12);
  const longS = sentences.filter((x) => x.words > 25);
  const lowerText = trLower(text);
  const fillerHits = [];
  for (const f of FILLERS) {
    const fr = new RegExp(`(^|[^\\p{L}])(${f})(?=[^\\p{L}]|$)`, 'gu');
    let mm;
    while ((mm = fr.exec(lowerText))) fillerHits.push({ start: mm.index + mm[1].length, end: mm.index + mm[1].length + mm[2].length, w: f });
  }
  const paras = $$('p, li, h1, h2, h3, h4, blockquote, pre, td, th', page()).filter((b) => b.textContent.trim()).length;
  const [, ph] = PAGE_SIZES[doc.page.size] || PAGE_SIZES.A4;
  const pages = Math.max(1, Math.ceil(page().scrollHeight / ph));
  return {
    idx, words: nw, chars: text.replace(/\n/g, '').length, charsNoSpace: text.replace(/\s/g, '').length, sentences: sentences.length, sentenceList: sentences,
    paras, syl, atesman, top, longS, fillerHits, pages, avg: nw / ns,
  };
}
function scheduleStats() { clearTimeout(statsTimer); statsTimer = setTimeout(updateStats, 350); }
const minutes = (n) => (n < 1 ? '< 1 dk' : `${Math.round(n)} dk`);
function readLabel(a) {
  if (a === null) return ['—', 'var(--muted)'];
  if (a >= 90) return ['Çok kolay', '#16a34a'];
  if (a >= 70) return ['Kolay', '#22c55e'];
  if (a >= 50) return ['Orta', '#f59e0b'];
  if (a >= 30) return ['Zor', '#f97316'];
  return ['Çok zor', '#dc2626'];
}
function updateStats() {
  if (!doc) return;
  refreshStructure();
  const st = computeStats();
  lastStats = st;
  $('#stWords').textContent = `${st.words.toLocaleString('tr-TR')} kelime`;
  $('#stRead').textContent = `📖 ${minutes(st.words / 200)} okuma`;
  $('#stPages').textContent = `${st.pages} sayfa`;
  const goal = doc.goal || 0;
  const g = $('#stGoal');
  g.hidden = !goal;
  if (goal) {
    const p = clamp((st.words / goal) * 100, 0, 100);
    g.style.setProperty('--p', `${p}%`);
    $('b', g).textContent = `Hedef ${st.words}/${goal}${p >= 100 ? ' 🎉' : ''}`;
    if (p >= 100 && !doc._goalCelebrated) { doc._goalCelebrated = true; P.toast('Tebrikler! Yazma hedefinize ulaştınız. 🎉'); }
  }
  const [label, color] = readLabel(st.atesman);
  const session = st.words - sessionStartWords;
  $('#insights').innerHTML = `
    <h3>İstatistikler</h3>
    <div class="stat-grid">
      <div class="stat"><b>${st.words.toLocaleString('tr-TR')}</b><small>kelime</small></div>
      <div class="stat"><b>${st.chars.toLocaleString('tr-TR')}</b><small>karakter (${st.charsNoSpace.toLocaleString('tr-TR')} boşluksuz)</small></div>
      <div class="stat"><b>${st.sentences}</b><small>cümle</small></div>
      <div class="stat"><b>${st.paras}</b><small>paragraf</small></div>
      <div class="stat"><b>${minutes(st.words / 200)}</b><small>okuma süresi</small></div>
      <div class="stat"><b>${minutes(st.words / 130)}</b><small>konuşma süresi</small></div>
      <div class="stat"><b>${st.avg.toFixed(1)}</b><small>kelime / cümle</small></div>
      <div class="stat"><b>${session >= 0 ? '+' : ''}${session}</b><small>bu oturumda</small></div>
    </div>
    <h3>Okunabilirlik (Ateşman)</h3>
    <div class="ins-row"><b style="font-size:22px;color:${color}">${st.atesman === null ? '—' : Math.round(st.atesman)}</b><span class="grow"><b style="color:${color}">${label}</b><div class="meter"><i style="width:${st.atesman || 0}%;background:${color}"></i></div></span></div>
    <p class="hint">Türkçe metinler için geliştirilmiş formül: kısa cümleler ve kısa kelimeler puanı yükseltir.</p>
    <h3>Yazım analizi</h3>
    <div class="ins-row"><span class="grow">Uzun cümleler (25+ kelime)</span><b>${st.longS.length}</b><button class="ins-toggle${hl.long ? ' on' : ''}" data-hl="long">Vurgula</button></div>
    <div class="ins-row"><span class="grow">Dolgu kelimeler (aslında, yani…)</span><b>${st.fillerHits.length}</b><button class="ins-toggle${hl.filler ? ' on' : ''}" data-hl="filler">Vurgula</button></div>
    ${st.top.length ? `<p class="hint">Sık tekrarlanan kelimeler — vurgulamak için tıklayın:</p><div>${st.top.map(([w, c]) => `<span class="word-chip${hl.word === w ? ' on' : ''}" data-word="${esc(w)}">${esc(w)} <i>${c}</i></span>`).join('')}</div>` : '<p class="hint">Sık tekrar eden kelime yok. 👍</p>'}
    ${HAS_HL ? '' : '<p class="hint">Bu tarayıcı vurgulamayı desteklemiyor.</p>'}
    <h3>Hedef</h3>
    ${goal ? `<div class="meter"><i style="width:${clamp((st.words / goal) * 100, 0, 100)}%;background:var(--ok)"></i></div><p class="hint">${st.words} / ${goal} kelime</p>` : ''}
    <button class="btn sm" data-ins="goal">${goal ? 'Hedefi değiştir' : 'Kelime hedefi koy'}</button>`;
  applyAnalysisHighlights();
}
function applyAnalysisHighlights() {
  const st = lastStats;
  if (!st || !HAS_HL) return;
  setHighlight('ew-long', hl.long ? st.longS.map((x) => rangeFor(st.idx, x.start, x.end)).filter(Boolean) : null);
  setHighlight('ew-filler', hl.filler ? st.fillerHits.map((x) => rangeFor(st.idx, x.start, x.end)).filter(Boolean) : null);
  if (hl.word) {
    const ranges = [];
    const lower = trLower(st.idx.s);
    const re = new RegExp(`(^|[^\\p{L}\\p{N}])(${hl.word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})(?=[^\\p{L}\\p{N}]|$)`, 'gu');
    let m;
    while ((m = re.exec(lower))) { const s0 = m.index + m[1].length; const r = rangeFor(st.idx, s0, s0 + m[2].length); if (r) ranges.push(r); }
    setHighlight('ew-repeat', ranges);
  } else setHighlight('ew-repeat', null);
}

/* ------------------------------------------------------------------ */
/* Bul ve değiştir                                                     */
/* ------------------------------------------------------------------ */

let finds = [];
let findCur = -1;
function openFind(replace = false) {
  const fb = $('#findbar');
  fb.hidden = false;
  const sel = getSelection().toString();
  if (sel && sel.length < 80 && !sel.includes('\n')) $('#findText').value = sel;
  (replace ? $('#replText') : $('#findText')).focus();
  $('#findText').select();
  runFind();
}
function closeFind() {
  $('#findbar').hidden = true;
  finds = [];
  setHighlight('ew-find', null);
  setHighlight('ew-find-cur', null);
}
function runFind(keepIndex = false) {
  const q = $('#findText').value;
  finds = [];
  if (q) {
    const idx = textIndex();
    const cs = $('#findCase').checked;
    const hay = cs ? idx.s : trLower(idx.s);
    const needle = cs ? q : trLower(q);
    let i = hay.indexOf(needle);
    while (i >= 0 && finds.length < 5000) {
      const r = rangeFor(idx, i, i + needle.length);
      if (r) finds.push(r);
      i = hay.indexOf(needle, i + Math.max(1, needle.length));
    }
  }
  if (!keepIndex) findCur = finds.length ? 0 : -1;
  findCur = Math.min(findCur, finds.length - 1);
  $('#findCount').textContent = q ? (finds.length ? `${findCur + 1} / ${finds.length}` : 'Bulunamadı') : '';
  setHighlight('ew-find', finds);
  showCurFind();
}
function showCurFind() {
  const r = finds[findCur];
  setHighlight('ew-find-cur', r ? [r] : null);
  if (!r) return;
  $('#findCount').textContent = `${findCur + 1} / ${finds.length}`;
  const rect = r.getBoundingClientRect();
  const wrap = $('#pageWrap');
  const wr = wrap.getBoundingClientRect();
  if (rect.top < wr.top + 40 || rect.bottom > wr.bottom - 40) wrap.scrollTop += rect.top - wr.top - wr.height / 3;
  if (!HAS_HL) { const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
}
function findStep(d) {
  if (!finds.length) return;
  findCur = (findCur + d + finds.length) % finds.length;
  showCurFind();
}
function replaceRange(r, text) {
  const pg = page();
  pg.focus({ preventScroll: true });
  const s = getSelection();
  s.removeAllRanges();
  s.addRange(r);
  document.execCommand('insertText', false, text);
}
function replaceOne() {
  if (!finds[findCur]) return;
  replaceRange(finds[findCur], $('#replText').value);
  changed();
  runFind(true);
  $('#replText').focus();
}
function replaceAll() {
  const n = finds.length;
  if (!n) return;
  const t = $('#replText').value;
  for (let i = finds.length - 1; i >= 0; i--) replaceRange(finds[i], t);
  changed();
  runFind();
  P.toast(`${n} eşleşme değiştirildi.`);
}

/* ------------------------------------------------------------------ */
/* Metin araçları                                                      */
/* ------------------------------------------------------------------ */

function transformSelection(fn) {
  restoreSel();
  const s = getSelection();
  if (!s.rangeCount || s.isCollapsed) { P.toast('Önce dönüştürmek istediğiniz metni seçin.'); return; }
  const r = s.getRangeAt(0);
  // Biçimi korumak için seçimdeki metin düğümlerini tek tek dönüştür
  const nodes = [];
  const w = document.createTreeWalker(r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) if (r.intersectsNode(n)) nodes.push(n);
  let pos = 0;
  const full = r.toString();
  const out = fn(full);
  if (out.length !== full.length) { document.execCommand('insertText', false, out); changed(); return; }
  for (const t of nodes) {
    const a = t === r.startContainer ? r.startOffset : 0;
    const b = t === r.endContainer ? r.endOffset : t.data.length;
    const seg = out.slice(pos, pos + (b - a));
    pos += b - a;
    t.data = t.data.slice(0, a) + seg + t.data.slice(b);
  }
  changed();
}
const titleCase = (s) => trLower(s).replace(/(^|[\s"“‘(\-–—/])(\p{L})/gu, (m, a, c) => a + trUpper(c));
const sentenceCase = (s) => trLower(s).replace(/(^\s*|[.!?…]\s+)(\p{L})/gu, (m, a, c) => a + trUpper(c));

async function cleanText() {
  await saveSnap('Temizlik öncesi');
  const w = document.createTreeWalker(page(), NodeFilter.SHOW_TEXT);
  let n;
  let count = 0;
  while ((n = w.nextNode())) {
    if (n.parentElement.closest('pre, code')) continue;
    const t = n.data.replace(/[  ]{2,}/g, ' ').replace(/ +([,.;:!?])/g, '$1').replace(/([,;:])(?=\p{L})/gu, '$1 ').replace(/,,+/g, ',').replace(/([!?])\1{2,}/g, '$1');
    if (t !== n.data) { n.data = t; count++; }
  }
  const empties = $$(':scope > p', page()).filter((p) => !p.textContent.trim() && !p.querySelector('img'));
  let removed = 0;
  empties.forEach((p) => { const prev = p.previousElementSibling; if (prev && prev.tagName === 'P' && !prev.textContent.trim() && !prev.querySelector('img')) { p.remove(); removed++; } });
  changed();
  P.toast(`Metin temizlendi: ${count} düzeltme, ${removed} fazla boş paragraf. (Önceki hâl sürüm geçmişinde)`);
}
async function sortList() {
  restoreSel();
  const ctx = caretCtx();
  const s = getSelection();
  const list = ctx ? ctx.block.closest('ul, ol') : (s.rangeCount ? closestBlock(s.anchorNode)?.closest('ul, ol') : null);
  await saveSnap('Sıralama öncesi');
  if (list) {
    const lis = $$(':scope > li', list).sort((a, b) => a.textContent.localeCompare(b.textContent, 'tr', { numeric: true }));
    lis.forEach((li) => list.appendChild(li));
    changed();
    P.toast('Liste sıralandı.');
    return;
  }
  const blocks = selectedBlocks().filter((b) => b.tagName === 'P' && b.parentNode === page());
  if (blocks.length < 2) { P.toast('Sıralamak için imleci bir listeye koyun veya birden çok paragraf seçin.'); return; }
  const anchor = blocks[blocks.length - 1].nextSibling;
  blocks.sort((a, b) => a.textContent.localeCompare(b.textContent, 'tr', { numeric: true })).forEach((b) => page().insertBefore(b, anchor));
  changed();
  P.toast('Paragraflar sıralandı.');
}

/* ------------------------------------------------------------------ */
/* Odak, daktilo, dikte, sesli okuma, pomodoro, hedef                  */
/* ------------------------------------------------------------------ */

function toggleFocus() {
  const on = document.body.classList.toggle('focus-mode');
  let btn = $('.focus-exit');
  if (!btn) {
    btn = document.createElement('button');
    btn.className = 'btn focus-exit';
    btn.textContent = '✕ Odak modundan çık';
    btn.addEventListener('click', toggleFocus);
    document.body.appendChild(btn);
  }
  if (!on) $$('.is-active', page()).forEach((n) => n.classList.remove('is-active'));
  else { P.toast('Odak modu: yalnızca yazdığınız paragraf parlak. Çıkmak için Esc.'); updateToolbar(); }
}
function typewriterScroll() {
  if (!settings.typewriter || !doc) return;
  const s = getSelection();
  if (!s.rangeCount || !inPage(s.anchorNode)) return;
  const r = s.getRangeAt(0).cloneRange();
  let rect = r.getBoundingClientRect();
  if (!rect.height) { const b = closestBlock(s.anchorNode); if (b) rect = b.getBoundingClientRect(); }
  const wrap = $('#pageWrap');
  const wr = wrap.getBoundingClientRect();
  const target = wr.top + wr.height * 0.45;
  if (Math.abs(rect.top - target) > 24) wrap.scrollTop += rect.top - target;
}

let dictation = null;
function toggleDictation() {
  if (dictation) { stopDictation(); P.toast('Dikte durduruldu.'); return; }
  restoreSel();
  const st = $('#stDictate');
  dictation = P.speech({
    onText: (fin, tmp) => {
      if (fin) {
        restoreSel();
        const ctx = caretCtx();
        const pre = ctx && ctx.before && !/\s$/.test(ctx.before) ? ' ' : '';
        let t = fin.trim();
        if (!ctx || !ctx.before.trim() || /[.!?…]\s*$/.test(ctx.before)) t = t.charAt(0).toLocaleUpperCase('tr') + t.slice(1);
        document.execCommand('insertText', false, pre + t);
        saveSel();
        changed();
      }
      st.hidden = false;
      st.textContent = `🎙 ${tmp || 'Dinleniyor…'}`;
    },
    onEnd: () => { st.hidden = true; dictation = null; },
  });
  if (!dictation) { P.toast('Bu tarayıcı konuşma tanımayı desteklemiyor (Chrome veya Edge deneyin).'); return; }
  st.hidden = false;
  st.textContent = '🎙 Dinleniyor…';
  P.toast('Dikte açık: konuşun, metin imlece yazılsın. Durdurmak için tekrar seçin.');
}
function stopDictation() { if (dictation) { dictation.stop(); dictation = null; } $('#stDictate').hidden = true; }

function toggleSpeak() {
  if (!('speechSynthesis' in window)) { P.toast('Bu tarayıcı sesli okumayı desteklemiyor.'); return; }
  if (speechSynthesis.speaking) { speechSynthesis.cancel(); P.toast('Sesli okuma durduruldu.'); return; }
  const sel = getSelection().toString().trim();
  const text = sel || page().innerText;
  const voice = speechSynthesis.getVoices().find((v) => /^tr/i.test(v.lang));
  const chunks = text.match(/[^.!?…\n]+[.!?…]*|\n+/g) || [];
  let buf = '';
  const queue = [];
  for (const c of chunks) { if ((buf + c).length > 220) { queue.push(buf); buf = ''; } buf += c; }
  if (buf.trim()) queue.push(buf);
  queue.filter((q) => q.trim()).forEach((q) => {
    const u = new SpeechSynthesisUtterance(q);
    u.lang = 'tr-TR';
    if (voice) u.voice = voice;
    speechSynthesis.speak(u);
  });
  P.toast(sel ? 'Seçim okunuyor… (durdurmak için tekrar seçin)' : 'Belge okunuyor… (durdurmak için tekrar seçin)');
}

let pomo = null;
function togglePomodoro() {
  const el = $('#stPomo');
  if (pomo) { clearInterval(pomo.t); pomo = null; el.hidden = true; P.toast('Pomodoro durduruldu.'); return; }
  const start = (mode) => {
    pomo = { mode, end: Date.now() + (mode === 'work' ? 25 : 5) * 60000, t: pomo ? pomo.t : null };
    el.classList.toggle('break', mode === 'break');
  };
  start('work');
  el.hidden = false;
  const tick = () => {
    if (!pomo) return;
    const left = Math.max(0, pomo.end - Date.now());
    el.textContent = `${pomo.mode === 'work' ? '🍅 Odak' : '☕ Mola'} ${String(Math.floor(left / 60000)).padStart(2, '0')}:${String(Math.floor(left / 1000) % 60).padStart(2, '0')}`;
    if (!left) {
      beep();
      const next = pomo.mode === 'work' ? 'break' : 'work';
      P.toast(next === 'break' ? '25 dakika doldu! 5 dakika mola verin. ☕' : 'Mola bitti, yazmaya devam! 🍅');
      start(next);
    }
  };
  pomo.t = setInterval(tick, 1000);
  tick();
  P.toast('Pomodoro başladı: 25 dk odaklanın. Durdurmak için zamanlayıcıya tıklayın.');
}
function beep() {
  try {
    const ac = new (window.AudioContext || window.webkitAudioContext)();
    [0, 0.25, 0.5].forEach((t) => {
      const o = ac.createOscillator();
      const g = ac.createGain();
      o.frequency.value = 880;
      g.gain.setValueAtTime(0.15, ac.currentTime + t);
      g.gain.exponentialRampToValueAtTime(0.001, ac.currentTime + t + 0.2);
      o.connect(g).connect(ac.destination);
      o.start(ac.currentTime + t);
      o.stop(ac.currentTime + t + 0.2);
    });
  } catch (e) { /* ses yok */ }
}
async function setGoal() {
  const v = await P.prompt('Yazma hedefi', String(doc.goal || 500), 'Bu belge için kelime hedefi (0 = kapalı). İlerleme durum çubuğunda görünür.');
  if (v === null) return;
  doc.goal = Math.max(0, parseInt(v, 10) || 0);
  delete doc._goalCelebrated;
  changed();
}
async function insertEmoji() {
  const e = await P.pickEmoji();
  if (e) exec('insertText', e);
}

/* ------------------------------------------------------------------ */
/* Sürümler ve karşılaştırma                                           */
/* ------------------------------------------------------------------ */

async function saveSnap(label = '') {
  if (!doc) return;
  doc.html = serialize();
  await P.snapshots.save('easywrite', doc.id, doc.title, { title: doc.title, html: doc.html, page: doc.page }, label);
}
async function snapshotNamed() {
  const label = await P.prompt('Sürüm kaydet', `Sürüm ${new Date().toLocaleString('tr-TR', { dateStyle: 'short', timeStyle: 'short' })}`, 'Adlandırılmış sürümler otomatik silinmez.');
  if (label === null) return;
  await saveSnap(label.trim() || 'Adsız sürüm');
  P.toast('Sürüm kaydedildi.');
}
const plainOf = (html) => {
  const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  $$('p, h1, h2, h3, h4, li, tr, blockquote, pre, div', d.body).forEach((b) => b.append('\n'));
  return d.body.textContent.replace(/\n{2,}/g, '\n').trim();
};
function diffHtml(a, b) {
  const tok = (s) => s.match(/\s+|[^\s]+/g) || [];
  let A = tok(a);
  let B = tok(b);
  if (A.length * B.length > 4e6) { A = a.split(/(?<=\n)/); B = b.split(/(?<=\n)/); }
  const n = A.length;
  const m = B.length;
  if (n * m > 4e6) return '<p>Belge karşılaştırma için çok büyük.</p>';
  const dp = new Uint32Array((n + 1) * (m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i * (m + 1) + j] = A[i] === B[j] ? dp[(i + 1) * (m + 1) + j + 1] + 1 : Math.max(dp[(i + 1) * (m + 1) + j], dp[i * (m + 1) + j + 1]);
  let i = 0;
  let j = 0;
  let out = '';
  let ins = 0;
  let del = 0;
  while (i < n && j < m) {
    if (A[i] === B[j]) { out += esc(A[i]); i++; j++; } else if (dp[(i + 1) * (m + 1) + j] >= dp[i * (m + 1) + j + 1]) { out += `<del>${esc(A[i])}</del>`; if (A[i].trim()) del++; i++; } else { out += `<ins>${esc(B[j])}</ins>`; if (B[j].trim()) ins++; j++; }
  }
  while (i < n) { out += `<del>${esc(A[i])}</del>`; if (A[i].trim()) del++; i++; }
  while (j < m) { out += `<ins>${esc(B[j])}</ins>`; if (B[j].trim()) ins++; j++; }
  return `<p class="muted"><span style="color:var(--ok)">+${ins} eklendi</span> · <span style="color:var(--danger)">−${del} silindi</span> (kelime)</p><div class="diff">${out.replace(/<\/(del|ins)><\1>/g, '')}</div>`;
}
async function versions() {
  await saveNow();
  const list = await P.snapshots.list('easywrite', doc.id);
  const draw = () => (list.length ? `<ul class="p365-list">${list.map((s, k) => `<li><span class="grow"><b>${esc(s.label || 'Otomatik kayıt')}</b><small>${new Date(s.at).toLocaleString('tr-TR', { dateStyle: 'medium', timeStyle: 'short' })}</small></span>
    <button class="btn sm" data-cmp="${k}">Karşılaştır</button><button class="btn sm primary" data-restore="${k}">Geri yükle</button><button class="btn sm danger" data-del="${k}">🗑</button></li>`).join('')}</ul>` : '<p class="muted">Henüz kayıtlı sürüm yok.</p>');
  const r = await P.dialog({
    title: 'Sürüm geçmişi',
    wide: true,
    html: `<p class="muted">Belgeyi her açtığınızda (30 dakikada bir) ve riskli işlemlerden önce otomatik sürüm alınır. “Karşılaştır” eklenen ve silinen kelimeleri gösterir.</p><div id="vlist">${draw()}</div><div id="vdiff"></div>`,
    actions: [{ label: 'Şimdi sürüm kaydet', value: 'save' }, { label: 'Kapat', primary: true, value: null }],
    onOpen: (body, close) => body.addEventListener('click', async (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.del) {
        await P.snapshots.del(list[+b.dataset.del].id);
        list.splice(+b.dataset.del, 1);
        $('#vlist', body).innerHTML = draw();
      } else if (b.dataset.cmp) {
        const snap = JSON.parse(list[+b.dataset.cmp].data);
        $('#vdiff', body).innerHTML = `<h3 style="font-size:14px">“${esc(list[+b.dataset.cmp].label || 'Otomatik kayıt')}” → şimdiki hâl</h3>${diffHtml(plainOf(snap.html), plainOf(serialize()))}`;
      } else if (b.dataset.restore) {
        const snap = JSON.parse(list[+b.dataset.restore].data);
        await saveSnap('');
        page().focus();
        const rr = document.createRange();
        rr.selectNodeContents(page());
        getSelection().removeAllRanges();
        getSelection().addRange(rr);
        document.execCommand('insertHTML', false, C.sanitize(snap.html) || '<p><br></p>');
        if (snap.page) { doc.page = { ...doc.page, ...snap.page }; applyPage(); }
        changed();
        refreshStructure(true);
        close();
        P.toast('Sürüm geri yüklendi. Geri almak için Ctrl+Z.');
      }
    }),
  });
  if (r === 'save') snapshotNamed();
}

/* ------------------------------------------------------------------ */
/* Dışa / içe aktarma                                                  */
/* ------------------------------------------------------------------ */

const EW_PART = 'easywrite/doc.json';
const EW_REL = 'https://platform365.app/2026/relationships/easywrite';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function docCss() {
  const sheet = [...document.styleSheets].find((s) => s.href && /easywrite\/styles\.css/.test(s.href));
  try { return sheet ? [...sheet.cssRules].map((r) => r.cssText).filter((t) => /^\.page|^@media print/.test(t) && !/focus-mode/.test(t)).join('\n') : ''; } catch (e) { return ''; }
}
async function exportAs(kind, d = doc) {
  if (d === doc && doc) await saveNow();
  const name = P.safeName(d.title, 'belge');
  try {
    if (kind === 'easywrite' || kind === 'docx') {
      P.toast('Dosya hazırlanıyor…');
      const zip = await Docx.build(d);
      if (kind === 'easywrite') {
        zip.file(EW_PART, JSON.stringify({ format: FILE_FORMAT, version: APP_VERSION, exported: new Date().toISOString(), doc: { ...d, _goalCelebrated: undefined } }));
        let rels = await zip.file('_rels/.rels').async('string');
        rels = rels.replace('</Relationships>', `<Relationship Id="rIdEasyWrite" Type="${EW_REL}" Target="${EW_PART}"/></Relationships>`);
        zip.file('_rels/.rels', rels);
      }
      const blob = await zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
      P.download(blob, `${name}.${kind}`);
      P.toast(kind === 'docx' ? 'Word dosyası indirildi.' : 'EasyWrite dosyası indirildi (Word ile de açılır).');
    } else if (kind === 'md') {
      P.download(new Blob([C.htmlToMd(d.html)], { type: 'text/markdown' }), `${name}.md`);
    } else if (kind === 'txt') {
      P.download(new Blob([plainOf(d.html)], { type: 'text/plain' }), `${name}.txt`);
    } else if (kind === 'html') {
      const p = d.page || {};
      const [w] = PAGE_SIZES[p.size] || PAGE_SIZES.A4;
      const html = `<!doctype html><html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(d.title)}</title><meta name="generator" content="Platform365 EasyWrite ${APP_VERSION}">
<style>${docCss()}
body{margin:0;background:#f4f5f9;padding:24px 12px}.page{margin:0 auto;max-width:100%;box-sizing:border-box}@media (max-width:840px){.page{padding:24px !important}}</style></head>
<body><article class="page" style="--page-w:${w}px;--margin:${MARGINS[p.margin] || 96}px;--doc-font:'${esc(p.font || 'Calibri')}';--doc-size:${p.fontSize || 16}px">${d.html}</article></body></html>`;
      P.download(new Blob([html], { type: 'text/html' }), `${name}.html`);
    } else if (kind === 'pdf') printDoc();
  } catch (e) {
    console.error(e);
    P.toast(`İndirme başarısız: ${e.message}`);
  }
}
function printDoc() {
  const area = $('#printArea');
  area.innerHTML = '';
  const c = page().cloneNode(true);
  c.removeAttribute('contenteditable');
  c.removeAttribute('id');
  c.className = 'page';
  area.appendChild(c);
  let st = $('#printPageSize');
  if (!st) { st = document.createElement('style'); st.id = 'printPageSize'; document.head.appendChild(st); }
  const size = { A4: 'A4', Letter: 'letter', A5: 'A5' }[doc.page.size] || 'A4';
  const mm = { narrow: '12.7mm', normal: '25.4mm', wide: '38.1mm' }[doc.page.margin] || '25.4mm';
  st.textContent = `@media print { @page { size: ${size}; margin: ${mm}; } }`;
  const prev = document.title;
  document.title = P.safeName(doc.title, 'belge');
  setTimeout(() => {
    window.print();
    document.title = prev;
    area.innerHTML = '';
  }, 50);
}

async function sniff(file) {
  const h = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  if (h[0] === 0x50 && h[1] === 0x4b) return 'zip';
  if (h[0] === 0xd0 && h[1] === 0xcf) return 'ole';
  return 'text';
}
async function importFile(file) {
  try {
    const kind = await sniff(file);
    const base = file.name.replace(/\.[^.]+$/, '') || 'Belge';
    let d;
    if (kind === 'ole') { P.toast('Eski .doc biçimi desteklenmiyor. Word\'de “Farklı Kaydet → .docx” ile kaydedip tekrar deneyin.'); return; }
    if (kind === 'zip') {
      const JSZip = await P.loadZip();
      const zip = await JSZip.loadAsync(await file.arrayBuffer());
      const own = zip.file(EW_PART);
      if (own) {
        const data = JSON.parse(await own.async('string'));
        const src = data.doc || {};
        d = newDoc(String(src.title || base), C.sanitize(src.html || ''));
        d.page = { ...d.page, ...(src.page || {}) };
        d.goal = +src.goal || 0;
      } else if (zip.file('word/document.xml')) {
        P.toast('Word dosyası açılıyor…');
        const res = await Docx.parse(zip);
        d = newDoc(res.title || base, C.sanitize(res.html) || '<p><br></p>');
        d.page.size = res.size;
      } else if (zip.file('ppt/presentation.xml')) {
        P.toast('Bu bir sunum dosyası — Present\'te açılıyor…');
        P.inbox.send('present', { kind: 'file', blob: file, name: file.name, type: file.type });
        return;
      } else throw new Error('Tanınmayan paket');
    } else {
      const text = await file.text();
      if (/\.(md|markdown)$/i.test(file.name)) d = newDoc(base, C.sanitize(C.mdToHtml(text)));
      else if (/\.html?$/i.test(file.name) || /^\s*<(!doctype|html)/i.test(text)) {
        const parsed = new DOMParser().parseFromString(text, 'text/html');
        d = newDoc(parsed.title || base, C.sanitize(parsed.body.innerHTML));
      } else {
        try {
          const j = JSON.parse(text);
          if (j && j.format === FILE_FORMAT && j.doc) d = newDoc(j.doc.title || base, C.sanitize(j.doc.html || ''));
        } catch (e) { /* düz metin */ }
        if (!d) d = newDoc(base, text.split(/\r?\n/).map((l) => `<p>${esc(l) || '<br>'}</p>`).join(''));
      }
    }
    if (!d.html) d.html = '<p><br></p>';
    await Store.put(d);
    P.toast(`“${d.title}” açıldı.`);
    openDoc(d.id);
  } catch (e) {
    console.error(e);
    P.toast(`Dosya açılamadı: ${e.message || 'bilinmeyen hata'}`);
  }
}

function toPresent() {
  saveNow();
  const outline = C.outlineFromHtml(serialize(), doc.title);
  if (!outline.length) { P.toast('Sunuma dönüştürülecek içerik yok. Başlıklar ekleyin.'); return; }
  P.inbox.send('present', { kind: 'outline', title: doc.title, slides: outline });
}

/* ------------------------------------------------------------------ */
/* Yapıştırma                                                          */
/* ------------------------------------------------------------------ */

let pastePlain = false;
const looksMd = (t) => /(^|\n)(#{1,6} |[-*] |\d+\. |> |```)|\*\*[^*]+\*\*|\[[^\]]+\]\([^)]+\)/.test(t);
function onPaste(e) {
  const cd = e.clipboardData;
  if (!cd) return;
  const img = [...(cd.files || [])].find((f) => f.type.startsWith('image/'));
  if (img) { e.preventDefault(); saveSel(); insertImageFile(img); return; }
  const html = cd.getData('text/html');
  const text = cd.getData('text/plain');
  e.preventDefault();
  if (pastePlain || (!html && !looksMd(text))) {
    pastePlain = false;
    document.execCommand('insertText', false, text);
  } else if (html) {
    const clean = new DOMParser().parseFromString(`<body>${C.sanitize(html)}</body>`, 'text/html').body;
    $$('[class]', clean).forEach((n) => { const keep = [...n.classList].filter((c) => c === 'checklist'); if (keep.length) n.className = keep.join(' '); else n.removeAttribute('class'); });
    $$('[id]', clean).forEach((n) => n.removeAttribute('id'));
    document.execCommand('insertHTML', false, clean.innerHTML);
  } else {
    document.execCommand('insertHTML', false, C.sanitize(C.mdToHtml(text)));
    P.toast('Markdown biçimli metin olarak yapıştırıldı. Düz metin için Ctrl+Shift+V.');
  }
  changed();
}

/* ------------------------------------------------------------------ */
/* Yardım                                                              */
/* ------------------------------------------------------------------ */

function keysHelp() {
  const rows = [
    ['/', 'Komut menüsü (başlık, tablo, kutu, tarih…)'], ['Ctrl+K', 'Platform365 komut paleti'], ['# + boşluk', 'Başlık 1 (## Başlık 2, ### Başlık 3)'],
    ['- + boşluk', 'Madde listesi'], ['1. + boşluk', 'Numaralı liste'], ['[] + boşluk', 'Yapılacaklar listesi'], ['> + boşluk', 'Alıntı'],
    ['**metin**', 'Kalın'], ['*metin*', 'İtalik'], ['~~metin~~', 'Üstü çizili'], ['==metin==', 'Vurgulu'], ['`kod`', 'Satır içi kod'],
    ['--- + Enter', 'Ayraç çizgisi'], ['``` + Enter', 'Kod bloğu'], ['[toc] + Enter', 'İçindekiler'], ['12*3+4=', 'Satır içi hesap: =40 yazar'],
    ['Ctrl+Enter', 'Sayfa sonu'], ['Alt+↑ / Alt+↓', 'Paragrafı yukarı / aşağı taşı'], ['Tab (tabloda)', 'Sonraki hücre (sonda yeni satır)'],
    ['Ctrl+F / Ctrl+H', 'Bul / değiştir'], ['Ctrl+Shift+F', 'Odak modu'], ['Ctrl+Shift+V', 'Biçimsiz yapıştır'], ['Ctrl+Alt+1…3 / 0', 'Başlık 1–3 / Normal'],
    ['Ctrl+Shift+7 / 8 / 9', 'Numaralı / madde / yapılacaklar listesi'], ['Ctrl+S', 'Kaydet'], ['Ctrl+P', 'Yazdır / PDF'],
  ];
  P.dialog({ title: 'Kısayollar ve hızlı yazma', wide: true, html: `<div class="p365-keys">${rows.map(([a, b]) => `<kbd>${esc(a)}</kbd><span>${esc(b)}</span>`).join('')}</div>` });
}

/* ------------------------------------------------------------------ */
/* Olaylar                                                             */
/* ------------------------------------------------------------------ */

function runTool(t) {
  const tools = {
    focus: toggleFocus,
    typewriter: () => { settings.typewriter = !settings.typewriter; saveSettings(); P.toast(`Daktilo kaydırma ${settings.typewriter ? 'açık' : 'kapalı'}.`); typewriterScroll(); },
    dictate: toggleDictation, speak: toggleSpeak, pomodoro: togglePomodoro, goal: setGoal,
    find: () => openFind(false),
    'case-upper': () => transformSelection(trUpper), 'case-lower': () => transformSelection(trLower),
    'case-title': () => transformSelection(titleCase), 'case-sentence': () => transformSelection(sentenceCase),
    clean: cleanText, sort: sortList, toc: insertToc, 'to-present': toPresent, snapshot: snapshotNamed, page: pageSettings, keys: keysHelp,
  };
  tools[t] && tools[t]();
}

function bindEvents() {
  P.mountSwitcher('easywrite');
  $('#btnHome').addEventListener('click', () => { location.hash = ''; });
  $('#btnNewDoc').addEventListener('click', () => createDoc('blank'));
  $('#btnImport').addEventListener('click', () => $('#fileImport').click());
  $('#tplGrid').addEventListener('click', (e) => { const b = e.target.closest('[data-tpl]'); if (b) createDoc(b.dataset.tpl); });
  $('#fileImport').addEventListener('change', (e) => { const f = e.target.files[0]; if (f) importFile(f); e.target.value = ''; });
  $('#fileImage').addEventListener('change', (e) => { const f = e.target.files[0]; if (f) insertImageFile(f); e.target.value = ''; });

  $('#docTitle').addEventListener('input', (e) => { doc.title = e.target.value; scheduleSave(); });
  $('#docTitle').addEventListener('change', (e) => { doc.title = e.target.value.trim() || 'Adsız belge'; e.target.value = doc.title; scheduleSave(); });
  $('#docTitle').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); page().focus(); } });

  $('#btnUndo').addEventListener('click', () => exec('undo'));
  $('#btnRedo').addEventListener('click', () => exec('redo'));
  $('#btnPrint').addEventListener('click', async () => { await saveNow(); printDoc(); });
  $('#btnVersions').addEventListener('click', versions);
  $('#btnToPresent').addEventListener('click', toPresent);

  const menus = [['#btnTools', '#toolsMenu'], ['#btnDownload', '#downloadMenu']];
  for (const [b, m] of menus) {
    $(b).addEventListener('mousedown', saveSel);
    $(b).addEventListener('click', (e) => { e.stopPropagation(); menus.forEach(([, mm]) => { if (mm !== m) $(mm).hidden = true; }); $(m).hidden = !$(m).hidden; });
  }
  document.addEventListener('click', (e) => { if (!e.target.closest('.menu')) menus.forEach(([, m]) => { $(m).hidden = true; }); });
  $('#toolsMenu').addEventListener('click', (e) => { const b = e.target.closest('[data-tool]'); if (b) { $('#toolsMenu').hidden = true; runTool(b.dataset.tool); } });
  $('#downloadMenu').addEventListener('click', (e) => { const b = e.target.closest('[data-export]'); if (b) { $('#downloadMenu').hidden = true; exportAs(b.dataset.export); } });

  // Şerit
  const ribbon = $('#ribbon');
  ribbon.addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); else saveSel(); });
  ribbon.addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.cmd === 'insertUnorderedList' || b.dataset.cmd === 'insertOrderedList') { listify(b.dataset.cmd === 'insertOrderedList' ? 'ol' : 'ul'); return; }
    if (b.dataset.cmd) { exec(b.dataset.cmd); return; }
    const acts = {
      checklist: toggleChecklist, link: insertLink, image: () => { saveSel(); $('#fileImage').click(); }, table: insertTable,
      callout: () => insertCallout('info'), details: () => insertHtmlCaret('<details open><summary>{c}Başlık</summary><p><br></p></details><p><br></p>'),
      hr: () => exec('insertHorizontalRule'), pagebreak: () => insertHtmlCaret('<hr data-page-break><p>{c}</p>'), emoji: insertEmoji,
      find: () => openFind(false), focus: toggleFocus,
      panel: () => { const eb = $('.ed-body'); if (window.innerWidth <= 800) eb.classList.toggle('show-insights'); else eb.classList.toggle('no-panels'); },
    };
    acts[b.dataset.act] && acts[b.dataset.act]();
  });
  $('#blockStyle').addEventListener('change', (e) => formatBlock(e.target.value));
  $('#fontName').addEventListener('change', (e) => exec('fontName', e.target.value));
  $('#fontSize').addEventListener('change', (e) => setFontSize(+e.target.value));
  $('#lineHeight').addEventListener('change', (e) => { setLineHeight(e.target.value); e.target.value = ''; });
  $('#foreColor').addEventListener('input', (e) => { e.target.parentNode.style.setProperty('--c', e.target.value); });
  $('#foreColor').addEventListener('change', (e) => exec('foreColor', e.target.value));
  $('#hiliteColor').addEventListener('input', (e) => { e.target.parentNode.style.setProperty('--c', e.target.value); });
  $('#hiliteColor').addEventListener('change', (e) => setHilite(e.target.value));
  $$('.rb-color').forEach((l) => l.addEventListener('mousedown', saveSel));

  // Düzenleyici
  const pg = page();
  try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* eski tarayıcı */ }
  pg.addEventListener('keydown', onEditorKey);
  pg.addEventListener('input', () => {
    changed();
    if (slash) drawSlash();
    if (!pg.firstElementChild) { pg.innerHTML = '<p><br></p>'; placeCaret(pg.firstElementChild); }
    typewriterScroll();
  });
  pg.addEventListener('paste', onPaste);
  pg.addEventListener('click', (e) => {
    const t = e.target;
    selImg = null;
    $$('img.sel', pg).forEach((i) => i.classList.remove('sel'));
    if (t.tagName === 'IMG') { selImg = t; t.classList.add('sel'); updateCtxBar(); return; }
    const tocA = t.closest('nav[data-toc] a');
    if (tocA) { e.preventDefault(); jumpTo(tocA.dataset.target); return; }
    const a = t.closest('a[href]');
    if (a && (e.ctrlKey || e.metaKey)) { window.open(a.href, '_blank', 'noopener'); return; }
    const li = t.closest('ul.checklist > li');
    if (li) {
      const r = li.getBoundingClientRect();
      const em = parseFloat(getComputedStyle(li).fontSize) || 16;
      if (e.clientX - r.left < em * 1.6) { li.dataset.check = li.dataset.check === '1' ? '0' : '1'; changed(); }
    }
    const co = t.closest('[data-callout]');
    if (co && e.target === co) {
      const r = co.getBoundingClientRect();
      if (e.clientX - r.left < 36 && e.clientY - r.top < 40) { co.dataset.callout = CALLOUTS[(CALLOUTS.indexOf(co.dataset.callout) + 1) % CALLOUTS.length]; changed(); }
    }
  });
  pg.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files')) e.preventDefault(); });
  pg.addEventListener('drop', (e) => {
    const f = [...e.dataTransfer.files].find((x) => x.type.startsWith('image/'));
    if (f) {
      e.preventDefault();
      const pos = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
      if (pos) { getSelection().removeAllRanges(); getSelection().addRange(pos); }
      saveSel();
      insertImageFile(f);
    }
  });
  document.addEventListener('selectionchange', () => {
    if (!doc) return;
    saveSel();
    clearTimeout(updateToolbar.t);
    updateToolbar.t = setTimeout(updateToolbar, 60);
    const s = getSelection();
    const st = $('#stSel');
    if (s.rangeCount && !s.isCollapsed && inPage(s.anchorNode)) {
      const n = (s.toString().match(/[\p{L}\p{N}]+/gu) || []).length;
      st.hidden = !n;
      st.textContent = `${n} kelime seçili`;
    } else st.hidden = true;
    if (slash && slashQuery() === null) closeSlash();
  });
  $('#slashMenu').addEventListener('mousedown', (e) => { e.preventDefault(); const it = e.target.closest('[data-i]'); if (it) chooseSlash(+it.dataset.i); });
  $('#ctxBar').addEventListener('mousedown', (e) => { e.preventDefault(); const b = e.target.closest('[data-ctx]'); if (b) onCtx(b.dataset.ctx); });
  $('#pageWrap').addEventListener('scroll', () => { closeSlash(); if (!$('#ctxBar').hidden) updateCtxBar(); }, { passive: true });
  window.addEventListener('resize', () => { if (doc) { closeSlash(); updateCtxBar(); clearTimeout(autoZoom.t); autoZoom.t = setTimeout(autoZoom, 200); } });

  // Ana hat ve içgörüler
  $('#outline').addEventListener('click', (e) => { const a = e.target.closest('[data-target]'); if (a) jumpTo(a.dataset.target); });
  $('#insights').addEventListener('click', (e) => {
    const b = e.target.closest('[data-hl]');
    if (b) { hl[b.dataset.hl] = !hl[b.dataset.hl]; updateStats(); if (!HAS_HL) P.toast('Bu tarayıcı vurgulamayı desteklemiyor.'); return; }
    const w = e.target.closest('[data-word]');
    if (w) { hl.word = hl.word === w.dataset.word ? null : w.dataset.word; updateStats(); return; }
    if (e.target.closest('[data-ins="goal"]')) setGoal();
  });

  // Bul / değiştir
  $('#findText').addEventListener('input', () => runFind());
  $('#findCase').addEventListener('change', () => runFind());
  $('#findText').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); findStep(e.shiftKey ? -1 : 1); }
    if (e.key === 'Escape') { e.preventDefault(); closeFind(); page().focus(); }
  });
  $('#replText').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); replaceOne(); }
    if (e.key === 'Escape') { e.preventDefault(); closeFind(); page().focus(); }
  });
  $('#findNext').addEventListener('click', () => findStep(1));
  $('#findPrev').addEventListener('click', () => findStep(-1));
  $('#replOne').addEventListener('click', replaceOne);
  $('#replAll').addEventListener('click', replaceAll);
  $('#findClose').addEventListener('click', () => { closeFind(); page().focus(); });

  $('#zoom').addEventListener('input', (e) => setZoom(+e.target.value));

  // Ana sayfa sürükle-bırak
  const home = $('#homeView');
  home.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files')) e.preventDefault(); });
  home.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) { e.preventDefault(); importFile(f); } });

  document.addEventListener('keydown', onKey);
  window.addEventListener('hashchange', route);
  window.addEventListener('beforeunload', () => { if (doc) saveNow(); });
  document.addEventListener('visibilitychange', () => { if (document.hidden && doc) saveNow(); });

  P.registerCommands([
    { id: 'new', group: 'EasyWrite', title: 'Yeni boş belge', run: () => createDoc('blank') },
    ...TEMPLATES.filter((t) => t.id !== 'blank').map((t) => ({ id: `tpl-${t.id}`, group: 'Şablon', title: `Yeni: ${t.name}`, keywords: t.desc, run: () => createDoc(t.id) })),
    { id: 'open', group: 'Dosya', title: 'Dosya aç (.docx, .md, .txt, .html)', run: () => $('#fileImport').click() },
    { id: 'home', group: 'EasyWrite', title: 'Belgelerim', run: () => { location.hash = ''; } },
    ...[['easywrite', 'EasyWrite dosyası (.easywrite)'], ['docx', 'Word (.docx)'], ['pdf', 'PDF / Yazdır'], ['md', 'Markdown (.md)'], ['html', 'Web sayfası (.html)'], ['txt', 'Düz metin (.txt)']]
      .map(([k, t]) => ({ id: `dl-${k}`, group: 'İndir', title: t, when: () => !!doc, run: () => exportAs(k) })),
    { id: 'find', group: 'Düzen', title: 'Bul ve değiştir', key: 'Ctrl+F', when: () => !!doc, run: () => openFind(false) },
    { id: 'to-present', group: 'Dönüştür', title: 'Sunuma dönüştür (Present)', when: () => !!doc, run: toPresent },
    { id: 'versions', group: 'Dosya', title: 'Sürüm geçmişi ve karşılaştırma', when: () => !!doc, run: versions },
    { id: 'snapshot', group: 'Dosya', title: 'Sürüm kaydet', when: () => !!doc, run: snapshotNamed },
    { id: 'page', group: 'Görünüm', title: 'Sayfa düzeni (kağıt, kenar boşluğu, yazı tipi)', when: () => !!doc, run: pageSettings },
    { id: 'focus', group: 'Görünüm', title: 'Odak modu', key: 'Ctrl+Shift+F', when: () => !!doc, run: toggleFocus },
    { id: 'typewriter', group: 'Görünüm', title: 'Daktilo kaydırma aç/kapat', when: () => !!doc, run: () => runTool('typewriter') },
    { id: 'dictate', group: 'Araçlar', title: 'Sesle yaz (dikte)', when: () => !!doc, run: toggleDictation },
    { id: 'speak', group: 'Araçlar', title: 'Sesli oku', when: () => !!doc, run: toggleSpeak },
    { id: 'pomodoro', group: 'Araçlar', title: 'Pomodoro zamanlayıcı', when: () => !!doc, run: togglePomodoro },
    { id: 'goal', group: 'Araçlar', title: 'Yazma hedefi', when: () => !!doc, run: setGoal },
    { id: 'toc', group: 'Ekle', title: 'İçindekiler', when: () => !!doc, run: insertToc },
    { id: 'table', group: 'Ekle', title: 'Tablo', when: () => !!doc, run: insertTable },
    { id: 'image', group: 'Ekle', title: 'Görsel', when: () => !!doc, run: () => $('#fileImage').click() },
    { id: 'emoji', group: 'Ekle', title: 'Emoji', when: () => !!doc, run: insertEmoji },
    { id: 'clean', group: 'Metin', title: 'Metni temizle', when: () => !!doc, run: cleanText },
    { id: 'sort', group: 'Metin', title: 'Listeyi sırala (A→Z)', when: () => !!doc, run: sortList },
    { id: 'upper', group: 'Metin', title: 'BÜYÜK HARF (seçim)', when: () => !!doc, run: () => runTool('case-upper') },
    { id: 'lower', group: 'Metin', title: 'küçük harf (seçim)', when: () => !!doc, run: () => runTool('case-lower') },
    { id: 'titlec', group: 'Metin', title: 'Başlık Durumu (seçim)', when: () => !!doc, run: () => runTool('case-title') },
    { id: 'dark', group: 'Görünüm', title: 'Uygulama temasını değiştir (aydınlık / karanlık)', run: P.cycleTheme },
    { id: 'keys', group: 'Yardım', title: 'Kısayollar ve Markdown', run: keysHelp },
    { id: 'go-hub', group: 'Platform365', title: 'Ana ekran', run: () => { location.href = P.ROOT; } },
    { id: 'go-present', group: 'Platform365', title: 'Present\'i aç', run: () => { location.href = `${P.ROOT}present/`; } },
  ]);
}

function onKey(e) {
  if (!doc || P.hasDialog()) return;
  const mod = e.ctrlKey || e.metaKey;
  const k = e.key.toLowerCase();
  if (e.key === 'Escape') {
    if (!$('#findbar').hidden) { closeFind(); page().focus(); return; }
    if (document.body.classList.contains('focus-mode')) { toggleFocus(); return; }
    if (selImg) { selImg.classList.remove('sel'); selImg = null; hideCtx(); }
    return;
  }
  if (!mod) return;
  if (k === 's') { e.preventDefault(); saveNow(); P.toast('Kaydedildi.'); return; }
  if (k === 'p') { e.preventDefault(); saveNow().then(printDoc); return; }
  if (k === 'f' && e.shiftKey) { e.preventDefault(); toggleFocus(); return; }
  if (k === 'f') { e.preventDefault(); openFind(false); return; }
  if (k === 'h') { e.preventDefault(); openFind(true); return; }
  if (k === 'v' && e.shiftKey) { pastePlain = true; setTimeout(() => { pastePlain = false; }, 500); return; }
  if (e.altKey && /^[0-3]$/.test(e.key)) { e.preventDefault(); formatBlock(e.key === '0' ? 'p' : `h${e.key}`); return; }
  if (e.shiftKey && (e.code === 'Digit7' || e.code === 'Digit8' || e.code === 'Digit9')) {
    e.preventDefault();
    if (e.code === 'Digit7') listify('ol');
    else if (e.code === 'Digit8') listify('ul');
    else toggleChecklist();
  }
}

/* ------------------------------------------------------------------ */
/* PWA ve başlat                                                       */
/* ------------------------------------------------------------------ */

function setupPwa() {
  let deferred = null;
  const btn = $('#btnInstall');
  window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferred = e; btn.hidden = false; });
  btn.addEventListener('click', async () => {
    if (!deferred) return;
    deferred.prompt();
    await deferred.userChoice;
    deferred = null;
    btn.hidden = true;
  });
  window.addEventListener('appinstalled', () => { btn.hidden = true; P.toast('Platform365 yüklendi!'); });
  P.registerSw(async () => { if (doc) await saveNow(); });
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async (params) => { for (const h of params.files || []) importFile(await h.getFile()); });
  }
}

(async function init() {
  $$('.app-version').forEach((n) => { n.textContent = `v${APP_VERSION}`; });
  await Store.open();
  renderTemplates();
  bindEvents();
  setupPwa();
  await route();
})();
