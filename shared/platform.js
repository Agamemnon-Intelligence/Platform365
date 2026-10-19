/* Platform365 — Present ve EasyWrite için ortak çekirdek
 * Tema, uygulama değiştirici, iletişim kutuları, komut paleti (Ctrl+K),
 * uygulamalar arası gelen kutusu ve sürüm geçmişi (anlık görüntüler). */
'use strict';

(function () {
  const ROOT = (document.currentScript && document.currentScript.dataset.root) || './';
  const VERSION = '3.0.0';
  const NAME = 'Platform365';
  const I = window.P365_I18N || { lang: 'tr', chosen: true, locale: 'tr-TR', t: (s) => s, L: (tr) => tr, setLang() {}, browserLang: () => 'tr' };
  const { t, L } = I;

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

  /* ---------------- Tema ---------------- */
  const THEME_KEY = 'p365-theme';
  const readTheme = () => { try { return localStorage.getItem(THEME_KEY) || 'system'; } catch (e) { return 'system'; } };
  function applyTheme(mode) {
    const html = document.documentElement;
    if (mode === 'light' || mode === 'dark') html.dataset.theme = mode; else delete html.dataset.theme;
  }
  function setTheme(mode) {
    try { localStorage.setItem(THEME_KEY, mode); } catch (e) { /* özel pencere */ }
    applyTheme(mode);
  }
  function cycleTheme() {
    const order = ['system', 'light', 'dark'];
    const next = order[(order.indexOf(readTheme()) + 1) % order.length];
    setTheme(next);
    toast(t('Tema: {name}', { name: t({ system: 'Sistem', light: 'Aydınlık', dark: 'Karanlık' }[next]) }));
    return next;
  }
  applyTheme(readTheme());
  window.addEventListener('storage', (e) => { if (e.key === THEME_KEY) applyTheme(readTheme()); });

  /* ---------------- IndexedDB (ortak) ---------------- */
  let dbp = null;
  function db() {
    if (dbp) return dbp;
    dbp = new Promise((res, rej) => {
      if (!('indexedDB' in window)) { rej(new Error('IndexedDB yok')); return; }
      const r = indexedDB.open('platform365-db', 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        if (!d.objectStoreNames.contains('inbox')) d.createObjectStore('inbox', { keyPath: 'id' });
        if (!d.objectStoreNames.contains('snapshots')) {
          const s = d.createObjectStore('snapshots', { keyPath: 'id' });
          s.createIndex('doc', 'docKey');
        }
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    dbp.catch(() => { dbp = null; });
    return dbp;
  }
  async function tx(store, mode, fn) {
    const d = await db();
    return new Promise((res, rej) => {
      const r = fn(d.transaction(store, mode).objectStore(store));
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }

  /** Başka bir uygulamanın veritabanını açar (ana ekran ve dönüşümler için) */
  function openAppDb(name, store) {
    return new Promise((res) => {
      if (!('indexedDB' in window)) { res(null); return; }
      const r = indexedDB.open(name, 1);
      r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains(store)) r.result.createObjectStore(store, { keyPath: 'id' }); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => res(null);
    });
  }
  async function appStore(name, store) {
    const d = await openAppDb(name, store);
    const run = (mode, fn) => new Promise((res, rej) => {
      if (!d) { res(mode === 'readonly' ? [] : null); return; }
      const r = fn(d.transaction(store, mode).objectStore(store));
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return {
      all: () => run('readonly', (s) => s.getAll()),
      get: (id) => run('readonly', (s) => s.get(id)),
      put: (v) => run('readwrite', (s) => s.put(v)),
      del: (id) => run('readwrite', (s) => s.delete(id)),
      close: () => d && d.close(),
    };
  }
  const APPS_DB = {
    present: { db: 'present-db', store: 'decks' },
    easywrite: { db: 'easywrite-db', store: 'docs' },
  };

  /* Gelen kutusu: uygulamalar arasında dosya / içerik aktarımı */
  const inbox = {
    async put(item) { const id = uid(); await tx('inbox', 'readwrite', (s) => s.put({ ...item, id, at: Date.now() })); return id; },
    async take(id) {
      const v = await tx('inbox', 'readonly', (s) => s.get(id));
      if (v) await tx('inbox', 'readwrite', (s) => s.delete(id));
      return v || null;
    },
    /** Öğeyi kutuya koyup hedef uygulamayı açar */
    async send(app, item) {
      const id = await this.put(item);
      location.href = `${ROOT}${app}/#/inbox/${id}`;
    },
  };

  /* Sürüm geçmişi */
  const snapshots = {
    async save(app, docId, title, data, label = '') {
      const snap = { id: uid(), docKey: `${app}:${docId}`, app, docId, title, label, at: Date.now(), data: JSON.stringify(data) };
      await tx('snapshots', 'readwrite', (s) => s.put(snap));
      // En fazla 40 sürüm sakla (etiketsiz olanlardan en eskiler silinir)
      const list = await this.list(app, docId);
      const extra = list.filter((x) => !x.label).slice(40);
      for (const x of extra) await this.del(x.id);
      return snap;
    },
    async list(app, docId) {
      const all = await tx('snapshots', 'readonly', (s) => s.index('doc').getAll(`${app}:${docId}`));
      return all.sort((a, b) => b.at - a.at);
    },
    async all() { return tx('snapshots', 'readonly', (s) => s.getAll()); },
    async get(id) { return tx('snapshots', 'readonly', (s) => s.get(id)); },
    async del(id) { return tx('snapshots', 'readwrite', (s) => s.delete(id)); },
    async put(v) { return tx('snapshots', 'readwrite', (s) => s.put(v)); },
  };

  /* ---------------- Bildirim ---------------- */
  let toastTimer = null;
  function toast(msg) {
    let t = document.getElementById('toast');
    if (!t) { t = document.createElement('div'); t.id = 'toast'; t.className = 'toast'; document.body.appendChild(t); }
    msg = I.t(msg);
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, Math.max(2600, msg.length * 70));
  }

  /* ---------------- İletişim kutusu ---------------- */
  const openDialogs = [];
  /**
   * dialog({ title, html | node, wide, actions: [{ label, value, primary, danger }], onOpen(body, close) })
   * Kapatılınca seçilen eylemin değeriyle (veya null) çözülür.
   */
  function dialog({ title = '', html = '', node = null, wide = false, actions = [{ label: 'Kapat', value: null }], onOpen = null, className = '' } = {}) {
    title = t(title);
    actions = actions.map((a) => ({ ...a, label: t(a.label) }));
    return new Promise((resolve) => {
      const back = document.createElement('div');
      back.className = 'p365-backdrop';
      back.innerHTML = `<div class="p365-dialog${wide ? ' wide' : ''}${className ? ` ${className}` : ''}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
        <header><h2>${esc(title)}</h2><button data-x aria-label="${t('Kapat')}">✕</button></header>
        <div class="body"></div>
        ${actions.length ? `<footer>${actions.map((a, i) => `<button class="btn${a.primary ? ' primary' : ''}${a.danger ? ' danger' : ''}" data-i="${i}">${esc(a.label)}</button>`).join('')}</footer>` : ''}
      </div>`;
      const body = back.querySelector('.body');
      if (node) body.appendChild(node); else body.innerHTML = html;
      let done = false;
      const close = (v = null) => {
        if (done) return;
        done = true;
        back.remove();
        openDialogs.splice(openDialogs.indexOf(close), 1);
        resolve(v);
      };
      openDialogs.push(close);
      back.addEventListener('pointerdown', (e) => { if (e.target === back) close(null); });
      back.querySelector('[data-x]').addEventListener('click', () => close(null));
      back.querySelectorAll('footer [data-i]').forEach((b) => b.addEventListener('click', () => {
        const a = actions[+b.dataset.i];
        close(typeof a.value === 'function' ? a.value(body) : a.value);
      }));
      back.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') { e.stopPropagation(); close(null); }
        e.stopPropagation();
      });
      document.body.appendChild(back);
      const f = back.querySelector('input, textarea, select') || back.querySelector('footer .primary') || back.querySelector('[data-x]');
      setTimeout(() => f && f.focus(), 0);
      if (onOpen) onOpen(body, close);
    });
  }
  const hasDialog = () => openDialogs.length > 0;

  async function prompt(title, value = '', label = '') {
    const r = await dialog({
      title,
      html: `${label ? `<p class="muted">${esc(t(label))}</p>` : ''}<input type="text" value="${esc(value)}">`,
      actions: [{ label: 'Vazgeç', value: null }, { label: 'Tamam', primary: true, value: (b) => b.querySelector('input').value }],
      onOpen: (b, close) => b.querySelector('input').addEventListener('keydown', (e) => { if (e.key === 'Enter') close(e.target.value); }),
    });
    return r;
  }

  /* ---------------- Komut paleti ---------------- */
  const commands = [];
  function registerCommands(list) {
    for (const src of list) {
      // Başlık ve grup arayüz dilinde gösterilir; Türkçe aslı her iki dilde de aranabilir kalır
      const c = { ...src, title: t(src.title), group: src.group && t(src.group), keywords: [src.keywords, src.title !== t(src.title) ? src.title : ''].filter(Boolean).join(' ') };
      const i = commands.findIndex((x) => x.id === c.id);
      if (i >= 0) commands[i] = c; else commands.push(c);
    }
  }
  const norm = (s) => String(s).toLocaleLowerCase('tr').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ı/g, 'i');
  function score(q, text) {
    if (!q) return 1;
    const t = norm(text);
    if (t.includes(q)) return 100 - t.indexOf(q);
    let ti = 0;
    for (const ch of q) {
      ti = t.indexOf(ch, ti);
      if (ti < 0) return 0;
      ti += 1;
    }
    return 10;
  }
  function openPalette() {
    if (document.querySelector('.p365-palette')) return;
    const back = document.createElement('div');
    back.className = 'p365-backdrop p365-palette';
    back.innerHTML = `<div class="p365-dialog" role="dialog" aria-label="${t('Komut paleti')}">
      <input type="text" placeholder="${t('Bir komut yazın… (ör. “sun”, “pdf”, “tema”)')}" aria-label="${t('Komut ara')}">
      <ul class="p365-cmds"></ul></div>`;
    const input = back.querySelector('input');
    const ul = back.querySelector('ul');
    let items = [];
    let on = 0;
    const visible = () => commands.filter((c) => !c.when || c.when());
    const draw = () => {
      const q = norm(input.value.trim());
      items = visible().map((c) => ({ c, s: score(q, `${c.group || ''} ${c.title} ${c.keywords || ''}`) }))
        .filter((x) => x.s > 0).sort((a, b) => b.s - a.s).slice(0, 40).map((x) => x.c);
      on = Math.min(on, Math.max(0, items.length - 1));
      ul.innerHTML = items.length ? items.map((c, i) => `<li data-i="${i}" class="${i === on ? 'on' : ''}"><span class="g">${esc(c.group || '')}</span><span>${esc(c.title)}</span>${c.key ? `<span class="k">${esc(c.key)}</span>` : ''}</li>`).join('')
        : `<li>${t('Eşleşen komut yok')}</li>`;
      ul.querySelector('.on')?.scrollIntoView({ block: 'nearest' });
    };
    const close = () => back.remove();
    const run = (c) => { close(); if (c) setTimeout(() => c.run(), 0); };
    input.addEventListener('input', () => { on = 0; draw(); });
    back.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') close();
      else if (e.key === 'ArrowDown') { e.preventDefault(); on = Math.min(items.length - 1, on + 1); draw(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); on = Math.max(0, on - 1); draw(); }
      else if (e.key === 'Enter') { e.preventDefault(); run(items[on]); }
    });
    ul.addEventListener('click', (e) => { const li = e.target.closest('[data-i]'); if (li) run(items[+li.dataset.i]); });
    back.addEventListener('pointerdown', (e) => { if (e.target === back) close(); });
    document.body.appendChild(back);
    draw();
    input.focus();
  }

  /* ---------------- Uygulama değiştirici ---------------- */
  const APPS = [
    { id: 'hub', name: NAME, desc: t('Ana ekran ve tüm dosyalar'), href: ROOT, ico: '365' },
    { id: 'present', name: 'Present', desc: t('Sunumlar'), href: `${ROOT}present/`, ico: '▶' },
    { id: 'easywrite', name: 'EasyWrite', desc: t('Belgeler'), href: `${ROOT}easywrite/`, ico: 'W' },
  ];
  function mountSwitcher(app) {
    const bar = document.querySelector('.topbar');
    if (!bar || bar.querySelector('.p365-switch')) return;
    const wrap = document.createElement('div');
    wrap.className = 'p365-switch';
    wrap.innerHTML = `<button class="p365-waffle" aria-haspopup="true" title="${t('{name} uygulamaları', { name: NAME })}">${'<i></i>'.repeat(9)}</button>
      <div class="menu-list p365-apps" hidden>
        <div class="menu-label">${NAME}</div>
        ${APPS.map((a) => `<a href="${a.href}" class="${a.id === app ? 'on' : ''}"><span class="ico ico-${a.id}">${a.ico}</span><span><b>${a.name}</b><small>${a.desc}</small></span></a>`).join('')}
        <hr>
        <button data-p365="palette"><b>${t('Komut paleti')}</b><small>${t('Ctrl+K — tüm komutları arayın')}</small></button>
        <button data-p365="theme"><b>${t('Temayı değiştir')}</b><small>${t('Sistem / Aydınlık / Karanlık')}</small></button>
        <button data-p365="lang"><b>🌐 ${L('Dil: Türkçe', 'Language: English')}</b><small>${L('Türkçe / English', 'English / Türkçe')}</small></button>
        <button data-p365="about"><b>${t('Platform365 hakkında')}</b><small>${t('Sürüm {v}', { v: VERSION })} · ${t('Sürüm, lisans ve bağlantılar')}</small></button>
      </div>`;
    bar.insertBefore(wrap, bar.firstChild);
    const list = wrap.querySelector('.p365-apps');
    wrap.querySelector('.p365-waffle').addEventListener('click', (e) => { e.stopPropagation(); list.hidden = !list.hidden; });
    list.addEventListener('click', (e) => {
      const b = e.target.closest('[data-p365]');
      if (!b) return;
      list.hidden = true;
      ({ palette: openPalette, theme: cycleTheme, lang: chooseLanguage, about: showAbout })[b.dataset.p365]();
    });
    document.addEventListener('click', (e) => { if (!e.target.closest('.p365-switch')) list.hidden = true; });
  }

  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'k' && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      openPalette();
    }
  }, true);

  /* ---------------- Yardımcılar ---------------- */
  const TR_MAP = { ı: 'i', İ: 'I', ş: 's', Ş: 'S', ğ: 'g', Ğ: 'G', ü: 'u', Ü: 'U', ö: 'o', Ö: 'O', ç: 'c', Ç: 'C' };
  const safeName = (s, def = 'dosya') => (s || '')
    .replace(/[ıİşŞğĞüÜöÖçÇ]/g, (c) => TR_MAP[c])
    .normalize('NFKD').replace(/[^\x20-\x7e]/g, '')
    .replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, ' ').trim() || def;
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  let zipPromise = null;
  function loadZip() {
    if (window.JSZip) return Promise.resolve(window.JSZip);
    if (zipPromise) return zipPromise;
    zipPromise = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = `${ROOT}vendor/pptxgen.bundle.js`;
      s.onload = () => res(window.JSZip);
      s.onerror = () => { zipPromise = null; rej(new Error(t('ZIP kütüphanesi yüklenemedi'))); };
      document.head.appendChild(s);
    });
    return zipPromise;
  }
  const relTime = (t) => {
    const s = (Date.now() - t) / 1000;
    if (s < 60) return I.t('az önce');
    if (s < 3600) return I.t('{n} dk önce', { n: Math.floor(s / 60) });
    if (s < 86400) return I.t('{n} sa önce', { n: Math.floor(s / 3600) });
    if (s < 86400 * 7) return I.t('{n} gün önce', { n: Math.floor(s / 86400) });
    return new Date(t).toLocaleDateString(I.locale, { dateStyle: 'medium' });
  };

  const EMOJI = ('😀 😃 😄 😁 😆 😅 😂 🙂 😉 😊 😍 🤩 😎 🤔 😮 😢 😡 👍 👎 👏 🙌 🙏 💪 👀 ✅ ❌ ⚠️ ❗ ❓ 💡 🔥 ⭐ 🌟 ✨ 🎯 🚀 📈 📉 📊 📌 📎 📝 📚 📅 ⏰ ⌛ 💬 🔔 🔒 🔑 🧠 💼 🏆 🎉 🎁 ❤️ 💙 💚 💛 🧡 💜 ☀️ 🌙 🌍 🌱 🍀 ☕ 🍎 🎵 🎨 📷 💻 📱 🖥️ ⚙️ 🔧 🧪 🧩 🏠 🚗 ✈️ 🇹🇷 ➡️ ⬅️ ⬆️ ⬇️ ✔️ ➕ ➖ 💯 🔴 🟠 🟡 🟢 🔵 🟣 ⚫ ⚪').split(' ');
  function pickEmoji() {
    return dialog({
      title: t('Emoji ekle'),
      html: `<div class="p365-emoji">${EMOJI.map((e) => `<button data-e="${e}">${e}</button>`).join('')}</div>`,
      actions: [],
      onOpen: (body, close) => body.addEventListener('click', (e) => { const b = e.target.closest('[data-e]'); if (b) close(b.dataset.e); }),
    });
  }

  /** Konuşma tanıma (arayüz dilinde). Destek yoksa null döner. */
  function speech({ onText, onEnd, interim = true, lang = I.locale }) {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) return null;
    const r = new SR();
    r.lang = lang;
    r.continuous = true;
    r.interimResults = interim;
    let stopped = false;
    r.onresult = (e) => {
      let fin = '';
      let tmp = '';
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) fin += e.results[i][0].transcript; else tmp += e.results[i][0].transcript;
      }
      onText(fin, tmp);
    };
    r.onend = () => { if (!stopped) { try { r.start(); } catch (err) { onEnd && onEnd(); } } else onEnd && onEnd(); };
    r.onerror = (e) => { if (e.error === 'not-allowed' || e.error === 'service-not-allowed') { stopped = true; toast('Mikrofon izni verilmedi.'); } };
    r.start();
    return { stop() { stopped = true; try { r.stop(); } catch (e) { /* zaten durdu */ } } };
  }

  /* ---------------- Masaüstü uygulaması / indirme ---------------- */
  const REPO = 'Agamemnon-Intelligence/Platform365';
  const RELEASES = `https://github.com/${REPO}/releases/latest`;
  const DOWNLOADS = {
    win: { label: t('Windows için indir (.exe)'), url: `${RELEASES}/download/Platform365-Setup.exe` },
    mac: { label: t('macOS için indir (.dmg)'), url: `${RELEASES}/download/Platform365-macOS.dmg` },
  };
  const DL_KEY = 'p365-download-seen';
  /** Electron içinde çalışıyorsa önyükleme betiğinin sunduğu köprü, aksi halde null */
  const desktop = window.p365Desktop || null;
  const userOs = () => {
    const p = ((navigator.userAgentData && navigator.userAgentData.platform) || navigator.platform || navigator.userAgent || '').toLowerCase();
    return /mac|iphone|ipad/.test(p) ? 'mac' : /win/.test(p) ? 'win' : '';
  };

  /** Masaüstü sürümünü indirme penceresi (Windows .exe / macOS .dmg) */
  function showDownload() {
    const os = userOs();
    const order = os === 'mac' ? ['mac', 'win'] : ['win', 'mac'];
    return dialog({
      title: 'Platform365\'i indirin',
      html: `<div class="p365-download">
        <img src="${ROOT}icons/platform365.svg" alt="" width="64" height="64">
        <p><b>${t('Platform365 masaüstü uygulaması')}</b>${t(" Windows ve macOS'ta çalışır: çevrimdışı, kendi penceresinde ve .pptx / .docx / .present / .easywrite dosyalarını çift tıklayarak açar.")}</p>
        <div class="p365-dl-buttons">${order.map((k, i) => `<a class="btn lg${i === 0 && os ? ' primary' : ''}" href="${DOWNLOADS[k].url}" data-dl>${DOWNLOADS[k].label}</a>`).join('')}</div>
        <p class="muted">${t('Tüm sürümler:')} <a href="${RELEASES}" target="_blank" rel="noopener">GitHub Releases</a></p>
      </div>`,
      actions: [{ label: 'Tarayıcıda devam et', value: null }],
      onOpen: (body, close) => body.querySelectorAll('[data-dl]').forEach((a) => a.addEventListener('click', () => setTimeout(() => close('download'), 300))),
    });
  }

  /* ---------------- Dil seçimi ---------------- */
  /** İki dilli dil seçim penceresi; seçim değişirse sayfayı yeni dilde yeniden yükler */
  async function chooseLanguage() {
    const pre = I.chosen ? I.lang : I.browserLang();
    const opt = (code, flag, name, sub) => `<button class="p365-lang-opt${code === pre ? ' on' : ''}" data-lang="${code}"><span class="flag">${flag}</span><span><b>${name}</b><small>${sub}</small></span></button>`;
    const picked = await dialog({
      title: 'Dil seçin · Choose your language',
      className: 'p365-lang-dialog',
      html: `<div class="p365-lang">
        <img src="${ROOT}icons/platform365.svg" alt="" width="56" height="56">
        <p>Platform365'e hoş geldiniz! Arayüz dilini seçin.<br><span class="muted">Welcome to Platform365! Choose the interface language.</span></p>
        <div class="p365-lang-opts">${opt('tr', '🇹🇷', 'Türkçe', 'Arayüz Türkçe olur')}${opt('en', '🇬🇧', 'English', 'Use the interface in English')}</div>
        <p class="muted small">Daha sonra uygulama menüsünden (▦) değiştirebilirsiniz · You can change this later from the app menu (▦).</p>
      </div>`,
      actions: [],
      onOpen: (body, close) => {
        body.addEventListener('click', (e) => { const b = e.target.closest('[data-lang]'); if (b) close(b.dataset.lang); });
        setTimeout(() => body.querySelector('.p365-lang-opt.on')?.focus(), 0);
      },
    });
    const next = picked || pre;
    I.setLang(next);
    if (desktop && desktop.setLanguage) desktop.setLanguage(next);
    if (next !== I.lang) { location.reload(); return true; }
    return false;
  }

  /* ---------------- Hakkında ---------------- */
  const LINKS = {
    repo: `https://github.com/${REPO}`,
    notes: `https://github.com/${REPO}/blob/main/CHANGELOG.md`,
    releases: `https://github.com/${REPO}/releases`,
  };
  /** Hakkında içeriği: ana ekrandaki bölüm ve "Hakkında" penceresi aynı içeriği kullanır */
  function aboutHtml() {
    const env = desktop ? `${t('Masaüstü uygulaması')} (${{ win: 'Windows', mac: 'macOS' }[desktop.platform] || 'Linux'})` : t('Web tarayıcısı');
    const row = (k, v) => `<dt>${t(k)}</dt><dd>${v}</dd>`;
    return `<div class="p365-about">
      <div class="p365-about-head">
        <img src="${ROOT}icons/platform365.svg" alt="" width="64" height="64">
        <div><h3>${NAME}</h3><p class="p365-about-ver"><span class="p365-ver-pill">v${VERSION}</span> ${t('Sürüm {v}', { v: VERSION })}</p></div>
      </div>
      <p>${t('Present (sunum) ve EasyWrite (belge) uygulamalarını tek çatıda birleştiren, hesap gerektirmeyen ve çevrimdışı çalışan ofis paketi.')}</p>
      <div class="p365-about-apps">
        <a href="${ROOT}present/"><span class="ico ico-present">▶</span><span><b>Present</b><small>${t('Sunum hazırlama ve sunma')}</small></span></a>
        <a href="${ROOT}easywrite/"><span class="ico ico-easywrite">W</span><span><b>EasyWrite</b><small>${t('Belge ve kelime işlemci')}</small></span></a>
      </div>
      <dl class="p365-about-facts">
        ${row('Yayıncı', 'Agamemnon Intelligence')}
        ${row('Lisans', 'MIT')}
        ${row('Çalışma ortamı', env)}
        ${row('Arayüz dili', L('Türkçe', 'English'))}
        ${row('Verileriniz', t('Yalnızca bu cihazda'))}
      </dl>
      <div class="p365-about-new"><b>${t('3.0.0 ile gelenler')}</b><ul>
        <li>${t('İngilizce arayüz; ilk açılışta dil seçimi')}</li>
        <li>${t('Yenilenen ana ekran ve hızlı başlangıç')}</li>
        <li>${t('Platform365 hakkında bölümü')}</li>
        <li>${t('İngilizce metinler için Flesch okunabilirlik puanı')}</li>
      </ul></div>
      <p class="p365-about-links">
        <a href="${LINKS.repo}" target="_blank" rel="noopener">${t('Kaynak kodu (GitHub)')}</a>
        <a href="${LINKS.notes}" target="_blank" rel="noopener">${t('Sürüm notları')}</a>
        <a href="${LINKS.releases}" target="_blank" rel="noopener">${t('Tüm sürümler')}</a>
      </p>
      <p class="muted small">© 2026 Platform365 (Athene Softworks and Platform365 Technologies) · MIT<br>${t('Açık kaynak bileşenler: PptxGenJS ve JSZip (MIT).')}</p>
    </div>`;
  }
  const showAbout = () => dialog({ title: t('Platform365 hakkında'), html: aboutHtml() });

  /**
   * Uygulama başlangıcı: eski PWA önbelleğini temizler, ilk açılışta dil seçimini sorar,
   * web'de ilk açılışta indirme penceresini gösterir, masaüstünde işletim sisteminden açılan dosyaları onFile'a iletir.
   */
  function initApp({ onFile } = {}) {
    // Önceki PWA sürümünden kalan service worker ve önbellekleri kaldır
    if ('serviceWorker' in navigator) navigator.serviceWorker.getRegistrations().then((rs) => rs.forEach((r) => r.unregister())).catch(() => {});
    if (window.caches) caches.keys().then((ks) => ks.filter((k) => /^(present|platform365)-/.test(k)).forEach((k) => caches.delete(k))).catch(() => {});

    registerCommands([
      { id: 'p365-lang', group: 'Görünüm', title: 'Dili değiştir (Türkçe / English)', keywords: 'dil language english türkçe', run: chooseLanguage },
      { id: 'p365-about', group: 'Yardım', title: 'Platform365 hakkında', keywords: 'about sürüm version', run: showAbout },
    ]);
    const btn = document.getElementById('btnInstall');
    if (desktop) {
      document.documentElement.classList.add('p365-desktop', `p365-${desktop.platform}`);
      if (btn) btn.remove();
      if (desktop.setLanguage) desktop.setLanguage(I.lang);
      if (desktop.onMenu) desktop.onMenu((cmd) => { if (cmd === 'about') showAbout(); else if (cmd === 'lang-tr' || cmd === 'lang-en') { I.setLang(cmd.slice(5)); location.reload(); } });
      if (onFile) desktop.onOpenFile((f) => onFile(new File([f.data], f.name)));
      if (!I.chosen) chooseLanguage();
      return;
    }
    if (btn) { btn.hidden = false; btn.addEventListener('click', showDownload); }
    (async () => {
      if (!I.chosen && await chooseLanguage()) return; // yeni dilde yeniden yükleniyor
      let seen = false;
      try { seen = !!localStorage.getItem(DL_KEY); localStorage.setItem(DL_KEY, '1'); } catch (e) { /* özel pencere */ }
      if (!seen && location.protocol !== 'file:') showDownload();
    })();
  }

  window.P365 = {
    NAME, VERSION, ROOT, APPS_DB, esc, uid, toast, dialog, prompt, hasDialog,
    registerCommands, openPalette, mountSwitcher, setTheme, cycleTheme, readTheme,
    inbox, snapshots, appStore, safeName, download, loadZip, relTime, EMOJI, pickEmoji, speech, desktop, showDownload, initApp,
    t, L, lang: I.lang, locale: I.locale, chooseLanguage, aboutHtml, showAbout,
  };
})();
