/* Platform365 — Present ve EasyWrite için ortak çekirdek
 * Tema, uygulama değiştirici, iletişim kutuları, komut paleti (Ctrl+K),
 * uygulamalar arası gelen kutusu ve sürüm geçmişi (anlık görüntüler). */
'use strict';

(function () {
  const ROOT = (document.currentScript && document.currentScript.dataset.root) || './';
  const VERSION = '2.0.0';
  const NAME = 'Platform365';

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
    toast(`Tema: ${{ system: 'Sistem', light: 'Aydınlık', dark: 'Karanlık' }[next]}`);
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
  function dialog({ title = '', html = '', node = null, wide = false, actions = [{ label: 'Kapat', value: null }], onOpen = null } = {}) {
    return new Promise((resolve) => {
      const back = document.createElement('div');
      back.className = 'p365-backdrop';
      back.innerHTML = `<div class="p365-dialog${wide ? ' wide' : ''}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
        <header><h2>${esc(title)}</h2><button data-x aria-label="Kapat">✕</button></header>
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
      html: `${label ? `<p class="muted">${esc(label)}</p>` : ''}<input type="text" value="${esc(value)}">`,
      actions: [{ label: 'Vazgeç', value: null }, { label: 'Tamam', primary: true, value: (b) => b.querySelector('input').value }],
      onOpen: (b, close) => b.querySelector('input').addEventListener('keydown', (e) => { if (e.key === 'Enter') close(e.target.value); }),
    });
    return r;
  }

  /* ---------------- Komut paleti ---------------- */
  const commands = [];
  function registerCommands(list) {
    for (const c of list) {
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
    back.innerHTML = `<div class="p365-dialog" role="dialog" aria-label="Komut paleti">
      <input type="text" placeholder="Bir komut yazın… (ör. “sun”, “pdf”, “tema”)" aria-label="Komut ara">
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
        : '<li>Eşleşen komut yok</li>';
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
    { id: 'hub', name: NAME, desc: 'Ana ekran ve tüm dosyalar', href: ROOT, ico: '365' },
    { id: 'present', name: 'Present', desc: 'Sunumlar', href: `${ROOT}present/`, ico: '▶' },
    { id: 'easywrite', name: 'EasyWrite', desc: 'Belgeler', href: `${ROOT}easywrite/`, ico: 'W' },
  ];
  function mountSwitcher(app) {
    const bar = document.querySelector('.topbar');
    if (!bar || bar.querySelector('.p365-switch')) return;
    const wrap = document.createElement('div');
    wrap.className = 'p365-switch';
    wrap.innerHTML = `<button class="p365-waffle" aria-haspopup="true" title="${NAME} uygulamaları">${'<i></i>'.repeat(9)}</button>
      <div class="menu-list p365-apps" hidden>
        <div class="menu-label">${NAME}</div>
        ${APPS.map((a) => `<a href="${a.href}" class="${a.id === app ? 'on' : ''}"><span class="ico ico-${a.id}">${a.ico}</span><span><b>${a.name}</b><small>${a.desc}</small></span></a>`).join('')}
        <hr>
        <button data-p365="palette"><b>Komut paleti</b><small>Ctrl+K — tüm komutları arayın</small></button>
        <button data-p365="theme"><b>Temayı değiştir</b><small>Sistem / Aydınlık / Karanlık</small></button>
      </div>`;
    bar.insertBefore(wrap, bar.firstChild);
    const list = wrap.querySelector('.p365-apps');
    wrap.querySelector('.p365-waffle').addEventListener('click', (e) => { e.stopPropagation(); list.hidden = !list.hidden; });
    list.addEventListener('click', (e) => {
      const b = e.target.closest('[data-p365]');
      if (!b) return;
      list.hidden = true;
      if (b.dataset.p365 === 'palette') openPalette(); else cycleTheme();
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
      s.onerror = () => { zipPromise = null; rej(new Error('ZIP kütüphanesi yüklenemedi')); };
      document.head.appendChild(s);
    });
    return zipPromise;
  }
  const relTime = (t) => {
    const s = (Date.now() - t) / 1000;
    if (s < 60) return 'az önce';
    if (s < 3600) return `${Math.floor(s / 60)} dk önce`;
    if (s < 86400) return `${Math.floor(s / 3600)} sa önce`;
    if (s < 86400 * 7) return `${Math.floor(s / 86400)} gün önce`;
    return new Date(t).toLocaleDateString('tr-TR', { dateStyle: 'medium' });
  };

  const EMOJI = ('😀 😃 😄 😁 😆 😅 😂 🙂 😉 😊 😍 🤩 😎 🤔 😮 😢 😡 👍 👎 👏 🙌 🙏 💪 👀 ✅ ❌ ⚠️ ❗ ❓ 💡 🔥 ⭐ 🌟 ✨ 🎯 🚀 📈 📉 📊 📌 📎 📝 📚 📅 ⏰ ⌛ 💬 🔔 🔒 🔑 🧠 💼 🏆 🎉 🎁 ❤️ 💙 💚 💛 🧡 💜 ☀️ 🌙 🌍 🌱 🍀 ☕ 🍎 🎵 🎨 📷 💻 📱 🖥️ ⚙️ 🔧 🧪 🧩 🏠 🚗 ✈️ 🇹🇷 ➡️ ⬅️ ⬆️ ⬇️ ✔️ ➕ ➖ 💯 🔴 🟠 🟡 🟢 🔵 🟣 ⚫ ⚪').split(' ');
  function pickEmoji() {
    return dialog({
      title: 'Emoji ekle',
      html: `<div class="p365-emoji">${EMOJI.map((e) => `<button data-e="${e}">${e}</button>`).join('')}</div>`,
      actions: [],
      onOpen: (body, close) => body.addEventListener('click', (e) => { const b = e.target.closest('[data-e]'); if (b) close(b.dataset.e); }),
    });
  }

  /** Konuşma tanıma (Türkçe). Destek yoksa null döner. */
  function speech({ onText, onEnd, interim = true, lang = 'tr-TR' }) {
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

  /** Service worker kaydı (tüm platform tek önbellek) */
  function registerSw(beforeReload) {
    if (!('serviceWorker' in navigator) || location.protocol === 'file:') return;
    const hadController = !!navigator.serviceWorker.controller;
    let reloading = false;
    navigator.serviceWorker.addEventListener('controllerchange', async () => {
      if (!hadController || reloading) return;
      reloading = true;
      if (beforeReload) await beforeReload();
      location.reload();
    });
    navigator.serviceWorker.register(`${ROOT}sw.js`, { scope: ROOT, updateViaCache: 'none' })
      .then((reg) => reg.update())
      .catch((e) => console.warn('Service worker kaydedilemedi', e));
  }

  window.P365 = {
    NAME, VERSION, ROOT, APPS_DB, esc, uid, toast, dialog, prompt, hasDialog,
    registerCommands, openPalette, mountSwitcher, setTheme, cycleTheme, readTheme,
    inbox, snapshots, appStore, safeName, download, loadZip, relTime, EMOJI, pickEmoji, speech, registerSw,
  };
})();
