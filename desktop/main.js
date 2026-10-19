/* Platform365 masaüstü uygulaması (Windows / macOS) — Electron ana süreci
 * Web uygulamasını app://platform365/ adresinden sunar; böylece IndexedDB verileri
 * kalıcı bir kökene bağlı kalır ve göreli bağlantılar web sürümündeki gibi çalışır. */
'use strict';

const { app, BrowserWindow, Menu, protocol, net, shell, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { pathToFileURL } = require('url');

const ROOT = path.join(__dirname, '..');
const ORIGIN = 'app://platform365';
const OPENABLE = /\.(present|pptx|easywrite|docx|md|markdown|txt|html?)$/i;

protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } },
]);

// Tek örnek: ikinci kez açılan dosyalar mevcut pencereye gider
if (!app.requestSingleInstanceLock()) { app.quit(); process.exit(0); }

let win = null;
const pending = []; // Pencere hazır olmadan açılan dosyalar
let lang = 'tr'; // Arayüz dili; sayfa kendi tercihini bildirince menü yeniden kurulur

// Dosyayı açabilecek uygulama: Present (.present/.pptx) veya EasyWrite (diğerleri)
const appFor = (name) => (/\.(present|pptx)$/i.test(name) ? 'present' : 'easywrite');

function sendFile(file) {
  if (!file || !OPENABLE.test(file)) return;
  let data;
  try { data = fs.readFileSync(file); } catch (e) { return; }
  const item = { name: path.basename(file), data };
  if (!win) { pending.push(item); return; }
  if (win.isMinimized()) win.restore();
  win.focus();
  const wc = win.webContents;
  const target = appFor(item.name);
  let page = '';
  try { page = new URL(wc.getURL()).pathname; } catch (e) { /* henüz yüklenmedi */ }
  // Ana ekran her türü yönlendirir; açık uygulama yalnızca kendi türünü açar
  if (wc.listenerReady && (page === '/' || page === '/index.html' || page.startsWith(`/${target}/`))) {
    wc.send('p365:open-file', item);
  } else {
    pending.push(item);
    if (wc.listenerReady) win.loadURL(`${ORIGIN}/${target}/`);
  }
}
const filesFromArgv = (argv) => argv.slice(app.isPackaged ? 1 : 2).filter((a) => !a.startsWith('-') && OPENABLE.test(a));

// macOS: Finder'dan dosya açma
app.on('open-file', (e, file) => { e.preventDefault(); sendFile(file); });
app.on('second-instance', (e, argv) => {
  filesFromArgv(argv).forEach(sendFile);
  if (win) { if (win.isMinimized()) win.restore(); win.focus(); }
});

// Önyükleme betiği dinlemeye başladığında bekleyen dosyaları teslim et
ipcMain.handle('p365:ready', (e) => {
  e.sender.listenerReady = true;
  return pending.splice(0);
});
// Sayfa değişince yeniden dinlemesini bekle
const resetReady = (wc) => wc.on('did-start-navigation', (e, url, inPlace, isMainFrame) => { if (isMainFrame && !inPlace) wc.listenerReady = false; });

function serve() {
  protocol.handle('app', (req) => {
    const { pathname } = new URL(req.url);
    let rel = decodeURIComponent(pathname);
    if (rel.endsWith('/')) rel += 'index.html';
    const file = path.normalize(path.join(ROOT, rel));
    if (!file.startsWith(ROOT + path.sep)) return new Response('Yasak', { status: 403 });
    return net.fetch(pathToFileURL(file).toString());
  });
}

function createWindow() {
  win = new BrowserWindow({
    width: 1360,
    height: 880,
    minWidth: 760,
    minHeight: 520,
    title: 'Platform365',
    backgroundColor: '#f4f5f9',
    icon: path.join(ROOT, 'build', 'icon.png'),
    autoHideMenuBar: process.platform !== 'darwin',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });
  resetReady(win.webContents);

  // Dış bağlantılar varsayılan tarayıcıda, uygulama içi pencereler (ör. sunucu görünümü) Electron'da
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith(ORIGIN) || url === 'about:blank') return { action: 'allow' };
    if (/^(https?|mailto):/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith(ORIGIN)) { e.preventDefault(); if (/^(https?|mailto):/.test(url)) shell.openExternal(url); }
  });
  // Mikrofon (dikte, canlı altyazı) ve tam ekran izinleri
  win.webContents.session.setPermissionRequestHandler((wc, perm, cb) => cb(['media', 'fullscreen', 'clipboard-sanitized-write', 'notifications'].includes(perm)));

  win.loadURL(`${ORIGIN}/`);
  win.on('closed', () => { win = null; });
}

const MENU_TEXT = {
  tr: {
    file: 'Dosya', home: 'Ana ekran', newDeck: 'Yeni sunum', newDoc: 'Yeni belge', close: 'Pencereyi kapat', quit: 'Çıkış',
    edit: 'Düzen', undo: 'Geri al', redo: 'Yinele', cut: 'Kes', copy: 'Kopyala', paste: 'Yapıştır', selectAll: 'Tümünü seç',
    view: 'Görünüm', reload: 'Yenile', devtools: 'Geliştirici araçları', resetZoom: 'Gerçek boyut', zoomIn: 'Yakınlaştır', zoomOut: 'Uzaklaştır', full: 'Tam ekran',
    language: 'Dil / Language', window: 'Pencere', help: 'Yardım', about: 'Platform365 hakkında', releases: 'Sürümler (GitHub)',
  },
  en: {
    file: 'File', home: 'Home', newDeck: 'New presentation', newDoc: 'New document', close: 'Close window', quit: 'Quit',
    edit: 'Edit', undo: 'Undo', redo: 'Redo', cut: 'Cut', copy: 'Copy', paste: 'Paste', selectAll: 'Select all',
    view: 'View', reload: 'Reload', devtools: 'Developer tools', resetZoom: 'Actual size', zoomIn: 'Zoom in', zoomOut: 'Zoom out', full: 'Full screen',
    language: 'Language / Dil', window: 'Window', help: 'Help', about: 'About Platform365', releases: 'Releases (GitHub)',
  },
};

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const m = MENU_TEXT[lang] || MENU_TEXT.tr;
  const go = (p) => () => win && win.loadURL(`${ORIGIN}/${p}`);
  const menuCmd = (cmd) => () => win && win.webContents.send('p365:menu', cmd);
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: m.file,
      submenu: [
        { label: m.home, accelerator: 'CmdOrCtrl+Shift+H', click: go('') },
        { label: m.newDeck, accelerator: 'CmdOrCtrl+Shift+P', click: go('present/#/new') },
        { label: m.newDoc, accelerator: 'CmdOrCtrl+Shift+N', click: go('easywrite/#/new') },
        { type: 'separator' },
        isMac ? { role: 'close', label: m.close } : { role: 'quit', label: m.quit },
      ],
    },
    { label: m.edit, submenu: [
      { role: 'undo', label: m.undo }, { role: 'redo', label: m.redo }, { type: 'separator' },
      { role: 'cut', label: m.cut }, { role: 'copy', label: m.copy }, { role: 'paste', label: m.paste },
      { role: 'selectAll', label: m.selectAll },
    ] },
    { label: m.view, submenu: [
      { role: 'reload', label: m.reload }, { role: 'toggleDevTools', label: m.devtools }, { type: 'separator' },
      { role: 'resetZoom', label: m.resetZoom }, { role: 'zoomIn', label: m.zoomIn }, { role: 'zoomOut', label: m.zoomOut },
      { type: 'separator' }, { role: 'togglefullscreen', label: m.full }, { type: 'separator' },
      { label: m.language, submenu: [
        { label: 'Türkçe', type: 'radio', checked: lang === 'tr', click: menuCmd('lang-tr') },
        { label: 'English', type: 'radio', checked: lang === 'en', click: menuCmd('lang-en') },
      ] },
    ] },
    { role: 'windowMenu', label: m.window },
    { role: 'help', label: m.help, submenu: [
      { label: m.about, click: menuCmd('about') },
      { label: m.releases, click: () => shell.openExternal('https://github.com/Agamemnon-Intelligence/Platform365/releases') },
    ] },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// Sayfa arayüz dilini bildirince menüyü o dilde yeniden kur
ipcMain.on('p365:lang', (e, l) => {
  if ((l === 'tr' || l === 'en') && l !== lang) { lang = l; buildMenu(); }
});

app.whenReady().then(() => {
  app.setAboutPanelOptions({ applicationName: 'Platform365', applicationVersion: app.getVersion(), copyright: '© 2026 Platform365 (Athene Softworks and Platform365 Technologies) · MIT' });
  serve();
  buildMenu();
  createWindow();
  filesFromArgv(process.argv).forEach(sendFile);
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
