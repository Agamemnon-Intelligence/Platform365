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

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const go = (p) => () => win && win.loadURL(`${ORIGIN}/${p}`);
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'Dosya',
      submenu: [
        { label: 'Ana ekran', accelerator: 'CmdOrCtrl+Shift+H', click: go('') },
        { label: 'Yeni sunum', accelerator: 'CmdOrCtrl+Shift+P', click: go('present/#/new') },
        { label: 'Yeni belge', accelerator: 'CmdOrCtrl+Shift+N', click: go('easywrite/#/new') },
        { type: 'separator' },
        isMac ? { role: 'close', label: 'Pencereyi kapat' } : { role: 'quit', label: 'Çıkış' },
      ],
    },
    { label: 'Düzen', submenu: [
      { role: 'undo', label: 'Geri al' }, { role: 'redo', label: 'Yinele' }, { type: 'separator' },
      { role: 'cut', label: 'Kes' }, { role: 'copy', label: 'Kopyala' }, { role: 'paste', label: 'Yapıştır' },
      { role: 'selectAll', label: 'Tümünü seç' },
    ] },
    { label: 'Görünüm', submenu: [
      { role: 'reload', label: 'Yenile' }, { role: 'toggleDevTools', label: 'Geliştirici araçları' }, { type: 'separator' },
      { role: 'resetZoom', label: 'Gerçek boyut' }, { role: 'zoomIn', label: 'Yakınlaştır' }, { role: 'zoomOut', label: 'Uzaklaştır' },
      { type: 'separator' }, { role: 'togglefullscreen', label: 'Tam ekran' },
    ] },
    { role: 'windowMenu', label: 'Pencere' },
    { role: 'help', label: 'Yardım', submenu: [
      { label: 'Sürümler (GitHub)', click: () => shell.openExternal('https://github.com/Agamemnon-Intelligence/Platform365/releases') },
    ] },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

app.whenReady().then(() => {
  serve();
  buildMenu();
  createWindow();
  filesFromArgv(process.argv).forEach(sendFile);
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
