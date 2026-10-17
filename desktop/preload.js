/* Platform365 masaüstü — sayfaya güvenli köprü (window.p365Desktop) */
'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('p365Desktop', {
  platform: process.platform === 'darwin' ? 'mac' : process.platform === 'win32' ? 'win' : 'linux',
  /** İşletim sisteminden açılan dosyaları alır: cb({ name, data: Uint8Array }) */
  onOpenFile(cb) {
    ipcRenderer.on('p365:open-file', (e, f) => cb(f));
    ipcRenderer.invoke('p365:ready').then((files) => files.forEach(cb));
  },
});
