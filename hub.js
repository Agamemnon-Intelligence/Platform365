/* Platform365 ana ekranı: uygulamalar, ortak dosya listesi, arama, yedekleme */
'use strict';

(() => {
  const P = window.P365;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = P.esc;
  const { t } = P;

  const FEATURES = [
    ['✨', 'Metinden sunum', 'Ana hattı yapıştırın; başlıklar slayta, maddeler slayt maddesine dönüşür, kalabalık slaytlar otomatik bölünür.', 'present'],
    ['🔁', 'Belge ⇄ Sunum', 'EasyWrite belgesini tek tıkla sunuma, sunumu düzenlenebilir belgeye çevirin.', 'both'],
    ['🩺', 'Sunum sağlık kontrolü', 'Düşük kontrast, küçük yazı, kalabalık slayt, taşan öğe ve yazı tipi karmaşasını puanlar.', 'present'],
    ['✎', 'Sunarken kalem, lazer, büyüteç', 'Slayt üstüne çizin, vurgulayın, lazerle gösterin, bir bölgeyi büyütün — eklenti gerekmez.', 'present'],
    ['CC', 'Canlı altyazı', 'Sunum sırasında konuşmanız ekranın altında altyazıya dönüşür.', 'present'],
    ['⏱', 'Prova modu', 'Slayt başına süreyi ölçer, uzun kalan slaytları işaretler, süreleri notlara yazar.', 'present'],
    ['🌐', 'Tek dosya web sunumu', 'Sunumu programsız açılan tek bir .html dosyası olarak paylaşın.', 'present'],
    ['/', 'Eğik çizgi komutları', '“/” yazın: başlık, tablo, bilgi kutusu, yapılacaklar, tarih, içindekiler…', 'easywrite'],
    ['#', 'Markdown ile yazma', '# başlık, - madde, **kalın**, [] görev, --- çizgi; Markdown yapıştırınca biçimlenir.', 'easywrite'],
    ['📖', 'Okunabilirlik analizi', 'Okunabilirlik puanı (Türkçe için Ateşman, İngilizce için Flesch), uzun cümle ve dolgu kelime vurgusu, sık tekrarlanan kelimeler.', 'easywrite'],
    ['🧮', 'Satır içi hesap', '“1250*12%=” yazın, sonucu kendisi yazar.', 'easywrite'],
    ['◎', 'Odak + daktilo modu', 'Yalnızca yazdığınız paragraf parlak kalır, imleç ekranın ortasında durur.', 'easywrite'],
    ['🍅', 'Pomodoro ve yazma hedefi', 'Kelime hedefi, oturum ilerlemesi ve 25/5 dakikalık odak zamanlayıcısı.', 'easywrite'],
    ['🎙', 'Dikte ve sesli okuma', 'Konuşarak yazın, belgeyi sesli dinleyin.', 'easywrite'],
    ['🔍', 'Sürüm karşılaştırma', 'Otomatik sürümler; eski hâl ile şimdiki hâl arasında eklenen/silinen kelimeleri görün.', 'both'],
    ['⌘', 'Komut paleti', 'Ctrl+K ile her komuta yazarak ulaşın — menülerde aramaya son.', 'both'],
    ['💾', 'Tek tıkla tam yedek', 'Tüm sunum, belge ve sürümleri tek dosyada yedekleyin, başka cihaza taşıyın.', 'both'],
    ['🌐', 'İki dil', 'Arayüz Türkçe veya İngilizce; ilk açılışta sorulur, istediğiniz an değiştirilir.', 'both'],
    ['📴', 'Hesapsız ve çevrimdışı', 'Kayıt yok, bulut yok: her şey cihazınızda kalır, internet olmadan çalışır.', 'both'],
  ];

  // Hızlı başlangıç: [uygulama, simge, ad, adres veya eylem]
  const QUICK = [
    ['present', '▶', 'Boş sunum', 'present/#/new'],
    ['present', '✨', 'Metinden sunum', 'outline'],
    ['easywrite', '📄', 'Boş belge', 'easywrite/#/new'],
    ['easywrite', '✉️', 'Mektup', 'easywrite/#/new/letter'],
    ['easywrite', '🧑‍💼', 'Özgeçmiş', 'easywrite/#/new/cv'],
    ['easywrite', '🗒️', 'Toplantı notu', 'easywrite/#/new/meeting'],
    ['easywrite', '📊', 'Rapor', 'easywrite/#/new/report'],
    ['easywrite', '🎓', 'Ödev / Makale', 'easywrite/#/new/essay'],
    ['easywrite', '🚀', 'Proje planı', 'easywrite/#/new/plan'],
    ['easywrite', '📔', 'Günlük', 'easywrite/#/new/journal'],
  ];

  const plainDeck = (d) => (d.slides || []).map((s) => [
    ...(s.elements || []).flatMap((e) => (e.type === 'table' ? e.rows.flatMap((r) => r.cells.map((c) => c.text || '')) : [e.text || ''])),
    s.notes || '',
  ].join(' ')).join(' ');
  const plainDoc = (d) => (d.html || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
  const norm = (s) => String(s).toLocaleLowerCase('tr');

  let items = [];
  let filter = 'all';

  async function load() {
    const [ps, es] = await Promise.all([P.appStore('present-db', 'decks'), P.appStore('easywrite-db', 'docs')]);
    const [decks, docs] = await Promise.all([ps.all(), es.all()]);
    ps.close();
    es.close();
    items = [
      ...decks.map((d) => ({ app: 'present', id: d.id, title: d.title || t('Adsız sunum'), updated: d.updated, meta: t('{n} slayt', { n: (d.slides || []).length }), text: plainDeck(d).replace(/\s+/g, ' ').trim() })),
      ...docs.map((d) => {
        const text = plainDoc(d).replace(/\s+/g, ' ').trim();
        return { app: 'easywrite', id: d.id, title: d.title || t('Adsız belge'), updated: d.updated, meta: t('{n} kelime', { n: (text.match(/[\p{L}\p{N}]+/gu) || []).length }), text };
      }),
    ].sort((a, b) => b.updated - a.updated);
    $('#stDecks').textContent = t('{n} sunum', { n: decks.length });
    $('#stDocs').textContent = t('{n} belge', { n: docs.length });
    render();
  }

  function render() {
    const q = norm($('#search').value.trim());
    const list = items.filter((it) => (filter === 'all' || it.app === filter) && (!q || norm(it.title).includes(q) || norm(it.text).includes(q)));
    const box = $('#files');
    $('#noFiles').hidden = list.length > 0;
    $('#noFiles').textContent = t(q ? 'Aramanızla eşleşen dosya yok.' : 'Henüz dosya yok. Yukarıdan bir sunum veya belge oluşturun.');
    const mark = (s) => {
      if (!q) return esc(s);
      const i = norm(s).indexOf(q);
      return i < 0 ? esc(s) : `${esc(s.slice(0, i))}<mark>${esc(s.slice(i, i + q.length))}</mark>${esc(s.slice(i + q.length))}`;
    };
    box.innerHTML = list.slice(0, 60).map((it) => {
      let snip = it.text.slice(0, 140);
      if (q) { const i = norm(it.text).indexOf(q); if (i > 40) snip = `…${it.text.slice(i - 30, i + 110)}`; }
      return `<a class="file" href="${it.app}/#/d/${encodeURIComponent(it.id)}"><span class="ico ${it.app}">${it.app === 'present' ? '▶' : 'W'}</span>
        <span class="txt"><b translate="no">${mark(it.title)}</b><small>${it.app === 'present' ? 'Present' : 'EasyWrite'} · ${it.meta} · ${P.relTime(it.updated)}</small>${snip ? `<span class="snip" translate="no">${mark(snip)}</span>` : ''}</span></a>`;
    }).join('');
  }

  /* Dosyayı doğru uygulamaya gönder */
  async function openFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    let app = null;
    if (['present', 'pptx', 'ppt'].includes(ext)) app = 'present';
    else if (['easywrite', 'docx', 'doc', 'md', 'markdown', 'txt', 'html', 'htm'].includes(ext)) app = 'easywrite';
    else if (ext === 'json') {
      try { const j = JSON.parse(await file.text()); app = j.format === 'easywrite' ? 'easywrite' : 'present'; } catch (e) { /* değil */ }
    }
    if (!app) { P.toast('Bu dosya türü desteklenmiyor.'); return; }
    try {
      await P.inbox.send(app, { kind: 'file', blob: file, name: file.name, type: file.type });
    } catch (e) {
      P.toast(t('Dosya aktarılamadı: {msg}', { msg: e.message }));
    }
  }

  /* Yedekleme */
  async function backup() {
    const [ps, es] = await Promise.all([P.appStore('present-db', 'decks'), P.appStore('easywrite-db', 'docs')]);
    const data = { format: 'platform365-backup', version: P.VERSION, created: new Date().toISOString(), decks: await ps.all(), docs: await es.all(), snapshots: [] };
    try { data.snapshots = await P.snapshots.all(); } catch (e) { /* sürüm yok */ }
    ps.close();
    es.close();
    const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
    P.download(blob, `Platform365-${P.L('yedek', 'backup')}-${new Date().toISOString().slice(0, 10)}.p365`);
    P.toast(t('{a} sunum, {b} belge ve {c} sürüm yedeklendi.', { a: data.decks.length, b: data.docs.length, c: data.snapshots.length }));
  }
  async function restore(file) {
    try {
      const data = JSON.parse(await file.text());
      if (data.format !== 'platform365-backup') throw new Error(t('Platform365 yedeği değil'));
      const ok = await P.dialog({
        title: 'Yedekten geri yükle',
        html: `<p>${t('{a} sunum, {b} belge ve {c} sürüm geri yüklenecek.', { a: (data.decks || []).length, b: (data.docs || []).length, c: (data.snapshots || []).length })}</p><p class="muted">Aynı kimliğe sahip dosyalar yedektekiyle değiştirilir; diğer dosyalarınız korunur.</p>`,
        actions: [{ label: 'Vazgeç', value: false }, { label: 'Geri yükle', primary: true, value: true }],
      });
      if (!ok) return;
      const [ps, es] = await Promise.all([P.appStore('present-db', 'decks'), P.appStore('easywrite-db', 'docs')]);
      for (const d of data.decks || []) if (d && d.id && Array.isArray(d.slides)) await ps.put(d);
      for (const d of data.docs || []) if (d && d.id && typeof d.html === 'string') await es.put({ ...d, html: window.Convert ? window.Convert.sanitize(d.html) : d.html });
      for (const s of data.snapshots || []) if (s && s.id && s.docKey) await P.snapshots.put(s);
      ps.close();
      es.close();
      P.toast('Yedek geri yüklendi.');
      load();
      storage();
    } catch (e) {
      P.toast(t('Geri yüklenemedi: {msg}', { msg: e.message }));
    }
  }
  async function storage() {
    const el = $('#storage');
    if (!navigator.storage || !navigator.storage.estimate) { el.textContent = 'Depolama bilgisi alınamadı.'; return; }
    const { usage = 0, quota = 0 } = await navigator.storage.estimate();
    const mb = (n) => (n / 1048576).toLocaleString(P.locale, { maximumFractionDigits: 1 });
    let persisted = false;
    try { persisted = navigator.storage.persisted ? await navigator.storage.persisted() : false; } catch (e) { /* yok */ }
    el.innerHTML = `${t('{used} MB kullanılıyor (kullanılabilir {quota} MB).', { used: mb(usage), quota: mb(quota) })} ${persisted ? t('🔒 Kalıcı depolama açık.') : `<a href="#" id="persist">${t('Kalıcı depolama iste')}</a>${t(' — tarayıcı yer açarken dosyalarınızı silmesin.')}`}`;
    $('#persist')?.addEventListener('click', async (e) => {
      e.preventDefault();
      const ok = navigator.storage.persist ? await navigator.storage.persist() : false;
      P.toast(ok ? 'Kalıcı depolama açıldı.' : 'Tarayıcı kalıcı depolamaya izin vermedi (uygulamayı yüklemek yardımcı olur).');
      storage();
    });
  }

  function init() {
    P.mountSwitcher('hub');
    $('#ver').textContent = `v${P.VERSION}`;
    $('#verBadge').textContent = `v${P.VERSION}`;
    $('#langLbl').textContent = P.lang.toUpperCase();
    const h = new Date().getHours();
    $('#greet').textContent = `${t(h < 6 ? 'İyi geceler' : h < 12 ? 'Günaydın' : h < 18 ? 'İyi günler' : 'İyi akşamlar')} · ${new Date().toLocaleDateString(P.locale, { weekday: 'long', day: 'numeric', month: 'long' })}`;
    $('#features').innerHTML = FEATURES.map(([i, ti, d, a]) => `<div class="feat"><div class="fi">${i}</div><b>${esc(t(ti))}</b><p>${esc(t(d))}</p><span class="tag ${a}">${a === 'both' ? 'Platform365' : a === 'present' ? 'Present' : 'EasyWrite'}</span></div>`).join('');
    $('#quick').innerHTML = QUICK.map(([a, i, n, href]) => {
      const inner = `<span class="q-ico">${i}</span><b>${esc(t(n))}</b><small>${a === 'present' ? 'Present' : 'EasyWrite'}</small>`;
      return href === 'outline' ? `<button class="q-tile ${a}" data-new-outline>${inner}</button>` : `<a class="q-tile ${a}" href="${href}">${inner}</a>`;
    }).join('');
    $('#aboutCard').innerHTML = P.aboutHtml();
    $('#search').addEventListener('input', render);
    $('#tabs').addEventListener('click', (e) => {
      const b = e.target.closest('[data-f]');
      if (!b) return;
      filter = b.dataset.f;
      $$('#tabs button').forEach((x) => x.classList.toggle('on', x === b));
      render();
    });
    const dz = $('#dropZone');
    dz.addEventListener('click', () => $('#fileOpen').click());
    dz.addEventListener('keydown', (e) => { if (e.key === 'Enter') $('#fileOpen').click(); });
    $('#fileOpen').addEventListener('change', (e) => { const f = e.target.files[0]; if (f) openFile(f); e.target.value = ''; });
    document.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files')) { e.preventDefault(); dz.classList.add('over'); } });
    document.addEventListener('dragleave', (e) => { if (!e.relatedTarget) dz.classList.remove('over'); });
    document.addEventListener('drop', (e) => { dz.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) { e.preventDefault(); openFile(f); } });
    $$('[data-new-outline]').forEach((b) => b.addEventListener('click', () => P.inbox.send('present', { kind: 'outline-dialog' })));
    $('#btnBackup').addEventListener('click', backup);
    $('#btnRestore').addEventListener('click', () => $('#fileRestore').click());
    $('#fileRestore').addEventListener('change', (e) => { const f = e.target.files[0]; if (f) restore(f); e.target.value = ''; });
    $('#btnPalette').addEventListener('click', P.openPalette);
    $('#btnTheme').addEventListener('click', P.cycleTheme);
    $('#btnLang').addEventListener('click', P.chooseLanguage);
    $('#btnAbout').addEventListener('click', P.showAbout);

    P.initApp({ onFile: openFile });

    P.registerCommands([
      { id: 'new-deck', group: 'Present', title: 'Yeni sunum', run: () => { location.href = 'present/#/new'; } },
      { id: 'new-outline', group: 'Present', title: 'Metinden sunum oluştur', run: () => P.inbox.send('present', { kind: 'outline-dialog' }) },
      { id: 'new-doc', group: 'EasyWrite', title: 'Yeni belge', run: () => { location.href = 'easywrite/#/new'; } },
      { id: 'new-meeting', group: 'EasyWrite', title: 'Yeni toplantı notu', run: () => { location.href = 'easywrite/#/new/meeting'; } },
      { id: 'new-report', group: 'EasyWrite', title: 'Yeni rapor', run: () => { location.href = 'easywrite/#/new/report'; } },
      { id: 'open', group: 'Dosya', title: 'Dosya aç', run: () => $('#fileOpen').click() },
      { id: 'search', group: 'Dosya', title: 'Dosyalarda ara', run: () => $('#search').focus() },
      { id: 'backup', group: 'Veri', title: 'Tümünü yedekle', run: backup },
      { id: 'restore', group: 'Veri', title: 'Yedekten geri yükle', run: () => $('#fileRestore').click() },
      { id: 'theme', group: 'Görünüm', title: 'Temayı değiştir', run: P.cycleTheme },
    ]);
    document.addEventListener('keydown', (e) => {
      if (e.key === '/' && document.activeElement === document.body) { e.preventDefault(); $('#search').focus(); }
    });
    load();
    storage();
    if (location.hash === '#about') P.showAbout();
  }
  init();
})();
