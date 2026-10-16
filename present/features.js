/* Platform365 Present — PowerPoint'te olmayan ek özellikler
 * Metinden slayt, EasyWrite dönüşümleri, sağlık kontrolü, prova modu, sürüm geçmişi,
 * sunumda kalem / vurgulayıcı / lazer / büyüteç / ızgara / canlı altyazı, web sayfası ve Markdown dışa aktarma. */
'use strict';

const Fx = (() => {
  const P = window.P365;
  const C = window.Convert;

  /* ---------------- Ana hat → slaytlar ---------------- */
  function slidesFromOutline(outline, t) {
    const out = [];
    for (const s of outline) {
      if (!s.bullets.length) {
        const sl = makeSlide('section', t);
        sl.elements[1].text = s.title || ' ';
        sl.elements.splice(2, 1);
        sl.notes = s.notes || '';
        out.push(sl);
        continue;
      }
      // Kalabalık slaytları böl (en fazla 7 madde)
      for (let k = 0; k < s.bullets.length; k += 7) {
        const part = s.bullets.slice(k, k + 7);
        const sl = makeSlide('content', t);
        sl.elements[0].text = (k ? `${s.title} (devam)` : s.title) || ' ';
        const body = sl.elements[2];
        const chars = part.reduce((n, b) => n + b.text.length, 0);
        body.fontSize = Math.round(clamp(26 - Math.max(0, part.length - 5) * 2 - Math.max(0, (chars - 280) / 50), 14, 26));
        body.paras = part.map((b) => ({ bullet: 'char', buChar: '•', ...(b.lvl ? { lvl: b.lvl } : {}), runs: [{ text: b.text }] }));
        simplifyText(body);
        sl.notes = k ? '' : s.notes || '';
        out.push(sl);
      }
    }
    return out;
  }

  async function createDeckFromOutline(title, outline) {
    const d = newDeck(title || 'Yeni sunum');
    const t = themeById(d.theme);
    const first = d.slides[0];
    first.elements[0].text = d.title;
    first.elements[2].text = new Date().toLocaleDateString('tr-TR', { dateStyle: 'long' });
    d.slides.push(...slidesFromOutline(outline, t));
    await Store.put(d);
    location.hash = `#/d/${d.id}`;
    P.toast(`${d.slides.length} slaytlık sunum hazırlandı.`);
  }

  function appendOutline(outline) {
    const added = slidesFromOutline(outline, themeById(deck.theme));
    if (!added.length) { P.toast('Slayta dönüştürülecek içerik bulunamadı.'); return; }
    deck.slides.splice(cur + 1, 0, ...added);
    cur += 1;
    selId = null;
    commit();
    P.toast(`${added.length} slayt eklendi.`);
  }

  async function outlineDialog(newDeckMode = !deck) {
    const sample = '# Proje planı\n- Hedefler\n- Zaman çizelgesi\n  - Ekim: tasarım\n  - Kasım: geliştirme\n> Konuşmacı notu: süreyi vurgula\n\n# Sonuç\n- Sorular?';
    const r = await P.dialog({
      title: newDeckMode ? 'Metinden sunum oluştur' : 'Metinden slayt ekle',
      wide: true,
      html: `<p class="muted">Her <b>başlık satırı</b> (veya <code>#</code> ile başlayan satır) yeni bir slayt olur. <code>-</code> ile başlayan ya da girintili satırlar madde olur, <code>&gt;</code> ile başlayanlar konuşmacı notuna gider. Yedi maddeden uzun slaytlar otomatik bölünür ve yazı boyutu içeriğe göre ayarlanır.</p>
        ${newDeckMode ? '<p><input type="text" id="olTitle" placeholder="Sunum adı" value="Yeni sunum"></p>' : ''}
        <textarea id="olText" rows="14" placeholder="${P.esc(sample)}"></textarea>
        <p><button class="btn sm" type="button" data-sample>Örnek doldur</button></p>`,
      actions: [{ label: 'Vazgeç', value: null }, { label: 'Oluştur', primary: true, value: (b) => ({ text: b.querySelector('#olText').value, title: b.querySelector('#olTitle')?.value }) }],
      onOpen: (body) => body.querySelector('[data-sample]').addEventListener('click', () => { body.querySelector('#olText').value = sample; }),
    });
    if (r) finishOutline(newDeckMode, r.text, r.title);
  }
  function finishOutline(newDeckMode, text, title) {
    const outline = C.outlineFromText(text);
    if (!outline.length) { P.toast('Metin boş görünüyor.'); return; }
    if (newDeckMode) createDeckFromOutline((title || '').trim() || outline[0].title || 'Yeni sunum', outline);
    else appendOutline(outline);
  }

  /* ---------------- EasyWrite ile alışveriş ---------------- */
  async function fromEasyWrite() {
    const st = await P.appStore('easywrite-db', 'docs');
    const docs = (await st.all()).sort((a, b) => b.updated - a.updated);
    st.close();
    if (!docs.length) { P.toast('Henüz EasyWrite belgesi yok.'); return; }
    const id = await P.dialog({
      title: 'EasyWrite belgesinden slayt',
      html: `<p class="muted">Belgedeki Başlık 1 / Başlık 2'ler slayt başlığı, altındaki paragraflar ve maddeler slayt maddesi olur.</p>
        <ul class="p365-list">${docs.map((d) => `<li><span class="grow"><b>${P.esc(d.title)}</b><small>${P.relTime(d.updated)}</small></span><button class="btn sm primary" data-id="${d.id}">Ekle</button></li>`).join('')}</ul>`,
      actions: [{ label: 'Vazgeç', value: null }],
      onOpen: (body, close) => body.addEventListener('click', (e) => { const b = e.target.closest('[data-id]'); if (b) close(b.dataset.id); }),
    });
    const d = docs.find((x) => x.id === id);
    if (d) appendOutline(C.outlineFromHtml(d.html, d.title));
  }

  /** Sunumdan ana hat: en büyük yazılı metin başlık, diğer metinler madde */
  const plain = (t) => (t.paras ? t.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n') : t.text || '');
  function deckOutline(d = deck) {
    return d.slides.map((s) => {
      const texts = s.elements.filter((e) => e.type === 'text' && plain(e).trim()).sort((a, b) => a.y - b.y || a.x - b.x);
      let titleEl = null;
      for (const e of texts) if (!titleEl || e.fontSize > titleEl.fontSize + 2) titleEl = e;
      const bullets = [];
      for (const e of texts) {
        if (e === titleEl) continue;
        if (e.paras) e.paras.forEach((p) => { const tx = p.runs.map((r) => r.text).join('').trim(); if (tx) bullets.push({ text: tx.replace(/^[•\-–*]\s*/, ''), lvl: p.lvl || 0 }); });
        else plain(e).split('\n').forEach((l) => { if (l.trim()) bullets.push({ text: l.trim().replace(/^[•\-–*]\s*/, ''), lvl: 0 }); });
      }
      for (const tb of s.elements.filter((e) => e.type === 'table')) {
        tb.rows.forEach((r) => bullets.push({ text: r.cells.filter((c) => !c.hidden).map((c) => plain(c).trim()).join(' · '), lvl: 0 }));
      }
      return { title: titleEl ? plain(titleEl).replace(/\n/g, ' ').trim() : '', bullets, notes: s.notes || '' };
    });
  }
  async function toEasyWrite() {
    await saveNow();
    const html = C.outlineToHtml(deck.title, deckOutline());
    P.inbox.send('easywrite', { kind: 'html', title: deck.title, html });
  }

  /* ---------------- Sağlık kontrolü ---------------- */
  const lum = (hex) => {
    const h = toHex(hex);
    if (!h) return null;
    const c = [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255).map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
  };
  const contrast = (a, b) => {
    const la = lum(a);
    const lb = lum(b);
    if (la === null || lb === null) return null;
    return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
  };
  function healthCheck() {
    const issues = [];
    const fonts = new Set();
    deck.slides.forEach((s, i) => {
      let words = 0;
      let hasText = false;
      for (const e of s.elements) {
        const name = { text: 'Metin', image: 'Görsel', table: 'Tablo', shape: 'Şekil', rect: 'Dikdörtgen', ellipse: 'Elips' }[e.type] || 'Öğe';
        if (e.x + e.w < 4 || e.y + e.h < 4 || e.x > W - 4 || e.y > H - 4) issues.push({ sev: 'err', i, id: e.id, msg: `${name} slaytın tamamen dışında` });
        else if (e.x < -2 || e.y < -2 || e.x + e.w > W + 2 || e.y + e.h > H + 2) issues.push({ sev: 'warn', i, id: e.id, msg: `${name} slayt kenarından taşıyor` });
        for (const t of textTargets(e)) {
          const txt = plain(t);
          if (t.font) fonts.add(t.font);
          (t.paras || []).forEach((p) => p.runs.forEach((r) => r.font && fonts.add(r.font)));
          if (!txt.trim()) { if (e.type === 'text') issues.push({ sev: 'info', i, id: e.id, msg: 'Boş metin kutusu' }); continue; }
          hasText = true;
          words += txt.split(/\s+/).filter(Boolean).length;
          const size = Math.min(t.fontSize, ...(t.paras || []).flatMap((p) => p.runs.map((r) => r.size || p.size || t.fontSize)));
          if (size < 14) issues.push({ sev: 'warn', i, id: e.id, msg: `Çok küçük yazı (${Math.round(size)} px) — arkadan okunmaz` });
          const bg = t.fill || (e.type === 'text' && e.fill) || (!s.bgImg && !(s.bgGrad && s.bgGrad.stops) ? s.bg : null);
          const ratio = bg ? contrast(t.color, bg) : null;
          if (ratio !== null) {
            const need = t.fontSize >= 24 || t.bold ? 3 : 4.5;
            if (ratio < need) issues.push({ sev: ratio < 2 ? 'err' : 'warn', i, id: e.id, msg: `Düşük kontrast ${ratio.toFixed(1)}:1 (en az ${need}:1 önerilir) — “${txt.trim().slice(0, 30)}”` });
          }
        }
      }
      if (words > 80) issues.push({ sev: 'warn', i, msg: `Kalabalık slayt: ${words} kelime (40'ın altı önerilir)` });
      if (!hasText && !s.elements.some((e) => e.type === 'image')) issues.push({ sev: 'info', i, msg: 'Slaytta metin veya görsel yok' });
    });
    if (fonts.size > 4) issues.push({ sev: 'warn', i: 0, msg: `Sunumda ${fonts.size} farklı yazı tipi var (${[...fonts].slice(0, 6).join(', ')}) — tutarlılık için 2–3 önerilir` });
    if (deck.slides.length > 1 && deck.slides.every((s) => !s.notes)) issues.push({ sev: 'info', i: 0, msg: 'Hiçbir slaytta konuşmacı notu yok' });
    const score = Math.max(0, 100 - issues.reduce((n, x) => n + ({ err: 12, warn: 5, info: 1 }[x.sev]), 0));
    const icon = { err: '<span class="sev-err">●</span>', warn: '<span class="sev-warn">●</span>', info: '<span class="muted">○</span>' };
    P.dialog({
      title: 'Sunum sağlık kontrolü',
      wide: true,
      html: `<p><b style="font-size:28px" class="${score >= 85 ? 'sev-ok' : score >= 60 ? 'sev-warn' : 'sev-err'}">${score}</b><span class="muted"> / 100 puan · ${deck.slides.length} slayt · ${issues.length} bulgu</span></p>
        ${issues.length ? `<ul class="p365-list">${issues.map((x, k) => `<li>${icon[x.sev]}<span class="grow">${P.esc(x.msg)}<small>Slayt ${x.i + 1}</small></span><button class="btn sm" data-k="${k}">Git</button></li>`).join('')}</ul>` : '<p class="sev-ok">Harika! Sorun bulunamadı. 🎉</p>'}`,
      onOpen: (body, close) => body.addEventListener('click', (e) => {
        const b = e.target.closest('[data-k]');
        if (!b) return;
        const x = issues[+b.dataset.k];
        close();
        selectSlide(x.i);
        if (x.id) { selId = x.id; renderStage(); renderProps(); }
      }),
    });
  }

  /* ---------------- Emoji ---------------- */
  async function insertEmoji() {
    const e = await P.pickEmoji();
    if (!e || !deck) return;
    const t = themeById(deck.theme);
    const el = textEl(t, { x: Math.round((W - 180) / 2), y: Math.round((H - 160) / 2), w: 180, h: 160, text: e, fontSize: 110, align: 'center', valign: 'middle', font: 'Segoe UI Emoji' });
    curSlide().elements.push(el);
    selId = el.id;
    commit();
  }

  /* ---------------- Sürüm geçmişi ---------------- */
  async function onOpen(d) {
    try {
      const list = await P.snapshots.list('present', d.id);
      if (!list.length || Date.now() - list[0].at > 30 * 60 * 1000) await P.snapshots.save('present', d.id, d.title, d);
    } catch (e) { /* depolama yoksa sürüm yok */ }
  }
  async function saveSnapshot() {
    const label = await P.prompt('Sürüm kaydet', `Sürüm ${new Date().toLocaleString('tr-TR', { dateStyle: 'short', timeStyle: 'short' })}`, 'Adlandırılmış sürümler otomatik silinmez.');
    if (label === null) return;
    await saveNow();
    await P.snapshots.save('present', deck.id, deck.title, deck, label.trim() || 'Adsız sürüm');
    P.toast('Sürüm kaydedildi.');
  }
  async function versions() {
    await saveNow();
    const list = await P.snapshots.list('present', deck.id);
    const draw = () => (list.length ? `<ul class="p365-list">${list.map((s, k) => {
      const n = (() => { try { return JSON.parse(s.data).slides.length; } catch (e) { return '?'; } })();
      return `<li><span class="grow"><b>${P.esc(s.label || 'Otomatik kayıt')}</b><small>${new Date(s.at).toLocaleString('tr-TR', { dateStyle: 'medium', timeStyle: 'short' })} · ${n} slayt</small></span>
        <button class="btn sm" data-prev="${k}">Önizle</button><button class="btn sm primary" data-restore="${k}">Geri yükle</button><button class="btn sm danger" data-del="${k}">🗑</button></li>`;
    }).join('')}</ul>` : '<p class="muted">Henüz kayıtlı sürüm yok.</p>');
    await P.dialog({
      title: 'Sürüm geçmişi',
      wide: true,
      html: `<p class="muted">Sunumu her açtığınızda (30 dakikada bir) otomatik sürüm alınır. Bir sürümü geri yüklemek, şimdiki hâli de yeni bir sürüm olarak saklar — hiçbir şey kaybolmaz.</p><div id="vlist">${draw()}</div><div id="vprev" class="ver-prev"></div>`,
      actions: [{ label: 'Şimdi sürüm kaydet', value: 'save' }, { label: 'Kapat', primary: true, value: null }],
      onOpen: (body, close) => body.addEventListener('click', async (e) => {
        const b = e.target.closest('button');
        if (!b) return;
        if (b.dataset.del) {
          await P.snapshots.del(list[+b.dataset.del].id);
          list.splice(+b.dataset.del, 1);
          body.querySelector('#vlist').innerHTML = draw();
        } else if (b.dataset.prev) {
          const data = JSON.parse(list[+b.dataset.prev].data);
          const box = body.querySelector('#vprev');
          box.innerHTML = '';
          data.slides.slice(0, 12).forEach((s) => { const th = document.createElement('div'); th.className = 'thumb'; box.appendChild(th); mountThumb(th, s, data.h || 540); });
        } else if (b.dataset.restore) {
          const snap = list[+b.dataset.restore];
          await P.snapshots.save('present', deck.id, deck.title, deck, '');
          const data = JSON.parse(snap.data);
          deck = { ...data, id: deck.id, updated: Date.now() };
          H = deck.h || 540;
          cur = 0;
          selId = null;
          $('#deckTitle').value = deck.title;
          commit();
          fitStage();
          close();
          P.toast('Sürüm geri yüklendi. Geri almak için Ctrl+Z.');
        }
      }),
    }).then((r) => { if (r === 'save') saveSnapshot(); });
  }

  /* ---------------- Dışa aktarma: Web sayfası ve Markdown ---------------- */
  function presentCss() {
    const sheet = [...document.styleSheets].find((s) => s.href && /present\/styles\.css/.test(s.href));
    try { return sheet ? [...sheet.cssRules].map((r) => r.cssText).join('\n') : ''; } catch (e) { return ''; }
  }
  function exportHtml() {
    const slides = deck.slides.map((s) => { const n = slideNode(s); return n.outerHTML; }).join('\n');
    const notes = JSON.stringify(deck.slides.map((s) => s.notes || ''));
    const html = `<!doctype html><html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${P.esc(deck.title)}</title><meta name="generator" content="Platform365 Present ${APP_VERSION}">
<style>${presentCss()}
html,body{margin:0;height:100%;background:#000;overflow:hidden;font-family:system-ui,sans-serif}
#deck .slide{position:absolute;left:0;top:0;transform-origin:0 0;display:none}#deck .slide.on{display:block;animation:fadeIn .3s ease}
#bar{position:fixed;left:50%;bottom:14px;transform:translateX(-50%);color:#fff;background:rgba(20,20,28,.75);padding:4px 12px;border-radius:99px;font-size:13px;opacity:.0;transition:opacity .2s}
body:hover #bar{opacity:1}#bar button{background:none;border:0;color:#fff;font-size:18px;cursor:pointer}
#prog{position:fixed;left:0;bottom:0;height:3px;background:#818cf8;transition:width .3s}</style></head>
<body><div id="deck">${slides}</div><div id="prog"></div><div id="bar"><button id="b1">‹</button> <span id="cnt"></span> <button id="b2">›</button></div>
<script>(function(){var W=${W},H=${H},S=[].slice.call(document.querySelectorAll('#deck .slide')),i=Math.max(0,Math.min(S.length-1,(parseInt(location.hash.slice(1))||1)-1)),notes=${notes};
function fit(){var k=Math.min(innerWidth/W,innerHeight/H);S.forEach(function(s){s.style.transform='translate('+((innerWidth-W*k)/2)+'px,'+((innerHeight-H*k)/2)+'px) scale('+k+')'})}
function show(){S.forEach(function(s,k){s.classList.toggle('on',k===i)});document.getElementById('cnt').textContent=(i+1)+' / '+S.length;document.getElementById('prog').style.width=((i+1)/S.length*100)+'%';history.replaceState(null,'','#'+(i+1));if(notes[i])console.log('Not '+(i+1)+': '+notes[i])}
function go(d){i=Math.max(0,Math.min(S.length-1,i+d));show()}
addEventListener('keydown',function(e){if(['ArrowRight','ArrowDown','PageDown',' ','Enter'].indexOf(e.key)>=0){e.preventDefault();go(1)}else if(['ArrowLeft','ArrowUp','PageUp','Backspace'].indexOf(e.key)>=0){e.preventDefault();go(-1)}else if(e.key==='f'){document.documentElement.requestFullscreen&&document.documentElement.requestFullscreen()}});
document.getElementById('b1').onclick=function(e){e.stopPropagation();go(-1)};document.getElementById('b2').onclick=function(e){e.stopPropagation();go(1)};
document.getElementById('deck').onclick=function(){go(1)};var tx=null;addEventListener('touchstart',function(e){tx=e.touches[0].clientX},{passive:true});addEventListener('touchend',function(e){if(tx===null)return;var dx=e.changedTouches[0].clientX-tx;if(Math.abs(dx)>50)go(dx<0?1:-1);tx=null});
addEventListener('resize',fit);fit();show()})();<\/script></body></html>`;
    P.download(new Blob([html], { type: 'text/html' }), `${P.safeName(deck.title, 'sunum')}.html`);
    P.toast('Web sayfası indirildi. Tek dosyadır; her tarayıcıda açılır (F: tam ekran).');
  }
  function exportMarkdown() {
    const ol = deckOutline();
    let md = `# ${deck.title}\n\n`;
    ol.forEach((s, i) => {
      md += `## ${s.title || `Slayt ${i + 1}`}\n\n`;
      s.bullets.forEach((b) => { md += `${'  '.repeat(b.lvl)}- ${b.text}\n`; });
      if (s.bullets.length) md += '\n';
      if (s.notes) md += `${s.notes.split('\n').map((l) => `> ${l}`).join('\n')}\n\n`;
    });
    P.download(new Blob([md], { type: 'text/markdown' }), `${P.safeName(deck.title, 'sunum')}.md`);
    P.toast('Ana hat Markdown olarak indirildi.');
  }

  /* ---------------- Sunum modu araçları ---------------- */
  const ink = {};          // slayt dizini → çizgiler (slayt koordinatında)
  let stroke = null;
  let timerStart = 0;
  let timerInt = null;
  let slideStart = 0;
  let lastIdx = -1;
  let timings = [];
  let rehearsing = false;
  let caption = null;
  let digits = '';
  let zoomed = false;

  const pres = () => $('#presenter');
  const geom = () => {
    const k = Math.min(window.innerWidth / W, window.innerHeight / H);
    return { k, ox: (window.innerWidth - W * k) / 2, oy: (window.innerHeight - H * k) / 2 };
  };
  function sizeCanvas() {
    const c = $('#inkCanvas');
    const dpr = window.devicePixelRatio || 1;
    c.width = Math.round(window.innerWidth * dpr);
    c.height = Math.round(window.innerHeight * dpr);
    drawInk();
  }
  function drawInk() {
    const c = $('#inkCanvas');
    if (!c) return;
    const ctx = c.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, c.width, c.height);
    const { k, ox, oy } = geom();
    ctx.setTransform(dpr * k, 0, 0, dpr * k, dpr * ox, dpr * oy);
    for (const s of ink[pIndex] || []) {
      ctx.strokeStyle = s.color;
      ctx.lineWidth = s.width;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.globalCompositeOperation = s.tool === 'hl' ? 'multiply' : 'source-over';
      ctx.beginPath();
      s.points.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)));
      if (s.points.length === 1) ctx.lineTo(s.points[0][0] + 0.1, s.points[0][1]);
      ctx.stroke();
    }
    ctx.globalCompositeOperation = 'source-over';
  }
  const PEN_COLORS = ['#ef4444', '#2563eb', '#16a34a', '#111827', '#f59e0b'];
  let penColor = 0;

  function setTool(t) {
    if (zoomed) unzoom();
    Fx.tool = Fx.tool === t ? null : t;
    const p = pres();
    p.classList.toggle('tool-draw', Fx.tool === 'pen' || Fx.tool === 'hl');
    p.classList.toggle('tool-laser', Fx.tool === 'laser');
    p.classList.toggle('tool-zoom', Fx.tool === 'zoom');
    $$('[data-ptool]', p).forEach((b) => b.classList.toggle('on', b.dataset.ptool === Fx.tool));
    $('#laserDot').hidden = Fx.tool !== 'laser';
    if (Fx.tool === 'pen') P.toast(`Kalem: çizmek için sürükleyin. Renk değiştirmek için D'ye tekrar basın. Silmek: E`);
  }
  function unzoom() {
    zoomed = false;
    const s = $('#presenterScaler');
    s.style.transition = 'transform .25s ease';
    s.style.transform = '';
    $('#inkCanvas').style.visibility = '';
  }
  function curtain(color) {
    const c = $('#pCurtain');
    if (!c.hidden && c.dataset.color === color) { c.hidden = true; return; }
    c.dataset.color = color;
    c.style.background = color;
    c.hidden = false;
  }
  function toggleGrid() {
    const g = $('#pOverview');
    if (!g.hidden) { g.hidden = true; return; }
    g.innerHTML = '';
    deck.slides.forEach((s, i) => {
      const item = document.createElement('button');
      item.className = `po-item${i === pIndex ? ' on' : ''}`;
      item.innerHTML = `<div class="thumb"></div><span>${i + 1}</span>`;
      item.addEventListener('click', (e) => { e.stopPropagation(); g.hidden = true; pIndex = i; showPresented(); });
      g.appendChild(item);
      mountThumb($('.thumb', item), s);
    });
    g.hidden = false;
  }
  function toggleCaptions() {
    const box = $('#pCaption');
    if (caption) { caption.stop(); caption = null; box.hidden = true; P.toast('Canlı altyazı kapatıldı.'); return; }
    let finalText = '';
    caption = P.speech({
      onText: (fin, tmp) => {
        finalText = `${finalText} ${fin}`.trim().slice(-220);
        const show = `${finalText} ${tmp}`.trim();
        box.textContent = show.length > 160 ? `…${show.slice(-160)}` : show;
        box.hidden = !show;
      },
      onEnd: () => { box.hidden = true; },
    });
    if (!caption) P.toast('Bu tarayıcı konuşma tanımayı desteklemiyor (Chrome veya Edge deneyin).');
    else P.toast('Canlı altyazı açık — konuşmanız ekranda yazıya dökülür.');
  }
  const fmt = (ms) => { const s = Math.floor(ms / 1000); return `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; };
  function tick() { $('#pTimer').textContent = fmt(Date.now() - timerStart); }

  function onPresentStart() {
    timerStart = Date.now();
    slideStart = Date.now();
    lastIdx = -1;
    timings = deck.slides.map(() => 0);
    Object.keys(ink).forEach((k) => delete ink[k]);
    clearInterval(timerInt);
    timerInt = setInterval(tick, 1000);
    tick();
    sizeCanvas();
    pres().classList.toggle('rehearsing', rehearsing);
  }
  function onSlideShown(i) {
    if (lastIdx >= 0 && lastIdx < timings.length) timings[lastIdx] += Date.now() - slideStart;
    slideStart = Date.now();
    lastIdx = i;
    if (zoomed) unzoom();
    $('#pProgress i').style.width = `${((i + 1) / deck.slides.length) * 100}%`;
    drawInk();
  }
  function onPresentEnd() {
    if (lastIdx >= 0 && lastIdx < timings.length) timings[lastIdx] += Date.now() - slideStart;
    clearInterval(timerInt);
    if (caption) { caption.stop(); caption = null; }
    Fx.tool = null;
    setTool(null);
    Fx.tool = null;
    $$('[data-ptool]').forEach((b) => b.classList.remove('on'));
    pres().classList.remove('tool-draw', 'tool-laser', 'tool-zoom', 'rehearsing');
    $('#pCurtain').hidden = true;
    $('#pOverview').hidden = true;
    $('#pCaption').hidden = true;
    $('#laserDot').hidden = true;
    if (zoomed) unzoom();
    if (rehearsing) { rehearsing = false; setTimeout(rehearsalReport, 300); }
  }
  function rehearsalReport() {
    const total = timings.reduce((a, b) => a + b, 0);
    const max = Math.max(...timings, 1);
    const ol = deckOutline();
    P.dialog({
      title: `Prova sonucu — toplam ${fmt(total)}`,
      wide: true,
      html: `<p class="muted">Ortalama slayt süresi ${fmt(total / Math.max(1, timings.filter(Boolean).length))}. Uzun kalan slaytlar kırmızıyla işaretlidir.</p>
        <ul class="p365-list">${timings.map((t, i) => `<li><span style="width:28px" class="muted">${i + 1}</span><span class="grow">${P.esc(ol[i]?.title || `Slayt ${i + 1}`)}
        <span class="rh-bar"><i style="width:${(t / max) * 100}%;background:${t > (total / timings.length) * 1.8 ? 'var(--danger)' : 'var(--accent)'}"></i></span></span><b>${fmt(t)}</b></li>`).join('')}</ul>`,
      actions: [{ label: 'Sürelerini notlara yaz', value: 'notes' }, { label: 'Kapat', primary: true, value: null }],
    }).then((r) => {
      if (r !== 'notes') return;
      deck.slides.forEach((s, i) => { s.notes = `${(s.notes || '').replace(/\n?⏱ Prova: \d\d:\d\d$/, '')}\n⏱ Prova: ${fmt(timings[i])}`.trim(); });
      commit();
      P.toast('Süreler konuşmacı notlarına eklendi.');
    });
  }
  function rehearse() { rehearsing = true; startPresentation(true); P.toast('Prova modu: her slaytta geçen süre ölçülüyor.'); }

  function onPresentKey(e) {
    if (e.ctrlKey || e.metaKey || e.altKey) return false;
    const k = e.key.toLowerCase();
    if (/^[0-9]$/.test(e.key)) { digits += e.key; return true; }
    if (e.key === 'Enter' && digits) {
      pIndex = clamp(+digits - 1, 0, deck.slides.length - 1);
      digits = '';
      showPresented();
      e.preventDefault();
      return true;
    }
    digits = '';
    if (e.key === 'Escape') {
      if (!$('#pOverview').hidden) { $('#pOverview').hidden = true; return true; }
      if (!$('#pCurtain').hidden) { $('#pCurtain').hidden = true; return true; }
      if (Fx.tool) { setTool(Fx.tool); return true; }
      return false;
    }
    const map = {
      d: () => { if (Fx.tool === 'pen') { penColor = (penColor + 1) % PEN_COLORS.length; P.toast('Kalem rengi değişti'); } else setTool('pen'); },
      h: () => setTool('hl'), l: () => setTool('laser'), z: () => setTool('zoom'),
      e: () => { delete ink[pIndex]; drawInk(); }, b: () => curtain('#000'), w: () => curtain('#fff'), '.': () => curtain('#000'),
      g: toggleGrid, o: toggleGrid, c: toggleCaptions,
      t: () => { timerStart = Date.now(); tick(); P.toast('Zamanlayıcı sıfırlandı'); },
      '?': () => keysHelp(),
    };
    if (map[k]) { e.preventDefault(); map[k](); return true; }
    return false;
  }

  function bindPresenter() {
    const p = pres();
    p.addEventListener('pointerdown', (e) => {
      if (e.target.closest('.presenter-bar, .p-overview')) return;
      if (Fx.tool === 'pen' || Fx.tool === 'hl') {
        const { k, ox, oy } = geom();
        const pt = [(e.clientX - ox) / k, (e.clientY - oy) / k];
        stroke = Fx.tool === 'hl'
          ? { tool: 'hl', color: 'rgba(250, 204, 21, .45)', width: 22, points: [pt] }
          : { tool: 'pen', color: PEN_COLORS[penColor], width: 4, points: [pt] };
        (ink[pIndex] = ink[pIndex] || []).push(stroke);
        p.setPointerCapture?.(e.pointerId);
        drawInk();
      } else if (Fx.tool === 'zoom') {
        const s = $('#presenterScaler');
        if (zoomed) { unzoom(); return; }
        zoomed = true;
        const r = s.getBoundingClientRect();
        s.style.transition = 'transform .25s ease';
        s.style.transformOrigin = `${e.clientX - r.left}px ${e.clientY - r.top}px`;
        s.style.transform = 'scale(2.2)';
        $('#inkCanvas').style.visibility = 'hidden';
      }
    });
    p.addEventListener('pointermove', (e) => {
      if (Fx.tool === 'laser') {
        const d = $('#laserDot');
        d.style.left = `${e.clientX}px`;
        d.style.top = `${e.clientY}px`;
      }
      if (stroke) {
        const { k, ox, oy } = geom();
        stroke.points.push([(e.clientX - ox) / k, (e.clientY - oy) / k]);
        drawInk();
      }
    });
    const end = () => { stroke = null; };
    p.addEventListener('pointerup', end);
    p.addEventListener('pointercancel', end);
    $$('[data-ptool]', p).forEach((b) => b.addEventListener('click', (e) => { e.stopPropagation(); setTool(b.dataset.ptool); }));
    $$('[data-pact]', p).forEach((b) => b.addEventListener('click', (e) => {
      e.stopPropagation();
      ({ clear: () => { delete ink[pIndex]; drawInk(); }, black: () => curtain('#000'), grid: toggleGrid, cc: toggleCaptions })[b.dataset.pact]();
    }));
    $('#pCurtain').addEventListener('click', (e) => { e.stopPropagation(); $('#pCurtain').hidden = true; });
    window.addEventListener('resize', () => { if (!p.hidden) sizeCanvas(); });
  }

  /* ---------------- Kısayollar ---------------- */
  function keysHelp() {
    const rows = [
      ['Ctrl+K', 'Komut paleti (tüm Platform365 komutları)'], ['F5 / Shift+F5', 'Baştan / geçerli slayttan sun'],
      ['Çift tıklama, Enter', 'Metni düzenle'], ['Ctrl+Z / Ctrl+Y', 'Geri al / yinele'], ['Ctrl+D', 'Çoğalt'], ['Ok tuşları', 'Öğeyi kaydır (Shift: 10 px)'],
      ['— Sunum modu —', ''], ['D', 'Kalem (tekrar: renk değiştir)'], ['H', 'Vurgulayıcı'], ['L', 'Lazer işaretçi'], ['Z', 'Büyüteç (tıklanan yeri büyütür)'],
      ['E', 'Bu slayttaki çizimleri sil'], ['B / W', 'Siyah / beyaz ekran'], ['G', 'Slayt ızgarası'], ['C', 'Canlı altyazı (konuşma → yazı)'],
      ['T', 'Zamanlayıcıyı sıfırla'], ['Sayı + Enter', 'O slayta git'], ['Esc', 'Aracı kapat / sunumdan çık'],
    ];
    P.dialog({ title: 'Klavye kısayolları', html: `<div class="p365-keys">${rows.map(([a, b]) => (b ? `<kbd>${P.esc(a)}</kbd><span>${P.esc(b)}</span>` : `<b style="grid-column:1/3;margin-top:8px">${P.esc(a)}</b>`)).join('')}</div>` });
  }

  /* ---------------- Gelen kutusu ---------------- */
  async function handleInbox(id) {
    let item = null;
    try { item = await P.inbox.take(id); } catch (e) { /* yok */ }
    if (!item) { location.replace('#'); return; }
    if (item.kind === 'file') {
      window.history.replaceState(null, '', location.pathname);
      showHome();
      importFile(new File([item.blob], item.name, { type: item.type || '' }));
    } else if (item.kind === 'outline') {
      window.history.replaceState(null, '', location.pathname);
      createDeckFromOutline(item.title, item.slides);
    } else if (item.kind === 'outline-dialog') {
      window.history.replaceState(null, '', location.pathname);
      showHome();
      outlineDialog(true);
    } else if (item.kind === 'new') {
      location.replace('#/new');
    } else location.replace('#');
  }

  /* ---------------- Başlat ---------------- */
  function init() {
    P.mountSwitcher('present');
    const menu = $('#toolsMenu');
    $('#btnTools').addEventListener('click', (e) => { e.stopPropagation(); menu.hidden = !menu.hidden; });
    document.addEventListener('click', (e) => { if (!e.target.closest('#btnTools')) menu.hidden = true; });
    const tools = {
      outline: () => outlineDialog(false), 'from-doc': fromEasyWrite, 'to-doc': toEasyWrite, check: healthCheck,
      emoji: insertEmoji, rehearse, snapshot: saveSnapshot, keys: keysHelp,
    };
    menu.addEventListener('click', (e) => { const b = e.target.closest('[data-tool]'); if (b) { menu.hidden = true; tools[b.dataset.tool](); } });
    $('#btnVersions').addEventListener('click', versions);
    $('#btnOutlineNew').addEventListener('click', () => outlineDialog(true));
    bindPresenter();

    const inEd = () => !!deck;
    P.registerCommands([
      { id: 'new', group: 'Present', title: 'Yeni sunum', run: () => { location.hash = '#/new'; } },
      { id: 'new-outline', group: 'Present', title: 'Metinden yeni sunum oluştur', keywords: 'outline ana hat otomatik', run: () => outlineDialog(true) },
      { id: 'open', group: 'Dosya', title: 'Dosya aç (.present, .pptx)', run: () => $('#fileImport').click() },
      { id: 'home', group: 'Present', title: 'Sunumlarım', run: () => { location.hash = ''; } },
      { id: 'play', group: 'Sunum', title: 'Baştan sun', key: 'F5', when: inEd, run: () => startPresentation(true) },
      { id: 'play-cur', group: 'Sunum', title: 'Geçerli slayttan sun', key: 'Shift+F5', when: inEd, run: () => startPresentation(false) },
      { id: 'rehearse', group: 'Sunum', title: 'Prova modu (süre ölç)', when: inEd, run: rehearse },
      { id: 'add-slide', group: 'Slayt', title: 'Slayt ekle (başlık ve içerik)', when: inEd, run: () => addSlide('content') },
      { id: 'add-text', group: 'Ekle', title: 'Metin kutusu', when: inEd, run: () => addElement('text') },
      { id: 'add-img', group: 'Ekle', title: 'Görsel', when: inEd, run: () => addElement('image') },
      { id: 'add-table', group: 'Ekle', title: 'Tablo', when: inEd, run: () => addElement('table') },
      { id: 'add-emoji', group: 'Ekle', title: 'Emoji / ikon', when: inEd, run: insertEmoji },
      { id: 'outline', group: 'Araçlar', title: 'Metinden slayt ekle', when: inEd, run: () => outlineDialog(false) },
      { id: 'from-doc', group: 'Araçlar', title: 'EasyWrite belgesinden slayt', when: inEd, run: fromEasyWrite },
      { id: 'to-doc', group: 'Araçlar', title: 'EasyWrite belgesine dönüştür', when: inEd, run: toEasyWrite },
      { id: 'check', group: 'Araçlar', title: 'Sunum sağlık kontrolü', keywords: 'erişilebilirlik kontrast', when: inEd, run: healthCheck },
      { id: 'versions', group: 'Dosya', title: 'Sürüm geçmişi', when: inEd, run: versions },
      { id: 'snapshot', group: 'Dosya', title: 'Sürüm kaydet', when: inEd, run: saveSnapshot },
      { id: 'pdf', group: 'Dosya', title: 'Yazdır / PDF', when: inEd, run: async () => { await saveNow(); exportPdf(); } },
      { id: 'dl-present', group: 'İndir', title: 'Present dosyası (.present)', when: inEd, run: async () => { await saveNow(); exportFile(deck, 'present'); } },
      { id: 'dl-pptx', group: 'İndir', title: 'PowerPoint (.pptx)', when: inEd, run: async () => { await saveNow(); exportFile(deck, 'pptx'); } },
      { id: 'dl-html', group: 'İndir', title: 'Web sayfası (.html)', when: inEd, run: exportHtml },
      { id: 'dl-md', group: 'İndir', title: 'Ana hat (.md)', when: inEd, run: exportMarkdown },
      ...THEMES.map((t) => ({ id: `theme-${t.id}`, group: 'Tema', title: `Slayt teması: ${t.name}`, when: inEd, run: () => applyTheme(t.id) })),
      { id: 'dark', group: 'Görünüm', title: 'Uygulama temasını değiştir (aydınlık / karanlık)', run: P.cycleTheme },
      { id: 'keys', group: 'Yardım', title: 'Klavye kısayolları', run: keysHelp },
      { id: 'go-hub', group: 'Platform365', title: 'Ana ekran', run: () => { location.href = P.ROOT; } },
      { id: 'go-ew', group: 'Platform365', title: 'EasyWrite\'ı aç', run: () => { location.href = `${P.ROOT}easywrite/`; } },
    ]);
  }

  return {
    tool: null, init, onOpen, handleInbox, insertEmoji, exportHtml, exportMarkdown,
    onPresentStart, onSlideShown, onPresentEnd, onPresentKey,
  };
})();
