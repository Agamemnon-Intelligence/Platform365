/* Platform365 — ortak dönüştürücüler
 * Markdown ⇄ HTML, güvenli HTML temizleme, belge/metin → sunum ana hattı. */
'use strict';

(function () {
  const tx = (s, v) => (window.P365 ? window.P365.t(s, v) : s);
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* ---------------- Güvenli HTML ---------------- */
  const ALLOWED = new Set(['P', 'DIV', 'BR', 'HR', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL', 'SUB', 'SUP',
    'SPAN', 'FONT', 'MARK', 'A', 'UL', 'OL', 'LI', 'BLOCKQUOTE', 'PRE', 'CODE', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'TD', 'TH',
    'COLGROUP', 'COL', 'IMG', 'DETAILS', 'SUMMARY', 'INPUT', 'LABEL', 'FIGURE', 'FIGCAPTION', 'ASIDE', 'NAV', 'SECTION', 'SMALL', 'KBD']);
  const DROP = new Set(['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT', 'EMBED', 'LINK', 'META', 'TITLE', 'HEAD', 'NOSCRIPT', 'TEMPLATE', 'SVG', 'MATH', 'FORM', 'BUTTON', 'SELECT', 'TEXTAREA', 'VIDEO', 'AUDIO', 'CANVAS']);
  const ATTRS = new Set(['href', 'src', 'alt', 'title', 'colspan', 'rowspan', 'style', 'class', 'type', 'checked', 'open', 'width', 'height', 'color', 'face', 'data-callout', 'data-toc', 'data-check', 'data-page-break', 'contenteditable', 'id', 'start']);
  const STYLE_OK = /^(color|background-color|background|font-size|font-family|font-weight|font-style|text-decoration|text-decoration-line|text-align|line-height|margin-left|padding-left|text-indent|width|height|vertical-align|border|border-collapse)$/;

  const safeUrl = (u, img) => {
    const v = String(u || '').trim();
    if (img) return /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml);/i.test(v) && !/^data:image\/svg/i.test(v) ? v : (/^https?:\/\//i.test(v) ? v : '');
    return /^(https?:|mailto:|tel:|#)/i.test(v) ? v : '';
  };

  function cleanStyle(st) {
    return String(st || '').split(';').map((d) => d.trim()).filter(Boolean).map((d) => {
      const i = d.indexOf(':');
      if (i < 0) return null;
      const k = d.slice(0, i).trim().toLowerCase();
      const v = d.slice(i + 1).trim();
      if (!STYLE_OK.test(k) || /url\(|expression|javascript:/i.test(v)) return null;
      return `${k}: ${v}`;
    }).filter(Boolean).join('; ');
  }

  function sanitizeNode(root) {
    const walk = (node) => {
      for (const c of Array.from(node.childNodes)) {
        if (c.nodeType === 8) { c.remove(); continue; }
        if (c.nodeType !== 1) continue;
        const tag = c.tagName.toUpperCase();
        if (DROP.has(tag)) { c.remove(); continue; }
        if (!ALLOWED.has(tag)) {
          walk(c);
          c.replaceWith(...Array.from(c.childNodes));
          continue;
        }
        for (const a of Array.from(c.attributes)) {
          const n = a.name.toLowerCase();
          if (!ATTRS.has(n)) { c.removeAttribute(a.name); continue; }
          if (n === 'href') { const u = safeUrl(a.value, false); if (u) c.setAttribute('href', u); else c.removeAttribute('href'); }
          if (n === 'src') { const u = safeUrl(a.value, true); if (u) c.setAttribute('src', u); else c.removeAttribute('src'); }
          if (n === 'style') { const v = cleanStyle(a.value); if (v) c.setAttribute('style', v); else c.removeAttribute('style'); }
          if (n === 'contenteditable' && a.value !== 'false') c.removeAttribute(a.name);
        }
        if (tag === 'INPUT' && c.getAttribute('type') !== 'checkbox') { c.remove(); continue; }
        if (tag === 'A') { c.setAttribute('target', '_blank'); c.setAttribute('rel', 'noopener noreferrer'); }
        if (tag === 'IMG' && !c.getAttribute('src')) { c.remove(); continue; }
        walk(c);
      }
    };
    walk(root);
    return root;
  }

  function sanitize(html) {
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    sanitizeNode(doc.body);
    return doc.body.innerHTML;
  }

  /* ---------------- Markdown → HTML ---------------- */
  function inlineMd(s) {
    let t = esc(s);
    const codes = [];
    t = t.replace(/`([^`]+)`/g, (m, c) => { codes.push(c); return `\u0000${codes.length - 1}\u0000`; });
    t = t.replace(/!\[([^\]]*)\]\(([^)\s]+)\)/g, (m, a, u) => (safeUrl(u, true) ? `<img alt="${a}" src="${u}">` : a));
    t = t.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (m, a, u) => (safeUrl(u) ? `<a href="${u}">${a}</a>` : a));
    t = t.replace(/\*\*\*(.+?)\*\*\*/g, '<b><i>$1</i></b>');
    t = t.replace(/\*\*(.+?)\*\*|__(.+?)__/g, (m, a, b) => `<b>${a || b}</b>`);
    t = t.replace(/(^|[^*\w])\*(?!\s)(.+?)\*(?!\w)/g, '$1<i>$2</i>');
    t = t.replace(/(^|[^_\w])_(?!\s)(.+?)_(?!\w)/g, '$1<i>$2</i>');
    t = t.replace(/~~(.+?)~~/g, '<s>$1</s>');
    t = t.replace(/==(.+?)==/g, '<mark>$1</mark>');
    t = t.replace(/\u0000(\d+)\u0000/g, (m, i) => `<code>${codes[+i]}</code>`);
    return t;
  }

  function mdToHtml(md) {
    const lines = String(md || '').replace(/\r\n?/g, '\n').split('\n');
    const out = [];
    let i = 0;
    const isTableSep = (l) => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(l);
    const cells = (l) => l.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
    while (i < lines.length) {
      const l = lines[i];
      if (!l.trim()) { i++; continue; }
      let m;
      if ((m = /^```/.exec(l))) {
        const buf = [];
        i++;
        while (i < lines.length && !/^```/.test(lines[i])) buf.push(lines[i++]);
        i++;
        out.push(`<pre><code>${esc(buf.join('\n'))}</code></pre>`);
        continue;
      }
      if ((m = /^(#{1,6})\s+(.*)$/.exec(l))) { out.push(`<h${m[1].length}>${inlineMd(m[2])}</h${m[1].length}>`); i++; continue; }
      if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(l)) { out.push('<hr>'); i++; continue; }
      if (/^\s*>/.test(l)) {
        const buf = [];
        while (i < lines.length && /^\s*>/.test(lines[i])) buf.push(lines[i++].replace(/^\s*>\s?/, ''));
        const first = buf[0] || '';
        const cm = /^\[!(NOTE|INFO|TIP|WARNING|DANGER|SUCCESS)\]\s*/i.exec(first);
        if (cm) {
          const kind = { NOTE: 'info', INFO: 'info', TIP: 'success', SUCCESS: 'success', WARNING: 'warning', DANGER: 'danger' }[cm[1].toUpperCase()];
          buf[0] = first.slice(cm[0].length);
          out.push(`<div data-callout="${kind}">${mdToHtml(buf.join('\n')) || '<p><br></p>'}</div>`);
        } else out.push(`<blockquote>${mdToHtml(buf.join('\n'))}</blockquote>`);
        continue;
      }
      if (l.includes('|') && i + 1 < lines.length && isTableSep(lines[i + 1])) {
        const head = cells(l);
        i += 2;
        const rows = [];
        while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(cells(lines[i++]));
        out.push(`<table><thead><tr>${head.map((c) => `<th>${inlineMd(c)}</th>`).join('')}</tr></thead><tbody>${rows.map((r) => `<tr>${head.map((_, k) => `<td>${inlineMd(r[k] || '')}</td>`).join('')}</tr>`).join('')}</tbody></table>`);
        continue;
      }
      if (/^\s*([-*+]|\d+[.)])\s+/.test(l)) {
        // İç içe listeler (2 boşluk / sekme girintisi)
        const items = [];
        while (i < lines.length && (/^\s*([-*+]|\d+[.)])\s+/.test(lines[i]))) {
          const mm = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(lines[i]);
          items.push({ depth: Math.floor(mm[1].replace(/\t/g, '  ').length / 2), ordered: /\d/.test(mm[2]), text: mm[3] });
          i++;
        }
        const build = (start, depth) => {
          let html = '';
          let k = start;
          const ordered = items[start].ordered;
          let task = false;
          const lis = [];
          while (k < items.length && items[k].depth >= depth) {
            if (items[k].depth > depth) { const [sub, nk] = build(k, items[k].depth); lis[lis.length - 1] += sub; k = nk; continue; }
            let t = items[k].text;
            const tm = /^\[( |x|X)\]\s+(.*)$/.exec(t);
            if (tm) { task = true; lis.push(`<li data-check="${tm[1] === ' ' ? '0' : '1'}">${inlineMd(tm[2])}`); } else lis.push(`<li>${inlineMd(t)}`);
            k++;
          }
          const tag = ordered ? 'ol' : 'ul';
          html = `<${tag}${task ? ' class="checklist"' : ''}>${lis.map((x) => `${x}</li>`).join('')}</${tag}>`;
          return [html, k];
        };
        let k = 0;
        while (k < items.length) { const [h, nk] = build(k, items[k].depth); out.push(h); k = nk; }
        continue;
      }
      const buf = [];
      while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|\s*>|\s*([-*+]|\d+[.)])\s+)/.test(lines[i]) && !(lines[i].includes('|') && i + 1 < lines.length && isTableSep(lines[i + 1]))) buf.push(lines[i++]);
      out.push(`<p>${buf.map((b) => inlineMd(b.replace(/\s{2,}$/, ''))).join('<br>')}</p>`);
    }
    return out.join('');
  }

  /* ---------------- HTML → Markdown ---------------- */
  function htmlToMd(html) {
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    const inl = (n) => Array.from(n.childNodes).map((c) => {
      if (c.nodeType === 3) return c.data.replace(/ /g, ' ').replace(/([*_`])/g, '\\$1');
      if (c.nodeType !== 1) return '';
      const t = c.tagName;
      const inner = inl(c);
      if (!inner.trim() && t !== 'BR' && t !== 'IMG') return inner;
      if (t === 'BR') return '  \n';
      if (t === 'B' || t === 'STRONG') return `**${inner}**`;
      if (t === 'I' || t === 'EM') return `*${inner}*`;
      if (t === 'S' || t === 'DEL' || t === 'STRIKE') return `~~${inner}~~`;
      if (t === 'MARK') return `==${inner}==`;
      if (t === 'CODE') return `\`${c.textContent}\``;
      if (t === 'A') return `[${inner}](${c.getAttribute('href') || ''})`;
      if (t === 'IMG') return `![${c.getAttribute('alt') || ''}](${c.getAttribute('src') || ''})`;
      if (t === 'INPUT') return '';
      if (t === 'SPAN' && c.style.fontWeight && +c.style.fontWeight >= 600) return `**${inner}**`;
      return inner;
    }).join('');
    const list = (el, depth) => {
      const ordered = el.tagName === 'OL';
      let n = 0;
      return Array.from(el.children).filter((li) => li.tagName === 'LI').map((li) => {
        n++;
        const subs = Array.from(li.children).filter((x) => x.tagName === 'UL' || x.tagName === 'OL');
        const clone = li.cloneNode(true);
        Array.from(clone.children).filter((x) => x.tagName === 'UL' || x.tagName === 'OL').forEach((x) => x.remove());
        const chk = li.dataset.check !== undefined ? `[${li.dataset.check === '1' ? 'x' : ' '}] ` : '';
        const line = `${'  '.repeat(depth)}${ordered ? `${n}.` : '-'} ${chk}${inl(clone).trim()}`;
        return [line, ...subs.map((s) => list(s, depth + 1))].join('\n');
      }).join('\n');
    };
    const block = (n) => Array.from(n.childNodes).map((c) => {
      if (c.nodeType === 3) return c.data.trim() ? `${c.data.trim()}\n\n` : '';
      if (c.nodeType !== 1) return '';
      const t = c.tagName;
      if (/^H[1-6]$/.test(t)) return `${'#'.repeat(+t[1])} ${inl(c).trim()}\n\n`;
      if (t === 'P') return inl(c).trim() ? `${inl(c).trim()}\n\n` : '';
      if (t === 'UL' || t === 'OL') return `${list(c, 0)}\n\n`;
      if (t === 'BLOCKQUOTE') return `${block(c).trim().split('\n').map((x) => `> ${x}`).join('\n')}\n\n`;
      if (t === 'PRE') return `\`\`\`\n${c.textContent.replace(/\n$/, '')}\n\`\`\`\n\n`;
      if (t === 'HR') return c.dataset.pageBreak !== undefined ? '\n---\n\n' : '---\n\n';
      if (t === 'DIV' && c.dataset.callout) {
        const tag = { info: 'NOTE', success: 'TIP', warning: 'WARNING', danger: 'DANGER' }[c.dataset.callout] || 'NOTE';
        return `> [!${tag}]\n${block(c).trim().split('\n').map((x) => `> ${x}`).join('\n')}\n\n`;
      }
      if (t === 'TABLE') {
        const rows = Array.from(c.querySelectorAll('tr')).map((r) => Array.from(r.children).map((td) => inl(td).trim().replace(/\|/g, '\\|').replace(/\n/g, ' ')));
        if (!rows.length) return '';
        const w = Math.max(...rows.map((r) => r.length));
        const fill = (r) => `| ${Array.from({ length: w }, (_, k) => r[k] || '').join(' | ')} |`;
        return `${[fill(rows[0]), `|${' --- |'.repeat(w)}`, ...rows.slice(1).map(fill)].join('\n')}\n\n`;
      }
      if (t === 'DETAILS') {
        const sum = c.querySelector('summary');
        const rest = c.cloneNode(true);
        rest.querySelector('summary')?.remove();
        return `**${sum ? inl(sum).trim() : ''}**\n\n${block(rest)}`;
      }
      if (t === 'IMG') return `${inl({ childNodes: [c] })}\n\n`;
      if (t === 'DIV' || t === 'SECTION' || t === 'ASIDE' || t === 'FIGURE' || t === 'NAV') return block(c);
      return inl(c).trim() ? `${inl(c).trim()}\n\n` : '';
    }).join('');
    return block(doc.body).replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  /* ---------------- Ana hat (sunum için) ---------------- */
  /** Düz metin / Markdown ana hattını slayt listesine çevirir: [{ title, bullets: [{ text, lvl }], notes }] */
  function outlineFromText(text) {
    const slides = [];
    let curr = null;
    const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
    const start = (title) => { curr = { title: title.trim(), bullets: [], notes: '' }; slides.push(curr); };
    for (const raw of lines) {
      if (!raw.trim()) continue;
      let m;
      if ((m = /^\s*#{1,6}\s+(.*)$/.exec(raw))) { start(m[1]); continue; }
      if (/^\s*---+\s*$/.test(raw)) { curr = null; continue; }
      if ((m = /^\s*>\s?(.*)$/.exec(raw))) { if (!curr) start(''); curr.notes += (curr.notes ? '\n' : '') + m[1]; continue; }
      const bm = /^(\s*)([-*+•]|\d+[.)])\s+(.*)$/.exec(raw);
      const indent = (raw.match(/^\s*/)[0] || '').replace(/\t/g, '  ').length;
      if (bm || indent >= 2) {
        if (!curr) start('');
        const t = bm ? bm[3] : raw.trim();
        const lvl = Math.min(4, Math.floor((bm ? bm[1] : raw.match(/^\s*/)[0]).replace(/\t/g, '  ').length / 2));
        curr.bullets.push({ text: t.replace(/\*\*|__|`/g, ''), lvl });
      } else start(raw.replace(/\*\*|__|`/g, ''));
    }
    return slides.filter((s) => s.title || s.bullets.length);
  }

  /** HTML belgesinden ana hat: H1/H2 başlıkları slayt, altındaki paragraflar ve maddeler madde olur */
  function outlineFromHtml(html, docTitle = '') {
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    const slides = [];
    let curr = null;
    const hasH = doc.body.querySelector('h1, h2');
    const start = (title) => { curr = { title, bullets: [], notes: '' }; slides.push(curr); };
    const shorten = (t) => {
      const s = t.replace(/\s+/g, ' ').trim();
      if (s.length <= 140) return s;
      const first = s.split(/(?<=[.!?])\s/)[0];
      return first.length <= 160 ? first : `${s.slice(0, 137)}…`;
    };
    const push = (text, lvl = 0) => {
      const t = shorten(text);
      if (!t) return;
      if (!curr || (!hasH && curr.bullets.length >= 5)) start(curr ? tx('{title} (devam)', { title: curr.title.replace(/ \((devam|cont\.)\)$/, '') }) : docTitle || tx('Giriş'));
      curr.bullets.push({ text: t, lvl });
    };
    const visit = (n, lvl) => {
      for (const c of Array.from(n.children)) {
        const t = c.tagName;
        if (t === 'H1' || t === 'H2') start(c.textContent.trim());
        else if (/^H[3-6]$/.test(t)) push(c.textContent, 0);
        else if (t === 'UL' || t === 'OL') {
          for (const li of Array.from(c.children)) {
            const clone = li.cloneNode(true);
            Array.from(clone.querySelectorAll('ul, ol')).forEach((x) => x.remove());
            push(clone.textContent, lvl);
            Array.from(li.children).filter((x) => x.tagName === 'UL' || x.tagName === 'OL').forEach((sub) => visit({ children: [sub] }, Math.min(4, lvl + 1)));
          }
        } else if (t === 'P' || t === 'BLOCKQUOTE' || t === 'PRE') push(c.textContent, lvl);
        else if (t === 'TABLE') {
          Array.from(c.querySelectorAll('tr')).slice(0, 6).forEach((r) => push(Array.from(r.children).map((x) => x.textContent.trim()).join(' · '), lvl));
        } else if (t === 'DIV' || t === 'DETAILS' || t === 'SECTION') visit(c, lvl);
      }
    };
    visit(doc.body, 0);
    return slides.filter((s) => s.title || s.bullets.length);
  }

  function bulletList(items, start, lvl) {
    let h = '<ul>';
    let k = start;
    while (k < items.length && items[k].lvl >= lvl) {
      if (items[k].lvl > lvl) {
        const [sub, nk] = bulletList(items, k, items[k].lvl);
        h = h.endsWith('</li>') ? `${h.slice(0, -5)}${sub}</li>` : `${h}<li>${sub}</li>`;
        k = nk;
        continue;
      }
      h += `<li>${esc(items[k].text)}</li>`;
      k++;
    }
    return [`${h}</ul>`, k];
  }

  /** Ana hattı EasyWrite HTML'ine çevirir */
  function outlineToHtml(title, slides) {
    let html = title ? `<h1>${esc(title)}</h1>` : '';
    for (const s of slides) {
      if (s.title) html += `<h2>${esc(s.title)}</h2>`;
      let k = 0;
      while (k < s.bullets.length) { const [h, nk] = bulletList(s.bullets, k, s.bullets[k].lvl); html += h; k = nk; }
      if (s.notes) html += `<div data-callout="info"><p><b>${tx('Konuşmacı notu:')}</b> ${esc(s.notes).replace(/\n/g, '<br>')}</p></div>`;
    }
    return html;
  }

  window.Convert = { sanitize, sanitizeNode, mdToHtml, htmlToMd, inlineMd, outlineFromText, outlineFromHtml, outlineToHtml, esc };
})();
