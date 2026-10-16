/* Platform365 EasyWrite — Word (.docx / Office Open XML) okuma ve yazma
 * Bağımlılık: JSZip (vendor/pptxgen.bundle.js içinde). */
'use strict';

const Docx = (() => {
  const xe = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const PAGE = { A4: [11906, 16838], Letter: [12240, 15840], A5: [8391, 11906] };
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  const cssColor = (c) => {
    if (!c) return null;
    c = c.trim();
    let m = /^#([0-9a-f]{3})$/i.exec(c);
    if (m) return [...m[1]].map((x) => x + x).join('').toUpperCase();
    m = /^#([0-9a-f]{6})/i.exec(c);
    if (m) return m[1].toUpperCase();
    m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/i.exec(c);
    if (m) { if (m[4] !== undefined && +m[4] === 0) return null; return [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase(); }
    const named = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080' };
    return named[c.toLowerCase()] || null;
  };
  const pxOf = (v, base = 16) => {
    if (!v) return null;
    const n = parseFloat(v);
    if (Number.isNaN(n)) return null;
    if (/pt$/.test(v)) return n / 0.75;
    if (/em$/.test(v)) return n * base;
    if (/%$/.test(v)) return (n / 100) * base;
    return n;
  };
  const FONT_SIZES = { 1: 10, 2: 13, 3: 16, 4: 18, 5: 24, 6: 32, 7: 48 };

  /* ================= YAZMA ================= */
  async function build(doc) {
    const JSZip = await P365.loadZip();
    const zip = new JSZip();
    const rels = [];      // { id, type, target, external }
    const media = [];     // { name, data(base64), ext }
    let relN = 10;
    let numN = 2;         // 1: madde işaretli (paylaşılan); 2+: numaralı listeler (her biri yeniden başlar)
    const nums = [];      // { id, abs }
    const addRel = (type, target, external) => {
      const id = `rId${++relN}`;
      rels.push({ id, type, target, external });
      return id;
    };
    const page = PAGE[doc.page?.size] || PAGE.A4;
    const marginPx = { narrow: 48, normal: 96, wide: 144 }[doc.page?.margin || 'normal'] || 96;
    const margin = Math.round(marginPx * 15);
    const contentPx = page[0] / 15 - marginPx * 2;
    const baseFont = doc.page?.font || 'Calibri';
    const baseSize = doc.page?.fontSize || 16; // px

    const root = new DOMParser().parseFromString(`<body>${doc.html || ''}</body>`, 'text/html').body;

    // Görsellerin doğal boyutlarını önceden öğren
    const imgSize = new Map();
    await Promise.all(Array.from(root.querySelectorAll('img')).map((img) => new Promise((res) => {
      const im = new Image();
      im.onload = () => { imgSize.set(img, [im.naturalWidth, im.naturalHeight]); res(); };
      im.onerror = () => res();
      im.src = img.getAttribute('src');
    })));

    const rPr = (f) => {
      let x = '';
      if (f.style) x += `<w:rStyle w:val="${f.style}"/>`;
      if (f.font) x += `<w:rFonts w:ascii="${xe(f.font)}" w:hAnsi="${xe(f.font)}" w:cs="${xe(f.font)}"/>`;
      if (f.bold) x += '<w:b/><w:bCs/>';
      if (f.italic) x += '<w:i/><w:iCs/>';
      if (f.strike) x += '<w:strike/>';
      if (f.color) x += `<w:color w:val="${f.color}"/>`;
      if (f.size) x += `<w:sz w:val="${Math.round(f.size * 1.5)}"/><w:szCs w:val="${Math.round(f.size * 1.5)}"/>`;
      if (f.mark) x += '<w:highlight w:val="yellow"/>';
      if (f.underline) x += '<w:u w:val="single"/>';
      if (f.shd) x += `<w:shd w:val="clear" w:color="auto" w:fill="${f.shd}"/>`;
      if (f.vert) x += `<w:vertAlign w:val="${f.vert}"/>`;
      return x ? `<w:rPr>${x}</w:rPr>` : '';
    };
    const textRun = (t, f) => `<w:r>${rPr(f)}<w:t xml:space="preserve">${xe(t)}</w:t></w:r>`;

    function imageRun(img) {
      const src = img.getAttribute('src') || '';
      const m = /^data:image\/(png|jpe?g|gif|bmp|webp);base64,(.*)$/i.exec(src);
      if (!m) return '';
      const ext = m[1].toLowerCase() === 'jpeg' ? 'jpg' : m[1].toLowerCase();
      const name = `image${media.length + 1}.${ext}`;
      media.push({ name, data: m[2], ext });
      const rid = addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', `media/${name}`);
      const [nw, nh] = imgSize.get(img) || [400, 300];
      let w = nw;
      const sw = img.style.width;
      if (sw && /%$/.test(sw)) w = (parseFloat(sw) / 100) * contentPx;
      else if (sw) w = pxOf(sw) || nw;
      else if (img.getAttribute('width')) w = +img.getAttribute('width');
      w = Math.min(w, contentPx);
      const h = (w * nh) / (nw || 1);
      const cx = Math.round(w * 9525);
      const cy = Math.round(h * 9525);
      const id = media.length;
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Resim ${id}" descr="${xe(img.getAttribute('alt') || '')}"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    }

    function fmtOf(n, f) {
      const g = { ...f };
      const t = n.tagName;
      if (t === 'B' || t === 'STRONG' || t === 'TH') g.bold = true;
      if (t === 'I' || t === 'EM') g.italic = true;
      if (t === 'U') g.underline = true;
      if (t === 'S' || t === 'STRIKE' || t === 'DEL') g.strike = true;
      if (t === 'SUP') g.vert = 'superscript';
      if (t === 'SUB') g.vert = 'subscript';
      if (t === 'MARK') g.mark = true;
      if (t === 'CODE' || t === 'KBD') g.font = 'Consolas';
      if (t === 'SMALL') g.size = (g.size || baseSize) * 0.83;
      if (t === 'FONT') {
        const c = cssColor(n.getAttribute('color'));
        if (c) g.color = c;
        if (n.getAttribute('face')) g.font = n.getAttribute('face').split(',')[0].replace(/["']/g, '').trim();
        if (n.getAttribute('size')) g.size = FONT_SIZES[n.getAttribute('size')] || g.size;
      }
      const st = n.style;
      if (st) {
        if (st.fontWeight) g.bold = st.fontWeight === 'bold' || +st.fontWeight >= 600;
        if (st.fontStyle) g.italic = st.fontStyle === 'italic';
        const td = st.textDecorationLine || st.textDecoration;
        if (td) { if (td.includes('underline')) g.underline = true; if (td.includes('line-through')) g.strike = true; }
        if (st.color) { const c = cssColor(st.color); if (c) g.color = c; }
        if (st.backgroundColor) { const c = cssColor(st.backgroundColor); if (c) g.shd = c; }
        if (st.fontSize) { const p = pxOf(st.fontSize, g.size || baseSize); if (p) g.size = p; }
        if (st.fontFamily) g.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      }
      return g;
    }

    function runs(node, f) {
      let x = '';
      for (const c of Array.from(node.childNodes)) {
        if (c.nodeType === 3) {
          const t = c.data.replace(/[\n\r]+/g, ' ');
          if (t) x += textRun(t, f);
        } else if (c.nodeType === 1) {
          const t = c.tagName;
          if (t === 'BR') x += '<w:r><w:br/></w:r>';
          else if (t === 'IMG') x += imageRun(c);
          else if (t === 'INPUT') continue;
          else if (t === 'A' && c.getAttribute('href') && !c.getAttribute('href').startsWith('#')) {
            const rid = addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', c.getAttribute('href'), true);
            x += `<w:hyperlink r:id="${rid}">${runs(c, { ...fmtOf(c, f), style: 'Hyperlink' })}</w:hyperlink>`;
          } else if (/^(UL|OL|TABLE|P|DIV|H[1-6]|BLOCKQUOTE|PRE)$/.test(t)) continue; // blok öğeler ayrı işlenir
          else x += runs(c, fmtOf(c, f));
        }
      }
      return x;
    }

    function pPr(n, o = {}) {
      let x = '';
      if (o.style) x += `<w:pStyle w:val="${o.style}"/>`;
      if (o.keepNext) x += '<w:keepNext/>';
      if (o.num) x += `<w:numPr><w:ilvl w:val="${o.num.lvl}"/><w:numId w:val="${o.num.id}"/></w:numPr>`;
      if (o.border === 'bottom') x += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A0A4B0"/></w:pBdr>';
      if (o.callout) {
        const col = { info: ['DBEAFE', '3B82F6'], success: ['DCFCE7', '22C55E'], warning: ['FEF3C7', 'F59E0B'], danger: ['FEE2E2', 'EF4444'] }[o.callout] || ['F3F4F6', '9CA3AF'];
        x += `<w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="${col[1]}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${col[0]}"/>`;
      }
      const st = n && n.style;
      const lh = st && st.lineHeight ? parseFloat(st.lineHeight) : null;
      if (lh && lh < 5) x += `<w:spacing w:line="${Math.round(lh * 240)}" w:lineRule="auto"/>`;
      const ml = (st && pxOf(st.marginLeft)) || (st && pxOf(st.paddingLeft)) || 0;
      if (o.indent || ml) x += `<w:ind w:left="${Math.round((o.indent || 0) + ml * 15)}"/>`;
      const al = (st && st.textAlign) || (n && n.getAttribute && n.getAttribute('align'));
      const jc = { center: 'center', right: 'right', justify: 'both', left: 'left' }[al];
      if (jc) x += `<w:jc w:val="${jc}"/>`;
      return x ? `<w:pPr>${x}</w:pPr>` : '';
    }
    const para = (n, inner, o) => `<w:p>${pPr(n, o)}${inner}</w:p>`;

    function listXml(list, lvl, numId) {
      let x = '';
      const ordered = list.tagName === 'OL';
      const check = list.classList.contains('checklist');
      if (lvl === 0) {
        if (ordered) { numId = ++numN; nums.push({ id: numId, abs: 1 }); } else numId = 1;
      }
      for (const li of Array.from(list.children)) {
        if (li.tagName !== 'LI') continue;
        const blocksInside = Array.from(li.children).filter((c) => /^(P|DIV|H[1-6])$/.test(c.tagName));
        let inner = '';
        if (check) inner += textRun(li.dataset.check === '1' ? '☒ ' : '☐ ', {});
        inner += runs(li, li.dataset.check === '1' ? { strike: false } : {});
        blocksInside.forEach((b) => { inner += runs(b, fmtOf(b, {})); });
        x += check ? para(li, inner, { indent: 360 * (lvl + 1) }) : para(li, inner, { num: { id: numId, lvl: Math.min(8, lvl) } });
        for (const sub of Array.from(li.children)) if (sub.tagName === 'UL' || sub.tagName === 'OL') x += listXml(sub, lvl + 1, sub.tagName === 'OL' && !ordered ? (() => { const id = ++numN; nums.push({ id, abs: 1 }); return id; })() : numId);
      }
      return x;
    }

    function tableXml(tbl) {
      const rows = Array.from(tbl.querySelectorAll(':scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr'));
      if (!rows.length) return '';
      const ncol = Math.max(...rows.map((r) => Array.from(r.children).reduce((s, c) => s + (+c.getAttribute('colspan') || 1), 0)), 1);
      const tw = Math.round(contentPx * 15);
      const cw = Math.floor(tw / ncol);
      let x = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${tw}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="A0A4B0"/>`).join('')}</w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`;
      x += `<w:tblGrid>${`<w:gridCol w:w="${cw}"/>`.repeat(ncol)}</w:tblGrid>`;
      for (const r of rows) {
        x += '<w:tr>';
        for (const c of Array.from(r.children)) {
          const span = +c.getAttribute('colspan') || 1;
          const th = c.tagName === 'TH';
          const bg = cssColor(c.style.backgroundColor) || (th ? 'EEF0F6' : null);
          let inner = '';
          const subBlocks = Array.from(c.children).filter((b) => /^(P|DIV|H[1-6]|UL|OL)$/.test(b.tagName));
          inner += para(c, runs(c, th ? { bold: true } : {}));
          subBlocks.forEach((b) => { inner += b.tagName === 'UL' || b.tagName === 'OL' ? listXml(b, 0) : para(b, runs(b, fmtOf(b, th ? { bold: true } : {}))); });
          x += `<w:tc><w:tcPr><w:tcW w:w="${cw * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : ''}</w:tcPr>${inner}</w:tc>`;
        }
        x += '</w:tr>';
      }
      return `${x}</w:tbl><w:p/>`;
    }

    function blocks(node, ctx = {}) {
      let x = '';
      let inline = [];
      const flush = () => {
        if (!inline.length) return;
        const wrap = document.createElement('span');
        inline.forEach((c) => wrap.appendChild(c.cloneNode(true)));
        if (wrap.textContent.trim() || wrap.querySelector('img')) x += para(null, runs(wrap, {}), ctx);
        inline = [];
      };
      for (const c of Array.from(node.childNodes)) {
        if (c.nodeType === 3 || (c.nodeType === 1 && !/^(P|DIV|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|PRE|HR|DETAILS|SUMMARY|NAV|SECTION|ASIDE|FIGURE|FIGCAPTION)$/.test(c.tagName))) {
          inline.push(c);
          continue;
        }
        flush();
        if (c.nodeType !== 1) continue;
        const t = c.tagName;
        if (/^H[1-6]$/.test(t)) x += para(c, runs(c, {}), { ...ctx, style: `Heading${Math.min(+t[1], 4)}`, keepNext: true });
        else if (t === 'P' || t === 'FIGCAPTION' || t === 'SUMMARY') {
          const hasBlock = Array.from(c.children).some((k) => /^(UL|OL|TABLE)$/.test(k.tagName));
          x += para(c, runs(c, t === 'SUMMARY' ? { bold: true } : {}), ctx);
          if (hasBlock) x += blocks({ childNodes: Array.from(c.children).filter((k) => /^(UL|OL|TABLE)$/.test(k.tagName)) }, ctx);
        } else if (t === 'UL' || t === 'OL') x += listXml(c, 0);
        else if (t === 'TABLE') x += tableXml(c);
        else if (t === 'BLOCKQUOTE') x += blocks(c, { ...ctx, style: 'Quote' });
        else if (t === 'PRE') {
          c.textContent.replace(/\n$/, '').split('\n').forEach((line) => { x += para(null, textRun(line || ' ', { font: 'Consolas' }), { ...ctx, style: 'Code' }); });
        } else if (t === 'HR') {
          if (c.dataset.pageBreak !== undefined) x += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
          else x += para(null, '', { border: 'bottom' });
        } else if (t === 'DIV' && c.dataset.callout) x += blocks(c, { ...ctx, callout: c.dataset.callout });
        else if (t === 'NAV' && c.dataset.toc !== undefined) {
          x += para(null, textRun('İçindekiler', { bold: true }), { style: 'TOCHeading' });
          c.querySelectorAll('a').forEach((a) => { x += para(null, textRun(a.textContent, {}), { indent: (+(a.dataset.lvl || 1) - 1) * 360 }); });
        } else x += blocks(c, ctx);
      }
      flush();
      return x;
    }

    let body = blocks(root);
    if (!body) body = '<w:p/>';
    const sect = `<w:sectPr><w:pgSz w:w="${page[0]}" w:h="${page[1]}"/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;
    const NSDECL = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NSDECL}><w:body>${body}${sect}</w:body></w:document>`);

    const half = Math.round(baseSize * 1.5);
    const hd = (id, name, size, color, extra = '') => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="${+id.slice(-1) - 1}"/>${extra}</w:pPr><w:rPr><w:b/><w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
    zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${xe(baseFont)}" w:hAnsi="${xe(baseFont)}" w:cs="${xe(baseFont)}" w:eastAsia="${xe(baseFont)}"/><w:sz w:val="${half}"/><w:szCs w:val="${half}"/><w:lang w:val="tr-TR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
${hd('Heading1', 'heading 1', Math.round(half * 2), '1F2937')}${hd('Heading2', 'heading 2', Math.round(half * 1.6), '1F2937')}${hd('Heading3', 'heading 3', Math.round(half * 1.3), '374151')}${hd('Heading4', 'heading 4', Math.round(half * 1.1), '374151')}
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="A0A4B0"/></w:pBdr><w:ind w:left="360"/></w:pPr><w:rPr><w:i/><w:color w:val="4B5563"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="${Math.round(half * 0.9)}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="${Math.round(half * 1.3)}"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="2563EB"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`);

    const lvlXml = (bullet) => Array.from({ length: 9 }, (_, i) => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${bullet ? 'bullet' : ['decimal', 'lowerLetter', 'lowerRoman'][i % 3]}"/><w:lvlText w:val="${bullet ? ['•', '◦', '▪'][i % 3] : `%${i + 1}.`}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + i * 360}" w:hanging="360"/></w:pPr>${bullet ? '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' : ''}</w:lvl>`).join('');
    zip.file('word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvlXml(true)}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvlXml(false)}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>${nums.map((n) => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="${n.abs}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('')}</w:numbering>`);

    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>${rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${xe(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`);
    media.forEach((m) => zip.file(`word/media/${m.name}`, m.data, { base64: true }));

    const exts = [...new Set(media.map((m) => m.ext))];
    const mime = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/>${exts.map((e) => `<Default Extension="${e}" ContentType="${mime[e]}"/>`).join('')}
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    const iso = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xe(doc.title || '')}</dc:title><dc:creator>Platform365 EasyWrite</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>`);
    zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Platform365 EasyWrite</Application></Properties>`);
    return zip;
  }

  /* ================= OKUMA ================= */
  const kids = (n, name) => Array.from(n.childNodes).filter((c) => c.nodeType === 1 && (!name || c.localName === name));
  const kid = (n, name) => kids(n, name)[0] || null;
  const wv = (n) => (n ? n.getAttributeNS(W_NS, 'val') ?? n.getAttribute('w:val') : null);
  const on = (n) => n && !['0', 'false', 'none'].includes(String(wv(n)));
  const HL = { yellow: '#fff59d', green: '#a5f3a5', cyan: '#a5f3fc', magenta: '#f5a5f3', blue: '#93c5fd', red: '#fca5a5', darkYellow: '#d4c34a', lightGray: '#e5e7eb', darkGray: '#9ca3af' };

  async function parse(zip) {
    const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : null; };
    const xml = (s) => (s ? new DOMParser().parseFromString(s, 'application/xml') : null);
    const docXml = xml(await read('word/document.xml'));
    if (!docXml) throw new Error('word/document.xml bulunamadı');
    const relsX = xml(await read('word/_rels/document.xml.rels'));
    const rels = {};
    if (relsX) Array.from(relsX.getElementsByTagName('Relationship')).forEach((r) => { rels[r.getAttribute('Id')] = { target: r.getAttribute('Target'), external: r.getAttribute('TargetMode') === 'External' }; });
    const stylesX = xml(await read('word/styles.xml'));
    const styleName = {};
    if (stylesX) Array.from(stylesX.getElementsByTagNameNS(W_NS, 'style')).forEach((s) => { const n = kid(s, 'name'); styleName[s.getAttributeNS(W_NS, 'styleId')] = (wv(n) || '').toLowerCase(); });
    const numX = xml(await read('word/numbering.xml'));
    const numFmt = {}; // numId → ilvl → fmt
    if (numX) {
      const abs = {};
      Array.from(numX.getElementsByTagNameNS(W_NS, 'abstractNum')).forEach((a) => {
        abs[a.getAttributeNS(W_NS, 'abstractNumId')] = Object.fromEntries(kids(a, 'lvl').map((l) => [l.getAttributeNS(W_NS, 'ilvl'), wv(kid(l, 'numFmt'))]));
      });
      Array.from(numX.getElementsByTagNameNS(W_NS, 'num')).forEach((n) => { numFmt[n.getAttributeNS(W_NS, 'numId')] = abs[wv(kid(n, 'abstractNumId'))] || {}; });
    }
    const mediaCache = {};
    async function mediaUrl(rid) {
      const r = rels[rid];
      if (!r || r.external) return null;
      const path = r.target.startsWith('/') ? r.target.slice(1) : `word/${r.target}`.replace(/\/[^/]+\/\.\.\//g, '/');
      if (mediaCache[path] !== undefined) return mediaCache[path];
      const f = zip.file(path);
      if (!f) return (mediaCache[path] = null);
      const ext = path.split('.').pop().toLowerCase();
      const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[ext];
      if (!mime) return (mediaCache[path] = null);
      mediaCache[path] = `data:${mime};base64,${await f.async('base64')}`;
      return mediaCache[path];
    }
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

    async function runHtml(r) {
      const pr = kid(r, 'rPr');
      let html = '';
      for (const c of kids(r)) {
        const n = c.localName;
        if (n === 't') html += esc(c.textContent);
        else if (n === 'tab') html += '&emsp;';
        else if (n === 'br') html += c.getAttributeNS(W_NS, 'type') === 'page' ? '\u0001PB\u0001' : '<br>';
        else if (n === 'drawing' || n === 'pict') {
          const blip = c.getElementsByTagNameNS('*', 'blip')[0] || c.getElementsByTagNameNS('*', 'imagedata')[0];
          const rid = blip && (blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
          const ext = c.getElementsByTagNameNS('*', 'extent')[0];
          const url = rid ? await mediaUrl(rid) : null;
          if (url) {
            const w = ext ? Math.round(+ext.getAttribute('cx') / 9525) : null;
            html += `<img src="${url}"${w ? ` style="width: ${w}px"` : ''}>`;
          }
        } else if (n === 'sym') html += '•';
      }
      if (!pr || !html) return html;
      const st = [];
      const color = wv(kid(pr, 'color'));
      if (color && color !== 'auto' && color !== '000000') st.push(`color: #${color}`);
      const sz = wv(kid(pr, 'sz'));
      if (sz) st.push(`font-size: ${+sz / 2}pt`);
      const fonts = kid(pr, 'rFonts');
      const face = fonts && (fonts.getAttributeNS(W_NS, 'ascii') || fonts.getAttributeNS(W_NS, 'hAnsi'));
      if (face) st.push(`font-family: ${face}`);
      const hl = wv(kid(pr, 'highlight'));
      const shd = kid(pr, 'shd') && kid(pr, 'shd').getAttributeNS(W_NS, 'fill');
      if (hl && HL[hl]) st.push(`background-color: ${HL[hl]}`);
      else if (shd && shd !== 'auto' && shd !== 'FFFFFF') st.push(`background-color: #${shd}`);
      if (st.length) html = `<span style="${st.join('; ')}">${html}</span>`;
      if (on(kid(pr, 'b'))) html = `<b>${html}</b>`;
      if (on(kid(pr, 'i'))) html = `<i>${html}</i>`;
      const u = kid(pr, 'u');
      if (u && wv(u) !== 'none') html = `<u>${html}</u>`;
      if (on(kid(pr, 'strike')) || on(kid(pr, 'dstrike'))) html = `<s>${html}</s>`;
      const va = wv(kid(pr, 'vertAlign'));
      if (va === 'superscript') html = `<sup>${html}</sup>`;
      if (va === 'subscript') html = `<sub>${html}</sub>`;
      return html;
    }

    async function inlineHtml(p) {
      let html = '';
      for (const c of kids(p)) {
        const n = c.localName;
        if (n === 'r') html += await runHtml(c);
        else if (n === 'hyperlink') {
          const rid = c.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
          const href = rid && rels[rid] ? rels[rid].target : null;
          const inner = await inlineHtml(c);
          html += href ? `<a href="${esc(href)}">${inner}</a>` : inner;
        } else if (n === 'ins' || n === 'smartTag' || n === 'sdt' || n === 'sdtContent' || n === 'fldSimple' || n === 'customXml') html += await inlineHtml(c);
      }
      return html;
    }

    const out = [];
    let listStack = []; // { tag, numId }
    const closeLists = (to = 0) => { while (listStack.length > to) out.push(`</li></${listStack.pop().tag}>`); };

    async function para(p) {
      const pr = kid(p, 'pPr');
      const sid = pr ? wv(kid(pr, 'pStyle')) : null;
      const sname = (sid && styleName[sid]) || (sid || '').toLowerCase();
      let inner = await inlineHtml(p);
      const pageBreak = inner.includes('\u0001PB\u0001') || (pr && on(kid(pr, 'pageBreakBefore')));
      inner = inner.replace(/\u0001PB\u0001/g, '');
      const jc = pr ? wv(kid(pr, 'jc')) : null;
      const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc];
      const style = align ? ` style="text-align: ${align}"` : '';
      const numPr = pr && kid(pr, 'numPr');
      const numId = numPr ? wv(kid(numPr, 'numId')) : null;
      if (pageBreak && pr && on(kid(pr, 'pageBreakBefore'))) { closeLists(); out.push('<hr data-page-break>'); }
      if (numId && numId !== '0') {
        const lvl = +(wv(kid(numPr, 'ilvl')) || 0);
        const fmt = (numFmt[numId] || {})[String(lvl)] || 'bullet';
        const tag = fmt === 'bullet' || fmt === 'none' ? 'ul' : 'ol';
        while (listStack.length > lvl + 1) out.push(`</li></${listStack.pop().tag}>`);
        if (listStack.length === lvl + 1 && listStack[lvl].tag !== tag) { out.push(`</li></${listStack.pop().tag}>`); }
        if (listStack.length === lvl + 1) out.push('</li>');
        while (listStack.length < lvl + 1) { out.push(`<${tag}>`); listStack.push({ tag, numId }); }
        out.push(`<li${style}>${inner || '<br>'}`);
      } else {
        closeLists();
        let tag = 'p';
        const hm = /^(heading|başlık)\s*(\d)/.exec(sname);
        if (hm) tag = `h${Math.min(6, +hm[2])}`;
        else if (sname === 'title') tag = 'h1';
        else if (sname === 'subtitle') tag = 'h2';
        else if (/quote|alıntı/.test(sname)) tag = 'blockquote';
        if (tag === 'blockquote') out.push(`<blockquote><p${style}>${inner || '<br>'}</p></blockquote>`);
        else out.push(`<${tag}${style}>${inner || '<br>'}</${tag}>`);
      }
      if (pageBreak && !(pr && on(kid(pr, 'pageBreakBefore')))) { closeLists(); out.push('<hr data-page-break>'); }
    }

    async function table(t) {
      closeLists();
      let html = '<table>';
      const rows = kids(t, 'tr');
      for (let ri = 0; ri < rows.length; ri++) {
        html += '<tr>';
        for (const tc of kids(rows[ri], 'tc')) {
          const pr = kid(tc, 'tcPr');
          const span = pr ? +(wv(kid(pr, 'gridSpan')) || 1) : 1;
          const vm = pr && kid(pr, 'vMerge');
          if (vm && wv(vm) !== 'restart') continue;
          const shd = pr && kid(pr, 'shd') && kid(pr, 'shd').getAttributeNS(W_NS, 'fill');
          const inner = [];
          for (const c of kids(tc)) {
            if (c.localName === 'p') inner.push(await inlineHtml(c));
            else if (c.localName === 'tbl') inner.push(await (async () => { const save = out.length; await table(c); return out.splice(save).join(''); })());
          }
          html += `<td${span > 1 ? ` colspan="${span}"` : ''}${shd && shd !== 'auto' && shd !== 'FFFFFF' ? ` style="background-color: #${shd}"` : ''}>${inner.filter(Boolean).join('<br>') || '<br>'}</td>`;
        }
        html += '</tr>';
      }
      out.push(`${html}</table>`);
    }

    async function walk(node) {
      for (const c of kids(node)) {
        if (c.localName === 'p') await para(c);
        else if (c.localName === 'tbl') await table(c);
        else if (c.localName === 'sdt') { const content = kid(c, 'sdtContent'); if (content) await walk(content); }
        else if (c.localName === 'customXml') await walk(c);
      }
    }
    const body = docXml.getElementsByTagNameNS(W_NS, 'body')[0];
    await walk(body);
    closeLists();
    let title = '';
    const core = xml(await read('docProps/core.xml'));
    if (core) { const t = core.getElementsByTagNameNS('http://purl.org/dc/elements/1.1/', 'title')[0]; if (t) title = t.textContent.trim(); }
    const sect = body.getElementsByTagNameNS(W_NS, 'pgSz')[0];
    let size = 'A4';
    if (sect) { const w = +sect.getAttributeNS(W_NS, 'w'); if (Math.abs(w - 12240) < 200) size = 'Letter'; else if (Math.abs(w - 8391) < 200) size = 'A5'; }
    // EasyWrite'ın yazdığı ☐ / ☒ paragraflarını yeniden yapılacaklar listesine çevir
    const tmp = new DOMParser().parseFromString(`<body>${out.join('')}</body>`, 'text/html').body;
    let group = null;
    for (const p of Array.from(tmp.children)) {
      const m = p.tagName === 'P' && /^\s*([☐☒☑])\s?/.exec(p.textContent);
      if (!m) { group = null; continue; }
      const li = tmp.ownerDocument.createElement('li');
      li.dataset.check = m[1] === '☐' ? '0' : '1';
      li.innerHTML = p.innerHTML.replace(/[☐☒☑]\s?/, '');
      if (!group) { group = tmp.ownerDocument.createElement('ul'); group.className = 'checklist'; p.before(group); }
      group.appendChild(li);
      p.remove();
    }
    return { html: tmp.innerHTML, title, size };
  }

  return { build, parse };
})();
