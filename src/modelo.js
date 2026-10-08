/* Formata ABNT — modo "Modelo da faculdade".
   Para trabalhos feitos num modelo pronto (plano de aula, fichas, relatórios em tabela).
   Em vez de remontar o documento, ajusta o próprio .docx: o cabeçalho, o logo, as
   tabelas e a orientação da folha ficam como estão. Corrige só a formatação:
   - junta tabelas e frases que foram cortadas entre páginas (comum em arquivo convertido de PDF);
   - junta quebras de linha no meio da frase;
   - unifica fonte e tamanho e remove o espaçamento comprimido entre letras;
   - padroniza espaçamentos e recuos dos parágrafos e a margem interna das células;
   - formata as referências (à esquerda, espaço simples, 6 pt antes e depois, ordem alfabética);
   - corrige espaços antes de pontuação e o ° usado no lugar de º.
   Nenhuma palavra é trocada. O que parecer errado no conteúdo vira aviso. */
(function (root) {
  'use strict';

  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  // ordem dos elementos exigida pelo formato do Word (fora de ordem, o Word reclama do arquivo)
  const ORDER = {
    rPr: ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'],
    pPr: ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'],
    tblPr: ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'],
    tcPr: ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'],
    trPr: ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange'],
    sectPr: ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'],
    tblCellMar: ['top', 'left', 'start', 'bottom', 'right', 'end'],
  };
  const FIRST_CHILD = { r: 'rPr', p: 'pPr', tbl: 'tblPr', tc: 'tcPr', tr: 'trPr' };

  const SYMBOL_FONT = /symbol|wingdings|webdings|zapf|dingbat/i;
  const SANS = /arial|helvet|calibri|roboto|aptos|verdana|tahoma|segoe|open sans|lato|montserrat|liberation sans|carlito|noto sans|trebuchet|century gothic|gill/i;
  const SERIF = /times|georgia|cambria|garamond|book antiqua|palatino|liberation serif|tinos|caladea|constantia|baskerville/i;

  // ---------------------------------------------------------------- DOM helpers
  const isW = (n, name) => n && n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name);
  const kids = (el, name) => (el ? Array.from(el.childNodes).filter((n) => isW(n, name)) : []);
  const kid = (el, name) => kids(el, name)[0] || null;
  const desc = (el, name) => (el ? Array.from(el.getElementsByTagNameNS(W, name)) : []);
  const val = (el, name) => (el ? el.getAttributeNS(W, name || 'val') : null);
  const num = (el, name) => { const v = val(el, name); return v == null || v === '' ? null : Number(v); };
  const setVal = (el, name, v) => el.setAttributeNS(W, 'w:' + name, String(v));
  const delAttr = (el, name) => { if (el.hasAttributeNS(W, name)) el.removeAttributeNS(W, name); };
  const mk = (doc, name) => doc.createElementNS(W, 'w:' + name);

  function insertOrdered(parent, el, order) {
    const i = order.indexOf(el.localName);
    for (const c of Array.from(parent.childNodes)) {
      if (c.nodeType !== 1) continue;
      // elementos de outros esquemas (w14:, w15:…) vêm sempre depois dos do Word
      if (c.namespaceURI !== W || order.indexOf(c.localName) > i) { parent.insertBefore(el, c); return el; }
    }
    parent.appendChild(el);
    return el;
  }
  function getOrMake(parent, name, order) {
    const found = kid(parent, name);
    if (found) return found;
    return insertOrdered(parent, mk(parent.ownerDocument, name), order);
  }
  function propsOf(el) {
    const pn = FIRST_CHILD[el.localName];
    let p = kid(el, pn);
    if (!p) {
      p = mk(el.ownerDocument, pn);
      const ex = el.localName === 'tr' ? kid(el, 'tblPrEx') : null;
      el.insertBefore(p, ex ? ex.nextSibling : el.firstChild);
    }
    return p;
  }
  function removeKids(parent, names) {
    let n = 0;
    for (const c of kids(parent)) if (names.indexOf(c.localName) >= 0) { parent.removeChild(c); n++; }
    return n;
  }

  // ---------------------------------------------------------------- texto
  function pText(p) {
    let s = '';
    const walk = (n) => {
      n.childNodes.forEach((c) => {
        if (c.nodeType !== 1) return;
        if (c.namespaceURI === W) {
          if (c.localName === 't') { s += c.textContent; return; }
          if (c.localName === 'tab') { s += '\t'; return; }
          if (c.localName === 'br' || c.localName === 'cr') { s += ' '; return; }
          if (c.localName === 'pPr' || c.localName === 'rPr' || c.localName === 'delText' || c.localName === 'instrText') return;
        }
        walk(c);
      });
    };
    walk(p);
    return s;
  }
  const hasDrawing = (p) => p.getElementsByTagNameNS('*', 'drawing').length + p.getElementsByTagNameNS('*', 'pict').length + p.getElementsByTagNameNS('*', 'object').length > 0;
  const isEmptyP = (p) => isW(p, 'p') && !pText(p).trim() && !hasDrawing(p);
  const isListItem = (p) => !!kid(kid(p, 'pPr'), 'numPr');

  const endsTerminal = (t) => /[.!?:;…]["'”’»)\]]*\s*$/.test(t);
  const startsLower = (t) => { const m = /^[\s"'“‘«(\[–—-]*([A-Za-zÀ-ÖØ-öø-ÿ])/.exec(t); return !!m && m[1] !== m[1].toUpperCase(); };
  const upperRatio = (t) => { const l = t.replace(/[^A-Za-zÀ-ÖØ-öø-ÿ]/g, ''); return l ? l.replace(/[^A-ZÀ-ÖØ-Þ]/g, '').length / l.length : 0; };
  const headingLike = (t) => { const s = t.trim(); return s.length <= 70 && (upperRatio(s) > 0.8 || /:\s*$/.test(s)); };
  const isMarker = (t) => /^\s*(\d{1,3}|[a-z])\s*[.)º°ª-]?\s*$/i.test(t);
  const REFS_LABEL = /^\s*refer[êe]ncias(\s+bibliogr[áa]ficas)?\s*:?\s*$/i;
  const looksLikeRef = (t) => /^[A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ'´`\- ]{1,60}[,.]/.test(t.trim());
  const shorten = (t, n) => { const s = t.replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; };
  const tail = (t, n) => { const s = t.replace(/\s+/g, ' ').trim(); return s.length > n ? '…' + s.slice(s.length - n + 1) : s; };

  function lastTextP(cell) { const ps = kids(cell, 'p').filter((p) => pText(p).trim()); return ps[ps.length - 1] || null; }
  function firstTextP(cell) { return kids(cell, 'p').find((p) => pText(p).trim()) || null; }

  // ---------------------------------------------------------------- edição de texto dentro do parágrafo
  // aplica substituições no texto do parágrafo inteiro, mesmo quando o trecho atravessa vários trechos de formatação
  function rewrite(p, re, rep) {
    const ts = desc(p, 't');
    if (!ts.length) return [];
    const full = ts.map((t) => t.textContent).join('');
    const edits = [];
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(full))) {
      const r = typeof rep === 'function' ? rep(m) : rep;
      if (r !== m[0]) {
        // trecho de exemplo, cortado em palavras inteiras: "curricular : Matemática → curricular: Matemática"
        let a = Math.max(0, m.index - 8), b = Math.min(full.length, m.index + m[0].length + 6);
        while (a > 0 && /\S/.test(full[a - 1])) a--;
        while (b < full.length && /\S/.test(full[b])) b++;
        const ctxBefore = full.slice(a, b).trim();
        const ctxAfter = (full.slice(a, m.index) + r + full.slice(m.index + m[0].length, b)).trim();
        edits.push({ start: m.index, end: m.index + m[0].length, rep: r, before: ctxBefore + ' → ' + ctxAfter });
      }
      if (m[0].length === 0) re.lastIndex++;
    }
    for (let k = edits.length - 1; k >= 0; k--) {
      const e = edits[k];
      let pos = 0, placed = false;
      for (const t of ts) {
        const s = t.textContent, a = pos, b = pos + s.length;
        pos = b;
        if (e.end <= a && !(e.start === e.end && e.start === a)) continue;
        if (e.start >= b && !(e.start === b && !placed && e.start === e.end)) continue;
        const ls = Math.max(0, e.start - a), le = Math.min(s.length, e.end - a);
        const ins = placed ? '' : e.rep;
        placed = true;
        t.textContent = s.slice(0, ls) + ins + s.slice(le);
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      }
    }
    return edits;
  }

  // ---------------------------------------------------------------- estilos e fontes
  function themeFonts(themeXml, parser) {
    if (!themeXml) return {};
    const d = parser.parseFromString(themeXml, 'application/xml');
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const get = (n) => { const el = d.getElementsByTagNameNS(A, n)[0]; const l = el && el.getElementsByTagNameNS(A, 'latin')[0]; return l ? l.getAttribute('typeface') : null; };
    return { minor: get('minorFont'), major: get('majorFont') };
  }
  function fontFromRFonts(rf, theme) {
    if (!rf) return null;
    const a = val(rf, 'ascii') || val(rf, 'hAnsi');
    if (a) return a;
    const th = val(rf, 'asciiTheme') || val(rf, 'hAnsiTheme');
    if (th) return /major/i.test(th) ? theme.major : theme.minor;
    return null;
  }
  function styleResolver(stylesDoc, theme) {
    const map = {};
    let def = null, defPara = null;
    if (stylesDoc) {
      const dd = desc(stylesDoc.documentElement, 'rPrDefault')[0];
      def = fontFromRFonts(kid(kid(dd, 'rPr'), 'rFonts'), theme);
      for (const st of desc(stylesDoc.documentElement, 'style')) {
        const id = val(st, 'styleId');
        map[id] = { font: fontFromRFonts(kid(kid(st, 'rPr'), 'rFonts'), theme), based: val(kid(st, 'basedOn')) };
        if (val(st, 'type') === 'paragraph' && val(st, 'default') === '1') defPara = id;
      }
    }
    const fontOf = (id, depth) => {
      if (!id || !map[id] || depth > 12) return null;
      return map[id].font || fontOf(map[id].based, (depth || 0) + 1);
    };
    return (pStyle) => fontOf(pStyle || defPara, 0) || fontOf(defPara, 0) || def || 'Times New Roman';
  }
  const cleanFontName = (f) => String(f || '').replace(/\s+(MT|PSMT|PS)$/i, '').replace(/,.*$/, '').trim();

  function setRunFont(rPr, font) {
    const rf = getOrMake(rPr, 'rFonts', ORDER.rPr);
    if (SYMBOL_FONT.test(val(rf, 'ascii') || '') || SYMBOL_FONT.test(val(rf, 'hAnsi') || '')) return false;
    ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'].forEach((a) => delAttr(rf, a));
    ['ascii', 'hAnsi', 'cs', 'eastAsia'].forEach((a) => setVal(rf, a, font));
    return true;
  }
  function setRunSize(rPr, halfPts) {
    setVal(getOrMake(rPr, 'sz', ORDER.rPr), 'val', halfPts);
    setVal(getOrMake(rPr, 'szCs', ORDER.rPr), 'val', halfPts);
  }

  // ---------------------------------------------------------------- tabelas
  function tableModel(tbl) {
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => num(g, 'w') || 0);
    const cum = [0];
    grid.forEach((w, i) => cum.push(cum[i] + w));
    const at = (c) => cum[Math.min(Math.max(c, 0), cum.length - 1)];
    const rows = kids(tbl, 'tr').map((tr) => {
      const trPr = kid(tr, 'trPr');
      const before = num(kid(trPr, 'gridBefore')) || 0;
      const after = num(kid(trPr, 'gridAfter')) || 0;
      let col = before;
      const cells = kids(tr, 'tc').map((tc) => {
        const span = num(kid(kid(tc, 'tcPr'), 'gridSpan')) || 1;
        const c = { tc, x0: at(col), x1: at(col + span) };
        col += span;
        return c;
      });
      return { tr, x0: at(0), xb: at(before), xa: at(col), xe: at(col + after), before, after, cells };
    });
    return { tbl, grid, cum, rows, width: cum[cum.length - 1] };
  }
  function clusterPoints(points, tol) {
    const s = points.slice().sort((a, b) => a - b);
    const out = [];
    for (const p of s) if (!out.length || p - out[out.length - 1] > tol) out.push(p);
    return out;
  }
  function nearestIdx(pts, x) {
    let best = 0, d = Infinity;
    pts.forEach((p, i) => { const e = Math.abs(p - x); if (e < d) { d = e; best = i; } });
    return best;
  }
  function setRowGrid(tr, name, n) {
    if (n > 0) { setVal(getOrMake(propsOf(tr), name, ORDER.trPr), 'val', n); return; }
    const trPr = kid(tr, 'trPr');
    const g = kid(trPr, name);
    if (g) trPr.removeChild(g);
  }
  function applyGrid(model, pts) {
    const idx = (x) => nearestIdx(pts, x);
    for (const r of model.rows) {
      setRowGrid(r.tr, 'gridBefore', idx(r.xb) - idx(r.x0));
      setRowGrid(r.tr, 'gridAfter', idx(r.xe) - idx(r.xa));
      for (const c of r.cells) {
        const span = Math.max(1, idx(c.x1) - idx(c.x0));
        const tcPr = propsOf(c.tc);
        if (span > 1) setVal(getOrMake(tcPr, 'gridSpan', ORDER.tcPr), 'val', span);
        else { const g = kid(tcPr, 'gridSpan'); if (g) tcPr.removeChild(g); }
      }
    }
  }
  function setGrid(tbl, pts) {
    const doc = tbl.ownerDocument;
    let tg = kid(tbl, 'tblGrid');
    if (!tg) { tg = mk(doc, 'tblGrid'); tbl.insertBefore(tg, kid(tbl, 'tr')); }
    while (tg.firstChild) tg.removeChild(tg.firstChild);
    for (let i = 1; i < pts.length; i++) {
      const g = mk(doc, 'gridCol');
      setVal(g, 'w', Math.max(1, Math.round(pts[i] - pts[i - 1])));
      tg.appendChild(g);
    }
  }
  const spansOf = (tr) => kids(tr, 'tc').map((tc) => num(kid(kid(tc, 'tcPr'), 'gridSpan')) || 1).join(',');

  function moveCellContent(from, to, afterNode) {
    const nodes = kids(from).filter((n) => n.localName !== 'tcPr');
    let ref = afterNode ? afterNode.nextSibling : null;
    for (const n of nodes) to.insertBefore(n, ref);
  }
  function trimTrailingEmpty(cell) {
    const ps = kids(cell, 'p');
    for (let k = ps.length - 1; k > 0; k--) { if (isEmptyP(ps[k])) cell.removeChild(ps[k]); else break; }
  }
  function spaceRun(p) {
    const doc = p.ownerDocument;
    const runs = desc(p, 'r');
    const r = mk(doc, 'r');
    const last = runs[runs.length - 1];
    const rp = last && kid(last, 'rPr');
    if (rp) r.appendChild(rp.cloneNode(true));
    const t = mk(doc, 't');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = ' ';
    r.appendChild(t);
    return r;
  }
  function joinParagraphs(p1, p2) {
    const t1 = pText(p1), t2 = pText(p2);
    if (!/\s$/.test(t1) && !/^\s/.test(t2) && !/[-‐]$/.test(t1)) p1.appendChild(spaceRun(p1));
    for (const n of Array.from(p2.childNodes)) if (!isW(n, 'pPr')) p1.appendChild(n);
    p2.parentNode.removeChild(p2);
  }

  // junta a última linha da tabela A com a primeira da tabela B quando a página cortou a linha no meio
  function joinBoundaryRows(rA, rB) {
    const ca = kids(rA, 'tc'), cb = kids(rB, 'tc');
    if (!ca.length || ca.length !== cb.length || spansOf(rA) !== spansOf(rB)) return null;
    const pairs = ca.map((a, k) => {
      const pa = lastTextP(a), pb = firstTextP(cb[k]);
      return { a, b: cb[k], pa, pb, ta: pa ? pText(pa).trim() : '', tb: pb ? pText(pb).trim() : '' };
    });
    if (pairs.some((x) => x.ta && x.tb && isMarker(x.ta) && isMarker(x.tb) && x.ta !== x.tb)) return null;
    const contText = (x) => x.tb && (startsLower(x.tb) || (x.ta.length >= 25 && !endsTerminal(x.ta) && !headingLike(x.ta)));
    const contRefs = (x) => x.tb && looksLikeRef(x.tb) && kids(x.a, 'p').some((p) => REFS_LABEL.test(pText(p)) || /^\s*refer[êe]ncias\s*:/i.test(pText(p)));
    if (!pairs.some((x) => contText(x) || contRefs(x))) return null;
    let example = null;
    for (const x of pairs) {
      if (!x.tb) continue;
      trimTrailingEmpty(x.a);
      if (contText(x) && x.pa) {
        example = example || tail(x.ta, 34) + ' | ' + shorten(x.tb, 34);
        // move o que vem antes do 1º parágrafo com texto (vazios) para fora, junta o 1º, e traz o resto
        const bNodes = kids(x.b).filter((n) => n.localName !== 'tcPr');
        const startIdx = bNodes.indexOf(x.pb);
        const rest = bNodes.slice(startIdx + 1);
        joinParagraphs(x.pa, x.pb);
        let ref = x.pa.nextSibling;
        for (const n of rest) x.a.insertBefore(n, ref);
      } else {
        example = example || tail(x.ta || '', 30) + ' | ' + shorten(x.tb, 30);
        const bNodes = kids(x.b).filter((n) => n.localName !== 'tcPr');
        let k = 0;
        while (k < bNodes.length && isEmptyP(bNodes[k])) k++;
        for (const n of bNodes.slice(k)) x.a.appendChild(n);
      }
    }
    rB.parentNode.removeChild(rB);
    return example || 'linha da tabela';
  }

  function mergeTables(A, B, report) {
    const mA = tableModel(A), mB = tableModel(B);
    const tol = Math.max(60, Math.round(Math.max(mA.width, mB.width) * 0.01));
    const pts = clusterPoints(mA.cum.concat(mB.cum), tol);
    applyGrid(mA, pts);
    applyGrid(mB, pts);
    setGrid(A, pts);
    const lastA = kids(A, 'tr').pop();
    const rowsB = kids(B, 'tr');
    for (const tr of rowsB) A.appendChild(tr);
    B.parentNode.removeChild(B);
    report.tablesMerged++;
    if (lastA && rowsB[0]) {
      const ex = joinBoundaryRows(lastA, rowsB[0]);
      if (ex) { report.splitJoins++; report.splitExamples.push(ex); report.joinedRows.add(lastA); }
    }
  }

  // ---------------------------------------------------------------- processo principal
  function readZip(bytes) {
    const ff = root.fflate;
    if (!ff) throw new Error('fflate não carregou');
    return ff.unzipSync(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  }

  function process(bytes, settings, env) {
    settings = settings || {};
    env = env || {};
    const Parser = env.DOMParser || root.DOMParser;
    const Serializer = env.XMLSerializer || root.XMLSerializer;
    const ff = root.fflate;
    const files = readZip(bytes);
    const dec = (k) => (files[k] ? ff.strFromU8(files[k]) : null);
    const parser = new Parser();
    const docXml = dec('word/document.xml');
    if (!docXml) throw new Error('Este arquivo não parece ser um .docx do Word.');
    const doc = parser.parseFromString(docXml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Não consegui ler o conteúdo deste .docx.');
    const stylesXml = dec('word/styles.xml');
    const stylesDoc = stylesXml ? parser.parseFromString(stylesXml, 'application/xml') : null;
    const theme = themeFonts(dec('word/theme/theme1.xml'), parser);
    const resolve = styleResolver(stylesDoc, theme);
    const body = kid(doc.documentElement, 'body');

    const R = {
      tablesMerged: 0, splitJoins: 0, splitExamples: [], breaksRemoved: 0,
      lineJoins: 0, lineExamples: [], compressedRuns: 0, paraFixed: 0,
      refs: 0, refsSorted: false, spaces: 0, spaceExamples: [], ordinals: 0, ordinalExamples: [],
      cellPadding: 0, rowHeights: 0, keepRows: 0, joinedRows: new Set(), fontsBefore: {}, warnings: [],
    };

    // ---------- estatísticas (antes de mudar)
    const allText = pText(body).replace(/\s+/g, '');
    const inTables = kids(body, 'tbl').reduce((a, t) => a + pText(t).replace(/\s+/g, '').length, 0);
    for (const r of desc(body, 'r')) {
      const t = desc(r, 't').map((x) => x.textContent).join('');
      if (!t.trim()) continue;
      const p = r.parentNode && (isW(r.parentNode, 'p') ? r.parentNode : r.parentNode.parentNode);
      const f = cleanFontName(fontFromRFonts(kid(kid(r, 'rPr'), 'rFonts'), theme) || resolve(val(kid(kid(p, 'pPr'), 'pStyle'))));
      if (SYMBOL_FONT.test(f)) continue;
      R.fontsBefore[f] = (R.fontsBefore[f] || 0) + t.replace(/\s+/g, '').length;
    }
    const sizesBefore = {};
    for (const s of desc(body, 'sz')) { const v = num(s); if (v) sizesBefore[v / 2] = 1; }
    const finalSect = kid(body, 'sectPr');
    const pgSz = kid(finalSect, 'pgSz');
    const landscape = !!pgSz && (val(pgSz, 'orient') === 'landscape' || (num(pgSz, 'w') || 0) > (num(pgSz, 'h') || 0));
    const fromPdf = /TableParagraph/.test(stylesXml || '') || /"Arial MT"|ArialMT|"TimesNewRomanPSMT"/.test((stylesXml || '') + docXml);
    const tables = kids(body, 'tbl').length;
    const ratio = allText.length ? inTables / allText.length : 0;
    const fonts = Object.keys(R.fontsBefore).sort((a, b) => R.fontsBefore[b] - R.fontsBefore[a]);
    const dominant = fonts[0] || '';
    const suggestedFont = SERIF.test(dominant) ? 'Times New Roman' : SANS.test(dominant) ? 'Arial' : 'Times New Roman';
    const font = settings.fonte || suggestedFont;
    const size = settings.tamanho || 12;
    const lineMult = Number(settings.entrelinha) || 1;
    const LINE = Math.round(240 * lineMult);
    const fixTypos = settings.corrigirDigitacao !== false;
    const sortRefs = settings.ordenarReferencias !== false;

    // ---------- 1. quebras de seção/página de cada página (mesmo tamanho de folha) viram uma só
    const seps = new Set();
    const inner = desc(body, 'sectPr').filter((s) => s !== finalSect && isW(s.parentNode, 'pPr'));
    if (inner.length && finalSect) {
      const sig = (s) => { const z = kid(s, 'pgSz'), m = kid(s, 'pgMar'); return [num(z, 'w'), num(z, 'h'), num(m, 'top'), num(m, 'bottom'), num(m, 'left'), num(m, 'right')]; };
      const base = sig(finalSect);
      const same = inner.every((s) => sig(s).every((v, i) => Math.abs((v || 0) - (base[i] || 0)) <= 20));
      if (same) {
        for (const kind of ['headerReference', 'footerReference']) {
          if (kids(finalSect, kind).length) continue;
          let src = null;
          for (const s of inner) if (kids(s, kind).length) src = s;
          if (src) kids(src, kind).forEach((ref) => insertOrdered(finalSect, ref.cloneNode(true), ORDER.sectPr));
        }
        for (const name of ['pgNumType', 'titlePg']) {
          if (kid(finalSect, name)) continue;
          const src = inner.find((s) => kid(s, name));
          if (src) insertOrdered(finalSect, kid(src, name).cloneNode(true), ORDER.sectPr);
        }
        for (const s of inner) {
          const p = s.parentNode.parentNode;
          s.parentNode.removeChild(s);
          seps.add(p);
          R.breaksRemoved++;
        }
      }
    }
    // quebras de página manuais em parágrafos vazios
    for (const br of desc(body, 'br')) {
      if (val(br, 'type') !== 'page') continue;
      const p = br.parentNode && br.parentNode.parentNode;
      if (!isW(p, 'p') || pText(p).trim()) continue;
      br.parentNode.removeChild(br);
      seps.add(p);
      R.breaksRemoved++;
    }

    // ---------- 2. tabelas cortadas entre páginas viram uma só
    let again = true;
    while (again) {
      again = false;
      const ch = kids(body);
      for (let i = 0; i < ch.length && !again; i++) {
        if (ch[i].localName !== 'tbl') continue;
        let j = i + 1, hasSep = false;
        const between = [];
        while (j < ch.length && isEmptyP(ch[j])) { between.push(ch[j]); if (seps.has(ch[j])) hasSep = true; j++; }
        if (!hasSep || j >= ch.length || ch[j].localName !== 'tbl') continue;
        if (kid(kid(ch[i], 'tblPr'), 'tblpPr') || kid(kid(ch[j], 'tblPr'), 'tblpPr')) continue; // tabelas flutuantes: não mexe
        between.forEach((p) => { p.parentNode.removeChild(p); seps.delete(p); });
        mergeTables(ch[i], ch[j], R);
        again = true;
      }
    }

    // ---------- 3. quebras de linha no meio da frase (parágrafos que continuam no seguinte)
    const containers = [body].concat(desc(body, 'tc'));
    for (const c of containers) {
      let list = kids(c);
      for (let k = 0; k < list.length - 1; k++) {
        const p1 = list[k], p2 = list[k + 1];
        if (!isW(p1, 'p') || !isW(p2, 'p')) continue;
        const t1 = pText(p1), t2 = pText(p2);
        if (!t1.trim() || !t2.trim() || hasDrawing(p1) || hasDrawing(p2)) continue;
        if (t1.trim().length < 40 || isListItem(p2) || endsTerminal(t1) || !startsLower(t2)) continue;
        if (/heading|t[íi]tulo/i.test(val(kid(kid(p1, 'pPr'), 'pStyle')) || '')) continue;
        R.lineJoins++;
        if (R.lineExamples.length < 4) R.lineExamples.push(tail(t1, 26) + ' / ' + shorten(t2, 26));
        joinParagraphs(p1, p2);
        list = kids(c);
        k--;
      }
    }

    // ---------- 4. referências
    const refParas = new Set();
    for (const c of containers) {
      const list = kids(c);
      for (let k = 0; k < list.length; k++) {
        const p = list[k];
        if (!isW(p, 'p') || !REFS_LABEL.test(pText(p))) continue;
        const group = [];
        let m = k + 1;
        for (; m < list.length; m++) {
          const q = list[m];
          if (!isW(q, 'p')) break;
          const t = pText(q).trim();
          if (!t) { group.push({ p: q, empty: true }); continue; }
          if (headingLike(t) && !looksLikeRef(t)) break;
          group.push({ p: q, t });
        }
        const refs = group.filter((g) => !g.empty);
        if (!refs.length) continue;
        group.filter((g) => g.empty).forEach((g) => g.p.parentNode && g.p.parentNode.removeChild(g.p));
        if (sortRefs && refs.length > 1) {
          const key = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w\s]/g, '').toLowerCase();
          const sorted = refs.slice().sort((a, b) => key(a.t).localeCompare(key(b.t), 'pt-BR'));
          if (sorted.some((g, i) => g !== refs[i])) {
            R.refsSorted = true;
            const anchor = refs[refs.length - 1].p.nextSibling;
            const parent = refs[0].p.parentNode;
            sorted.forEach((g) => parent.insertBefore(g.p, anchor));
          }
        }
        refs.forEach((g) => { refParas.add(g.p); R.refs++; });
      }
    }

    // ---------- 5. letras, fontes e tamanhos
    for (const r of desc(body, 'r')) {
      const rPr = propsOf(r);
      if (removeKids(rPr, ['spacing', 'w', 'kern', 'position'])) R.compressedRuns++;
      setRunFont(rPr, font);
      setRunSize(rPr, size * 2);
    }

    // ---------- 6. parágrafos: espaçamento, recuos e alinhamento
    for (const p of desc(body, 'p')) {
      const pPr = propsOf(p);
      const t = pText(p).trim();
      const mark = getOrMake(pPr, 'rPr', ORDER.pPr);
      removeKids(mark, ['spacing', 'w', 'kern', 'position']);
      setRunFont(mark, font);
      setRunSize(mark, size * 2);
      const sp = getOrMake(pPr, 'spacing', ORDER.pPr);
      const before0 = num(sp, 'before') || 0, after0 = num(sp, 'after') || 0, line0 = num(sp, 'line');
      let changed = false;
      const isRef = refParas.has(p);
      const nb = isRef ? 120 : before0 <= 120 ? 0 : before0;
      const na = isRef ? 120 : after0 <= 120 ? 0 : after0;
      ['beforeAutospacing', 'afterAutospacing', 'beforeLines', 'afterLines'].forEach((a) => { if (sp.hasAttributeNS(W, a)) { delAttr(sp, a); changed = true; } });
      if (nb !== before0 || na !== after0 || line0 !== (isRef ? 240 : LINE) || val(sp, 'lineRule') !== 'auto') changed = true;
      setVal(sp, 'before', nb);
      setVal(sp, 'after', na);
      setVal(sp, 'line', isRef ? 240 : LINE);
      setVal(sp, 'lineRule', 'auto');
      const ind = kid(pPr, 'ind');
      if (ind && (isRef || !isListItem(p))) {
        for (const a of ['left', 'start', 'right', 'end', 'firstLine', 'hanging']) {
          const v = num(ind, a);
          if (v != null && (isRef || Math.abs(v) <= 200)) { delAttr(ind, a); changed = true; }
        }
        if (!ind.attributes.length) pPr.removeChild(ind);
      }
      const jc = kid(pPr, 'jc');
      const align = val(jc);
      if (isRef) {
        if (align !== 'left' && align !== 'start') { setVal(getOrMake(pPr, 'jc', ORDER.pPr), 'val', 'left'); changed = true; }
      } else if (t.length >= 80 && align !== 'center' && align !== 'right' && align !== 'end' && align !== 'both') {
        setVal(getOrMake(pPr, 'jc', ORDER.pPr), 'val', 'both');
        changed = true;
      }
      if (changed && t) R.paraFixed++;
    }

    // ---------- 7. margem interna das células, igual em todas as tabelas
    for (const tbl of desc(body, 'tbl')) {
      const tblPr = propsOf(tbl);
      const mar = getOrMake(tblPr, 'tblCellMar', ORDER.tblPr);
      while (mar.firstChild) mar.removeChild(mar.firstChild);
      [['top', 28], ['left', 108], ['bottom', 28], ['right', 108]].forEach(([side, w]) => {
        const e = mk(doc, side);
        setVal(e, 'w', w);
        setVal(e, 'type', 'dxa');
        mar.appendChild(e);
      });
      for (const tc of desc(tbl, 'tc')) { const tcPr = kid(tc, 'tcPr'); if (tcPr) removeKids(tcPr, ['tcMar']); }
      R.cellPadding++;
      // linhas curtas não se dividem entre páginas (no Word, "manter com o próximo" dentro de tabela prenderia as linhas umas às outras)
      const pageH = (num(pgSz, 'h') || 16838) - (num(kid(finalSect, 'pgMar'), 'top') || 1440) - (num(kid(finalSect, 'pgMar'), 'bottom') || 1440);
      const model = tableModel(tbl);
      for (const r of model.rows) {
        let rowH = 0;
        for (const c of r.cells) {
          const cw = Math.max(400, c.x1 - c.x0 - 216);
          const perLine = Math.max(8, Math.floor(cw / (size * 20 * 0.5)));
          let lines = 0;
          for (const p of kids(c.tc, 'p')) lines += Math.max(1, Math.ceil(pText(p).length / perLine));
          rowH = Math.max(rowH, lines * size * 20 * 1.2 * lineMult);
        }
        if (rowH > 0 && rowH < pageH * 0.35) { getOrMake(propsOf(r.tr), 'cantSplit', ORDER.trPr); R.keepRows++; }
      }
      // altura fixa das linhas: no arquivo convertido de PDF ela só copia o desenho antigo
      // e deixa buracos (ou corta texto). Linhas com texto passam a crescer conforme o conteúdo.
      for (const tr of kids(tbl, 'tr')) {
        const trPr = kid(tr, 'trPr');
        const h = kid(trPr, 'trHeight');
        if (!h || !pText(tr).trim()) continue;
        if (fromPdf || R.joinedRows.has(tr) || val(h, 'hRule') === 'exact') { trPr.removeChild(h); R.rowHeights++; }
      }
    }

    // ---------- 8. digitação: espaço antes de pontuação, espaços duplos, ° no lugar de º
    if (fixTypos) {
      for (const p of desc(body, 'p')) {
        const before = pText(p);
        if (!before.trim()) continue;
        const e1 = rewrite(p, /[  ]+(?=[,;:!?)]|\.(?!\.))/g, '');
        const e2 = rewrite(p, /\( +/g, '(');
        const e3 = rewrite(p, /(\S) {2,}(?=\S)/g, (m) => m[1] + ' ');
        const n = e1.length + e2.length + e3.length;
        if (n) {
          R.spaces += n;
          const ex = e1[0] || e2[0] || e3[0];
          if (R.spaceExamples.length < 3 && ex) R.spaceExamples.push(shorten(ex.before, 90));
        }
        const e4 = rewrite(p, /(\d)[°˚](?=\s*(ano|anos|série|serie|período|periodo|semestre|bimestre|trimestre|etapa|ciclo|lugar|grau)\b)/gi, (m) => m[1] + (/^(s[ée]rie|etapa)$/i.test(m[2]) ? 'ª' : 'º'));
        if (e4.length) { R.ordinals += e4.length; if (R.ordinalExamples.length < 2) R.ordinalExamples.push(shorten(e4[0].before, 60)); }
      }
    }

    // ---------- 9. estilos: fonte padrão do documento (o que não tem fonte própria)
    if (stylesDoc) {
      const dd = desc(stylesDoc.documentElement, 'rPrDefault')[0];
      if (dd) {
        const rPr = kid(dd, 'rPr') || dd.appendChild(mk(stylesDoc, 'rPr'));
        setRunFont(rPr, font);
        setRunSize(rPr, size * 2);
      }
      for (const st of desc(stylesDoc.documentElement, 'style')) {
        const rPr = kid(st, 'rPr');
        if (rPr && kid(rPr, 'rFonts')) setRunFont(rPr, font);
        if (rPr) removeKids(rPr, ['spacing', 'w', 'kern', 'position']);
      }
    }

    // ---------- 10. avisos sobre o conteúdo (nada é mudado)
    const seen = new Set();
    const warn = (kind, text, msg) => { const k = kind + text; if (seen.has(k)) return; seen.add(k); R.warnings.push({ kind, msg, text }); };
    for (const p of desc(body, 'p')) {
      const t = pText(p).replace(/\s+/g, ' ').trim();
      if (t.length < 12 || headingLike(t)) continue;
      if (/\b(de|da|do|das|dos|e|a|o|as|os|em|no|na|nos|nas|com|para|por|pela|pelo|que|um|uma|ao|à|aos|às|sem|sob|entre)$/i.test(t)) {
        warn('cortado', tail(t, 60), 'Parece terminar no meio da frase. Confira se faltou um pedaço.');
      } else if (/,$/.test(t)) {
        warn('virgula', tail(t, 60), 'Termina com vírgula.');
      }
    }

    // ---------- relatório
    const fontList = fonts.filter((f) => R.fontsBefore[f] > 0);
    const fixes = [];
    if (R.splitJoins || R.tablesMerged) fixes.push({ key: 'split', label: 'Texto cortado entre páginas, juntado de volta', count: R.splitJoins || R.tablesMerged, examples: R.splitExamples.slice(0, 4) });
    if (R.lineJoins) fixes.push({ key: 'lines', label: 'Quebras de linha no meio da frase, corrigidas', count: R.lineJoins, examples: R.lineExamples });
    const sizesList = Object.keys(sizesBefore).map(Number).sort((a, b) => a - b);
    const fontChange = fontList.length > 1 || (fontList[0] && fontList[0] !== font) || sizesList.length > 1 || (sizesList[0] && sizesList[0] !== size);
    if (fontChange) fixes.push({ key: 'font', label: 'Fonte unificada', detail: (fontList.join(', ') || '—') + (sizesList.length ? ' (' + sizesList.join(', ') + ' pt)' : '') + ' → ' + font + ' ' + size + ' pt' });
    if (R.compressedRuns) fixes.push({ key: 'compressed', label: 'Letras comprimidas (resto de conversão de PDF), normalizadas', count: R.compressedRuns });
    if (R.paraFixed) fixes.push({ key: 'para', label: 'Espaçamento e recuos dos parágrafos padronizados', count: R.paraFixed, detail: 'entrelinha ' + (lineMult === 1 ? 'simples' : String(lineMult).replace('.', ',')) });
    if (R.cellPadding) fixes.push({ key: 'cells', label: 'Margem interna das células igual em todas as tabelas', count: R.cellPadding });
    if (R.rowHeights) fixes.push({ key: 'rows', label: 'Linhas da tabela com altura ajustada ao texto', count: R.rowHeights });
    if (R.refs) fixes.push({ key: 'refs', label: 'Referências à esquerda, espaço simples, 6 pt antes e depois' + (R.refsSorted ? ', em ordem alfabética' : ''), count: R.refs });
    if (R.spaces) fixes.push({ key: 'spaces', label: 'Espaços sobrando antes de pontuação, removidos', count: R.spaces, examples: R.spaceExamples });
    if (R.ordinals) fixes.push({ key: 'ordinal', label: 'Ordinal corrigido: sinal de grau (°) trocado por º', count: R.ordinals, examples: R.ordinalExamples });

    const report = {
      fixes,
      warnings: R.warnings,
      stats: {
        tables, landscape, fromPdf, textInTables: Math.round(ratio * 100), fonts: fontList, sizes: sizesList,
        suggestedFont, font, size, lineMult, sections: inner.length + 1, chars: allText.length,
        templateLike: (tables >= 1 && ratio >= 0.6) || (landscape && tables >= 1),
      },
    };
    if (settings.dryRun) return { report };

    // ---------- grava o .docx (todo o resto do pacote fica igual)
    const ser = new Serializer();
    const out = (d) => { let s = ser.serializeToString(d); if (!/^<\?xml/.test(s)) s = XML_DECL + s; return ff.strToU8(s); };
    files['word/document.xml'] = out(doc);
    if (stylesDoc) files['word/styles.xml'] = out(stylesDoc);
    const zipped = ff.zipSync(files, { level: 6 });
    return { bytes: zipped, report };
  }

  // só a análise (para mostrar na tela o que será ajustado)
  function analyze(bytes, settings, env) {
    return process(bytes, Object.assign({}, settings, { dryRun: true }), env).report;
  }

  root.FormataModelo = { process, analyze };
})(typeof window !== 'undefined' ? window : globalThis);
