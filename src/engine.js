/* Formata ABNT — motor de formatação.
   Não usa IA: lê o texto, pré-classifica cada parágrafo por regras simples
   e monta o .docx aplicando as regras do perfil escolhido. */
(function (root) {
  'use strict';

  // ------------------------------------------------------------------
  // PERFIS — cada faculdade/curso vira um objeto destes
  // ------------------------------------------------------------------
  const PROFILES = {
    // Regras da Biblioteca Prof. Lúcio de Sousa (UNICID), "Normas para elaboração
    // de artigo científico" (2020), completadas pela ABNT NBR 14724:2024
    'unicid-pedagogia': {
      id: 'unicid-pedagogia',
      nome: 'UNICID · Pedagogia',
      norma: 'Normas da Biblioteca UNICID + ABNT NBR 14724:2024',
      instituicao: 'Universidade Cidade de São Paulo',
      curso: 'Pedagogia',
      cidade: 'São Paulo',
      natureza:
        'Trabalho de Conclusão de Curso apresentado ao Curso de Pedagogia da Universidade Cidade de São Paulo, como requisito parcial para obtenção do título de Licenciado(a) em Pedagogia.',
      fonte: 'Times New Roman', // o manual aceita Arial ou Times; os exemplos e as referências usam Times
      tamanho: 12, // pt
      tamanhoMenor: 10, // citações longas, notas, legendas, fontes, paginação
      margens: { sup: 3, esq: 3, inf: 2, dir: 2 }, // cm
      entrelinha: 1.5,
      recuoParagrafo: 2, // cm — recomendação do manual UNICID
      recuoCitacao: 4, // cm
      recuoNatureza: 8, // cm
      refEspaco: { antes: 6, depois: 6 }, // pt — manual UNICID
      palavrasChave: 'ponto', // "separadas por ponto, seguido de inicial maiúscula"
      titulos: {
        1: { bold: true, caps: true },
        2: { bold: false, caps: true },
        3: { bold: true, caps: false },
        4: { bold: false, caps: false },
        5: { bold: false, caps: false, italic: true },
      },
      ordenarReferencias: true,
    },
  };

  // Tipos que o usuário pode escolher para cada parágrafo
  const TYPES = {
    h1: { label: 'Seção primária', short: 'Título 1', group: 'titulo' },
    h2: { label: 'Seção secundária', short: 'Título 2', group: 'titulo' },
    h3: { label: 'Seção terciária', short: 'Título 3', group: 'titulo' },
    h4: { label: 'Seção quaternária', short: 'Título 4', group: 'titulo' },
    h5: { label: 'Seção quinária', short: 'Título 5', group: 'titulo' },
    post: { label: 'Título sem número', short: 'Sem nº', group: 'titulo', hint: 'Referências, Apêndice, Anexo' },
    p: { label: 'Parágrafo', short: 'Texto', group: 'texto' },
    quote: { label: 'Citação longa', short: 'Citação', group: 'citacao', hint: 'Mais de 3 linhas, recuo de 4 cm' },
    item: { label: 'Alínea (item de lista)', short: 'Alínea', group: 'texto' },
    caption: { label: 'Legenda', short: 'Legenda', group: 'ilustra', hint: 'Figura 1 – …, Quadro 1 – …' },
    source: { label: 'Fonte da ilustração', short: 'Fonte', group: 'ilustra', hint: 'Fonte: …' },
    ref: { label: 'Referência', short: 'Ref.', group: 'ref' },
    skip: { label: 'Ignorar', short: 'Ignorado', group: 'skip', hint: 'Capa, sumário antigo, etc.' },
    img: { label: 'Imagem', short: 'Imagem', group: 'ilustra' },
    table: { label: 'Tabela/Quadro', short: 'Tabela', group: 'ilustra' },
  };

  const cm = (v) => Math.round(v * 566.929); // cm → twips

  // ------------------------------------------------------------------
  // LEITURA: HTML (do .docx via mammoth, ou colado) → blocos
  // ------------------------------------------------------------------
  const BLOCK = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'TABLE', 'DIV', 'SECTION', 'ARTICLE', 'UL', 'OL', 'HEADER', 'FOOTER', 'MAIN', 'PRE', 'FIGURE', 'FIGCAPTION', 'HR', 'DL', 'DT', 'DD', 'ADDRESS', 'ASIDE', 'NAV']);
  let uid = 0;
  const nextId = () => 'b' + ++uid;

  function readFmt(el, fmt) {
    const f = Object.assign({}, fmt);
    const tag = el.tagName;
    if (tag === 'B' || tag === 'STRONG') f.b = true;
    if (tag === 'I' || tag === 'EM' || tag === 'CITE') f.i = true;
    if (tag === 'U') f.u = true;
    if (tag === 'SUP') f.sup = true;
    if (tag === 'SUB') f.sub = true;
    const st = el.getAttribute && el.getAttribute('style');
    if (st) {
      const fw = /font-weight\s*:\s*([^;]+)/i.exec(st);
      if (fw) {
        const v = fw[1].trim().toLowerCase();
        if (v === 'bold' || v === 'bolder' || parseInt(v, 10) >= 600) f.b = true;
        else if (v === 'normal' || v === 'lighter' || parseInt(v, 10) < 600) f.b = false;
      }
      const fs = /font-style\s*:\s*([^;]+)/i.exec(st);
      if (fs) f.i = /italic|oblique/i.test(fs[1]);
      if (/text-decoration[^;]*underline/i.test(st)) f.u = true;
      const va = /vertical-align\s*:\s*([^;]+)/i.exec(st);
      if (va) {
        if (/super/i.test(va[1])) f.sup = true;
        if (/sub/i.test(va[1])) f.sub = true;
      }
    }
    return f;
  }

  function sameFmt(a, b) {
    return !!a.b === !!b.b && !!a.i === !!b.i && !!a.u === !!b.u && !!a.sup === !!b.sup && !!a.sub === !!b.sub && !a.fn && !b.fn;
  }

  function finishRuns(runs) {
    // junta trechos com o mesmo estilo, normaliza espaços e apara as pontas
    const out = [];
    for (const r of runs) {
      if (r.fn) { out.push(r); continue; }
      const t = r.t.replace(/[ \t\r\n ]+/g, ' ');
      if (!t) continue;
      const last = out[out.length - 1];
      if (last && !last.fn && sameFmt(last, r)) last.t += t;
      else out.push(Object.assign({}, r, { t }));
    }
    while (out.length && !out[0].fn && !out[0].t.trim()) out.shift();
    while (out.length && !out[out.length - 1].fn && !out[out.length - 1].t.trim()) out.pop();
    if (out.length && !out[0].fn) out[0].t = out[0].t.replace(/^\s+/, '');
    const l = out[out.length - 1];
    if (l && !l.fn) l.t = l.t.replace(/\s+$/, '');
    // espaços duplos entre trechos
    for (let i = 1; i < out.length; i++) {
      const p = out[i - 1], c = out[i];
      if (!p.fn && !c.fn && /\s$/.test(p.t) && /^\s/.test(c.t)) c.t = c.t.replace(/^\s+/, '');
    }
    return out;
  }

  function runsText(runs) {
    return runs.map((r) => (r.fn ? '' : r.t)).join('');
  }

  function htmlToBlocks(html, parser) {
    const doc = parser.parseFromString(html, 'text/html');
    const body = doc.body;
    const notes = {}; // id → runs
    const blocks = [];

    // notas de rodapé geradas pelo mammoth
    body.querySelectorAll('li[id^="footnote-"], li[id^="endnote-"]').forEach((li) => {
      li.querySelectorAll('a[href^="#footnote-ref-"], a[href^="#endnote-ref-"]').forEach((a) => a.remove());
      const runs = [];
      collectInline(li, {}, runs);
      notes[li.id] = finishRuns(runs);
      const ol = li.parentElement;
      li.remove();
      if (ol && !ol.children.length) ol.remove();
    });

    function collectInline(node, fmt, runs) {
      node.childNodes.forEach((n) => {
        if (n.nodeType === 3) runs.push(Object.assign({ t: n.nodeValue }, fmt));
        else if (n.nodeType === 1) {
          if (n.tagName === 'BR') runs.push(Object.assign({ t: ' ' }, fmt));
          else collectInline(n, readFmt(n, fmt), runs);
        }
      });
    }

    let cur = null;
    function flush() {
      if (!cur) return;
      const runs = finishRuns(cur.runs);
      const text = runsText(runs).trim();
      if (text || runs.some((r) => r.fn)) {
        blocks.push({ id: nextId(), kind: 'p', meta: cur.meta, runs, text });
      }
      cur = null;
    }
    function open(meta) {
      flush();
      cur = { meta: meta || {}, runs: [] };
    }
    function push(run, ctx) {
      if (!cur) cur = { meta: { quote: ctx.quote, list: ctx.list }, runs: [] };
      cur.runs.push(run);
    }

    function walk(node, ctx, fmt) {
      node.childNodes.forEach((n) => {
        if (n.nodeType === 3) {
          if (cur || n.nodeValue.trim()) push(Object.assign({ t: n.nodeValue }, fmt), ctx);
          return;
        }
        if (n.nodeType !== 1) return;
        const tag = n.tagName;
        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'META' || tag === 'TITLE') return;
        if (tag === 'IMG') {
          const src = n.getAttribute('src') || '';
          if (/^data:image\//i.test(src)) {
            const meta = cur ? cur.meta : { quote: ctx.quote, list: ctx.list };
            flush();
            blocks.push({ id: nextId(), kind: 'img', src, alt: n.getAttribute('alt') || '', text: n.getAttribute('alt') || 'Imagem' });
            cur = { meta, runs: [] };
          }
          return;
        }
        if (tag === 'BR') { push(Object.assign({ t: ' ' }, fmt), ctx); return; }
        if (tag === 'A') {
          const href = n.getAttribute('href') || '';
          const m = /^#((?:foot|end)note-\d+)$/.exec(href);
          if (m && notes[m[1]]) { push({ fn: m[1], t: '' }, ctx); return; }
        }
        if (!BLOCK.has(tag)) {
          walk(n, ctx, readFmt(n, fmt));
          return;
        }
        // elementos de bloco
        if (tag === 'HR') { flush(); return; }
        if (tag === 'TABLE') { flush(); blocks.push(readTable(n)); return; }
        if (tag === 'UL' || tag === 'OL') {
          flush();
          const depth = ctx.list ? ctx.list.depth + 1 : 0;
          let idx = 0;
          n.childNodes.forEach((li) => {
            if (li.nodeType !== 1) return;
            if (li.tagName === 'LI') {
              idx++;
              const lctx = Object.assign({}, ctx, { list: { ordered: tag === 'OL', index: idx, depth } });
              open({ quote: ctx.quote, list: lctx.list });
              walk(li, lctx, fmt);
              flush();
            } else walk(li, ctx, fmt);
          });
          return;
        }
        if (tag === 'BLOCKQUOTE') {
          flush();
          walk(n, Object.assign({}, ctx, { quote: true }), fmt);
          flush();
          return;
        }
        if (/^(P|H[1-6]|DT|DD|PRE|FIGCAPTION|ADDRESS)$/.test(tag)) {
          const st = n.getAttribute('style') || '';
          const al = /text-align\s*:\s*(\w+)/i.exec(st);
          const centered = (n.getAttribute('class') || '').indexOf('abnt-center') >= 0;
          open({ tag: tag.toLowerCase(), quote: ctx.quote, list: ctx.list, align: centered ? 'center' : al ? al[1].toLowerCase() : undefined });
          walk(n, ctx, readFmt(n, fmt));
          flush();
          return;
        }
        // DIV, SECTION, LI solto etc.
        flush();
        walk(n, ctx, readFmt(n, fmt));
        flush();
      });
    }

    function readTable(table) {
      const rows = [];
      table.querySelectorAll('tr').forEach((tr) => {
        if (tr.closest('table') !== table) return;
        const cells = [];
        tr.childNodes.forEach((td) => {
          if (td.nodeType !== 1 || !/^(TD|TH)$/.test(td.tagName)) return;
          const paras = [];
          const ps = td.querySelectorAll('p');
          if (ps.length) ps.forEach((p) => { const r = []; collectInline(p, td.tagName === 'TH' ? { b: true } : {}, r); paras.push(finishRuns(r)); });
          else { const r = []; collectInline(td, td.tagName === 'TH' ? { b: true } : {}, r); paras.push(finishRuns(r)); }
          cells.push({ paras, colspan: parseInt(td.getAttribute('colspan') || '1', 10) || 1 });
        });
        if (cells.length) rows.push(cells);
      });
      const first = rows[0] ? rows[0].map((c) => c.paras.map(runsText).join(' ')).join(' | ') : '';
      return { id: nextId(), kind: 'table', rows, text: first || 'Tabela' };
    }

    walk(body, { quote: false, list: null }, {});
    flush();
    return { blocks, notes };
  }

  // opções do mammoth: detecta citações recuadas e parágrafos centralizados
  function mammothOptions(mammoth) {
    return {
      styleMap: [
        "p[style-name='Quote'] => blockquote:fresh",
        "p[style-name='Intense Quote'] => blockquote:fresh",
        "p[style-name='Citação'] => blockquote:fresh",
        "p[style-name='Citacao'] => blockquote:fresh",
        "p[style-name='Citação longa'] => blockquote:fresh",
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Título'] => h1:fresh",
        "p[style-name='__quote'] => blockquote:fresh",
        "p[style-name='__center'] => p.abnt-center:fresh",
      ],
      transformDocument: mammoth.transforms.paragraph(function (p) {
        const plain = !p.styleName || /^(normal|body text|corpo de texto|default|padrão|texto|list paragraph|parágrafo da lista)$/i.test(p.styleName);
        if (!plain) return p;
        const start = parseInt((p.indent && p.indent.start) || 0, 10) || 0;
        if (start >= 1700 && !p.numbering) return Object.assign({}, p, { styleName: '__quote', styleId: null });
        if (p.alignment === 'center' && !p.numbering) return Object.assign({}, p, { styleName: '__center', styleId: null });
        return p;
      }),
    };
  }

  function textToBlocks(text) {
    const blocks = [];
    String(text).split(/\r?\n/).forEach((line) => {
      const t = line.replace(/\s+/g, ' ').trim();
      if (!t) return;
      blocks.push({ id: nextId(), kind: 'p', meta: {}, runs: [{ t }], text: t });
    });
    return { blocks, notes: {} };
  }

  // ------------------------------------------------------------------
  // PRÉ-CLASSIFICAÇÃO POR REGRAS (sem IA)
  // ------------------------------------------------------------------
  const RX = {
    intro: /^(\d+(\.\d+)*\.?\s*[-–—]?\s*)?INTRODU[ÇC][ÃA]O$/i,
    refs: /^(REFER[ÊE]NCIAS(\s+BIBLIOGR[ÁA]FICAS)?|BIBLIOGRAFIA|REFER[ÊE]NCIAS\s+CONSULTADAS)$/i,
    post: /^(AP[ÊE]NDICES?|ANEXOS?|GLOSS[ÁA]RIO|[ÍI]NDICE)(\s+[A-Z0-9]{1,3})?(\s*[-–—:]\s*.*)?$/i,
    pre: /^(RESUMO|ABSTRACT|RESUMEN|SUM[ÁA]RIO|AGRADECIMENTOS?|DEDICAT[ÓO]RIA|EP[ÍI]GRAFE|ERRATA|FOLHA DE APROVA[ÇC][ÃA]O|BANCA EXAMINADORA|LISTA DE .+)$/i,
    kw: /^(palavras?[- ]chaves?)\s*:\s*(.*)$/i,
    kwEn: /^(key[- ]?words?)\s*:\s*(.*)$/i,
    caption: /^(Figura|Quadro|Tabela|Gr[áa]fico|Fotografia|Foto|Imagem|Ilustra[çc][ãa]o|Mapa|Esquema|Desenho|Organograma|Fluxograma|Diagrama|Planta|Retrato|Infogr[áa]fico)\s+\d+/i,
    source: /^Fonte\s*:/i,
    numbered: /^(\d{1,2}(?:\.\d{1,2}){0,4})\.?\s+(?:[-–—]\s*)?(\S.*)$/,
    longQuoteEnd: /\(([A-ZÀ-Ý][^()]{1,120}),\s*\d{4}[a-z]?,\s*(p|f)\.\s*[\dIVXLCivxlc]+(\s*[-–]\s*[\dIVXLCivxlc]+)?\)\.?$/,
    alinea: /^[a-z]\)\s+\S/,
    tocLine: /(\.{3,}|…{2,}|\t)\s*\d+\s*$/,
  };

  const upperRatio = (t) => {
    const letters = t.replace(/[^A-Za-zÀ-ÿ]/g, '');
    if (!letters) return 0;
    const up = letters.replace(/[^A-ZÀ-Þ]/g, '').length;
    return up / letters.length;
  };

  function headingLevelFromText(t) {
    const m = RX.numbered.exec(t);
    if (!m) return 0;
    const parts = m[1].split('.');
    if (parseInt(parts[0], 10) > 30) return 0;
    if (t.length > 160) return 0;
    if (/[.;:,]$/.test(t)) return 0;
    if (!/^[A-ZÀ-Ý"“]/.test(m[2])) return 0;
    return Math.min(parts.length, 5);
  }

  function classify(parsed) {
    const blocks = parsed.blocks;
    const extracted = {};
    let start = blocks.findIndex((b) => b.kind === 'p' && RX.intro.test(b.text) && !RX.tocLine.test(b.text));
    const hasIntro = start >= 0;
    if (!hasIntro) start = 0;

    // pré-textuais antes da introdução: ignorados, mas aproveita resumo/abstract/agradecimentos
    let mode = null;
    for (let i = 0; i < start; i++) {
      const b = blocks[i];
      b.type = 'skip';
      if (b.kind !== 'p') continue;
      const t = b.text;
      const up = t.toUpperCase();
      if (/^RESUMO$/.test(up)) { mode = 'resumo'; continue; }
      if (/^(ABSTRACT|RESUMEN)$/.test(up)) { mode = 'abstract'; continue; }
      if (/^AGRADECIMENTOS?$/.test(up)) { mode = 'agradecimentos'; continue; }
      if (RX.pre.test(t)) { mode = null; continue; }
      let m;
      if ((m = RX.kw.exec(t))) { extracted.palavrasChave = m[2]; mode = null; continue; }
      if ((m = RX.kwEn.exec(t))) { extracted.keywords = m[2]; mode = null; continue; }
      if (mode) extracted[mode] = (extracted[mode] ? extracted[mode] + '\n' : '') + t;
    }

    let inRefs = false;
    for (let i = start; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === 'img') { b.type = 'img'; continue; }
      if (b.kind === 'table') { b.type = 'table'; continue; }
      const t = b.text;
      const meta = b.meta || {};
      const short = t.length <= 120;
      const isUpper = upperRatio(t) > 0.85;
      const allBold = b.runs.length > 0 && b.runs.every((r) => r.fn || r.b || !r.t.trim());

      if (short && RX.refs.test(t)) { b.type = 'post'; inRefs = true; continue; }
      if (short && RX.post.test(t) && (isUpper || /^(ap[êe]ndice|anexo)\s+[a-z0-9]{1,3}\s*[-–—:]/i.test(t))) { b.type = 'post'; inRefs = false; continue; }
      if (short && RX.pre.test(t) && isUpper && !hasIntro) { b.type = 'skip'; continue; }
      if (RX.tocLine.test(t) && short) { b.type = 'skip'; continue; }

      if (inRefs) {
        const lvl = headingLevelFromText(t);
        if (lvl && isUpper) { b.type = 'h' + lvl; inRefs = false; continue; }
        b.type = 'ref';
        continue;
      }

      const tagLvl = meta.tag && /^h[1-6]$/.test(meta.tag) ? Math.min(parseInt(meta.tag[1], 10), 5) : 0;
      const txtLvl = headingLevelFromText(t);
      if (tagLvl) { b.type = 'h' + (txtLvl || tagLvl); continue; }
      if (RX.caption.test(t) && t.length < 260) { b.type = 'caption'; continue; }
      if (RX.source.test(t) && t.length < 400) { b.type = 'source'; continue; }
      if (meta.quote) { b.type = 'quote'; continue; }
      if (meta.list) { b.type = 'item'; continue; }
      if (txtLvl) { b.type = 'h' + txtLvl; continue; }
      if (short && isUpper && !/[.;,]$/.test(t) && /[A-ZÀ-Þ]{3}/.test(t)) { b.type = 'h1'; continue; }
      if (short && allBold && !/[.;,]$/.test(t)) { b.type = 'h2'; continue; }
      if (RX.alinea.test(t)) { b.type = 'item'; continue; }
      if (t.length >= 230 && !/[“”"«»]/.test(t) && RX.longQuoteEnd.test(t)) { b.type = 'quote'; continue; }
      b.type = 'p';
    }
    blocks.forEach((b) => { b.auto = b.type; });
    return { blocks, notes: parsed.notes, extracted, hasIntro };
  }

  // ------------------------------------------------------------------
  // PLANO: o que entra no documento, com numeração das seções
  // ------------------------------------------------------------------
  function cleanHeading(t) {
    const s = t.replace(/^\s*\d{1,2}(\.\d{1,2}){0,4}\.?\s*[-–—]?\s*/, '');
    return s.trim() ? s.trim() : t.trim();
  }
  const capsPT = (s) => s.toLocaleUpperCase('pt-BR');

  function buildPlan(blocks, prof, opts) {
    const items = [];
    const counters = [0, 0, 0, 0, 0];
    const list = blocks.filter((b) => b.type !== 'skip');
    for (const b of list) {
      const it = { block: b, type: b.type };
      if (/^h[1-5]$/.test(b.type)) {
        const lvl = parseInt(b.type[1], 10);
        counters[lvl - 1]++;
        for (let k = lvl; k < 5; k++) counters[k] = 0;
        for (let k = 0; k < lvl - 1; k++) if (!counters[k]) counters[k] = 1;
        it.level = lvl;
        it.number = counters.slice(0, lvl).join('.');
        const style = prof.titulos[lvl] || {};
        let title = cleanHeading(b.text);
        if (style.caps) title = capsPT(title);
        it.title = title;
        it.full = it.number + ' ' + title;
      } else if (b.type === 'post') {
        it.level = 1;
        it.title = capsPT(b.text.trim());
        it.full = it.title;
      }
      items.push(it);
    }
    // ordena cada grupo contínuo de referências
    if (opts.ordenarReferencias) {
      let i = 0;
      while (i < items.length) {
        if (items[i].type !== 'ref') { i++; continue; }
        let j = i;
        while (j < items.length && items[j].type === 'ref') j++;
        const group = items.slice(i, j).sort((a, b) => a.block.text.localeCompare(b.block.text, 'pt-BR', { sensitivity: 'base', ignorePunctuation: true }));
        items.splice(i, j - i, ...group);
        i = j;
      }
    }
    return items;
  }

  // ------------------------------------------------------------------
  // IMAGENS: lê tamanho direto dos bytes (PNG, JPEG, GIF, BMP)
  // ------------------------------------------------------------------
  function b64ToBytes(b64) {
    if (typeof atob === 'function') {
      const bin = atob(b64);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    return new Uint8Array(Buffer.from(b64, 'base64'));
  }

  function imageInfo(src) {
    const m = /^data:(image\/[\w.+-]+);base64,(.*)$/i.exec(src || '');
    if (!m) return null;
    const mime = m[1].toLowerCase();
    const bytes = b64ToBytes(m[2]);
    const u16be = (o) => (bytes[o] << 8) | bytes[o + 1];
    const u32be = (o) => ((bytes[o] << 24) | (bytes[o + 1] << 16) | (bytes[o + 2] << 8) | bytes[o + 3]) >>> 0;
    const u16le = (o) => bytes[o] | (bytes[o + 1] << 8);
    const i32le = (o) => bytes[o] | (bytes[o + 1] << 8) | (bytes[o + 2] << 16) | (bytes[o + 3] << 24);
    if (bytes[0] === 0x89 && bytes[1] === 0x50) return { type: 'png', bytes, w: u32be(16), h: u32be(20) };
    if (bytes[0] === 0x47 && bytes[1] === 0x49) return { type: 'gif', bytes, w: u16le(6), h: u16le(8) };
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return { type: 'bmp', bytes, w: Math.abs(i32le(18)), h: Math.abs(i32le(22)) };
    if (bytes[0] === 0xff && bytes[1] === 0xd8) {
      let o = 2;
      while (o < bytes.length) {
        if (bytes[o] !== 0xff) { o++; continue; }
        const marker = bytes[o + 1];
        const len = u16be(o + 2);
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { type: 'jpg', bytes, w: u16be(o + 7), h: u16be(o + 5) };
        }
        o += 2 + len;
      }
    }
    return { type: null, mime, bytes, w: 0, h: 0 };
  }

  const MAX_IMG_W = 600; // px (≈ 15,9 cm)
  const MAX_IMG_H = 640; // px (≈ 17 cm)
  function fitImage(info) {
    let w = info.w || 400, h = info.h || 300;
    const s = Math.min(1, MAX_IMG_W / w, MAX_IMG_H / h);
    return { w: Math.round(w * s), h: Math.round(h * s) };
  }

  // ------------------------------------------------------------------
  // ESTIMATIVA DE PÁGINAS (para o sumário já sair com números)
  // ------------------------------------------------------------------
  function makeLayout(prof, measure) {
    const F = prof.fonte;
    const PT = { W: (21 - prof.margens.esq - prof.margens.dir) / 2.54 * 72, H: (29.7 - prof.margens.sup - prof.margens.inf) / 2.54 * 72 };
    const cache = new Map();
    function width(str, size, b, i) {
      const k = size + (b ? 'b' : '') + (i ? 'i' : '') + '|' + str;
      let v = cache.get(k);
      if (v === undefined) { v = measure(str, { name: F, size, bold: !!b, italic: !!i }); cache.set(k, v); }
      return v;
    }
    const lh = (size, mult) => size * 1.15 * mult;

    // quantas linhas um parágrafo ocupa
    function countLines(runs, size, avail, firstIndent, opt) {
      const words = [];
      let curW = 0, curHas = false, pendingSpace = 0;
      for (const r of runs) {
        if (r.fn) { curW += width('9', size * 0.6, false, false); curHas = true; continue; }
        const b = opt && opt.bold != null ? opt.bold : r.b;
        const parts = r.t.split(/( +)/);
        for (const p of parts) {
          if (!p) continue;
          if (/^ +$/.test(p)) {
            if (curHas) { words.push({ w: curW, sp: 0 }); curW = 0; curHas = false; }
            pendingSpace = width(' ', size, b, r.i);
            if (words.length) words[words.length - 1].sp = pendingSpace;
          } else {
            curW += width(opt && opt.caps ? capsPT(p) : p, size, b, r.i);
            curHas = true;
          }
        }
      }
      if (curHas) words.push({ w: curW, sp: 0 });
      if (!words.length) return 1;
      let lines = 1, x = 0, lim = avail - firstIndent;
      for (let k = 0; k < words.length; k++) {
        const w = words[k];
        const prevSp = k > 0 ? words[k - 1].sp : 0;
        if (x > 0 && x + prevSp + w.w > lim + 0.01) { lines++; x = w.w; lim = avail; }
        else x += (x > 0 ? prevSp : 0) + w.w;
      }
      return lines;
    }

    return { PT, width, lh, countLines, F };
  }

  // ------------------------------------------------------------------
  // TIPOS DE TRABALHO
  // ------------------------------------------------------------------
  const WORK_TYPES = {
    trabalho: { id: 'trabalho', label: 'Trabalho do semestre', short: 'Trabalho', hint: 'Atividades e trabalhos das disciplinas', entrelinha: 1.5, novaPagina: true, capa: true, folhaRosto: false, sumario: true },
    artigo: { id: 'artigo', label: 'Artigo científico', short: 'Artigo', hint: 'Título, autores e resumo na 1ª página', entrelinha: 1, novaPagina: false, capa: false, folhaRosto: false, sumario: false },
    tcc: { id: 'tcc', label: 'TCC', short: 'TCC', hint: 'Trabalho de conclusão de curso', entrelinha: 1.5, novaPagina: true, capa: true, folhaRosto: true, sumario: true },
  };

  function resolveConfig(prof, info) {
    const T = WORK_TYPES[info.tipo] || WORK_TYPES.trabalho;
    const pick = (v, d) => (v === undefined || v === null || v === '' ? d : v);
    const tcc = T.id === 'tcc', trab = T.id === 'trabalho', art = T.id === 'artigo';
    const entrelinha = Number(pick(info.entrelinha, T.entrelinha)) || T.entrelinha;
    const txt = (k) => String(info[k] || '').trim();
    return {
      tipo: T.id, T, tcc, trab, art,
      fonte: info.fonte || prof.fonte,
      entrelinha,
      line: Math.round(240 * entrelinha),
      novaPagina: tcc ? true : !!pick(info.novaPagina, T.novaPagina),
      capa: tcc ? true : !!pick(info.capa, T.capa),
      folhaRosto: tcc ? true : trab ? !!pick(info.folhaRosto, T.folhaRosto) : false,
      sumario: tcc ? true : trab ? !!pick(info.sumario, T.sumario) : false,
      aprovacao: tcc && !!info.aprovacao,
      dedicatoria: tcc ? txt('dedicatoria') : '',
      epigrafe: tcc ? txt('epigrafe') : '',
      agradecimentos: trab ? '' : txt('agradecimentos'),
      resumo: txt('resumo'),
      abstract: txt('abstract'),
      refAntes: Math.round(((prof.refEspaco && prof.refEspaco.antes) || 0) * 20),
      refDepois: Math.round(((prof.refEspaco && prof.refEspaco.depois) || 12) * 20),
    };
  }

  const isAnnex = (t) => /^(AP[ÊE]NDICE|ANEXO)/i.test(t || '');

  // ------------------------------------------------------------------
  // ESTIMATIVA DE PÁGINAS (para o sumário já sair com números)
  // ------------------------------------------------------------------
  function estimatePages(plan, info, prof, cfg, measure) {
    if (!measure || !cfg.sumario) return null;
    const L = makeLayout(Object.assign({}, prof, { fonte: cfg.fonte }), measure);
    const S = prof.tamanho, s = prof.tamanhoMenor, E = cfg.entrelinha;
    const H = L.PT.H, W = L.PT.W;
    const ptOf = (tw) => tw / 20;
    const indentP = ptOf(cm(prof.recuoParagrafo));
    const indentQ = ptOf(cm(prof.recuoCitacao));
    const gap = ptOf(cfg.line);

    const paraPages = (paras, size, mult, extraTop) => {
      let y = extraTop || 0, pages = 1;
      const l = L.lh(size, mult);
      for (const t of paras) {
        const n = L.countLines([{ t }], size, W, 0);
        for (let k = 0; k < n; k++) { if (y + l > H) { pages++; y = 0; } y += l; }
      }
      return pages;
    };
    const top = L.lh(S, E) + 2 * gap;
    let pre = 0;
    if (cfg.folhaRosto) pre++;
    if (cfg.aprovacao) pre++;
    if (cfg.dedicatoria) pre++;
    if (cfg.tcc && cfg.agradecimentos) pre += paraPages(splitParas(cfg.agradecimentos), S, E, top);
    if (cfg.epigrafe) pre++;
    if (cfg.resumo) pre += paraPages(splitParas(cfg.resumo).concat(['Palavras-chave: ' + (info.palavrasChave || '')]), S, 1, top);
    if (cfg.abstract) pre += paraPages(splitParas(cfg.abstract).concat(['Keywords: ' + (info.keywords || '')]), S, 1, top);
    const tocLines = plan.filter((it) => it.full).map((it) => it.full + '   00');
    if (tocLines.length) pre += paraPages(tocLines, S, E, top);

    let page = pre + 1, y = 0, firstInSection = true;
    const pages = new Map();
    const newPage = () => { page++; y = 0; };

    function metrics(it) {
      const b = it.block;
      switch (it.type) {
        case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'post': {
          const st = it.type === 'post' ? { bold: true, caps: true } : prof.titulos[it.level] || {};
          const n = L.countLines([{ t: it.full, b: st.bold, i: st.italic }], S, W, 0);
          const brk = (it.type === 'h1' && cfg.novaPagina) || (it.type === 'post' && (cfg.novaPagina || isAnnex(it.title)));
          return { n, l: L.lh(S, E), before: brk ? 0 : gap, after: gap, keepNext: true, pageBreak: brk };
        }
        case 'quote': return { n: L.countLines(b.runs, s, W - indentQ, 0), l: L.lh(s, 1), before: 12, after: 12 };
        case 'item': return { n: L.countLines(b.runs, S, W - ptOf(cm(2)), 0), l: L.lh(S, E), before: 0, after: 0 };
        case 'caption': return { n: L.countLines(b.runs, s, W, 0), l: L.lh(s, 1), before: 12, after: 0, keepNext: true };
        case 'source': return { n: L.countLines(b.runs, s, W, 0), l: L.lh(s, 1), before: 0, after: 12 };
        case 'ref': return { n: L.countLines(b.runs, S, W, 0), l: L.lh(S, 1), before: ptOf(cfg.refAntes), after: ptOf(cfg.refDepois) };
        case 'img': {
          const inf = b._img || imageInfo(b.src);
          if (!inf || !inf.type) return { n: 0, l: 0, before: 0, after: 0 };
          const d = fitImage(inf);
          return { n: 1, l: d.h * 0.75 + 4, before: 0, after: 0, atomic: true };
        }
        case 'table': {
          const rowsH = b.rows.map((row) => {
            const cw = W / Math.max(1, row.reduce((a, c) => a + c.colspan, 0));
            let mx = 1;
            row.forEach((c) => {
              const n = c.paras.reduce((a, rr) => a + L.countLines(rr.length ? rr : [{ t: ' ' }], s, cw * c.colspan - 11, 0), 0);
              mx = Math.max(mx, n);
            });
            return mx * L.lh(s, 1) + 2;
          });
          return { rows: rowsH, before: 0, after: 0 };
        }
        default: return { n: L.countLines(b.runs, S, W, indentP), l: L.lh(S, E), before: 0, after: 0 };
      }
    }

    for (let idx = 0; idx < plan.length; idx++) {
      const it = plan[idx];
      const m = metrics(it);
      if (m.pageBreak && !firstInSection && y > 0) newPage();
      firstInSection = false;
      if (m.rows) {
        for (const rh of m.rows) { if (y + rh > H && y > 0) newPage(); y += rh; }
        continue;
      }
      if (!m.n) continue;
      if (y > 0) y += m.before;
      if (m.keepNext) {
        const nx = plan[idx + 1] ? metrics(plan[idx + 1]) : null;
        let need = m.n * m.l + m.after;
        if (nx) need += nx.rows ? nx.rows[0] || 0 : (nx.before || 0) + Math.min(nx.n || 0, nx.atomic ? 1 : 2) * (nx.l || 0);
        if (y + need > H && y > 0) newPage();
      }
      if (m.atomic) {
        if (y + m.l > H && y > 0) newPage();
        y += m.l;
      } else {
        let left = m.n;
        let fit = Math.floor((H - y + 0.01) / m.l);
        if (fit < left) {
          if (left <= 3 || fit < 2) { if (y > 0) newPage(); fit = Math.floor((H - y + 0.01) / m.l); }
          else if (left - fit < 2) fit = left - 2;
        }
        const startPage = page;
        while (left > 0) {
          const take = Math.min(left, Math.max(1, fit));
          y += take * m.l;
          left -= take;
          if (left > 0) { newPage(); fit = Math.floor(H / m.l); if (left - fit === 1 && fit > 2) fit--; }
        }
        if (it.full) pages.set(it, startPage);
      }
      y += m.after;
      if (y > H) { y = y - H > m.after ? 0 : y; if (y > H) newPage(); }
    }
    return { pages, total: page, firstTextPage: pre + 1 };
  }

  function splitParas(t) {
    return String(t || '').split(/\r?\n+/).map((s) => s.trim()).filter(Boolean);
  }

  // palavras-chave: UNICID pede separação por ponto e inicial maiúscula;
  // a NBR 6028:2021 usa ponto e vírgula
  function formatKeywords(s, estilo) {
    const parts = String(s || '').split(/[;,.]+/).map((x) => x.trim()).filter(Boolean);
    if (!parts.length) return '';
    if (estilo === 'ponto') return parts.map((p) => p.charAt(0).toLocaleUpperCase('pt-BR') + p.slice(1)).join('. ') + '.';
    return parts.join('; ') + '.';
  }

  function naturezaText(cfg, info, prof) {
    const custom = String(info.natureza || '').trim();
    const inst = String(info.instituicao || prof.instituicao).trim();
    const curso = String(info.curso || prof.curso).trim();
    if (cfg.tcc) return custom || prof.natureza;
    const disc = String(info.disciplina || '').trim();
    return (disc ? 'Trabalho apresentado à disciplina ' + disc + ' do Curso de ' : 'Trabalho apresentado ao Curso de ') + curso + ' da ' + inst + ', como requisito parcial de avaliação.';
  }

  // ------------------------------------------------------------------
  // GERAÇÃO DO .DOCX
  // ------------------------------------------------------------------
  function buildDocument(state, opts) {
    const D = root.docx;
    const prof = state.profile;
    const info = state.info;
    const cfg = resolveConfig(prof, info);
    const F = cfg.fonte;
    const S2 = prof.tamanho * 2; // meia-pontos
    const s2 = prof.tamanhoMenor * 2;
    const LINE = cfg.line;
    const AUTO = (D.LineRuleType && D.LineRuleType.AUTO) || 'auto';
    const W_TW = cm(21 - prof.margens.esq - prof.margens.dir);
    const warnings = [];
    const plan = buildPlan(state.blocks, prof, { ordenarReferencias: info.ordenarReferencias !== false });

    plan.forEach((it) => {
      if (it.type === 'img') {
        const inf = imageInfo(it.block.src);
        it.block._img = inf;
        if (!inf || !inf.type) warnings.push('Uma imagem está num formato que o Word do celular não lê (' + ((inf && inf.mime) || 'desconhecido') + ') e ficou de fora.');
      }
    });

    const est = opts && opts.measure ? estimatePages(plan, info, prof, cfg, opts.measure) : null;

    // notas de rodapé
    const footnotes = {};
    const fnIds = {};
    let fnCount = 0;
    const notes = state.notes || {};
    const run = (r, extra) => {
      if (r.fn) {
        if (!fnIds[r.fn]) {
          fnIds[r.fn] = ++fnCount;
          footnotes[fnCount] = { children: [new D.Paragraph({ style: 'NotaRodape', children: (notes[r.fn] || []).map((x) => run(x)) })] };
        }
        return new D.FootnoteReferenceRun(fnIds[r.fn]);
      }
      return new D.TextRun(Object.assign({
        text: r.t,
        bold: r.b || undefined,
        italics: r.i || undefined,
        underline: r.u ? {} : undefined,
        superScript: r.sup || undefined,
        subScript: r.sub || undefined,
      }, extra || {}));
    };
    const noteOf = (text) => {
      const id = ++fnCount;
      footnotes[id] = { children: [new D.Paragraph({ style: 'NotaRodape', children: [new D.TextRun(text)] })] };
      return new D.FootnoteReferenceRun(id);
    };
    const runsOf = (runs, extra) => runs.map((r) => run(r, extra));
    const P = (o) => new D.Paragraph(o);
    const T = (text, o) => new D.TextRun(Object.assign({ text }, o || {}));
    const center = D.AlignmentType.CENTER;

    // ---------- dados de identificação ----------
    const autores = splitParas(info.autores);
    const titulo = String(info.titulo || '').trim();
    const sub = String(info.subtitulo || '').trim();
    const cidade = String(info.cidade || prof.cidade).trim();
    const ano = String(info.ano || new Date().getFullYear()).trim();
    const inst = String(info.instituicao || prof.instituicao).trim();
    const curso = String(info.curso || prof.curso).trim();
    const orient = String(info.orientador || '').trim();
    const coorient = String(info.coorientador || '').trim();
    const disc = String(info.disciplina || '').trim();
    const pessoaLabel = cfg.trab ? 'Professor(a): ' : 'Orientador(a): ';

    // no artigo, as notas dos autores vêm antes das notas do texto (1, 2, …)
    const artNotes = { aut: [], ori: null };
    if (cfg.art) {
      const notaAut = String(info.notaAutores || '').trim() || 'Graduando(a) do Curso de ' + curso + ' da ' + inst + '.';
      const notaOri = String(info.notaOrientador || '').trim() || 'Professor(a) orientador(a) do Curso de ' + curso + ' da ' + inst + '.';
      artNotes.aut = (autores.length ? autores : ['Nome do(a) autor(a)']).map((a) => ({ a, ref: noteOf(notaAut) }));
      if (orient) artNotes.ori = noteOf(notaOri);
    }

    // ---------- textuais e pós-textuais ----------
    const body = [];
    const tocEntries = [];
    let firstBody = true;
    plan.forEach((it, idx) => {
      const b = it.block;
      const first = firstBody;
      firstBody = false;
      switch (it.type) {
        case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': {
          const lvl = it.level;
          body.push(P({
            heading: D.HeadingLevel['HEADING_' + lvl],
            pageBreakBefore: lvl === 1 && cfg.novaPagina && !first ? true : undefined,
            spacing: lvl === 1 && !cfg.novaPagina && !first ? { before: LINE, after: LINE, line: LINE, lineRule: AUTO } : undefined,
            children: [T(it.number + ' ' + it.title)],
          }));
          tocEntries.push({ title: it.full, level: lvl, page: est ? est.pages.get(it) : undefined });
          break;
        }
        case 'post': {
          const brk = (cfg.novaPagina || isAnnex(it.title)) && !first;
          body.push(P({ heading: D.HeadingLevel.HEADING_1, alignment: center, pageBreakBefore: brk || undefined, spacing: !brk && !first ? { before: LINE, after: LINE, line: LINE, lineRule: AUTO } : undefined, children: [T(it.title)] }));
          tocEntries.push({ title: it.title, level: 1, page: est ? est.pages.get(it) : undefined });
          break;
        }
        case 'quote':
          body.push(P({ style: 'CitacaoLonga', children: runsOf(b.runs) }));
          break;
        case 'item': {
          const lst = (b.meta && b.meta.list) || null;
          let label = null;
          if (lst && !RX.alinea.test(b.text)) label = lst.ordered ? String.fromCharCode(96 + Math.min(26, lst.index)) + ')' : '–';
          const kids = [];
          if (label) kids.push(T(label), new D.TextRun({ children: [new D.Tab()] }));
          kids.push(...runsOf(b.runs));
          const depth = lst ? lst.depth : 0;
          body.push(P({ style: 'Alinea', indent: depth ? { left: cm(2 + depth * 0.75), hanging: cm(0.75) } : undefined, children: kids }));
          break;
        }
        case 'caption':
          body.push(P({ style: 'Legenda', keepNext: true, children: runsOf(b.runs) }));
          break;
        case 'source':
          body.push(P({ style: 'FonteIlustracao', children: runsOf(b.runs) }));
          break;
        case 'ref':
          body.push(P({ style: 'Referencia', children: runsOf(b.runs) }));
          break;
        case 'img': {
          const inf = b._img;
          if (!inf || !inf.type) break;
          const d = fitImage(inf);
          const next = plan[idx + 1];
          body.push(P({
            style: 'Imagem',
            keepNext: next && next.type === 'source' ? true : undefined,
            children: [new D.ImageRun({ type: inf.type, data: inf.bytes, transformation: { width: d.w, height: d.h }, altText: b.alt ? { title: b.alt, description: b.alt, name: 'Imagem' } : undefined })],
          }));
          break;
        }
        case 'table': {
          const ncols = Math.max(1, ...b.rows.map((r) => r.reduce((a, c) => a + c.colspan, 0)));
          const border = { style: D.BorderStyle.SINGLE, size: 4, color: '000000' };
          const rows = b.rows.map((row) => new D.TableRow({
            children: row.map((c) => new D.TableCell({
              columnSpan: c.colspan > 1 ? c.colspan : undefined,
              width: { size: Math.round((W_TW / ncols) * c.colspan), type: D.WidthType.DXA },
              borders: { top: border, bottom: border, left: border, right: border },
              margins: { left: 80, right: 80, top: 20, bottom: 20 },
              children: (c.paras.length ? c.paras : [[]]).map((rr) => P({ style: 'Tabela', children: runsOf(rr) })),
            })),
          }));
          body.push(new D.Table({ width: { size: W_TW, type: D.WidthType.DXA }, columnWidths: Array.from({ length: ncols }, () => Math.round(W_TW / ncols)), rows }));
          body.push(P({ style: 'Tabela', children: [] }));
          break;
        }
        default:
          body.push(P({ style: 'Corpo', children: runsOf(b.runs) }));
      }
    });

    const titleRuns = () => {
      const r = [T(capsPT(titulo || 'TÍTULO DO TRABALHO'), { bold: true })];
      if (sub) r.push(T(': ' + sub, { bold: true }));
      return r;
    };
    const lineSp = { line: LINE, lineRule: AUTO, before: 0, after: 0 };
    const cityFooter = () => new D.Footer({ children: [
      P({ alignment: center, spacing: { line: LINE, lineRule: AUTO }, children: [T(capsPT(cidade))] }),
      P({ alignment: center, spacing: { line: LINE, lineRule: AUTO }, children: [T(ano)] }),
    ] });
    const emptyHeader = () => new D.Header({ children: [P({ children: [] })] });
    const emptyFooter = () => new D.Footer({ children: [P({ children: [] })] });
    const numberHeader = () => new D.Header({ children: [P({ alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ children: [D.PageNumber.CURRENT], size: s2 })] })] });
    let restartDone = false;
    const pageProps = (extra) => {
      const p = Object.assign({
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: cm(prof.margens.sup), left: cm(prof.margens.esq), bottom: cm(prof.margens.inf), right: cm(prof.margens.dir), header: cm(2), footer: cm(1.25) },
        },
      }, extra || {});
      return p;
    };
    // a primeira seção depois da capa recomeça a contagem em 1
    const counted = (props) => { if (!restartDone) { props.page.pageNumbers = { start: 1 }; restartDone = true; } return props; };

    const sections = [];
    const kwFmt = (s) => formatKeywords(s, prof.palavrasChave);

    // Capa (não é contada)
    if (cfg.capa) {
      const kids = [
        P({ alignment: center, spacing: lineSp, children: [T(capsPT(inst), { bold: true })] }),
        P({ alignment: center, spacing: lineSp, children: [T(capsPT('Curso de ' + curso))] }),
        P({ spacing: lineSp, children: [] }), P({ spacing: lineSp, children: [] }), P({ spacing: lineSp, children: [] }),
        ...autores.map((a) => P({ alignment: center, spacing: lineSp, children: [T(capsPT(a))] })),
        P({ alignment: center, spacing: { before: cm(4.5), line: LINE, lineRule: AUTO }, children: titleRuns() }),
      ];
      if (cfg.trab && !cfg.folhaRosto && (disc || orient)) {
        if (disc) kids.push(P({ alignment: center, spacing: { before: cm(2), line: LINE, lineRule: AUTO }, children: [T('Disciplina: ' + disc)] }));
        if (orient) kids.push(P({ alignment: center, spacing: { before: disc ? 0 : cm(2), line: LINE, lineRule: AUTO }, children: [T(pessoaLabel + orient)] }));
      }
      sections.push({ properties: pageProps(), headers: { default: emptyHeader() }, footers: { default: cityFooter() }, children: kids });
    }

    const preSection = (children, footer) => sections.push({
      properties: counted(pageProps()),
      headers: { default: emptyHeader() },
      footers: { default: footer || emptyFooter() },
      children,
    });
    const preTitle = (t) => P({ style: 'TituloPre', children: [T(t)] });
    const naturezaBlock = () => [
      P({ alignment: D.AlignmentType.JUSTIFIED, indent: { left: cm(prof.recuoNatureza) }, spacing: { before: cm(1.5), line: 240, lineRule: AUTO }, children: [T(naturezaText(cfg, info, prof))] }),
    ];
    const orientBlock = () => {
      const out = [];
      if (orient) out.push(P({ indent: { left: cm(prof.recuoNatureza) }, spacing: { before: 240, line: 240, lineRule: AUTO }, children: [T(pessoaLabel + orient)] }));
      if (coorient && !cfg.trab) out.push(P({ indent: { left: cm(prof.recuoNatureza) }, spacing: { before: 120, line: 240, lineRule: AUTO }, children: [T('Coorientador(a): ' + coorient)] }));
      return out;
    };
    const resumoBlock = (label, text, kwLabel, kw, lang) => {
      const l = lang ? { language: { value: lang } } : {};
      return [
        ...splitParas(text).map((t) => P({ style: 'Resumo', children: [T(t, l)] })),
        P({ style: 'Resumo', spacing: { before: cfg.art ? 0 : 240, line: 240, lineRule: AUTO }, children: [T(kwLabel + ': ', Object.assign({ bold: true }, l)), T(formatKeywords(kw, prof.palavrasChave), l)] }),
      ];
    };

    if (cfg.art) {
      // ---------- ARTIGO: tudo começa na 1ª página ----------
      const head = [
        P({ alignment: center, spacing: { line: 240, lineRule: AUTO, before: 0, after: 0 }, children: titleRuns() }),
        P({ spacing: { line: 360, lineRule: AUTO }, children: [] }),
      ];
      artNotes.aut.forEach((x) => head.push(P({ alignment: D.AlignmentType.RIGHT, spacing: { line: 240, lineRule: AUTO }, children: [T(x.a), x.ref] })));
      if (artNotes.ori) head.push(P({ alignment: D.AlignmentType.RIGHT, spacing: { line: 240, lineRule: AUTO }, children: [T(orient), artNotes.ori] }));
      if (cfg.resumo) {
        head.push(P({ alignment: center, spacing: { before: LINE + 240, after: 240, line: 240, lineRule: AUTO }, children: [T('RESUMO', { bold: true })] }));
        head.push(...resumoBlock('RESUMO', cfg.resumo, 'Palavras-chave', info.palavrasChave));
      }
      if (cfg.abstract) {
        head.push(P({ alignment: center, spacing: { before: LINE + 240, after: 240, line: 240, lineRule: AUTO }, children: [T('ABSTRACT', { bold: true, language: { value: 'en-US' } })] }));
        head.push(...resumoBlock('ABSTRACT', cfg.abstract, 'Keywords', info.keywords, 'en-US'));
      }
      head.push(P({ spacing: { line: LINE, lineRule: AUTO }, children: [] }));
      const tail = [];
      if (cfg.agradecimentos) {
        tail.push(P({ alignment: center, spacing: { before: LINE, after: LINE, line: LINE, lineRule: AUTO }, keepNext: true, children: [T('AGRADECIMENTOS', { bold: true })] }));
        splitParas(cfg.agradecimentos).forEach((t) => tail.push(P({ style: 'Corpo', children: [T(t)] })));
      }
      sections.push({
        properties: Object.assign(counted(pageProps()), { titlePage: true }),
        headers: { first: emptyHeader(), default: numberHeader() },
        footers: { first: emptyFooter(), default: emptyFooter() },
        children: [...head, ...body, ...tail],
      });
    } else {
      // ---------- TRABALHO / TCC ----------
      if (cfg.folhaRosto) {
        preSection([
          ...autores.map((a) => P({ alignment: center, spacing: lineSp, children: [T(capsPT(a))] })),
          P({ alignment: center, spacing: { before: cm(6), line: LINE, lineRule: AUTO }, children: titleRuns() }),
          ...naturezaBlock(),
          ...orientBlock(),
        ], cityFooter());
      }
      if (cfg.aprovacao) {
        const banca = splitParas(info.banca);
        const kids = [
          ...autores.map((a) => P({ alignment: center, spacing: lineSp, children: [T(capsPT(a))] })),
          P({ alignment: center, spacing: { before: cm(1.5), line: LINE, lineRule: AUTO }, children: titleRuns() }),
          ...naturezaBlock(),
          P({ spacing: { before: cm(1.2), line: LINE, lineRule: AUTO }, children: [T('Aprovado em: ____/____/________')] }),
          P({ alignment: center, spacing: { before: cm(1), after: 0, line: LINE, lineRule: AUTO }, children: [T('BANCA EXAMINADORA', { bold: true })] }),
        ];
        (banca.length ? banca : ['Prof(a). Dr(a). Nome – Instituição', 'Prof(a). Dr(a). Nome – Instituição', 'Prof(a). Dr(a). Nome – Instituição']).forEach((m) => {
          kids.push(P({ alignment: center, spacing: { before: cm(1.2), line: 240, lineRule: AUTO }, children: [T('_______________________________________')] }));
          kids.push(P({ alignment: center, spacing: { line: 240, lineRule: AUTO }, children: [T(m)] }));
        });
        preSection(kids);
      }
      // dedicatória e epígrafe ficam no pé da página: empurradas com linhas em branco
      // (o Word, o Google Docs e o LibreOffice respeitam; espaço antes do 1º parágrafo nem sempre)
      const pushDown = (lines) => {
        const n = lines.reduce((a, t) => a + Math.max(1, Math.ceil(t.length / 36)), 0);
        const lineCm = (prof.tamanho * 1.15 * cfg.entrelinha) / 72 * 2.54;
        const avail = 29.7 - prof.margens.sup - prof.margens.inf;
        const k = Math.max(0, Math.floor((avail - (n + 1) * lineCm) / lineCm) - 1);
        return Array.from({ length: k }, () => P({ spacing: { line: LINE, lineRule: AUTO }, children: [] }));
      };
      if (cfg.dedicatoria) {
        const ls = splitParas(cfg.dedicatoria);
        preSection([...pushDown(ls), ...ls.map((t) => P({ alignment: D.AlignmentType.JUSTIFIED, indent: { left: cm(prof.recuoNatureza) }, spacing: { line: LINE, lineRule: AUTO }, children: [T(t)] }))]);
      }
      if (cfg.tcc && cfg.agradecimentos) {
        preSection([preTitle('AGRADECIMENTOS'), ...splitParas(cfg.agradecimentos).map((t) => P({ style: 'Corpo', children: [T(t)] }))]);
      }
      if (cfg.epigrafe) {
        const ls = splitParas(cfg.epigrafe);
        preSection([...pushDown(ls), ...ls.map((t, k) => {
          const isAuthor = k === ls.length - 1 && ls.length > 1 && /^[(\-–—]/.test(t);
          return P({ alignment: isAuthor ? D.AlignmentType.RIGHT : D.AlignmentType.JUSTIFIED, indent: { left: cm(prof.recuoNatureza) }, spacing: { line: LINE, lineRule: AUTO }, children: [T(t, { italics: !isAuthor || undefined })] });
        })]);
      }
      if (cfg.resumo) preSection([preTitle('RESUMO'), ...resumoBlock('RESUMO', cfg.resumo, 'Palavras-chave', info.palavrasChave)]);
      if (cfg.abstract) preSection([preTitle('ABSTRACT'), ...resumoBlock('ABSTRACT', cfg.abstract, 'Keywords', info.keywords, 'en-US')]);
      if (cfg.sumario && tocEntries.length) {
        preSection([preTitle('SUMÁRIO'), new D.TableOfContents('Sumário', { headingStyleRange: '1-5', hyperlink: true, cachedEntries: tocEntries, beginDirty: true })]);
      }
      if (!body.length) body.push(P({ style: 'Corpo', children: [T('')] }));
      sections.push({
        properties: counted(pageProps()),
        headers: { default: numberHeader() },
        footers: { default: emptyFooter() },
        children: body,
      });
    }

    // ---------- estilos ----------
    const hStyle = (lvl) => {
      const st = prof.titulos[lvl] || {};
      const brk = lvl === 1 && cfg.novaPagina;
      return {
        run: { font: F, size: S2, bold: !!st.bold, italics: !!st.italic, color: '000000' },
        paragraph: { spacing: { before: brk ? 0 : LINE, after: LINE, line: LINE, lineRule: AUTO }, keepNext: true, keepLines: true, alignment: D.AlignmentType.LEFT },
      };
    };
    const tocStyle = (lvl) => {
      const st = prof.titulos[lvl] || {};
      return { id: 'TOC' + lvl, name: 'toc ' + lvl, basedOn: 'Normal', next: 'Normal', run: { bold: !!st.bold, italics: !!st.italic }, paragraph: { spacing: { line: LINE, lineRule: AUTO, before: 0, after: 0 }, indent: { left: 0 }, tabStops: [{ type: D.TabStopType.RIGHT, position: W_TW, leader: 'dot' }] } };
    };

    const doc = new D.Document({
      creator: autores[0] || 'Formata ABNT',
      title: titulo || 'Trabalho acadêmico',
      description: 'Formatado pelo Formata ABNT (' + prof.nome + ', ' + cfg.T.label + ')',
      features: cfg.sumario ? { updateFields: true } : undefined,
      footnotes,
      styles: {
        default: {
          document: { run: { font: F, size: S2, language: { value: 'pt-BR' } }, paragraph: { spacing: { line: LINE, lineRule: AUTO, before: 0, after: 0 } } },
          heading1: hStyle(1), heading2: hStyle(2), heading3: hStyle(3), heading4: hStyle(4), heading5: hStyle(5),
        },
        paragraphStyles: [
          { id: 'Corpo', name: 'Texto ABNT', basedOn: 'Normal', quickFormat: true, paragraph: { alignment: D.AlignmentType.JUSTIFIED, indent: { firstLine: cm(prof.recuoParagrafo) }, spacing: { line: LINE, lineRule: AUTO, before: 0, after: 0 } } },
          { id: 'CitacaoLonga', name: 'Citação longa ABNT', basedOn: 'Normal', quickFormat: true, run: { size: s2 }, paragraph: { alignment: D.AlignmentType.JUSTIFIED, indent: { left: cm(prof.recuoCitacao) }, spacing: { line: 240, lineRule: AUTO, before: 240, after: 240 } } },
          { id: 'Alinea', name: 'Alínea ABNT', basedOn: 'Normal', paragraph: { alignment: D.AlignmentType.JUSTIFIED, indent: { left: cm(2), hanging: cm(0.75) }, spacing: { line: LINE, lineRule: AUTO, before: 0, after: 0 }, tabStops: [{ type: D.TabStopType.LEFT, position: cm(2) }] } },
          { id: 'Legenda', name: 'Legenda ABNT', basedOn: 'Normal', run: { size: s2 }, paragraph: { alignment: center, spacing: { line: 240, lineRule: AUTO, before: 240, after: 0 }, keepNext: true } },
          { id: 'FonteIlustracao', name: 'Fonte da ilustração ABNT', basedOn: 'Normal', run: { size: s2 }, paragraph: { alignment: center, spacing: { line: 240, lineRule: AUTO, before: 0, after: 240 } } },
          { id: 'Imagem', name: 'Imagem ABNT', basedOn: 'Normal', paragraph: { alignment: center, spacing: { line: 240, lineRule: AUTO, before: 0, after: 0 } } },
          { id: 'Tabela', name: 'Texto de tabela ABNT', basedOn: 'Normal', run: { size: s2 }, paragraph: { spacing: { line: 240, lineRule: AUTO, before: 0, after: 0 } } },
          { id: 'Referencia', name: 'Referência ABNT', basedOn: 'Normal', paragraph: { alignment: D.AlignmentType.LEFT, spacing: { line: 240, lineRule: AUTO, before: cfg.refAntes, after: cfg.refDepois } } },
          { id: 'TituloPre', name: 'Título pré-textual ABNT', basedOn: 'Normal', run: { bold: true }, paragraph: { alignment: center, spacing: { line: LINE, lineRule: AUTO, before: 0, after: LINE * 2 } } },
          { id: 'Resumo', name: 'Resumo ABNT', basedOn: 'Normal', paragraph: { alignment: D.AlignmentType.JUSTIFIED, spacing: { line: 240, lineRule: AUTO, before: 0, after: 0 } } },
          { id: 'NotaRodape', name: 'Nota de rodapé ABNT', basedOn: 'Normal', run: { size: s2 }, paragraph: { alignment: D.AlignmentType.JUSTIFIED, spacing: { line: 240, lineRule: AUTO, before: 0, after: 0 } } },
          tocStyle(1), tocStyle(2), tocStyle(3), tocStyle(4), tocStyle(5),
        ],
      },
      sections,
    });

    const counts = {};
    plan.forEach((it) => { counts[it.type] = (counts[it.type] || 0) + 1; });
    return { doc, warnings, counts, toc: tocEntries, estimate: est, footnotes: fnCount, config: cfg };
  }

  root.FormataEngine = { PROFILES, TYPES, WORK_TYPES, resolveConfig, naturezaText, formatKeywords, mammothOptions, htmlToBlocks, textToBlocks, classify, buildPlan, buildDocument, estimatePages, imageInfo, capsPT };
})(typeof window !== 'undefined' ? window : globalThis);
