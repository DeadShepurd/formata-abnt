(function () {
  'use strict';
  const E = window.FormataEngine;
  const PROF = E.PROFILES['unicid-pedagogia'];
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => Array.from(document.querySelectorAll(s));
  const KEY = 'formata-abnt:v2';

  // capacidade de baixar arquivos (dentro do Claude); fora dele, link comum
  const downloadsP = window.claude && typeof window.claude.use === 'function' ? window.claude.use('downloads').catch(() => null) : Promise.resolve(null);

  const typeDefaults = (tipo) => {
    const T = E.WORK_TYPES[tipo] || E.WORK_TYPES.trabalho;
    return { tipo: T.id, entrelinha: T.entrelinha, novaPagina: T.novaPagina, capa: T.capa, folhaRosto: T.folhaRosto, sumario: T.sumario };
  };
  const defaults = () => Object.assign({
    instituicao: PROF.instituicao, curso: PROF.curso, cidade: PROF.cidade, natureza: PROF.natureza,
    ano: String(new Date().getFullYear()), autores: '', titulo: '', subtitulo: '', disciplina: '', orientador: '', coorientador: '',
    resumo: '', palavrasChave: '', abstract: '', keywords: '', dedicatoria: '', agradecimentos: '', epigrafe: '',
    notaAutores: '', notaOrientador: '',
    aprovacao: false, banca: '', fonte: PROF.fonte, ordenarReferencias: true,
  }, typeDefaults('trabalho'));

  const S = { info: defaults(), blocks: [], notes: {}, source: { kind: 'sample', name: '' }, view: 'texto', filter: 'all', sheetId: null };

  // ---------- TCC de exemplo (fictício) ----------
  const SAMPLE = `
<p>UNIVERSIDADE CIDADE DE SÃO PAULO</p><p>Nome da Aluna</p><p>TÍTULO DO TRABALHO</p><p>São Paulo 2026</p>
<p><strong>RESUMO</strong></p><p>Este trabalho investiga o papel do brincar nas práticas pedagógicas da Educação Infantil, a partir de observações realizadas durante o estágio supervisionado.</p><p>Palavras-chave: brincar, educação infantil, ludicidade</p>
<p><strong>SUMÁRIO</strong></p><p>1 INTRODUÇÃO ........ 7</p><p>2 O BRINCAR NA EDUCAÇÃO INFANTIL ........ 9</p>
<p><strong>1. INTRODUÇÃO</strong></p>
<p>A brincadeira ocupa lugar central no desenvolvimento infantil, pois é por meio dela que a criança experimenta papéis sociais, elabora regras de convivência e constrói significados sobre o mundo que a cerca.</p>
<p>Ao observar o cotidiano de uma turma de crianças de quatro e cinco anos, percebe-se que os momentos de brincadeira livre revelam interesses, conflitos e aprendizagens que nem sempre aparecem nas atividades dirigidas.</p>
<h1>2 O BRINCAR NA EDUCAÇÃO INFANTIL</h1>
<p>Nas instituições de Educação Infantil, o brincar deixa de ser apenas passatempo e passa a ser compreendido como eixo estruturante das práticas pedagógicas, conforme orientam os documentos curriculares nacionais (BRASIL, 2018).</p>
<h2>2.1 O lúdico como linguagem da infância</h2>
<p>Quando o adulto reconhece a brincadeira como linguagem da infância, amplia suas possibilidades de escuta e de planejamento.</p>
<blockquote><p>A brincadeira cria uma zona de desenvolvimento na criança, pois nela a criança se comporta além do comportamento habitual de sua idade, além de seu comportamento cotidiano; no brinquedo, é como se ela fosse maior do que é na realidade. Como no foco de uma lente de aumento, a brincadeira contém todas as tendências do desenvolvimento sob forma condensada (VIGOTSKI, 2008, p. 35).</p></blockquote>
<p>Entre as condições observadas no estágio, destacam-se:</p>
<ul><li>cantos de atividades com materiais variados;</li><li>tempo prolongado para brincadeira livre;</li><li>mediação do professor sem interromper o enredo das crianças.</li></ul>
<p>Quadro 1 – Autores de referência sobre o brincar</p>
<table><tr><th>Autor</th><th>Contribuição</th></tr><tr><td>Vigotski</td><td>Brincadeira e desenvolvimento</td></tr><tr><td>Kishimoto</td><td>Jogo, brinquedo e brincadeira</td></tr></table>
<p>Fonte: Elaborado pela autora (2026).</p>
<p><strong>3 METODOLOGIA</strong></p>
<p>Trata-se de uma pesquisa qualitativa, com observação participante e registro em diário de campo ao longo de dez semanas.</p>
<p><strong>4 CONSIDERAÇÕES FINAIS</strong></p>
<p>A organização intencional de tempos, espaços e materiais, somada à escuta atenta do professor, amplia as aprendizagens que as crianças constroem ao brincar.</p>
<p><strong>REFERÊNCIAS</strong></p>
<p>VIGOTSKI, L. S. A brincadeira e o seu papel no desenvolvimento psíquico da criança. <strong>Revista Virtual de Gestão de Iniciativas Sociais</strong>, Rio de Janeiro, n. 8, p. 23-36, jun. 2008.</p>
<p>BRASIL. Ministério da Educação. <strong>Base Nacional Comum Curricular</strong>. Brasília, DF: MEC, 2018.</p>
<p>KISHIMOTO, T. M. (org.). <strong>Jogo, brinquedo, brincadeira e a educação</strong>. 14. ed. São Paulo: Cortez, 2017.</p>
<p><strong>APÊNDICE A – Roteiro de observação</strong></p>
<p>Data, duração, número de crianças envolvidas e materiais utilizados em cada brincadeira observada.</p>`;

  // ---------- armazenamento local (conveniência) ----------
  function save() {
    try {
      const data = { info: S.info, blocks: S.source.kind === 'sample' ? null : S.blocks, notes: S.notes, source: S.source };
      localStorage.setItem(KEY, JSON.stringify(data));
    } catch (e) {
      try { localStorage.setItem(KEY, JSON.stringify({ info: S.info, source: { kind: 'sample', name: '' } })); } catch (e2) { /* sem armazenamento */ }
    }
  }
  function restore() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return false;
      const d = JSON.parse(raw);
      if (d.info) S.info = Object.assign(defaults(), d.info);
      if (d.blocks && d.blocks.length) { S.blocks = d.blocks; S.notes = d.notes || {}; S.source = d.source || { kind: 'docx', name: '' }; return true; }
    } catch (e) { /* ignora */ }
    return false;
  }

  // ---------- carregar texto ----------
  function load(parsed, source) {
    const cls = E.classify(parsed);
    S.blocks = cls.blocks;
    S.notes = cls.notes || {};
    S.source = source;
    const ex = cls.extracted || {};
    if (source.kind !== 'sample') {
      ['resumo', 'palavrasChave', 'abstract', 'keywords', 'agradecimentos'].forEach((k) => {
        if (ex[k] && !String(S.info[k] || '').trim()) S.info[k] = ex[k];
      });
    }
    fillForm();
    save();
    renderAll();
  }

  function loadSample() {
    load(E.htmlToBlocks(SAMPLE, new DOMParser()), { kind: 'sample', name: '' });
  }

  async function importFile(file) {
    const st = $('#fileStatus');
    st.hidden = false;
    if (!/\.docx$/i.test(file.name)) {
      st.textContent = 'Esse arquivo não é .docx. No Word use Salvar como › Documento do Word (.docx); no Google Docs, Arquivo › Baixar › Microsoft Word.';
      return;
    }
    st.textContent = 'Lendo ' + file.name + '…';
    try {
      if (!window.mammoth) throw new Error('leitor não carregou');
      const buf = await file.arrayBuffer();
      const res = await window.mammoth.convertToHtml({ arrayBuffer: buf }, E.mammothOptions(window.mammoth));
      const parsed = E.htmlToBlocks(res.value, new DOMParser());
      if (!parsed.blocks.length) { st.textContent = 'Não encontrei texto em ' + file.name + '.'; return; }
      load(parsed, { kind: 'docx', name: file.name });
      st.textContent = file.name;
      toast('Texto importado. Confira a estrutura.');
      go('estrutura');
    } catch (e) {
      st.textContent = 'Não consegui ler ' + file.name + '. Se ele estiver aberto em outro app, feche e tente de novo.';
    }
  }

  // ---------- navegação ----------
  function go(view) {
    S.view = view;
    $$('[data-view]').forEach((s) => { s.hidden = s.dataset.view !== view; });
    $$('.nav button').forEach((b) => { if (b.dataset.go === view) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
    if (view === 'estrutura') renderList();
    if (view === 'baixar') renderBaixar();
    window.scrollTo(0, 0);
    try { history.replaceState(null, '', '#' + view); } catch (e) { /* ok */ }
  }

  // ---------- estrutura ----------
  const GROUPS = [
    { id: 'all', label: 'Tudo' },
    { id: 'titulo', label: 'Títulos', cls: 'g-titulo' },
    { id: 'texto', label: 'Texto', cls: 'g-texto' },
    { id: 'citacao', label: 'Citações', cls: 'g-citacao' },
    { id: 'ilustra', label: 'Ilustrações', cls: 'g-ilustra' },
    { id: 'ref', label: 'Referências', cls: 'g-ref' },
    { id: 'skip', label: 'Ignorados', cls: 'g-skip' },
  ];
  const groupOf = (t) => (E.TYPES[t] ? E.TYPES[t].group : 'texto');

  function numbers() {
    const plan = E.buildPlan(S.blocks, PROF, { ordenarReferencias: false });
    const m = {};
    plan.forEach((it) => { if (it.number) m[it.block.id] = it.number; });
    return m;
  }

  function renderFilters() {
    const counts = { all: S.blocks.length };
    S.blocks.forEach((b) => { const g = groupOf(b.type); counts[g] = (counts[g] || 0) + 1; });
    const box = $('#filters');
    box.innerHTML = '';
    GROUPS.forEach((g) => {
      if (g.id !== 'all' && !counts[g.id]) return;
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'chip ' + (g.cls || '');
      b.setAttribute('aria-pressed', String(S.filter === g.id));
      b.innerHTML = (g.cls ? '<span class="dot" style="color:var(--c)"></span>' : '') + '<span></span><span class="n"></span>';
      b.children[g.cls ? 1 : 0].textContent = g.label;
      b.querySelector('.n').textContent = counts[g.id] || 0;
      b.addEventListener('click', () => { S.filter = g.id; renderList(); });
      box.appendChild(b);
    });
  }

  function renderList() {
    renderFilters();
    const ul = $('#blist');
    ul.innerHTML = '';
    const nums = numbers();
    const frag = document.createDocumentFragment();
    let shown = 0;
    S.blocks.forEach((b) => {
      const g = groupOf(b.type);
      if (S.filter !== 'all' && g !== S.filter) return;
      shown++;
      const li = document.createElement('li');
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'row';
      row.dataset.id = b.id;
      const T = E.TYPES[b.type] || E.TYPES.p;
      const tag = document.createElement('span');
      tag.className = 'tag g-' + g + (b.type !== b.auto ? ' changed' : '');
      const lbl = document.createElement('span');
      lbl.className = 'lbl';
      lbl.textContent = T.short;
      tag.appendChild(lbl);
      if (nums[b.id]) { const n = document.createElement('span'); n.className = 'num'; n.textContent = nums[b.id]; tag.appendChild(n); }
      const prev = document.createElement('span');
      prev.className = 'prev';
      if (/^h[1-5]$/.test(b.type)) {
        const st = PROF.titulos[b.type[1]] || {};
        prev.classList.add(st.bold ? 'h' : 'x');
        if (st.caps) prev.classList.add('caps');
        prev.textContent = b.text.replace(/^\s*\d{1,2}(\.\d{1,2}){0,4}\.?\s*[-–—]?\s*/, '') || b.text;
      } else if (b.type === 'post') { prev.classList.add('h', 'caps'); prev.textContent = b.text; }
      else if (b.type === 'quote') { prev.classList.add('quote'); prev.textContent = b.text; }
      else if (b.type === 'skip') { prev.classList.add('skip'); prev.textContent = b.text; }
      else if (b.type === 'ref') { prev.classList.add('ref'); prev.textContent = b.text; }
      else if (b.kind === 'img') { prev.classList.add('media'); prev.textContent = '[imagem] ' + (b.alt || ''); }
      else if (b.kind === 'table') { prev.classList.add('media'); prev.textContent = '[tabela, ' + b.rows.length + ' linhas] ' + b.text; }
      else prev.textContent = b.text;
      const sr = document.createElement('span');
      sr.className = 'sr';
      sr.textContent = '. Toque para mudar o tipo.';
      row.append(tag, prev, sr);
      row.addEventListener('click', () => openSheet(b.id));
      li.appendChild(row);
      frag.appendChild(li);
    });
    if (!shown) { const li = document.createElement('li'); li.className = 'empty'; li.textContent = 'Nada neste filtro.'; frag.appendChild(li); }
    ul.appendChild(frag);
  }

  // ---------- folha de tipos ----------
  let lastFocus = null;
  function openSheet(id) {
    const b = S.blocks.find((x) => x.id === id);
    if (!b) return;
    S.sheetId = id;
    lastFocus = document.activeElement;
    $('#sheetText').textContent = b.kind === 'img' ? '[imagem]' : b.kind === 'table' ? '[tabela] ' + b.text : b.text;
    $('#applyRun').checked = false;
    const grid = $('#typeGrid');
    grid.innerHTML = '';
    const opts = b.kind === 'img' ? ['img', 'skip'] : b.kind === 'table' ? ['table', 'skip'] : ['h1', 'h2', 'h3', 'h4', 'h5', 'post', 'p', 'quote', 'item', 'caption', 'source', 'ref', 'skip'];
    opts.forEach((t) => {
      const T = E.TYPES[t];
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'type-opt g-' + T.group;
      btn.style.setProperty('--c', 'var(--t-' + T.group + ')');
      btn.setAttribute('aria-pressed', String(b.type === t));
      const tt = document.createElement('span'); tt.className = 't';
      tt.innerHTML = '<span class="dot"></span>';
      tt.appendChild(document.createTextNode(T.label));
      btn.appendChild(tt);
      const hint = T.hint || (t === 'h1' ? 'Ex.: 1 INTRODUÇÃO' : t === 'h2' ? 'Ex.: 1.1 OBJETIVOS' : t === 'h3' ? 'Ex.: 1.1.1 Objetivo geral' : t === 'p' ? 'Texto corrido, recuo de 1,25 cm' : t === 'item' ? 'a), b), c)…' : t === 'ref' ? 'Alinhada à esquerda, espaço simples' : '');
      if (hint) { const h = document.createElement('span'); h.className = 'h'; h.textContent = hint; btn.appendChild(h); }
      btn.addEventListener('click', () => setType(id, t));
      grid.appendChild(btn);
    });
    const sh = $('#sheet'), sc = $('#scrim');
    sh.hidden = false; sc.hidden = false;
    requestAnimationFrame(() => { sh.classList.add('open'); sc.classList.add('open'); });
    setTimeout(() => { const cur = grid.querySelector('[aria-pressed="true"]'); (cur || $('#sheetClose')).focus(); }, 60);
  }
  function closeSheet() {
    const sh = $('#sheet'), sc = $('#scrim');
    sh.classList.remove('open'); sc.classList.remove('open');
    setTimeout(() => { sh.hidden = true; sc.hidden = true; }, 220);
    S.sheetId = null;
    if (lastFocus && document.contains(lastFocus)) lastFocus.focus();
  }
  function setType(id, t) {
    const i = S.blocks.findIndex((x) => x.id === id);
    if (i < 0) return;
    S.blocks[i].type = t;
    let n = 1;
    if ($('#applyRun').checked) {
      for (let k = i + 1; k < S.blocks.length; k++) {
        const nb = S.blocks[k];
        if (/^(h[1-5]|post)$/.test(nb.type) || nb.kind !== 'p') break;
        nb.type = t; n++;
      }
    }
    save();
    closeSheet();
    renderList();
    const row = document.querySelector('.row[data-id="' + id + '"]');
    if (row) row.focus();
    toast(n > 1 ? n + ' parágrafos marcados como ' + E.TYPES[t].label.toLowerCase() : 'Marcado como ' + E.TYPES[t].label.toLowerCase());
  }

  // ---------- formulário ----------
  function fillForm() {
    $$('#infoForm [data-k]').forEach((el) => {
      const k = el.dataset.k;
      if (el.type === 'checkbox') el.checked = !!S.info[k];
      else el.value = S.info[k] == null ? '' : S.info[k];
    });
    $('#bancaField').hidden = !S.info.aprovacao;
    $$('#fontSeg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.font === S.info.fonte)));
    $$('#lineSeg button').forEach((b) => b.setAttribute('aria-pressed', String(Number(b.dataset.ent) === Number(S.info.entrelinha))));
    updateTypeUI();
  }

  // ---------- tipo de trabalho ----------
  const KIND_ORDER = ['trabalho', 'artigo', 'tcc'];
  function renderKinds() {
    const box = $('#kinds');
    box.innerHTML = '';
    KIND_ORDER.forEach((id) => {
      const T = E.WORK_TYPES[id];
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'kind';
      b.setAttribute('role', 'radio');
      b.setAttribute('aria-checked', String(S.info.tipo === id));
      b.innerHTML = '<span class="rd" aria-hidden="true"></span><span><span class="nm"></span><span class="later"></span><br><span class="ht"></span></span>';
      b.querySelector('.nm').textContent = T.label;
      b.querySelector('.ht').textContent = T.hint;
      if (id === 'tcc') b.querySelector('.later').textContent = 'fim do curso';
      b.addEventListener('click', () => setKind(id));
      box.appendChild(b);
    });
  }
  function setKind(id) {
    if (S.info.tipo === id) return;
    Object.assign(S.info, typeDefaults(id));
    save();
    renderKinds();
    fillForm();
    toast('Tipo: ' + E.WORK_TYPES[id].label);
  }
  function updateTypeUI() {
    const t = S.info.tipo || 'trabalho';
    $$('[data-show]').forEach((el) => { el.hidden = el.dataset.show.split(' ').indexOf(t) < 0; });
    $('#kindName').textContent = E.WORK_TYPES[t].label;
    $('#orientLabel').textContent = t === 'trabalho' ? 'Professor(a) da disciplina' : 'Orientador(a)';
    $('#resumoSec').textContent = t === 'artigo' ? 'Resumo' : t === 'tcc' ? 'Resumo' : 'Resumo (se o professor pedir)';
    $('#dadosLede').textContent = t === 'artigo'
      ? 'No artigo, título, autores e resumo vão na primeira página, antes da introdução.'
      : t === 'tcc' ? 'Esses dados montam capa, folha de rosto e os elementos antes da introdução.'
      : 'Esses dados montam a capa e o que vem antes da introdução.';
    $('#lineHelp').textContent = t === 'artigo' ? 'O manual da UNICID recomenda espaço simples no artigo.' : 'A ABNT usa 1,5 em trabalhos acadêmicos. Use simples se o professor pedir.';
    wordCount();
  }
  function wordCount() {
    const n = (S.info.resumo || '').trim().split(/\s+/).filter(Boolean).length;
    const t = S.info.tipo;
    const alvo = t === 'artigo' ? ' · o manual da UNICID pede de 100 a 250' : t === 'tcc' ? ' · a NBR 6028 pede de 150 a 500' : '';
    $('#resumoCount').textContent = n ? n + (n === 1 ? ' palavra' : ' palavras') + alvo : (alvo ? alvo.replace(' · ', '').replace(/^./, (c) => c.toUpperCase()) + ' palavras.' : '');
  }
  function bindForm() {
    $('#infoForm').addEventListener('input', (e) => {
      const el = e.target;
      if (!el.dataset || !el.dataset.k) return;
      S.info[el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value;
      if (el.dataset.k === 'aprovacao') $('#bancaField').hidden = !el.checked;
      save();
    });
    $('#infoForm').addEventListener('submit', (e) => e.preventDefault());
    $$('#fontSeg button').forEach((b) => b.addEventListener('click', () => { S.info.fonte = b.dataset.font; fillForm(); save(); }));
    $$('#lineSeg button').forEach((b) => b.addEventListener('click', () => { S.info.entrelinha = Number(b.dataset.ent); fillForm(); save(); }));
    $('#f-resumo').addEventListener('input', wordCount);
    $('#kindChange').addEventListener('click', () => go('texto'));
  }

  // ---------- baixar ----------
  function rulesList() {
    const cfg = E.resolveConfig(PROF, S.info);
    const ent = cfg.entrelinha === 1 ? 'simples' : String(cfg.entrelinha).replace('.', ',');
    return [
      ['Papel', 'A4'],
      ['Fonte', cfg.fonte + ' ' + PROF.tamanho],
      ['Entrelinha do texto', ent],
      ['Recuo do parágrafo', String(PROF.recuoParagrafo).replace('.', ',') + ' cm'],
      ['Citação longa', '4 cm · ' + PROF.tamanhoMenor + ' pt · simples'],
      ['Notas, legendas, fontes', PROF.tamanhoMenor + ' pt'],
      ['Referências', 'simples · 6 pt antes e depois'],
      ['Palavras-chave', 'separadas por ponto'],
      ['Número da página', 'sup. direito, ' + PROF.tamanhoMenor + ' pt'],
      ['Numeração aparece', cfg.art ? 'da 2ª página' : 'a partir da introdução'],
    ];
  }
  const splitCount = (t) => String(t || '').split(/\n+/).map((x) => x.trim()).filter(Boolean).length;
  function analysis() {
    const plan = E.buildPlan(S.blocks, PROF, { ordenarReferencias: false });
    const c = {};
    plan.forEach((it) => { c[it.type] = (c[it.type] || 0) + 1; });
    const heads = ['h1', 'h2', 'h3', 'h4', 'h5'].reduce((a, k) => a + (c[k] || 0), 0);
    const post = plan.filter((it) => it.type === 'post').map((it) => it.title);
    return { c, heads, post, plan };
  }
  function renderBaixar() {
    const I = S.info, A = analysis();
    const cfg = E.resolveConfig(PROF, I);
    const has = (k) => !!String(I[k] || '').trim();
    const nT = A.heads + A.post.length;
    const parts = [];
    if (cfg.capa || !cfg.art) parts.push(['Capa', cfg.capa, has('titulo') ? '' : 'sem título']);
    if (cfg.trab) parts.push(['Folha de rosto', cfg.folhaRosto, '']);
    if (cfg.tcc) parts.push(['Folha de rosto', true, has('orientador') ? '' : 'sem orientador'], ['Folha de aprovação', cfg.aprovacao, ''], ['Dedicatória', !!cfg.dedicatoria, ''], ['Agradecimentos', !!cfg.agradecimentos, ''], ['Epígrafe', !!cfg.epigrafe, '']);
    if (cfg.art) parts.push(['Título e autores na 1ª página', true, (splitCount(I.autores) || 0) + (has('orientador') ? 1 : 0) + ' com nota']);
    parts.push(['Resumo', !!cfg.resumo, ''], ['Abstract', !!cfg.abstract, '']);
    if (!cfg.art) parts.push(['Sumário', cfg.sumario && nT > 0, nT + ' títulos']);
    parts.push(
      ['Texto', (A.c.p || 0) > 0, (A.c.p || 0) + ' parágrafos'],
      ['Citações longas', (A.c.quote || 0) > 0, String(A.c.quote || 0)],
      ['Ilustrações e tabelas', (A.c.img || 0) + (A.c.table || 0) > 0, String((A.c.img || 0) + (A.c.table || 0))],
      ['Referências', (A.c.ref || 0) > 0, String(A.c.ref || 0)]
    );
    if (cfg.art) parts.push(['Agradecimentos', !!cfg.agradecimentos, '']);
    const ul = $('#parts');
    ul.innerHTML = '';
    parts.forEach(([nm, yes, meta]) => {
      const li = document.createElement('li');
      li.className = yes ? 'yes' : 'no';
      li.innerHTML = '<span class="st" aria-hidden="true"></span><span class="nm"></span><span class="meta"></span>';
      li.querySelector('.st').textContent = yes ? '✓' : '–';
      li.querySelector('.nm').textContent = nm;
      li.querySelector('.meta').textContent = yes ? meta : 'não incluído';
      ul.appendChild(li);
    });

    const warns = [];
    if (S.source.kind === 'sample') warns.push(['Você ainda está com o TCC de exemplo.', 'Importar', 'texto']);
    if (!has('autores')) warns.push(['Falta o nome do autor.', 'Preencher', 'capa']);
    if (!has('titulo')) warns.push(['Falta o título do trabalho.', 'Preencher', 'capa']);
    if (!A.heads) warns.push(['Nenhum título de seção foi marcado.', 'Conferir', 'estrutura']);
    if (!A.c.ref) warns.push(['Nenhuma referência foi marcada.', 'Conferir', 'estrutura']);
    if (has('resumo') && !has('palavrasChave')) warns.push(['O resumo está sem palavras-chave.', 'Preencher', 'capa']);
    if (cfg.art && !has('resumo')) warns.push(['O manual da UNICID pede resumo no artigo.', 'Preencher', 'capa']);
    if (cfg.trab && cfg.capa && !has('disciplina')) warns.push(['Falta o nome da disciplina.', 'Preencher', 'capa']);
    if (!has('orientador')) warns.push([cfg.trab ? 'Falta o nome do professor(a).' : 'Falta o nome do orientador(a).', 'Preencher', 'capa']);
    const wl = $('#warns');
    wl.innerHTML = '';
    warns.forEach(([msg, act, view]) => {
      const li = document.createElement('li');
      const sp = document.createElement('span'); sp.textContent = msg;
      const bt = document.createElement('button'); bt.type = 'button'; bt.className = 'link'; bt.textContent = act;
      bt.addEventListener('click', () => go(view));
      li.append(sp, bt);
      wl.appendChild(li);
    });
    $('#warnWrap').hidden = !warns.length;

    const rl = $('#rules');
    rl.innerHTML = '';
    rulesList().forEach(([k, v]) => {
      const li = document.createElement('li');
      li.innerHTML = '<span></span><span></span>';
      li.children[0].textContent = k; li.children[1].textContent = v;
      rl.appendChild(li);
    });
    const tl = E.WORK_TYPES[cfg.tipo].label;
    $('#baixarLede').textContent = tl + (S.source.kind === 'docx' ? ' · ' + S.source.name : S.source.kind === 'paste' ? ' · texto colado' : ' · confira o que vai no arquivo e gere o .docx');
  }

  // medição de texto para calcular as páginas do sumário
  const ctx = document.createElement('canvas').getContext('2d');
  let lastFont = '';
  function measure(text, f) {
    const fam = f.name === 'Times New Roman' ? "Tinos, 'Times New Roman', 'Liberation Serif', serif" : "Arimo, Arial, 'Liberation Sans', Helvetica, sans-serif";
    const font = (f.italic ? 'italic ' : '') + (f.bold ? '700 ' : '400 ') + f.size + 'px ' + fam;
    if (font !== lastFont) { ctx.font = font; lastFont = font; }
    return ctx.measureText(text).width;
  }
  async function fontsReady(name) {
    if (!document.fonts || !document.fonts.load) return;
    const fam = name === 'Times New Roman' ? 'Tinos' : 'Arimo';
    const faces = ['400 12px ', '700 12px ', 'italic 400 12px ', 'italic 700 12px '].map((p) => document.fonts.load(p + fam));
    await Promise.race([Promise.all(faces).catch(() => null), new Promise((r) => setTimeout(r, 2500))]);
  }


  // ---------- app instalado (Android): salvar e compartilhar ----------
  const isNative = () => !!(window.Capacitor && typeof window.Capacitor.isNativePlatform === 'function' && window.Capacitor.isNativePlatform());
  const nativePlugin = (name) => {
    const cap = window.Capacitor;
    return (cap.Plugins && cap.Plugins[name]) || (typeof cap.registerPlugin === 'function' ? cap.registerPlugin(name) : null);
  };
  function blobToBase64(blob) {
    return new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => res(String(fr.result).split(',')[1] || '');
      fr.onerror = () => rej(fr.error);
      fr.readAsDataURL(blob);
    });
  }
  async function saveNative(blob, filename) {
    if (!isNative()) return false;
    const Fs = nativePlugin('Filesystem');
    const Share = nativePlugin('Share');
    if (!Fs) return false;
    const data = await blobToBase64(blob);
    let savedCopy = false;
    try {
      await Fs.writeFile({ path: 'Formata ABNT/' + filename, data, directory: 'DOCUMENTS', recursive: true });
      savedCopy = true;
    } catch (e) { /* sem acesso à pasta Documentos neste aparelho */ }
    const tmp = await Fs.writeFile({ path: filename, data, directory: 'CACHE' });
    if (Share) {
      try {
        await Share.share({ title: filename, files: [tmp.uri], dialogTitle: 'Abrir, salvar ou enviar o trabalho' });
      } catch (e) { /* a pessoa fechou a janela de compartilhar */ }
    }
    toast(savedCopy ? 'Arquivo salvo em Documentos › Formata ABNT' : 'Arquivo pronto: ' + filename);
    return true;
  }

  function slug(s) {
    return (s || 'trabalho').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'trabalho';
  }

  let busy = false;
  async function generate() {
    if (busy) return;
    if (!window.docx) { toast('O gerador ainda está carregando. Tente de novo em instantes.'); return; }
    busy = true;
    const btn = $('#genBtn');
    btn.disabled = true;
    const label = btn.querySelector('span');
    const old = label.textContent;
    label.textContent = 'Gerando…';
    try {
      await fontsReady(S.info.fonte);
      const res = E.buildDocument({ profile: PROF, info: S.info, blocks: S.blocks, notes: S.notes }, { measure: ctx ? measure : null });
      const blob = await window.docx.Packer.toBlob(res.doc);
      const filename = slug(S.info.titulo) + ' - ABNT.docx';
      res.warnings.forEach((w) => toast(w));
      const dl = await downloadsP;
      if (dl) {
        try {
          const r = await dl.save({ filename, data: blob });
          toast(r && r.status === 'delivered' ? 'Arquivo enviado.' : 'Arquivo salvo: ' + filename);
        } catch (e) {
          const code = e && e.code;
          if (code === 'declined') toast('Download cancelado.');
          else if (code === 'rate_limited') toast('Já há um download aguardando confirmação.');
          else toast('Não foi possível salvar o arquivo aqui.');
        }
      } else if (await saveNative(blob, filename)) {
        // salvo pelo Android (app instalado)
      } else {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
        toast('Arquivo gerado: ' + filename);
      }
    } catch (e) {
      console.error(e);
      toast('Algo deu errado ao montar o arquivo. Confira se há títulos e texto marcados.');
    } finally {
      busy = false;
      btn.disabled = false;
      label.textContent = old;
    }
  }

  // ---------- avisos rápidos ----------
  let toastT = null;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastT);
    toastT = setTimeout(() => { t.hidden = true; }, 3200);
  }

  function renderAll() {
    $('#sampleBanner').hidden = S.source.kind !== 'sample';
    const st = $('#fileStatus');
    if (S.source.kind === 'docx' && S.source.name) { st.hidden = false; st.textContent = S.source.name; }
    if (S.view === 'estrutura') renderList();
    if (S.view === 'baixar') renderBaixar();
  }

  // ---------- eventos ----------
  function bind() {
    $$('.nav button').forEach((b) => b.addEventListener('click', () => go(b.dataset.go)));
    $('#pickBtn').addEventListener('click', () => $('#fileInput').click());
    $('#fileInput').addEventListener('change', (e) => { const f = e.target.files && e.target.files[0]; if (f) importFile(f); e.target.value = ''; });
    const drop = $('#drop');
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) importFile(f); });

    let pastedHTML = null, pastedText = null;
    const pa = $('#pasteArea');
    pa.addEventListener('paste', (e) => {
      const cd = e.clipboardData;
      if (!cd) return;
      const h = cd.getData('text/html');
      pastedHTML = h && h.length > 20 ? h : null;
      setTimeout(() => { pastedText = pa.value; $('#usePaste').disabled = !pa.value.trim(); }, 0);
    });
    pa.addEventListener('input', () => { $('#usePaste').disabled = !pa.value.trim(); });
    $('#usePaste').addEventListener('click', () => {
      const txt = pa.value;
      if (!txt.trim()) return;
      const parsed = pastedHTML && pastedText === txt ? E.htmlToBlocks(pastedHTML, new DOMParser()) : E.textToBlocks(txt);
      if (!parsed.blocks.length) { toast('Não encontrei texto para usar.'); return; }
      load(parsed, { kind: 'paste', name: '' });
      toast('Texto carregado. Confira a estrutura.');
      go('estrutura');
    });

    $('#sheetClose').addEventListener('click', closeSheet);
    $('#scrim').addEventListener('click', closeSheet);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && S.sheetId) closeSheet(); });
    $('#genBtn').addEventListener('click', generate);
    bindForm();
  }

  // ---------- início ----------
  bind();
  if (!restore()) loadSample(); else { fillForm(); }
  renderKinds();
  fillForm();
  const h = (location.hash || '').replace('#', '');
  go(['texto', 'estrutura', 'capa', 'baixar'].includes(h) ? h : 'texto');
  renderAll();
})();
