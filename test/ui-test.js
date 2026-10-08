// Roda o app num navegador simulado (jsdom): exemplo, importação, troca de tipo e geração do .docx
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const root = path.join(__dirname, '..');
let html = fs.readFileSync(path.join(root, 'dist/formata-abnt.html'), 'utf8');
const mam = fs.readFileSync(path.join(root, 'node_modules/mammoth/mammoth.browser.min.js'), 'utf8');
const dx = fs.readFileSync(path.join(root, 'node_modules/docx/dist/index.iife.js'), 'utf8');
const ffl = fs.readFileSync(path.join(root, 'node_modules/fflate/umd/index.js'), 'utf8');
html = html.replace(/<script src="[^"]*mammoth[^"]*"><\/script>/, () => '<script>' + mam + '</script>')
  .replace(/<script src="[^"]*docx[^"]*"><\/script>/, () => '<script>' + dx + '</script>')
  .replace(/<script src="[^"]*fflate[^"]*"><\/script>/, () => '<script>' + ffl + '</script>')
  .replace(/<link[^>]*>/g, '');
const errors = [];
const dom = new JSDOM('<!doctype html><html><head><meta charset="utf-8"></head><body>' + html + '</body></html>', {
  runScripts: 'dangerously', pretendToBeVisual: true, virtualConsole: new (require('jsdom').VirtualConsole)().on('error', (e) => console.log('CONSOLE.ERROR', e && (e.stack || e.message || e))), url: 'https://exemplo.test/',
  beforeParse(w) {
    w.addEventListener('error', (e) => errors.push(e.message));
    w.scrollTo = () => {}; w.setImmediate = setImmediate; w.clearImmediate = clearImmediate;
    w.HTMLCanvasElement.prototype.getContext = () => null;
  },
});
const w = dom.window, d = w.document;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  await sleep(200);
  const rows = () => d.querySelectorAll('#blist .row');
  d.querySelector('[data-go="estrutura"]').click();
  console.log('exemplo: linhas', rows().length, '| filtros:', [...d.querySelectorAll('#filters .chip')].map((c) => c.textContent.trim()).join(' / '));
  console.log('primeiras:', [...rows()].slice(0, 16).map((r) => r.querySelector('.tag').textContent + ' » ' + r.querySelector('.prev').textContent.slice(0, 40)).join('\n  '));

  // importa o TCC de teste
  const buf = fs.readFileSync(path.join(__dirname, 'tcc-aluno.docx'));
  const file = new w.File([buf], 'tcc-aluno.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  if (!file.arrayBuffer) file.arrayBuffer = async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length);
  const inp = d.querySelector('#fileInput');
  Object.defineProperty(inp, 'files', { value: [file], configurable: true });
  inp.dispatchEvent(new w.Event('change'));
  await sleep(1500);
  console.log('após importar: visão estrutura?', !d.querySelector('#v-estrutura').hidden, '| linhas', rows().length, '| banner oculto?', d.querySelector('#sampleBanner').hidden);

  // troca tipo de um parágrafo e aplica em sequência
  const target = [...rows()].find((r) => r.querySelector('.prev').textContent.startsWith('Entre as condições'));
  target.click();
  await sleep(100);
  console.log('folha aberta?', !d.querySelector('#sheet').hidden, '| opções', d.querySelectorAll('#typeGrid .type-opt').length);
  const opt = [...d.querySelectorAll('#typeGrid .type-opt')].find((b) => b.textContent.includes('Citação longa'));
  opt.click();
  await sleep(300);
  console.log('toast:', d.querySelector('#toast').textContent);
  // desfaz
  [...rows()].find((r) => r.querySelector('.prev').textContent.startsWith('Entre as condições')).click();
  await sleep(50);
  [...d.querySelectorAll('#typeGrid .type-opt')].find((b) => b.textContent.startsWith('Parágrafo')).click();
  await sleep(300);

  // filtro de referências
  [...d.querySelectorAll('#filters .chip')].find((c) => c.textContent.includes('Referências')).click();
  console.log('filtro refs:', rows().length);

  // preenche capa
  d.querySelector('[data-go="capa"]').click();
  const set = (id, v) => { const el = d.getElementById(id); el.value = v; el.dispatchEvent(new w.Event('input', { bubbles: true })); };
  set('f-autores', 'Maria Aparecida dos Santos');
  set('f-titulo', 'O brincar na Educação Infantil');
  set('f-orientador', 'Profa. Dra. Ana Paula Ribeiro');
  console.log('resumo veio do arquivo?', d.getElementById('f-resumo').value.slice(0, 40));

  // baixar (trabalho do semestre, padrão)
  let captured = null;
  w.URL.createObjectURL = (b) => { captured = b; return 'blob:x'; };
  w.URL.revokeObjectURL = () => {};
  w.HTMLAnchorElement.prototype.click = function () { console.log('download:', this.download); };
  async function gen(out) {
    captured = null;
    d.querySelector('[data-go="baixar"]').click();
    console.log('lede:', d.getElementById('baixarLede').textContent);
    console.log('partes:', [...d.querySelectorAll('#parts li')].map((l) => l.textContent.replace(/\s+/g, ' ').trim()).join(' | '));
    console.log('avisos:', [...d.querySelectorAll('#warns li')].map((l) => l.textContent).join(' | ') || '(nenhum)');
    console.log('regras:', [...d.querySelectorAll('#rules li')].map((l) => l.textContent).join(' | '));
    d.querySelector('#genBtn').click();
    for (let i = 0; i < 200 && !captured; i++) await sleep(100);
    if (!captured) { console.log('NENHUM ARQUIVO'); process.exitCode = 1; return; }
    const ab = await new Promise((res) => { const fr = new w.FileReader(); fr.onload = () => res(fr.result); fr.readAsArrayBuffer(captured); });
    fs.writeFileSync(path.join(__dirname, out), Buffer.from(ab));
    console.log('docx gerado:', out, Buffer.from(ab).length, 'bytes');
  }
  set('f-disciplina', 'Fundamentos da Educação Infantil');
  console.log('tipos:', [...d.querySelectorAll('#kinds .kind')].map((k) => k.querySelector('.nm').textContent + (k.getAttribute('aria-checked') === 'true' ? '*' : '')).join(' / '));
  console.log('rótulo professor:', d.getElementById('orientLabel').textContent, '| disciplina visível?', !d.getElementById('f-disciplina').closest('[data-show]').hidden);
  await gen('saida-ui-trabalho.docx');
  // troca para artigo
  d.querySelector('[data-go="texto"]').click();
  [...d.querySelectorAll('#kinds .kind')].find((k) => k.textContent.includes('Artigo')).click();
  await sleep(50);
  d.querySelector('[data-go="capa"]').click();
  console.log('artigo → rótulo:', d.getElementById('orientLabel').textContent, '| notas visíveis?', !d.getElementById('f-notaAutores').closest('[data-show]').hidden, '| disciplina oculta?', d.getElementById('f-disciplina').closest('[data-show]').hidden, '| entrelinha:', [...d.querySelectorAll('#lineSeg button')].find((b) => b.getAttribute('aria-pressed') === 'true').textContent, '| contagem:', d.getElementById('resumoCount').textContent);
  await gen('saida-ui-artigo.docx');

  // ---------- modelo da faculdade: plano de aula em tabela ----------
  const pbuf = fs.readFileSync(path.join(__dirname, 'plano-aula.docx'));
  const pfile = new w.File([pbuf], 'Plano de aula.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  if (!pfile.arrayBuffer) pfile.arrayBuffer = async () => pbuf.buffer.slice(pbuf.byteOffset, pbuf.byteOffset + pbuf.length);
  d.querySelector('[data-go="texto"]').click();
  Object.defineProperty(inp, 'files', { value: [pfile], configurable: true });
  inp.dispatchEvent(new w.Event('change'));
  await sleep(1500);
  const kindNow = [...d.querySelectorAll('#kinds .kind')].find((k) => k.getAttribute('aria-checked') === 'true');
  console.log('plano: tipo escolhido sozinho:', kindNow && kindNow.querySelector('.nm').textContent, '| toast:', d.querySelector('#toast').textContent);
  console.log('plano: tela de ajustes?', !d.getElementById('estruturaModelo').hidden, '| lista de parágrafos oculta?', d.getElementById('estruturaTexto').hidden);
  console.log('plano: arquivo:', [...d.querySelectorAll('#modelFile span')].map((x) => x.textContent).join(' · '));
  console.log('plano: ajustes:', [...d.querySelectorAll('#fixes li')].map((l) => l.querySelector('.lb').textContent + ' ' + l.querySelector('.ct').textContent).join(' | '));
  console.log('plano: avisos:', [...d.querySelectorAll('#modelWarns li')].map((l) => l.textContent).join(' | '));
  d.querySelector('[data-go="capa"]').click();
  console.log('plano: dados visíveis:', [...d.querySelectorAll('#v-capa .lab, #v-capa .toggle > div > div')].filter((el) => !el.closest('[hidden]')).map((el) => el.textContent.trim()).join(' / '));
  console.log('plano: fonte marcada:', [...d.querySelectorAll('#fontSeg button')].find((b) => b.getAttribute('aria-pressed') === 'true').textContent);
  await gen('saida-ui-modelo.docx');
  if (captured) {
    const ab = Buffer.from(await new Promise((res) => { const fr = new w.FileReader(); fr.onload = () => res(fr.result); fr.readAsArrayBuffer(captured); }));
    const xml = require('fflate').strFromU8(require('fflate').unzipSync(new Uint8Array(ab))['word/document.xml']);
    const sects = (xml.match(/<w:sectPr/g) || []).length, tbls = (xml.match(/<w:tbl>/g) || []).length;
    console.log('plano: arquivo gerado tem', tbls, 'tabelas e', sects, 'seção');
    if (sects !== 1 || tbls !== 2) process.exitCode = 1;
  }
  if (!kindNow || !/Modelo/.test(kindNow.textContent)) process.exitCode = 1;
  console.log('erros:', errors.length ? errors : 'nenhum');
  process.exit(errors.length ? 1 : process.exitCode || 0);
})().catch((e) => { console.error(e); process.exit(1); });
