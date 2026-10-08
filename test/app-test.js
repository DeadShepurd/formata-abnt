// Abre a versão do app instalado (www/index.html) num navegador simulado e
// confere que as bibliotecas locais carregam sem internet.
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');
const file = path.join(__dirname, '..', 'www', 'index.html');
const errors = [];
const vc = new VirtualConsole();
vc.on('jsdomError', (e) => errors.push(e.message));
vc.on('error', (e) => errors.push(String(e && (e.message || e))));
JSDOM.fromFile(file, {
  runScripts: 'dangerously', resources: 'usable', pretendToBeVisual: true, virtualConsole: vc,
  beforeParse(w) { w.scrollTo = () => {}; w.setImmediate = setImmediate; w.HTMLCanvasElement.prototype.getContext = () => null; },
}).then(async (dom) => {
  const w = dom.window;
  await new Promise((r) => w.addEventListener('load', r));
  await new Promise((r) => setTimeout(r, 300));
  const ok = { docx: !!w.docx, mammoth: !!w.mammoth, motor: !!w.FormataEngine, tipos: w.document.querySelectorAll('#kinds .kind').length };
  w.document.querySelector('[data-go="estrutura"]').click();
  ok.paragrafos = w.document.querySelectorAll('#blist .row').length;
  console.log('app instalado:', JSON.stringify(ok), errors.length ? 'ERROS: ' + errors.join(' | ') : 'sem erros');
  const fail = !ok.docx || !ok.mammoth || !ok.motor || ok.tipos !== 3 || !ok.paragrafos || errors.length;
  process.exit(fail ? 1 : 0);
});
