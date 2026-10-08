// Testes automáticos do modo "Modelo da faculdade" com o plano de aula fictício.
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { JSDOM } = require('jsdom');
globalThis.fflate = require('fflate');
require('../src/modelo.js');
const M = globalThis.FormataModelo;
const { window } = new JSDOM('');
const env = { DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer };

const input = new Uint8Array(fs.readFileSync(path.join(__dirname, 'plano-aula.docx')));
const xmlOf = (bytes, name) => globalThis.fflate.strFromU8(globalThis.fflate.unzipSync(bytes)[name]);
const words = (xml) => (xml.replace(/<\/w:p>/g, ' ').match(/<w:t[^>]*>([^<]*)<\/w:t>|\s/g) || []).map((m) => m.replace(/<[^>]+>/g, '')).join('').match(/[\wÀ-ÿºª°]+/g) || [];
const count = (arr) => arr.reduce((m, w) => (m[w] = (m[w] || 0) + 1, m), {});

let ok = 0;
const check = (name, fn) => { fn(); ok++; console.log('  ✔', name); };

const dry = M.analyze(input, {}, env);
check('reconhece o arquivo como modelo em tabela', () => {
  assert.strictEqual(dry.stats.templateLike, true);
  assert.strictEqual(dry.stats.landscape, true);
  assert.strictEqual(dry.stats.suggestedFont, 'Arial');
});

for (const fonte of ['Arial', 'Times New Roman']) {
  const { bytes, report } = M.process(input, { fonte }, env);
  const doc = xmlOf(bytes, 'word/document.xml');
  const fix = (k) => report.fixes.find((f) => f.key === k);
  console.log(' fonte', fonte);
  check('XML válido', () => {
    const d = new window.DOMParser().parseFromString(doc, 'application/xml');
    assert.strictEqual(d.getElementsByTagName('parsererror').length, 0);
  });
  check('tabelas cortadas viram uma só e a página continua deitada', () => {
    assert.strictEqual((doc.match(/<w:tbl>/g) || []).length, 2);
    assert.strictEqual((doc.match(/<w:sectPr/g) || []).length, 1);
    assert.ok(/w:orient="landscape"/.test(doc));
    assert.ok(/<w:headerReference [^>]*w:type="default"/.test(doc), 'cabeçalho com o logo continua ligado');
  });
  check('frases cortadas entre páginas e no meio da linha foram juntadas', () => {
    assert.strictEqual(fix('split').count, 2);
    assert.strictEqual(fix('lines').count, 2);
    assert.ok(/e nas<\/w:t>.*?<w:t[^>]*> <\/w:t>.*?<w:t[^>]*>brincadeiras do recreio/s.test(doc));
  });
  check('linhas numeradas (1, 2) não são confundidas com continuação', () => {
    const text = words(doc).join(' ');
    assert.ok(/\b1 Conversa inicial/.test(text) && /\b2 Montagem do experimento/.test(text));
  });
  check('uma fonte e um tamanho só', () => {
    const fonts = new Set((doc.match(/w:ascii="([^"]+)"/g) || []).map((s) => s.slice(9, -1)));
    assert.deepStrictEqual([...fonts], [fonte]);
    assert.deepStrictEqual([...new Set((doc.match(/<w:sz w:val="(\d+)"/g) || []))], ['<w:sz w:val="24"']);
    assert.ok(!/<w:spacing w:val=/.test(doc), 'sem letras comprimidas');
  });
  check('referências em ordem alfabética e com espaçamento ABNT', () => {
    const iA = doc.indexOf('ALVES'), iB = doc.indexOf('BRASIL. Minist');
    assert.ok(iA > 0 && iB > iA);
    assert.strictEqual(fix('refs').count, 2);
  });
  check('nenhuma palavra some ou aparece (só ° → º)', () => {
    const a = count(words(xmlOf(input, 'word/document.xml'))), b = count(words(doc));
    const diff = Object.keys(Object.assign({}, a, b)).filter((k) => (a[k] || 0) !== (b[k] || 0));
    assert.deepStrictEqual(diff.sort(), ['2°', '2º'].sort());
  });
  check('avisa o trecho que termina no meio', () => {
    assert.ok(report.warnings.some((w) => w.kind === 'cortado' && /fita de$/.test(w.text)));
  });
  check('cabeçalho, imagens e numeração ficam iguais', () => {
    const o = globalThis.fflate.unzipSync(input), n = globalThis.fflate.unzipSync(bytes);
    for (const k of Object.keys(o)) if (/header|media|numbering|_rels/.test(k)) assert.ok(Buffer.from(o[k]).equals(Buffer.from(n[k])), k);
  });
}
console.log('modo modelo:', ok, 'verificações passaram');
