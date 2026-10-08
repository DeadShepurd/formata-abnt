// Roda o modo "Modelo da faculdade" num .docx e mostra o relatório.
// uso: node test/modelo-run.js entrada.docx saida.docx [Arial|"Times New Roman"]
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
globalThis.fflate = require('fflate');
require('../src/modelo.js');
const M = globalThis.FormataModelo;
const { window } = new JSDOM('');
const env = { DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer };

const input = process.argv[2] || path.join(__dirname, 'plano-aula.docx');
const out = process.argv[3] || path.join(__dirname, 'saida-modelo.docx');
const fonte = process.argv[4];
const bytes = new Uint8Array(fs.readFileSync(input));
const t0 = Date.now();
const { bytes: res, report } = M.process(bytes, { fonte }, env);
fs.writeFileSync(out, Buffer.from(res));
console.log('estatísticas:', JSON.stringify(report.stats));
for (const f of report.fixes) console.log('✔', f.label, f.count != null ? '(' + f.count + ')' : '', f.detail || '', f.examples && f.examples.length ? '\n     ex.: ' + f.examples.join('\n          ') : '');
for (const w of report.warnings) console.log('!', w.msg, '→', w.text);
console.log('gravado', out, res.length, 'bytes em', Date.now() - t0, 'ms');
