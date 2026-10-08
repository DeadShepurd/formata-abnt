// Passa o TCC de teste pelo motor e grava o .docx formatado
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const opentype = require('opentype.js');
globalThis.docx = require('docx');
require('../src/engine.js');
const E = globalThis.FormataEngine;
const mammoth = require('mammoth');

const FONTS = {
  Arial: { r: 'LiberationSans-Regular', b: 'LiberationSans-Bold', i: 'LiberationSans-Italic', bi: 'LiberationSans-BoldItalic' },
  'Times New Roman': { r: 'LiberationSerif-Regular', b: 'LiberationSerif-Bold', i: 'LiberationSerif-Italic', bi: 'LiberationSerif-BoldItalic' },
};
const loaded = {};
const FONT_DIR = '/usr/share/fonts/truetype/liberation/';
const haveFonts = fs.existsSync(FONT_DIR + 'LiberationSans-Regular.ttf');
function measure(text, f) {
  const set = FONTS[f.name] || FONTS.Arial;
  const key = f.bold && f.italic ? 'bi' : f.bold ? 'b' : f.italic ? 'i' : 'r';
  const file = set[key];
  if (!loaded[file]) loaded[file] = opentype.loadSync(FONT_DIR + file + '.ttf');
  return loaded[file].getAdvanceWidth(text, f.size, { kerning: false });
}

(async () => {
  const input = process.argv[2] || path.join(__dirname, 'tcc-aluno.docx');
  const out = process.argv[3] || path.join(__dirname, 'saida.docx');
  const fonte = process.argv[4] || undefined;
  const { value: html, messages } = await mammoth.convertToHtml({ path: input }, E.mammothOptions(mammoth));
  if (messages.length) console.log('mammoth:', messages.map((m) => m.message).slice(0, 5));
  const parser = new new JSDOM('').window.DOMParser();
  const parsed = E.htmlToBlocks(html, parser);
  const cls = E.classify(parsed);
  for (const b of cls.blocks) console.log(b.type.padEnd(8), (b.text || '').slice(0, 70));
  console.log('extraído:', cls.extracted);
  const prof = E.PROFILES['unicid-pedagogia'];
  const info = {
    autores: 'Maria Aparecida dos Santos', titulo: 'O brincar na Educação Infantil', subtitulo: 'práticas lúdicas e desenvolvimento',
    orientador: 'Profa. Dra. Ana Paula Ribeiro', ano: '2026', fonte, tipo: process.env.TIPO || 'trabalho', disciplina: process.env.DISC || 'Fundamentos da Educação Infantil',
    resumo: cls.extracted.resumo, palavrasChave: cls.extracted.palavrasChave, abstract: cls.extracted.abstract, keywords: cls.extracted.keywords,
    dedicatoria: 'Aos meus pais, que me ensinaram o valor da educação.', agradecimentos: 'Agradeço à minha orientadora pela paciência e pelas leituras atentas.\nÀs crianças e professoras da EMEI que me receberam durante o estágio.',
    epigrafe: '“Brincar com crianças não é perder tempo, é ganhá-lo.”\n(Carlos Drummond de Andrade)', aprovacao: true, banca: 'Profa. Dra. Ana Paula Ribeiro – UNICID (orientadora)\nProf. Me. Carlos Henrique Lima – UNICID',
  };
  const res = E.buildDocument({ profile: prof, info, blocks: cls.blocks, notes: cls.notes }, { measure: haveFonts ? measure : null });
  console.log('avisos:', res.warnings, 'contagem:', res.counts, 'notas:', res.footnotes);
  console.log('config:', JSON.stringify(Object.fromEntries(Object.entries(res.config).filter(([k]) => k !== 'T' && !/^(resumo|abstract|agradecimentos|dedicatoria|epigrafe)$/.test(k)))));
  console.log('sumário estimado:', res.toc.map((t) => t.page + '  ' + t.title).join('\n'));
  const buf = await docx.Packer.toBuffer(res.doc);
  fs.writeFileSync(out, buf);
  console.log('gravado', out, buf.length);
})().catch((e) => { console.error(e); process.exit(1); });
