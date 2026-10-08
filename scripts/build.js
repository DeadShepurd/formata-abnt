// Monta duas saídas a partir de src/:
//  - www/index.html  → app instalado (Android): tudo local, funciona sem internet
//  - dist/formata-abnt.html → versão web (artifact), com bibliotecas e fontes por CDN
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const r = (...p) => path.join(root, ...p);
const read = (p) => fs.readFileSync(r(p), 'utf8');

const shell = read('src/app.html');
const engine = read('src/engine.js');
const modelo = read('src/modelo.js');
const app = read('src/app.js');
if (/<\/script/i.test(engine + modelo + app)) throw new Error('"</script" dentro do JS');
const inject = (html) => html.replace('/*@ENGINE@*/', () => engine).replace('/*@MODELO@*/', () => modelo).replace('/*@APP@*/', () => app);

// ---------- versão web ----------
fs.mkdirSync(r('dist'), { recursive: true });
fs.writeFileSync(r('dist/formata-abnt.html'), inject(shell));

// ---------- app instalado ----------
const www = r('www');
fs.rmSync(www, { recursive: true, force: true });
fs.mkdirSync(path.join(www, 'vendor'), { recursive: true });
fs.mkdirSync(path.join(www, 'fonts'), { recursive: true });
fs.copyFileSync(r('node_modules/mammoth/mammoth.browser.min.js'), path.join(www, 'vendor/mammoth.browser.min.js'));
fs.copyFileSync(r('node_modules/docx/dist/index.iife.js'), path.join(www, 'vendor/docx.iife.js'));
fs.copyFileSync(r('node_modules/fflate/umd/index.js'), path.join(www, 'vendor/fflate.js'));

const faces = [
  ['Atkinson Hyperlegible', 'atkinson-hyperlegible', [[400, 'normal'], [700, 'normal'], [400, 'italic']]],
  ['IBM Plex Mono', 'ibm-plex-mono', [[400, 'normal'], [500, 'normal']]],
  ['Literata', 'literata', [[500, 'normal'], [700, 'normal']]],
  ['Arimo', 'arimo', [[400, 'normal'], [700, 'normal'], [400, 'italic'], [700, 'italic']]],
  ['Tinos', 'tinos', [[400, 'normal'], [700, 'normal'], [400, 'italic'], [700, 'italic']]],
];
let css = '/* Fontes locais (Fontsource, licenças OFL/Apache) */\n';
for (const [family, pkg, list] of faces) {
  for (const [w, st] of list) {
    const file = `${pkg}-latin-${w}-${st}.woff2`;
    fs.copyFileSync(r('node_modules/@fontsource', pkg, 'files', file), path.join(www, 'fonts', file));
    css += `@font-face{font-family:"${family}";font-style:${st};font-weight:${w};font-display:swap;src:url(${file}) format("woff2")}\n`;
  }
}
fs.writeFileSync(path.join(www, 'fonts/fonts.css'), css);

let body = shell
  .replace(/<title>[\s\S]*?<\/title>\s*/, '')
  .replace(/<link rel="preconnect"[^>]*>\s*/g, '')
  .replace(/<link rel="stylesheet" href="https:\/\/fonts\.googleapis\.com[^>]*>\s*/, '')
  .replace(/<script src="https:\/\/cdn\.jsdelivr\.net\/npm\/mammoth[^"]*"><\/script>/, '<script src="vendor/mammoth.browser.min.js"></script>')
  .replace(/<script src="https:\/\/cdn\.jsdelivr\.net\/npm\/docx[^"]*"><\/script>/, '<script src="vendor/docx.iife.js"></script>')
  .replace(/<script src="https:\/\/cdn\.jsdelivr\.net\/npm\/fflate[^"]*"><\/script>/, '<script src="vendor/fflate.js"></script>');
if (/https?:\/\/(cdn|fonts)\./.test(body)) throw new Error('ainda há recurso externo no app instalado');
body = inject(body);

const head = `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#f3f5fa" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#0e1222" media="(prefers-color-scheme: dark)">
<meta name="color-scheme" content="light dark">
<title>Formata ABNT</title>
<link rel="stylesheet" href="fonts/fonts.css">
<style>:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;font:14px system-ui,sans-serif;background:#f8f8f6}img{max-width:100%}[hidden]{display:none!important}</style>
</head>
<body>
`;
fs.writeFileSync(path.join(www, 'index.html'), head + body + '\n</body>\n</html>\n');
const kb = (p) => Math.round(fs.statSync(p).size / 1024) + ' KB';
console.log('web:', kb(r('dist/formata-abnt.html')), '| app:', kb(path.join(www, 'index.html')), '+ vendor + fontes');
