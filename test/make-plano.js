// Gera um "plano de aula" fictício com os defeitos típicos de arquivo convertido de PDF:
// folha deitada, uma seção e uma tabela por página, frases cortadas entre páginas,
// quebras de linha no meio da frase, fontes misturadas, letras comprimidas,
// espaço antes de dois-pontos, "1° ano" e um trecho que termina no meio.
const fs = require('fs');
const path = require('path');
const D = require('docx');

const logo = fs.readFileSync(path.join(__dirname, 'figura.png'));
const LAND = { page: { size: { width: 16840, height: 11910, orientation: D.PageOrientation.LANDSCAPE }, margin: { top: 1400, right: 1275, bottom: 280, left: 1133, header: 470, footer: 0 } } };
const W = 14144;

const run = (text, o = {}) => new D.TextRun({ text, bold: o.b, font: o.font, size: o.size || 24, characterSpacing: o.cs });
const para = (children, o = {}) => new D.Paragraph({ style: 'TableParagraph', alignment: D.AlignmentType.JUSTIFIED, indent: { left: o.ind || 68, right: 61 }, spacing: { before: o.before || 0 }, children: Array.isArray(children) ? children : [run(children, o)] });
const cell = (paras, width, span) => new D.TableCell({ width: { size: width, type: D.WidthType.DXA }, columnSpan: span, margins: { left: 0, right: 0 }, children: paras });
const row = (cells, h) => new D.TableRow({ height: { value: h || 600, rule: D.HeightRule.ATLEAST }, children: cells });
const table = (rows, widths) => new D.Table({ width: { size: W, type: D.WidthType.DXA }, columnWidths: widths, rows });
const empty = () => new D.Paragraph({ children: [] });

const p1 = [
  table([row([cell([para([run('PLANO DE AULA', { b: true })])], W)], 360), row([cell([para('')], W)], 355)], [W]),
  empty(), empty(),
  table([
    row([cell([para([run('Componente curricular ', { b: true }), run(': Ciências (Natureza)')])], W)]),
    row([cell([para([run('CIÊNCIAS – 2° ano: Ensino Fundamental I', { b: true })])], W)]),
    row([cell([
      para([run('OBJETIVOS GERAIS:', { b: true })]),
      para('Observar o ciclo da água no ambiente escolar e relacionar a evaporação, a condensação e a chuva com situações do dia a dia das'),
      para('crianças, ampliando o vocabulário científico da turma por meio de experiências simples e registros coletivos.'),
    ], W)], 1900),
    row([cell([
      para([run('OBJETIVOS ESPECÍFICOS:', { b: true })]),
      para([run('Conceitual', { b: true }), run(' : reconhecer os estados físicos da água e identificar onde eles aparecem na escola, no caminho de casa e nas')]),
    ], W)], 3000),
  ], [W]),
  empty(),
];
const p2 = [
  table([
    row([cell([
      para('brincadeiras do recreio, comparando as mudanças observadas ao longo da semana.'),
      empty(),
      para([run('Atitudinal', { b: true }), run(' : cuidar dos materiais coletivos e respeitar a vez de cada colega durante as observações em grupo.')]),
    ], W, 2)]),
    row([
      cell([para([run('Habilidades: EF02CI08:', { b: true }), run(' Comparar o efeito da radiação solar em diferentes tipos de superfície.')])], 5127),
      cell([
        para([run('Habilidades:', { b: true, font: 'Roboto', size: 26 })]),
        para([run('Observação', { b: true, font: 'Roboto', size: 26, cs: -16 }), run(' sistemática de um copo com água exposto ao sol e outro à sombra durante a manhã inteira', { font: 'Roboto', size: 26, cs: -15 })]),
        para([run('com registro em tabela simples feita pela própria turma.', { font: 'Roboto', size: 26 })]),
      ], 9018),
    ], 2900),
    row([cell([para('1')], 5127), cell([para('Conversa inicial com a turma sobre onde a água da chuva vai parar depois que cai no pátio da escola.')], 9018)]),
  ], [5127, 9018]),
  empty(),
];
const p3 = [
  table([
    row([cell([para('2')], 5127), cell([para('Montagem do experimento com dois copos, um no sol e outro na sombra, e combinação das regras de observação.')], 9018)]),
    row([cell([para([run('Recursos didáticos:', { b: true }), run(' copos transparentes, água, caneta, cartolina e fita de')])], W, 2)], 320),
    row([cell([para([run('Referências:')]), para('BRASIL. Ministério da Educação. Base Nacional Comum Curricular. Brasília: MEC, 2018.')], W, 2)], 1270),
  ], [5127, 9018]),
  empty(),
];
const p4 = [
  table([row([cell([para('ALVES, Rubem. A alegria de ensinar. Campinas: Papirus, 2000.')], W)], 635)], [W]),
  empty(),
];

const header = new D.Header({ children: [new D.Paragraph({ children: [new D.ImageRun({ type: 'png', data: logo, transformation: { width: 90, height: 52 } })] })] });
const doc = new D.Document({
  styles: {
    default: { document: { run: { font: 'Arial MT', size: 22 } } },
    paragraphStyles: [{ id: 'TableParagraph', name: 'Table Paragraph', basedOn: 'Normal', run: { font: 'Arial MT' } }],
  },
  sections: [
    { properties: Object.assign({ type: D.SectionType.CONTINUOUS }, LAND), headers: { default: header }, children: p1 },
    { properties: LAND, children: p2 },
    { properties: LAND, children: p3 },
    { properties: LAND, children: p4 },
  ],
});
D.Packer.toBuffer(doc).then((b) => { fs.writeFileSync(path.join(__dirname, 'plano-aula.docx'), b); console.log('plano-aula.docx', b.length); });
