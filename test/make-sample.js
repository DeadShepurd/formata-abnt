// Gera um "TCC de aluno" bagunçado (fora das regras) para testar o app
const fs = require('fs');
const D = require('docx');

const lorem = [
  'A brincadeira ocupa lugar central no desenvolvimento infantil, pois é por meio dela que a criança experimenta papéis sociais, elabora regras de convivência e constrói significados sobre o mundo que a cerca. Nas instituições de Educação Infantil, o brincar deixa de ser apenas passatempo e passa a ser compreendido como eixo estruturante das práticas pedagógicas, conforme orientam os documentos curriculares nacionais.',
  'Ao observar o cotidiano de uma turma de crianças de quatro e cinco anos, percebe-se que os momentos de brincadeira livre revelam interesses, conflitos e aprendizagens que nem sempre aparecem nas atividades dirigidas. Cabe ao professor organizar tempos, espaços e materiais que favoreçam essas experiências, sem transformar o brincar em mero instrumento para conteúdos escolares.',
  'Nesse sentido, a formação do pedagogo precisa contemplar a reflexão sobre a ludicidade como dimensão humana, e não apenas como recurso didático. Quando o adulto reconhece a brincadeira como linguagem da infância, amplia suas possibilidades de escuta e de planejamento, tornando a rotina mais significativa para as crianças (KISHIMOTO, 2017).',
  'Os dados produzidos ao longo do estágio supervisionado indicam que a disposição dos cantos de atividades interfere diretamente na qualidade das interações. Espaços com materiais não estruturados, como caixas, tecidos e elementos da natureza, estimularam brincadeiras de faz de conta mais longas e cooperativas do que os espaços com brinquedos prontos.',
];

const P = (text, o = {}) => new D.Paragraph({ children: [new D.TextRun({ text, bold: o.bold, size: 22 })], heading: o.heading, indent: o.indent, alignment: o.alignment, numbering: o.numbering });

const kids = [
  P('UNICID', { alignment: 'center' }),
  P('Maria Aparecida dos Santos', { alignment: 'center' }),
  P('O BRINCAR NA EDUCAÇÃO INFANTIL: práticas lúdicas e desenvolvimento', { alignment: 'center', bold: true }),
  P('São Paulo 2026', { alignment: 'center' }),
  P('RESUMO', { bold: true }),
  P('Este trabalho investiga o papel do brincar nas práticas pedagógicas da Educação Infantil, a partir de observações realizadas em uma escola municipal de São Paulo. Discute-se a ludicidade como linguagem da infância e analisam-se as condições de tempo, espaço e materiais que favorecem brincadeiras significativas. Conclui-se que a organização intencional do ambiente e a escuta atenta do professor ampliam as aprendizagens das crianças.'),
  P('Palavras-chave: brincar, educação infantil, ludicidade, prática pedagógica'),
  P('ABSTRACT', { bold: true }),
  P('This study investigates the role of play in Early Childhood Education practices, based on observations carried out in a public school in São Paulo.'),
  P('Keywords: play, early childhood education, playfulness'),
  P('SUMÁRIO', { bold: true }),
  new D.Paragraph({ children: [new D.TextRun({ text: '1 INTRODUÇÃO', size: 22 }), new D.TextRun({ children: [new D.Tab()] }), new D.TextRun({ text: '6', size: 22 })] }),
  new D.Paragraph({ children: [new D.TextRun({ text: '2 FUNDAMENTAÇÃO TEÓRICA', size: 22 }), new D.TextRun({ children: [new D.Tab()] }), new D.TextRun({ text: '8', size: 22 })] }),
  // textual
  P('1. INTRODUÇÃO', { bold: true }),
  P(lorem[0]), P(lorem[1]), P(lorem[2]),
  new D.Paragraph({ children: [new D.TextRun({ text: 'A escolha do tema surgiu durante o estágio', size: 22 }), new D.FootnoteReferenceRun(1), new D.TextRun({ text: ' na Educação Infantil, quando a observação das brincadeiras revelou questões que mobilizaram esta pesquisa.', size: 22 })] }),
  P('2 FUNDAMENTAÇÃO TEÓRICA', { heading: D.HeadingLevel.HEADING_1 }),
  P(lorem[3]), P(lorem[0]),
  P('2.1 O lúdico como linguagem da infância', { heading: D.HeadingLevel.HEADING_2 }),
  P(lorem[1]),
  P('Para Vigotski, a brincadeira cria uma zona de desenvolvimento iminente na criança, pois nela a criança se comporta além do comportamento habitual de sua idade, além de seu comportamento cotidiano; no brinquedo, é como se ela fosse maior do que é na realidade. Como no foco de uma lente de aumento, a brincadeira contém todas as tendências do desenvolvimento sob forma condensada (VIGOTSKI, 2008, p. 35).', { indent: { left: 2268 } }),
  P(lorem[2]),
  P('Entre as condições observadas, destacam-se:'),
  new D.Paragraph({ children: [new D.TextRun({ text: 'organização de cantos de atividades com materiais variados;', size: 22 })], numbering: { reference: 'lista', level: 0 } }),
  new D.Paragraph({ children: [new D.TextRun({ text: 'tempo prolongado para brincadeira livre na rotina diária;', size: 22 })], numbering: { reference: 'lista', level: 0 } }),
  new D.Paragraph({ children: [new D.TextRun({ text: 'mediação do professor sem interromper o enredo das crianças.', size: 22 })], numbering: { reference: 'lista', level: 0 } }),
  P('2.1.1 Brincadeira de faz de conta', { bold: true }),
  P(lorem[3]), P(lorem[1]),
  P('Figura 1 – Crianças em atividade de faz de conta no pátio', { alignment: 'center' }),
  new D.Paragraph({ alignment: 'center', children: [new D.ImageRun({ type: 'png', data: fs.readFileSync(__dirname + '/figura.png'), transformation: { width: 450, height: 260 } })] }),
  P('Fonte: Elaborado pela autora (2026).', { alignment: 'center' }),
  P(lorem[0]), P(lorem[2]),
  P('3 METODOLOGIA', { bold: true }),
  P(lorem[1]), P(lorem[3]),
  P('Quadro 1 – Autores de referência sobre o brincar'),
  new D.Table({ rows: [
    new D.TableRow({ children: ['Autor', 'Contribuição'].map((t) => new D.TableCell({ children: [new D.Paragraph({ children: [new D.TextRun({ text: t, bold: true })] })] })) }),
    new D.TableRow({ children: ['Vigotski', 'Brincadeira e zona de desenvolvimento iminente'].map((t) => new D.TableCell({ children: [new D.Paragraph(t)] })) }),
    new D.TableRow({ children: ['Kishimoto', 'Jogo, brinquedo e brincadeira na educação'].map((t) => new D.TableCell({ children: [new D.Paragraph(t)] })) }),
  ] }),
  P('Fonte: Elaborado pela autora (2026).'),
  P(lorem[0]), P(lorem[2]), P(lorem[1]),
  P('4 ANÁLISE DOS DADOS', { bold: true }),
  P(lorem[3]), P(lorem[0]), P(lorem[1]), P(lorem[2]), P(lorem[3]),
  P('4.1 Os espaços e os materiais', { heading: D.HeadingLevel.HEADING_2 }),
  P(lorem[0]), P(lorem[1]), P(lorem[2]),
  P('5 CONSIDERAÇÕES FINAIS', { bold: true }),
  P(lorem[2]), P(lorem[3]),
  P('REFERÊNCIAS', { bold: true }),
  new D.Paragraph({ children: [new D.TextRun({ text: 'VIGOTSKI, L. S. A brincadeira e o seu papel no desenvolvimento psíquico da criança. ', size: 22 }), new D.TextRun({ text: 'Revista Virtual de Gestão de Iniciativas Sociais', bold: true, size: 22 }), new D.TextRun({ text: ', Rio de Janeiro, n. 8, p. 23-36, jun. 2008.', size: 22 })] }),
  new D.Paragraph({ children: [new D.TextRun({ text: 'BRASIL. Ministério da Educação. ', size: 22 }), new D.TextRun({ text: 'Base Nacional Comum Curricular', bold: true, size: 22 }), new D.TextRun({ text: '. Brasília, DF: MEC, 2018.', size: 22 })] }),
  new D.Paragraph({ children: [new D.TextRun({ text: 'KISHIMOTO, T. M. (org.). ', size: 22 }), new D.TextRun({ text: 'Jogo, brinquedo, brincadeira e a educação', bold: true, size: 22 }), new D.TextRun({ text: '. 14. ed. São Paulo: Cortez, 2017.', size: 22 })] }),
  new D.Paragraph({ children: [new D.TextRun({ text: 'ÁRIES, P. ', size: 22 }), new D.TextRun({ text: 'História social da criança e da família', bold: true, size: 22 }), new D.TextRun({ text: '. 2. ed. Rio de Janeiro: LTC, 1981.', size: 22 })] }),
  P('APÊNDICE A – Roteiro de observação', { bold: true }),
  P('Registro das brincadeiras observadas, com data, duração, número de crianças envolvidas e materiais utilizados.'),
];

const doc = new D.Document({
  numbering: { config: [{ reference: 'lista', levels: [{ level: 0, format: 'bullet', text: '•', alignment: 'left' }] }] },
  footnotes: { 1: { children: [new D.Paragraph('Estágio supervisionado realizado no primeiro semestre de 2026 em uma EMEI da zona leste de São Paulo.')] } },
  styles: { default: { document: { run: { font: 'Times New Roman', size: 22 } } } },
  sections: [{ children: kids }],
});
D.Packer.toBuffer(doc).then((b) => { fs.writeFileSync(__dirname + '/tcc-aluno.docx', b); console.log('ok', b.length); });
