# Formata ABNT

Aplicativo para celular que coloca trabalhos acadêmicos já escritos nas regras da ABNT e do manual da UNICID. A pessoa importa o arquivo `.docx`, confere a estrutura tocando nos parágrafos e baixa o arquivo formatado. O texto nunca é alterado e não há IA envolvida: tudo funciona por regras, direto no aparelho e sem internet.

Tipos de trabalho: **trabalho do semestre**, **modelo da faculdade**, **artigo científico** e **TCC**.

### Modelo da faculdade (plano de aula e trabalhos em tabela)

Quando o arquivo segue um modelo pronto — folha deitada, conteúdo em tabelas, logo no cabeçalho — o app reconhece sozinho e **não remonta o documento**: ajusta o próprio `.docx` (`src/modelo.js`). O cabeçalho, as imagens, as bordas e a largura das colunas ficam como estão. O que ele corrige:

- tabelas e frases cortadas entre páginas (comum em arquivo convertido de PDF) voltam a ser uma só;
- quebras de linha no meio da frase;
- fonte e tamanho unificados (Arial ou Times New Roman 12) e letras comprimidas normalizadas;
- espaçamento, recuos e margem interna das células padronizados; altura das linhas ajustada ao texto;
- referências à esquerda, espaço simples, 6 pt antes e depois, em ordem alfabética;
- espaço antes de pontuação e `1°` no lugar de `1º`.

Nenhuma palavra é trocada. Trechos que parecem incompletos (terminando em "de", "a", vírgula…) aparecem como aviso para a pessoa conferir.

## Instalar no celular (Android)

1. No celular, abra a página **Releases** deste repositório e baixe o arquivo `formata-abnt.apk` da versão mais recente.
   Link direto, que sempre aponta para a última versão: `https://github.com/<seu-usuario>/formata-abnt/releases/latest/download/formata-abnt.apk`
2. Toque no arquivo baixado. Na primeira vez, o Android pede para permitir a instalação de apps dessa fonte (o navegador ou o app Arquivos). Permita e volte.
3. Toque em **Instalar**. O ícone **Formata ABNT** aparece na tela inicial.

Versões novas instalam por cima da antiga, sem perder os dados já preenchidos.

## Como o .apk é gerado

A cada envio para a branch `main`, o GitHub Actions (`.github/workflows/android.yml`):

1. instala as dependências e roda os testes do formatador;
2. monta a pasta `www/` a partir de `src/` (com as bibliotecas e fontes locais);
3. copia o app para o projeto Android (Capacitor) e compila o `.apk`;
4. publica o arquivo numa nova versão em **Releases**.

Também dá para rodar manualmente na aba **Actions › Gerar app Android › Run workflow**.

## Onde ficam as regras

Todas as regras ficam em `src/engine.js`, no objeto `PROFILES['unicid-pedagogia']` e em `WORK_TYPES`:

| Regra | Valor atual | Origem |
| --- | --- | --- |
| Margens | 3 cm (sup./esq.), 2 cm (inf./dir.) | Manual UNICID |
| Fonte | Times New Roman 12 (Arial opcional) | Manual UNICID |
| Recuo do parágrafo | 2 cm | Manual UNICID |
| Citação longa | recuo 4 cm, 10 pt, espaço simples | Manual UNICID |
| Notas, legendas, fontes, paginação | 10 pt | Manual UNICID |
| Referências | espaço simples, 6 pt antes e depois, à esquerda | Manual UNICID |
| Palavras-chave | separadas por ponto, inicial maiúscula | Manual UNICID |
| Títulos das seções | 1 NEGRITO MAIÚSCULO · 1.1 MAIÚSCULO · 1.1.1 Negrito · 1.1.1.1 Normal · 1.1.1.1.1 Itálico | Manual UNICID |
| Entrelinha | 1,5 (trabalho e TCC) · simples (artigo) | ABNT NBR 14724 / Manual UNICID |
| Capa, folha de rosto, sumário | NBR 14724:2024 | ABNT |

Para outra faculdade ou curso, basta criar um novo perfil com os mesmos campos.

## Estrutura

```
src/engine.js      motor: lê o .docx, pré-classifica os parágrafos e gera o .docx formatado
src/modelo.js      modo "Modelo da faculdade": ajusta o .docx original sem refazer o layout
src/app.js         interface (telas Texto, Estrutura, Dados e Baixar)
src/app.html       marcação e estilos
scripts/build.js   monta www/ (app instalado) e dist/ (versão web)
scripts/icons.py   gera ícone e tela de abertura do Android
android/           projeto Android (Capacitor)
test/              TCC e plano de aula fictícios e testes automáticos
```

## Desenvolvimento

```
npm ci
npm test              # monta o app e roda os testes
npx cap sync android  # copia www/ para o projeto Android
```

Para compilar no computador, abra a pasta `android/` no Android Studio.

A chave de assinatura (`android/app/formata-abnt.keystore`) está no repositório para que todas as versões saiam com a mesma assinatura. Antes de publicar na Play Store, troque por uma chave guardada nos *Secrets* do GitHub.
