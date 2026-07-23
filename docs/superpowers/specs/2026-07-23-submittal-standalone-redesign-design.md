# Submittal — Projeto Standalone + Catálogo Automático + Organização no Drive

**Data:** 2026-07-23
**Status:** Aprovado para plano de implementação (sub-projeto 1 de 6 vira `Submittal-GAS`; 2-6 evoluem dentro dele)

## Contexto

O `Submittal.gs` hoje vive em `Ferramentas-PLB-Sheets`, um repositório/projeto Apps Script
compartilhado por várias ferramentas (BOM, Request, SuperBusca, SheetManager, ColorConfig,
SummaryAll) e implantado, via `gas-deploy.ps1`, em 4 planilhas diferentes (`PLB Principal`,
`VISIONS RISER`, `submitall log`, `FASANO - PLB WORK`).

A planilha real usada pelo Submittal ("Submittal Items LOG.xlsx", inspecionada em
`C:\Users\thiag\Downloads\Submittal Items LOG.xlsx`) tem hoje:

- `DATA BASE SUBMITTAL` — 61 linhas já cadastradas manualmente. Colunas: `DISCIPLINE`,
  `FOLDER`, `ROOM`, `LOCATION`, `UPC`, `ITEM`, `PDF` (PDF = nome do arquivo, texto puro).
- `FINISHES TEMPLATE` — layout novo (16 colunas, ordem abaixo), com dados de exemplo
  clonados da DATA BASE SUBMITTAL. É o gabarito do layout daqui pra frente.
- `IQ LUX SUBMITTALS` — layout antigo (12 colunas, sem FOLDER/UPC/TITLE/LINK SUBMITTAL),
  com 1 linha real de produção (STATUS = NOT COMMITTED). **Será recriada manualmente pelo
  usuário no layout novo — fora do escopo deste projeto.**
- `5.COST.LIST`, `Sheet6`, `Pendencias` — de outras ferramentas, sem relação com Submittal.

Layout novo (`FINISHES TEMPLATE`), ordem oficial daqui pra frente:

```
#N | #SUBMITTAL AMBAR | SUBMITTAL DATE | SUBMITTAL TITLE | DISCIPLINE | FOLDER | ROOM |
LOCATION | UPC | ITEM | CREATOR | STATUS | STATUS DATE | NOTE | PDF | LINK SUBMITTAL
```

`UPC` é preenchido por fórmula própria da planilha — o script nunca lê nem escreve nela.

## Decomposição

| # | Sub-projeto | Depende de |
|---|---|---|
| 1 | Projeto Apps Script standalone (`Submittal-GAS`) | — |
| 2 | Modelo de colunas por nome de cabeçalho (não posição fixa) | 1 |
| 3 | DATA BASE SUBMITTAL — leitura e gravação | 2 |
| 4 | Autofill ao digitar ITEM (gatilho onEdit) | 2, 3 |
| 5 | Organização automática no Drive (02-PRECON) | 3, 4 |
| 6 | Troca de referência (link externo → cópia organizada) | 5 |

Este documento cobre o desenho de todos os 6. A implementação (plano) trata 1 como
pré-requisito de infraestrutura e 2-6 como as mudanças funcionais dentro do projeto novo.

---

## 1) Projeto Apps Script standalone

Mesmo padrão já usado para `SuperBusca-GAS`: pasta e repositório Git próprios, fora do
`Ferramentas-PLB-Sheets`, para que a evolução do Submittal nunca seja incluída sem querer
num push pros outros 3 projetos GAS (`PLB Principal`, `VISIONS RISER`, `FASANO`).

- Pasta local: `C:\DEV\Sheets\Submittal-GAS\`
- Repositório: `github.com/thiagobarretosn-hue/Submittal-GAS`
- `.clasp.json` → `scriptId` já salvo como `"submitall log"` em `gas-projects.json`
  (`1GFoToJkoXDTjHQ6ieixuN8Cdny7BQNr7q26hH2EgrsJ95fFDaKecS900`) — confirmar antes de usar
  que esse script já está de fato vinculado à planilha "Submittal Items LOG".
- Arquivos a copiar do `Ferramentas-PLB-Sheets` (sem trazer nada de BOM/Request/etc.):
  - `Submittal.gs`, `SubmittalSidebar.html` (evoluem para o layout novo)
  - `SharedStyles.html`, `SharedScripts.html` (cópia direta, sem mudança)
  - `lib/Html.gs` → só a função `include()`
  - `lib/Config.gs` → só `SharedConfig_createDocConfigService` (o resto do arquivo — `AppConfig`,
    `showConfigDialog` etc. — não é usado pelo Submittal, não precisa copiar)
  - `lib/Utils.gs` → só `SharedUtils_numberToColumnLetter`
  - `Menu.gs` novo, só com o item "📦 Montar Submittal"
  - `appsscript.json` próprio (`enabledAdvancedServices: Sheets v4`, timezone igual)

## 2) Modelo de colunas por nome de cabeçalho

Troca a constante `SUBMITTAL_CFG.COLS` (posições fixas) por uma função que lê a linha 1
da aba e monta um mapa `{ NOME_CABECALHO: indiceDaColuna }`, comparando por texto
(trim + uppercase, tolerando o espaço sobrando em `"FOLDER "`).

Cabeçalhos que o script efetivamente lê/escreve: `#SUBMITTAL AMBAR`, `SUBMITTAL DATE`,
`DISCIPLINE`, `FOLDER`, `ROOM`, `LOCATION`, `ITEM`, `STATUS`, `PDF`, `LINK SUBMITTAL`.
`#N`, `SUBMITTAL TITLE`, `CREATOR`, `STATUS DATE`, `NOTE`, `UPC` são só informativos /
preenchidos por fórmula ou pelo usuário — o script não precisa deles.

Isso resolve de vez o problema de "a ordem mudou": qualquer reordenação futura de coluna
não quebra o script, desde que o nome do cabeçalho continue igual.

## 3) DATA BASE SUBMITTAL — leitura e gravação

- Índice em memória: `ITEM` (trim + uppercase) → linha inteira da DATA BASE SUBMITTAL.
- **Leitura** (usada pelo autofill, sub-projeto 4): dado um texto de ITEM, procura no
  índice; se achar, retorna `{ discipline, folder, room, location, pdf }`.
- **Gravação** (usada pela ingestão, sub-projeto 4/5): quando o usuário cola um PDF/link
  para um ITEM que não existe no índice, acrescenta uma linha nova na DATA BASE SUBMITTAL
  com `DISCIPLINE, FOLDER, ROOM, LOCATION, ITEM, PDF` (na própria ordem de colunas da aba
  DATA BASE SUBMITTAL, resolvida por cabeçalho, não hardcoded).
- `UPC` nunca é lido nem escrito pelo script (fórmula própria da planilha, em ambas as abas).

## 4) Gatilho onEdit — autofill + ingestão

Duas colunas de disparo, um único gatilho **instalável** (`ScriptApp.newTrigger(...).onEdit()`,
registrado uma vez por um item de menu "🔌 Ativar automação" — precisa ser instalável, não
o `onEdit(e)` simples, porque mover/criar arquivo no Drive exige autorização completa):

- **Editou ITEM** → busca na DATA BASE SUBMITTAL; se achar, escreve
  DISCIPLINE/FOLDER/ROOM/LOCATION na mesma linha. Pra coluna PDF: a base só guarda o
  *nome* do arquivo (texto), não um link de verdade — então o autofill busca ao vivo
  dentro de `<RAIZ>/DISCIPLINE/FOLDER/ROOM/` um arquivo com esse nome exato. Se achar,
  escreve o hyperlink de verdade. **Se não achar** (caso dos itens legados da DATA BASE
  SUBMITTAL que foram cadastrados manualmente antes de qualquer automação existir, sem
  nunca terem sido organizados de fato), deixa PDF em branco — o item cai no mesmo fluxo
  de "item novo": o usuário cola o link uma vez, e a partir daí passa a ficar organizado.
- **Editou PDF** → roda a organização no Drive (sub-projeto 5) e, se o ITEM ainda não
  estava cadastrado, grava a linha nova na DATA BASE SUBMITTAL (sub-projeto 3).

**Multi-célula (paste de várias colunas/linhas de uma vez):** `e.range` de um paste pode
cobrir várias colunas e linhas ao mesmo tempo (ex: usuário cola uma linha inteira vinda do
Excel), não só uma célula. O handler precisa checar se o range editado **intersecta** a
coluna ITEM ou PDF (não comparar "é exatamente" essa coluna) e iterar por **cada linha**
do range, não só a primeira.

**Erros dentro do gatilho:** um trigger instalável não tem contexto de UI —
`SpreadsheetApp.getUi().alert()` não funciona aqui. Erros devem virar
`SpreadsheetApp.getActiveSpreadsheet().toast(...)` ou uma nota na própria célula, nunca
um `alert()` (que falharia silenciosamente e só geraria um e-mail de erro que ninguém vê).

**Por que não há risco de loop:** o handler só age quando a coluna editada é ITEM ou PDF.
As escritas do autofill (DISCIPLINE/FOLDER/ROOM/LOCATION) disparam onEdit de novo, mas
essas colunas não são gatilho — o handler entra, não reconhece a coluna, sai sem fazer nada.
A única coluna que é ao mesmo tempo alvo do autofill E gatilho é PDF; resolvido fazendo a
organização no Drive **idempotente** (só move/copia se o arquivo ainda não está no lugar
certo com o nome certo) — reprocessar o mesmo valor não tem efeito colateral.

**Concorrência:** com múltiplos usuários editando ao mesmo tempo, dois gatilhos podem, em
teoria, descobrir o mesmo ITEM novo simultaneamente e gravar duas linhas duplicadas na
DATA BASE SUBMITTAL. Baixa probabilidade (uso não é de alta frequência), mas vale usar
`LockService.getDocumentLock()` ao redor da gravação na base como proteção barata.

## 5) Organização automática no Drive (02-PRECON)

- Pasta raiz **fixa** (constante no código, como já é feito com `TEMPLATE_DOC_ID`):
  `1YqVuKzEDJTAl503zYCoaHXW7-vovs7ja` (Shared Drive "AMBAR US NEW", montado localmente
  como `H:\` via Google Drive for Desktop — o script usa o ID, não o caminho `H:\`).
  Repositório único e geral: todas as obras organizam seus PDFs aqui dentro, sem
  subpasta por obra.
- Caminho: `<RAIZ FIXA>/DISCIPLINE/FOLDER/ROOM/<ITEM>.pdf` — 3 níveis de subpasta
  (`DISCIPLINE` → `FOLDER` → `ROOM`), criando qualquer nível que não existir ainda
  (mesma lógica de "busca por nome, cria se não achar" já usada hoje pra pasta do pacote
  final). `LOCATION` nunca vira pasta — fica só como metadado na planilha.
- **Nome do arquivo salvo = texto exato da coluna ITEM** (+ `.pdf`), não o nome original
  do arquivo de origem.
- Fonte do arquivo: reaproveita `_sub_resolveBlob` já existente (Drive por ID, Google
  Doc/Sheet/Slide convertido pra PDF, ou download via `UrlFetchApp` se for link da web).
- **Idempotência / mudança de local:** se a célula PDF já é um hyperlink apontando pra um
  arquivo dentro do repositório fixo (extrai o `fileId` do link atual com
  `_sub_extractDriveId`, já existente), o script sabe que esse item já foi organizado antes.
  Compara a pasta atual desse arquivo (`file.getParents()`) com a pasta
  DISCIPLINE/FOLDER/ROOM que deveria ser agora: se for a mesma, não faz nada; se for
  diferente, **move** com `file.moveTo(novaPasta)` — não com o padrão antigo
  `addTo`/`removeFrom` (frágil em Shared Drive, onde um arquivo deve ter exatamente um
  pai; `moveTo` já cuida disso corretamente). Não precisa de coluna extra — o próprio link
  da célula já carrega o `fileId`.
- A pasta pai já configurável hoje (`config.parentFolderId`, usada só pelo pacote final
  "SUBMITTAL <num> - data") continua existindo e não tem nenhuma relação com este
  repositório fixo — são coisas diferentes por natureza (pacote final pro cliente vs.
  biblioteca permanente de cut-sheets).

## 6) Troca de referência

Depois que o arquivo está organizado no Drive, a célula PDF da linha passa a ter um
hyperlink apontando pro arquivo dentro de `DISCIPLINE/FOLDER/ROOM/`, substituindo o link
externo original (mesmo padrão de `RichTextValue` + `setLinkUrl` já usado hoje em
`saveSubmittalItemLink`).

## Convivência com "Montar Submittal" (já existente)

O fluxo de mesclagem/capa continua igual — só passa a ler o link a partir da coluna PDF
(resolvida por nome de cabeçalho) em vez da antiga coluna fixa L. Como o PDF, nesse ponto,
já é sempre o arquivo organizado no 02-PRECON (por causa do passo 6), o merge fica mais
confiável do que hoje (que às vezes depende de link externo).

## Fora de escopo deste desenho

- Migração da aba `IQ LUX SUBMITTALS` pro layout novo — o usuário recria manualmente.
- Qualquer mudança em `5.COST.LIST` / `Sheet6` / fórmula de UPC.
- Suporte a mais de um layout de coluna simultâneo.

## Riscos técnicos a validar cedo (antes de escrever o resto)

1. **Shared Drive + DriveApp:** confirmar que `DriveApp.getFolderById` /
   `folder.createFolder` / mover arquivo funcionam normalmente dentro de um Shared Drive
   (historicamente tem pegadinhas de permissão/API que não aparecem em Meu Drive normal).
2. **Gatilho instalável onEdit:** confirmar que primeira ativação exige autorização OAuth
   explícita do usuário (fluxo "Ativar automação" no menu) e que o gatilho sobrevive a
   um clasp push (não precisa reativar a cada deploy).
3. **`.clasp.json` do "submitall log"**: confirmar que o scriptId salvo já está mesmo
   vinculado à planilha "Submittal Items LOG" antes de criar o projeto novo em cima dele.
