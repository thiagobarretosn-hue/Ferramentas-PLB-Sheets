# Submittal Standalone + Catálogo Automático — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extrair o Submittal para um projeto Apps Script próprio (`Submittal-GAS`), trocar o mapeamento de colunas fixo por resolução via cabeçalho, e adicionar um catálogo (`DATA BASE SUBMITTAL`) com autofill automático e organização de PDFs no Drive, tudo disparado por um gatilho `onEdit` instalável.

**Architecture:** Apps Script V8, container-bound à planilha "Submittal Items LOG" (scriptId salvo como `"submitall log"`). Lógica dividida em módulos por responsabilidade: `Submittal.gs` (montagem/merge, já existente), `SubmittalColumns.gs` (resolução de colunas por cabeçalho), `SubmittalCatalog.gs` (leitura/escrita da DATA BASE SUBMITTAL), `SubmittalDriveOrg.gs` (organização de arquivos no Shared Drive fixo), `SubmittalTriggers.gs` (gatilho onEdit instalável). Funções puras (sem chamada a serviço do Apps Script) trazem um `module.exports` guardado por `typeof module !== 'undefined'`, o que permite rodá-las com `node` puro, sem depender do ambiente do Apps Script, para teste automatizado real.

**Tech Stack:** Google Apps Script V8 (IronPython não se aplica aqui — é JS puro), `clasp` para push/pull, Sheets API v4 avançada (já habilitada), DriveApp, LockService, ScriptApp (gatilhos instaláveis). Testes de lógica pura: Node.js simples (sem framework, sem npm deps).

## Global Constraints

- Nunca modificar `Ferramentas-PLB-Sheets` neste plano — todo o trabalho vive em `C:\DEV\Sheets\Submittal-GAS\`, projeto e repositório novos.
- `.gs`/`.html` não têm teste automatizado nativo — qualquer função que chame `SpreadsheetApp`/`DriveApp`/`ScriptApp`/`UrlFetchApp` é validada manualmente na planilha real ("Submittal Items LOG"), nunca assumida como correta só por "parecer certo".
- Toda função pura (sem chamada de serviço do Apps Script) leva teste automatizado real via `node`, seguindo o padrão de `module.exports` guardado descrito acima.
- `.format()`/concatenação de string em vez de qualquer coisa equivalente a f-string com expressão não se aplica aqui (isso é regra do Python/IronPython do outro projeto) — em JS, template literals normais estão liberados.
- Nome do arquivo salvo no Drive = texto exato da coluna `ITEM` + `.pdf` (nunca o nome original do arquivo de origem).
- Pasta raiz do repositório de cut-sheets é fixa: `1YqVuKzEDJTAl503zYCoaHXW7-vovs7ja` (Shared Drive "AMBAR US NEW", confirmado pelo usuário) — nunca configurável por aba/obra.
- Mover arquivo entre pastas sempre com `file.moveTo(novaPasta)` — nunca `addTo`/`removeFrom` (frágil em Shared Drive).
- Dentro do gatilho onEdit instalável, nunca chamar `SpreadsheetApp.getUi().alert()`/`prompt()` (sem contexto de UI) — erros viram `toast()` + `console.error`.
- `UPC` nunca é lido nem escrito pelo script (é fórmula própria da planilha, em ambas as abas envolvidas).

---

## Task 1: Esqueleto do projeto `Submittal-GAS` + confirmar o scriptId certo

**Files:**
- Create: `C:\DEV\Sheets\Submittal-GAS\.clasp.json`
- Create: `C:\DEV\Sheets\Submittal-GAS\.claspignore`
- Create: `C:\DEV\Sheets\Submittal-GAS\appsscript.json`
- Create: `C:\DEV\Sheets\Submittal-GAS\Menu.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\lib\Html.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\lib\Config.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\lib\Utils.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\SharedStyles.html`
- Create: `C:\DEV\Sheets\Submittal-GAS\SharedScripts.html`

**Interfaces:**
- Produces: `include(filename)` (global, de `lib/Html.gs`), `SharedConfig_createDocConfigService(propKey, getDefaults)` (global, de `lib/Config.gs`), `SharedUtils_numberToColumnLetter(n)` (global, de `lib/Utils.gs`). Tasks seguintes dependem desses três nomes exatamente assim.

- [ ] **Step 1: Confirmar que o scriptId salvo é o certo**

Rodar (usa a autenticação do `clasp` já configurada nesta máquina — a mesma usada por `gas-deploy.ps1`):

```bash
mkdir -p /tmp/verify-submittal-script
cd /tmp/verify-submittal-script
clasp clone 1GFoToJkoXDTjHQ6ieixuN8Cdny7BQNr7q26hH2EgrsJ95fFDaKecS900 --rootDir .
ls -la
```

Expected: a lista de arquivos baixados. Se aparecer `Submittal.gs`/`SubmittalSidebar.html` ou a pasta estiver praticamente vazia (só `Code.gs` padrão), é sinal de que o script ainda não tem nada de outra ferramenta — pode reaproveitar com segurança. **Se aparecer código de BOM/Request/SuperBusca**, PARE — esse scriptId não é dedicado, avise o usuário antes de continuar (ele pode ter reaproveitado o nome "submitall log" por engano). Apague `/tmp/verify-submittal-script` depois de conferir.

- [ ] **Step 2: Criar a pasta do projeto e o .clasp.json**

```bash
mkdir -p "C:/DEV/Sheets/Submittal-GAS/lib"
cd "C:/DEV/Sheets/Submittal-GAS"
```

Criar `.clasp.json`:

```json
{"scriptId":"1GFoToJkoXDTjHQ6ieixuN8Cdny7BQNr7q26hH2EgrsJ95fFDaKecS900","rootDir":"."}
```

- [ ] **Step 3: Criar .claspignore**

```
**/**
!appsscript.json
!*.gs
!*.html
!lib/**
```

- [ ] **Step 4: Criar appsscript.json**

```json
{
  "timeZone": "America/Fortaleza",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Sheets",
        "version": "v4",
        "serviceId": "sheets"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

- [ ] **Step 5: Copiar os 3 helpers compartilhados (versão mínima, sem o resto de Config.gs/Utils.gs)**

`lib/Html.gs`:

```javascript
/**
 * @fileoverview Helper de HtmlService — padrão oficial GAS pra compartilhar CSS/JS entre sidebars.
 * Requer que a sidebar seja aberta via createTemplateFromFile(...).evaluate().
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```

`lib/Config.gs`:

```javascript
/**
 * @fileoverview Factory de config por documento (persistida em DocumentProperties).
 * Copiado de Ferramentas-PLB-Sheets/lib/Shared/Config.gs — só a parte usada pelo Submittal.
 */
function SharedConfig_createDocConfigService(propKey, getDefaults) {
  return {
    getAll: function() {
      try {
        const raw = PropertiesService.getDocumentProperties().getProperty(propKey);
        if (raw) return Object.assign(getDefaults(), JSON.parse(raw));
      } catch (e) {
        console.error('[SharedConfig:' + propKey + '] ' + e.message);
      }
      return getDefaults();
    },
    get: function(key, defaultValue) {
      const all = this.getAll();
      const value = all[key];
      if (value === undefined || value === null || value === '') {
        return defaultValue === undefined ? '' : defaultValue;
      }
      return value;
    },
    saveAll: function(config) {
      try {
        PropertiesService.getDocumentProperties().setProperty(propKey, JSON.stringify(config));
        return { success: true };
      } catch (e) {
        return { success: false, message: e.message };
      }
    }
  };
}
```

`lib/Utils.gs`:

```javascript
/**
 * @fileoverview Copiado de Ferramentas-PLB-Sheets/lib/Shared/Utils.gs — só a função usada pelo Submittal.
 */
function SharedUtils_numberToColumnLetter(columnNumber) {
  if (!columnNumber || columnNumber < 1) return '';
  let result = '';
  let num = columnNumber;
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}
```

- [ ] **Step 6: Copiar SharedStyles.html e SharedScripts.html sem alteração**

Copiar o conteúdo exato de:
- `C:\DEV\Sheets\Ferramentas-PLB-Sheets\SharedStyles.html` → `C:\DEV\Sheets\Submittal-GAS\SharedStyles.html`
- `C:\DEV\Sheets\Ferramentas-PLB-Sheets\SharedScripts.html` → `C:\DEV\Sheets\Submittal-GAS\SharedScripts.html`

- [ ] **Step 7: Criar Menu.gs**

```javascript
/**
 * @fileoverview Menu Principal — Submittal-GAS
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Submittal')
    .addItem('📦 Montar Submittal', 'openSubmittalSidebar')
    .addToUi();
}
```

(O item "🔌 Ativar automação" é adicionado só na Task 6, junto com a função `_sub_trg_activate` que ele chama — referenciar uma função que ainda não existe daria erro se alguém clicasse no item entre as tasks.)

- [ ] **Step 8: Inicializar git e criar o repositório no GitHub**

```bash
cd "C:/DEV/Sheets/Submittal-GAS"
git init
git add .
git commit -m "$(cat <<'EOF'
chore: esqueleto inicial do projeto Submittal-GAS

Extraido de Ferramentas-PLB-Sheets para isolar o Submittal (em evolucao
ativa) dos outros 3 projetos GAS que compartilhavam o mesmo repositorio.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

**Antes de criar o repositório remoto e dar push, pare e confirme com o usuário** (ação visível/pública — criar repo no GitHub) — só depois:

```bash
gh repo create thiagobarretosn-hue/Submittal-GAS --private --source=. --remote=origin
git push -u origin main
```

---

## Task 2: Portar Submittal.gs + sidebar como estão (baseline antes do refactor)

**Files:**
- Create: `C:\DEV\Sheets\Submittal-GAS\Submittal.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalSidebar.html`

**Interfaces:**
- Consumes: `include()`, `SharedConfig_createDocConfigService()`, `SharedUtils_numberToColumnLetter()` (Task 1)
- Produces: todas as funções já existentes (`openSubmittalSidebar`, `getSubmittalInitData`, `getSubmittalItems`, `montarSubmittal`, etc.) — sem mudança de assinatura nesta task, só cópia.

- [ ] **Step 1: Copiar o conteúdo exato de `Ferramentas-PLB-Sheets/Submittal.gs` e `SubmittalSidebar.html`**

Copiar linha por linha (sem alterar nada ainda) para os arquivos novos. Isso estabelece uma base de comparação: se algo quebrar nas próximas tasks, dá pra voltar aqui e comparar.

- [ ] **Step 2: Push e teste manual de fumaça na planilha real**

```bash
cd "C:/DEV/Sheets/Submittal-GAS"
clasp push
```

Na planilha "Submittal Items LOG" (Google Sheets, no navegador): recarregar a página, abrir menu **📦 Submittal → 📦 Montar Submittal**, confirmar que a sidebar abre sem erro no console e que os dropdowns de aba/número carregam. **Não precisa rodar uma montagem completa ainda** — só confirmar que abriu sem erro (o layout de colunas hardcoded ainda é o antigo, então rodar uma montagem real pode falhar contra o layout novo — isso é esperado e corrigido na Task 3).

- [ ] **Step 3: Commit**

```bash
git add Submittal.gs SubmittalSidebar.html
git commit -m "feat: portar Submittal.gs e sidebar de Ferramentas-PLB-Sheets (baseline)"
```

---

## Task 3: Resolução de colunas por cabeçalho (sem mais posição fixa)

**Files:**
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalColumns.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalColumns.test.js`
- Modify: `C:\DEV\Sheets\Submittal-GAS\Submittal.gs`

**Interfaces:**
- Produces: `_sub_buildColumnMap(headerRow)` → `{ NOMECABECALHO: indice1Indexed }`; `_sub_getColumnMap(sheet)` → mesmo formato, lendo a linha 1 da aba; `_sub_requireColumns(colMap, requiredNames)` → lança erro se faltar alguma.
- Consumes (na Task 4/5): `_sub_getColumnMap`, `_sub_requireColumns`.

- [ ] **Step 1: Escrever o teste (função pura, roda com node puro)**

`SubmittalColumns.test.js`:

```javascript
const assert = require('assert');
const { _sub_buildColumnMap } = require('./SubmittalColumns.gs');

// header com espaço sobrando (bate com o dado real: "FOLDER " tem espaço no fim)
const headers = ['#N', '#SUBMITTAL AMBAR', 'SUBMITTAL DATE', 'SUBMITTAL TITLE',
  'DISCIPLINE', 'FOLDER ', 'ROOM', 'LOCATION', 'UPC', 'ITEM', 'CREATOR',
  'STATUS', 'STATUS DATE', 'NOTE', 'PDF', 'LINK SUBMITTAL'];

const map = _sub_buildColumnMap(headers);

assert.strictEqual(map['DISCIPLINE'], 5, 'DISCIPLINE deveria ser a coluna 5');
assert.strictEqual(map['FOLDER'], 6, 'FOLDER (com espaco sobrando no header) deveria virar coluna 6');
assert.strictEqual(map['ITEM'], 10, 'ITEM deveria ser a coluna 10');
assert.strictEqual(map['LINK SUBMITTAL'], 16, 'LINK SUBMITTAL deveria ser a coluna 16');
assert.strictEqual(map['NAO EXISTE'], undefined, 'cabecalho inexistente nao deveria aparecer no mapa');

// ordem diferente não quebra nada — é exatamente o problema que isso resolve
const headersReordenados = ['ITEM', 'DISCIPLINE', 'STATUS'];
const map2 = _sub_buildColumnMap(headersReordenados);
assert.strictEqual(map2['ITEM'], 1);
assert.strictEqual(map2['DISCIPLINE'], 2);
assert.strictEqual(map2['STATUS'], 3);

console.log('OK: SubmittalColumns.test.js — todos os asserts passaram');
```

- [ ] **Step 2: Rodar o teste e confirmar que falha (arquivo ainda não existe)**

```bash
cd "C:/DEV/Sheets/Submittal-GAS"
node SubmittalColumns.test.js
```

Expected: erro `Cannot find module './SubmittalColumns.gs'` ou similar.

- [ ] **Step 3: Implementar SubmittalColumns.gs**

```javascript
/**
 * @fileoverview Resolucao de colunas por nome de cabecalho — substitui posicao fixa.
 * Tolerante a reordenacao de coluna: só quebra se o NOME do cabecalho mudar.
 */

function _sub_buildColumnMap(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var name = String(headerRow[i] || '').trim().toUpperCase();
    if (name) map[name] = i + 1; // 1-indexed, igual Range do Sheets
  }
  return map;
}

function _sub_getColumnMap(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return _sub_buildColumnMap(headerRow);
}

function _sub_requireColumns(colMap, requiredNames) {
  var missing = requiredNames.filter(function(name) { return !colMap[name]; });
  if (missing.length > 0) {
    throw new Error('Colunas obrigatórias não encontradas: ' + missing.join(', '));
  }
}

if (typeof module !== 'undefined') {
  module.exports = { _sub_buildColumnMap: _sub_buildColumnMap };
}
```

- [ ] **Step 4: Rodar o teste de novo, confirmar que passa**

```bash
node SubmittalColumns.test.js
```

Expected: `OK: SubmittalColumns.test.js — todos os asserts passaram`

- [ ] **Step 5: Trocar `SUBMITTAL_CFG.COLS` fixo por `_sub_getColumnMap` em Submittal.gs**

Em `Submittal.gs`, remover o bloco `COLS: { NUM: 2, DATE: 3, ... }` de `SUBMITTAL_CFG` e trocar toda referência a `SUBMITTAL_CFG.COLS.X` por uma chamada a `_sub_getColumnMap(sheet)` no início de cada função pública que lê a planilha (`getSubmittalNumbers`, `getSubmittalItems`, `saveSubmittalItemLink`, `montarSubmittal`), usando os nomes de cabeçalho reais: `'#SUBMITTAL AMBAR'` (NUM), `'SUBMITTAL DATE'` (DATE), `'DISCIPLINE'`, `'ITEM'`, `'STATUS'`, `'PDF'` (LINK — a coluna que carrega o cut-sheet chama `PDF` no layout novo), `'LINK SUBMITTAL'` (LINK_OUT).

Exemplo da mudança em `getSubmittalNumbers`:

```javascript
function getSubmittalNumbers(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const colMap = _sub_getColumnMap(sheet);
  _sub_requireColumns(colMap, ['#SUBMITTAL AMBAR', 'STATUS']);

  const lastNeededCol = Math.max(colMap['#SUBMITTAL AMBAR'], colMap['STATUS']);
  const data = sheet.getRange(2, 1, lastRow - 1, lastNeededCol).getValues();
  const counts = {};
  data.forEach(function(row) {
    const status = String(row[colMap['STATUS'] - 1] || '').trim().toUpperCase();
    if (status !== SUBMITTAL_CFG.STATUS_FILTER) return;
    const num = String(row[colMap['#SUBMITTAL AMBAR'] - 1] || '').trim();
    if (!num) return;
    counts[num] = (counts[num] || 0) + 1;
  });

  return Object.keys(counts)
    .sort(function(a, b) { return a.localeCompare(b, undefined, { numeric: true }); })
    .map(function(num) { return { num: num, count: counts[num] }; });
}
```

Aplicar o mesmo padrão nas outras 4 funções que hoje usam `SUBMITTAL_CFG.COLS`. Código completo de cada uma:

`_sub_getLinksMap` — passa a receber a coluna do link como parâmetro (`linkCol`), em vez de ler `SUBMITTAL_CFG.COLS.LINK` fixo:

```javascript
function _sub_getLinksMap(sheetName, startRow, endRow, linkCol) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const colLetter = SharedUtils_numberToColumnLetter(linkCol);
  const range = "'" + sheetName + "'!" + colLetter + startRow + ':' + colLetter + endRow;

  const resp = Sheets.Spreadsheets.get(ssId, {
    ranges: [range],
    fields: 'sheets(data(rowData(values(hyperlink,formattedValue,chipRuns,textFormatRuns))))'
  });

  const map = {};
  try {
    const rows = resp.sheets[0].data[0].rowData || [];
    rows.forEach(function(rd, i) {
      const rowNum = startRow + i;
      const cell = (rd && rd.values && rd.values[0]) ? rd.values[0] : null;
      if (!cell) return;

      const label = String(cell.formattedValue || '').trim();

      if (cell.chipRuns && cell.chipRuns.length > 0) {
        for (let c = 0; c < cell.chipRuns.length; c++) {
          const chip = cell.chipRuns[c].chip;
          if (chip && chip.richLinkProperties && chip.richLinkProperties.uri) {
            map[rowNum] = { url: chip.richLinkProperties.uri, label: label || 'arquivo' };
            return;
          }
        }
      }

      if (cell.hyperlink) {
        map[rowNum] = { url: cell.hyperlink, label: label || cell.hyperlink };
        return;
      }

      if (cell.textFormatRuns && cell.textFormatRuns.length > 0) {
        for (let r = 0; r < cell.textFormatRuns.length; r++) {
          const fmt = cell.textFormatRuns[r].format;
          if (fmt && fmt.link && fmt.link.uri) {
            map[rowNum] = { url: fmt.link.uri, label: label || fmt.link.uri };
            return;
          }
        }
      }

      if (/^https?:\/\//i.test(label)) {
        map[rowNum] = { url: label, label: label };
      }
    });
  } catch (e) {
    console.error('[Submittal] Erro ao ler links: ' + e.message);
  }
  return map;
}
```

`getSubmittalItems`:

```javascript
function getSubmittalItems(num, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'Aba "' + sheetName + '" não encontrada.' };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Aba sem dados.' };

  const colMap = _sub_getColumnMap(sheet);
  _sub_requireColumns(colMap, ['#SUBMITTAL AMBAR', 'SUBMITTAL DATE', 'DISCIPLINE', 'ITEM', 'STATUS', 'PDF']);

  const data = sheet.getRange(2, 1, lastRow - 1, colMap['PDF']).getValues();
  const linksMap = _sub_getLinksMap(sheetName, 2, lastRow, colMap['PDF']);

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const items = [];
  data.forEach(function(row, i) {
    const rowNum = i + 2;
    const status = String(row[colMap['STATUS'] - 1] || '').trim().toUpperCase();
    const rowSubNum = String(row[colMap['#SUBMITTAL AMBAR'] - 1] || '').trim();
    if (status !== SUBMITTAL_CFG.STATUS_FILTER || rowSubNum !== String(num).trim()) return;

    let dateVal = row[colMap['SUBMITTAL DATE'] - 1];
    if (dateVal instanceof Date) dateVal = Utilities.formatDate(dateVal, tz, 'MM/dd/yyyy');

    items.push({
      row: rowNum,
      item: String(row[colMap['ITEM'] - 1] || '').trim(),
      discipline: String(row[colMap['DISCIPLINE'] - 1] || '').trim(),
      date: String(dateVal || '').trim(),
      link: linksMap[rowNum] || null
    });
  });

  return { success: true, items: items };
}
```

`saveSubmittalItemLink`:

```javascript
function saveSubmittalItemLink(rowNum, url, label, sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Aba não encontrada.' };
    const colMap = _sub_getColumnMap(sheet);
    _sub_requireColumns(colMap, ['PDF']);
    const cleanUrl = String(url || '').trim();
    if (!/^https?:\/\//i.test(cleanUrl)) {
      return { success: false, message: 'URL inválida (deve começar com http/https).' };
    }
    const text = String(label || '').trim() || cleanUrl;
    const rt = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(cleanUrl).build();
    sheet.getRange(rowNum, colMap['PDF']).setRichTextValue(rt);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
```

`montarSubmittal` — só a parte final muda (gravação da coluna `LINK SUBMITTAL`); todo o resto da função (coleta de PDFs, merge com pdf-lib, capa) fica **idêntico** ao que já existe hoje. Trecho final substituído:

```javascript
    // ── Coluna LINK SUBMITTAL em todas as linhas do submittal ──
    _sub_setProgress({ status: 'running', step: 'Gravando links na planilha...', done: loadedDocs.length + 1, total: loadedDocs.length + 1 });
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const colMap = _sub_getColumnMap(sheet);
    _sub_requireColumns(colMap, ['LINK SUBMITTAL']);

    const pdfUrl = pdfFile.getUrl();
    const linkRt = SpreadsheetApp.newRichTextValue().setText(pdfName).setLinkUrl(pdfUrl).build();
    items.forEach(function(it) {
      sheet.getRange(it.row, colMap['LINK SUBMITTAL']).setRichTextValue(linkRt);
    });
```

Isso substitui o trecho equivalente que hoje usa `SUBMITTAL_CFG.COLS.LINK_OUT` e o workaround de `insertColumnsAfter` (esse workaround existia porque o layout antigo não tinha a coluna M — no layout novo, `LINK SUBMITTAL` já é um cabeçalho real da aba, então a coluna sempre existe e o workaround não é mais necessário). O restante da função (da declaração `async function montarSubmittal` até a coleta/merge de PDFs) continua exatamente igual ao código já existente em `Ferramentas-PLB-Sheets/Submittal.gs` — só chame `getSubmittalItems`/`saveSubmittalItemLink` já atualizadas acima, que o resto flui sem mudança.

- [ ] **Step 6: Push e teste manual completo na planilha real**

```bash
clasp push
```

Na aba `FINISHES TEMPLATE` (ou uma cópia de teste dela) da "Submittal Items LOG": abrir a sidebar, selecionar essa aba, carregar os itens do número `000`, confirmar que aparecem os itens e os links corretamente (agora lendo pela coluna `PDF`, não mais pela antiga posição fixa `L`).

- [ ] **Step 7: Commit**

```bash
git add SubmittalColumns.gs SubmittalColumns.test.js Submittal.gs
git commit -m "feat: resolver colunas por nome de cabecalho em vez de posicao fixa"
```

---

## Task 4: Catálogo DATA BASE SUBMITTAL (leitura e gravação)

**Files:**
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalCatalog.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalCatalog.test.js`
- Modify: `C:\DEV\Sheets\Submittal-GAS\Submittal.gs` (adicionar `CATALOG_SHEET_NAME` em `SUBMITTAL_CFG`)

**Interfaces:**
- Consumes: `_sub_getColumnMap`, `_sub_requireColumns` (Task 3)
- Produces: `_sub_normalizeItemKey(text)` → string normalizada pra comparação; `_sub_catalog_lookup(itemName)` → `{ row, discipline, folder, room, location, item, pdf } | null`; `_sub_catalog_upsert(entry)` → atualiza a linha do item se já existe (classificação nova) ou acrescenta linha nova (com `entry = { discipline, folder, room, location, item }`). Task 6 (gatilho) consome as duas últimas.

- [ ] **Step 1: Escrever o teste da função pura de normalização**

`SubmittalCatalog.test.js`:

```javascript
const assert = require('assert');
const { _sub_normalizeItemKey } = require('./SubmittalCatalog.gs');

assert.strictEqual(
  _sub_normalizeItemKey('  Plastic Access Panel 6 in. x 6 in.  '),
  'PLASTIC ACCESS PANEL 6 IN. X 6 IN.',
  'deve remover espaco nas pontas e normalizar caixa'
);
assert.strictEqual(
  _sub_normalizeItemKey('shower   head trim'),
  'SHOWER HEAD TRIM',
  'espacos duplicados no meio devem virar um so espaco'
);
assert.strictEqual(_sub_normalizeItemKey(''), '', 'string vazia continua vazia');
assert.strictEqual(_sub_normalizeItemKey(null), '', 'null vira string vazia, nao erro');
assert.strictEqual(
  _sub_normalizeItemKey('Shower Head Trim') === _sub_normalizeItemKey('  shower  head  trim  '),
  true,
  'duas variacoes do mesmo item devem gerar a mesma chave'
);

console.log('OK: SubmittalCatalog.test.js — todos os asserts passaram');
```

- [ ] **Step 2: Rodar e confirmar que falha**

```bash
node SubmittalCatalog.test.js
```

Expected: erro de módulo não encontrado.

- [ ] **Step 3: Implementar SubmittalCatalog.gs**

```javascript
/**
 * @fileoverview DATA BASE SUBMITTAL — catalogo de itens (discipline/folder/room/location/pdf).
 * UPC nunca e lido/escrito aqui: e formula propria da planilha.
 */

function _sub_normalizeItemKey(text) {
  return String(text || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function _sub_catalog_getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMITTAL_CFG.CATALOG_SHEET_NAME);
  if (!sheet) throw new Error('Aba "' + SUBMITTAL_CFG.CATALOG_SHEET_NAME + '" não encontrada.');
  return sheet;
}

function _sub_catalog_readIndex() {
  var sheet = _sub_catalog_getSheet();
  var colMap = _sub_getColumnMap(sheet);
  _sub_requireColumns(colMap, ['DISCIPLINE', 'FOLDER', 'ROOM', 'LOCATION', 'ITEM', 'PDF']);

  var lastRow = sheet.getLastRow();
  var index = {};
  if (lastRow < 2) return { colMap: colMap, index: index };

  var numCols = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  data.forEach(function(row, i) {
    var item = String(row[colMap['ITEM'] - 1] || '').trim();
    if (!item) return;
    index[_sub_normalizeItemKey(item)] = {
      row:        i + 2, // linha real na aba — usado pelo upsert pra atualizar in-place
      discipline: String(row[colMap['DISCIPLINE'] - 1] || '').trim(),
      folder:     String(row[colMap['FOLDER'] - 1] || '').trim(),
      room:       String(row[colMap['ROOM'] - 1] || '').trim(),
      location:   String(row[colMap['LOCATION'] - 1] || '').trim(),
      item:       item,
      pdf:        String(row[colMap['PDF'] - 1] || '').trim()
    };
  });
  return { colMap: colMap, index: index };
}

/** Busca um item na base pelo nome (tolerante a espaco/caixa). Retorna null se nao achar. */
function _sub_catalog_lookup(itemName) {
  var built = _sub_catalog_readIndex();
  return built.index[_sub_normalizeItemKey(itemName)] || null;
}

/**
 * Atualiza a linha do item na DATA BASE SUBMITTAL (classificacao mudou) ou acrescenta
 * linha nova (item inedito). Manter o catalogo em sincronia com o Drive e essencial:
 * se o arquivo foi movido pra outra DISCIPLINE/FOLDER/ROOM mas o catalogo guardasse a
 * classificacao antiga, um autofill futuro procuraria o arquivo na pasta velha, nao
 * acharia, e uma nova colagem de link criaria DUPLICATA — exatamente o que o requisito
 * "mover, nunca duplicar" proibe.
 * entry: { discipline, folder, room, location, item }
 * Protegido por LockService — dois usuarios editando ao mesmo tempo nao duplicam linha.
 */
function _sub_catalog_upsert(entry) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    var built = _sub_catalog_readIndex();
    var sheet = _sub_catalog_getSheet();
    var colMap = built.colMap;
    var existing = built.index[_sub_normalizeItemKey(entry.item)];

    if (existing) {
      sheet.getRange(existing.row, colMap['DISCIPLINE']).setValue(entry.discipline || '');
      sheet.getRange(existing.row, colMap['FOLDER']).setValue(entry.folder || '');
      sheet.getRange(existing.row, colMap['ROOM']).setValue(entry.room || '');
      // LOCATION so atualiza se veio preenchida — nao apagar dado do catalogo com vazio
      if (String(entry.location || '').trim()) {
        sheet.getRange(existing.row, colMap['LOCATION']).setValue(entry.location);
      }
      return;
    }

    var row = new Array(sheet.getLastColumn()).fill('');
    row[colMap['DISCIPLINE'] - 1] = entry.discipline || '';
    row[colMap['FOLDER'] - 1]     = entry.folder || '';
    row[colMap['ROOM'] - 1]       = entry.room || '';
    row[colMap['LOCATION'] - 1]   = entry.location || '';
    row[colMap['ITEM'] - 1]       = entry.item || '';
    row[colMap['PDF'] - 1]        = (entry.item || '').trim() + '.pdf';
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

if (typeof module !== 'undefined') {
  module.exports = { _sub_normalizeItemKey: _sub_normalizeItemKey };
}
```

- [ ] **Step 4: Adicionar `CATALOG_SHEET_NAME` em `SUBMITTAL_CFG` (Submittal.gs)**

```javascript
const SUBMITTAL_CFG = {
  // ... campos ja existentes ...
  CATALOG_SHEET_NAME: 'DATA BASE SUBMITTAL'
};
```

- [ ] **Step 5: Rodar o teste, confirmar que passa**

```bash
node SubmittalCatalog.test.js
```

Expected: `OK: SubmittalCatalog.test.js — todos os asserts passaram`

- [ ] **Step 6: Teste manual na planilha real (via editor de Apps Script, executar uma função ad-hoc)**

No editor do Apps Script (`clasp open` ou pelo navegador), criar temporariamente uma função e rodar 1x pelo próprio editor (Executar):

```javascript
function _sub_debug_testCatalogLookup() {
  var r = _sub_catalog_lookup('Plastic Access Panel 6 in. X 6 in.');
  Logger.log(JSON.stringify(r));
}
```

Expected no log: objeto com `discipline: "PLUMBING"`, `folder: "FIXTURES"`, `room: "VARIES"`, `location: "C.A."`, `pdf: "PLASTIC ACCESS PANEL 6 IN. X 6 IN..pdf"` — bate com a linha 2 real da DATA BASE SUBMITTAL. Depois de confirmar, apagar essa função de debug (não fica no código final).

- [ ] **Step 7: Commit**

```bash
git add SubmittalCatalog.gs SubmittalCatalog.test.js Submittal.gs
git commit -m "feat: catalogo DATA BASE SUBMITTAL (leitura indexada + upsert com lock)"
```

---

## Task 5: Organização automática no Drive (02-PRECON)

**Files:**
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalDriveOrg.gs`
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalDriveOrg.test.js`
- Modify: `C:\DEV\Sheets\Submittal-GAS\Submittal.gs` (adicionar `PRECON_ROOT_FOLDER_ID`)

**Interfaces:**
- Consumes: `_sub_extractDriveId` (já existe em `Submittal.gs`), `_sub_resolveBlob` (já existe em `Submittal.gs`)
- Produces: `_sub_buildFolderPath(discipline, folder, room)` → `[disciplina, pasta, sala]` ou lança erro; `_sub_org_resolveExistingFile(richTextValue)` → `File | null`; `_sub_org_organizeItem(params)` → `{ file: File|null, created: boolean }`, onde `params = { itemName, discipline, folder, room, sourceUrl, existingFile }`. Task 6 consome `_sub_buildFolderPath` (indiretamente) e `_sub_org_organizeItem` + `_sub_org_resolveExistingFile`.

- [ ] **Step 1: Escrever o teste da função pura de caminho de pasta**

`SubmittalDriveOrg.test.js`:

```javascript
const assert = require('assert');
const { _sub_buildFolderPath } = require('./SubmittalDriveOrg.gs');

assert.deepStrictEqual(
  _sub_buildFolderPath('PLUMBING', 'FIXTURES', 'BATH'),
  ['PLUMBING', 'FIXTURES', 'BATH'],
  'deve retornar os 3 segmentos limpos, na ordem DISCIPLINE > FOLDER > ROOM'
);
assert.deepStrictEqual(
  _sub_buildFolderPath('  Plumbing  ', ' Fixtures ', ' Bath '),
  ['Plumbing', 'Fixtures', 'Bath'],
  'deve aparar espaco nas pontas, mas preservar a caixa original (nome de pasta legivel)'
);

assert.throws(
  function() { _sub_buildFolderPath('', 'FIXTURES', 'BATH'); },
  /incompletos/,
  'DISCIPLINE vazio deve lancar erro'
);
assert.throws(
  function() { _sub_buildFolderPath('PLUMBING', '', 'BATH'); },
  /incompletos/,
  'FOLDER vazio deve lancar erro'
);
assert.throws(
  function() { _sub_buildFolderPath('PLUMBING', 'FIXTURES', ''); },
  /incompletos/,
  'ROOM vazio deve lancar erro'
);

console.log('OK: SubmittalDriveOrg.test.js — todos os asserts passaram');
```

- [ ] **Step 2: Rodar e confirmar que falha**

```bash
node SubmittalDriveOrg.test.js
```

- [ ] **Step 3: Implementar SubmittalDriveOrg.gs**

```javascript
/**
 * @fileoverview Organizacao automatica de cut-sheets no Shared Drive fixo (02-PRECON).
 * Estrutura: <RAIZ FIXA>/DISCIPLINE/FOLDER/ROOM/<ITEM>.pdf
 * Raiz fixa e global — nao configuravel por obra/aba (repositorio unico e geral).
 */

function _sub_buildFolderPath(discipline, folder, room) {
  var clean = function(s) { return String(s || '').trim(); };
  var segments = [clean(discipline), clean(folder), clean(room)];
  if (segments.some(function(s) { return !s; })) {
    throw new Error('DISCIPLINE/FOLDER/ROOM incompletos — não é possível organizar o arquivo.');
  }
  return segments;
}

function _sub_getOrCreateNestedFolder(rootFolder, segments) {
  var folder = rootFolder;
  for (var i = 0; i < segments.length; i++) {
    var name = segments[i];
    var it = folder.getFoldersByName(name);
    folder = it.hasNext() ? it.next() : folder.createFolder(name);
  }
  return folder;
}

/** Extrai o arquivo do Drive de uma RichTextValue de celula, se ela ja for um link do Drive. */
function _sub_org_resolveExistingFile(richTextValue) {
  var url = richTextValue && richTextValue.getLinkUrl && richTextValue.getLinkUrl();
  if (!url) return null;
  var fileId = _sub_extractDriveId(url);
  if (!fileId) return null;
  try {
    return DriveApp.getFileById(fileId);
  } catch (e) {
    return null; // link quebrado/arquivo apagado — trata como se nao existisse
  }
}

/** Move o arquivo pra pasta certa SE ainda nao estiver la (Shared Drive: 1 pai so). */
function _sub_org_ensureFileLocation(file, targetFolder) {
  var parents = file.getParents();
  var currentParent = parents.hasNext() ? parents.next() : null;
  if (currentParent && currentParent.getId() === targetFolder.getId()) return file;
  return file.moveTo(targetFolder);
}

/**
 * Garante que o item esta organizado na pasta certa. Idempotente: reprocessar o mesmo
 * item nao duplica nem falha.
 *
 * params:
 *   itemName     (string, obrigatorio) — vira o nome do arquivo (+ .pdf)
 *   discipline, folder, room (string, obrigatorios)
 *   sourceUrl    (string|null) — link novo a baixar/copiar (usado quando ainda nao ha arquivo)
 *   existingFile (File|null)   — arquivo ja resolvido da celula atual, se houver
 *
 * Retorna { file: File|null, created: boolean }.
 *   file === null significa "nada a organizar ainda" (sem link novo e nao achou nada existente
 *   com esse nome na pasta certa) — quem chama decide o que fazer (ex: deixar celula em branco).
 */
function _sub_org_organizeItem(params) {
  var segments = _sub_buildFolderPath(params.discipline, params.folder, params.room);
  var root = DriveApp.getFolderById(SUBMITTAL_CFG.PRECON_ROOT_FOLDER_ID);
  var targetFolder = _sub_getOrCreateNestedFolder(root, segments);
  var fileName = String(params.itemName || '').trim() + '.pdf';

  if (params.existingFile) {
    return { file: _sub_org_ensureFileLocation(params.existingFile, targetFolder), created: false };
  }

  var byName = targetFolder.getFilesByName(fileName);
  if (byName.hasNext()) {
    return { file: byName.next(), created: false };
  }

  if (!params.sourceUrl) return { file: null, created: false };

  var resolved = _sub_resolveBlob(params.sourceUrl);
  if (resolved.kind !== 'pdf') {
    throw new Error('Arquivo "' + resolved.name + '" não é PDF — organização de cut-sheet exige PDF.');
  }
  var created = targetFolder.createFile(resolved.blob.setName(fileName));
  return { file: created, created: true };
}

if (typeof module !== 'undefined') {
  module.exports = { _sub_buildFolderPath: _sub_buildFolderPath };
}
```

- [ ] **Step 4: Rodar o teste, confirmar que passa**

```bash
node SubmittalDriveOrg.test.js
```

- [ ] **Step 5: Adicionar `PRECON_ROOT_FOLDER_ID` em `SUBMITTAL_CFG` (Submittal.gs)**

```javascript
const SUBMITTAL_CFG = {
  // ... campos ja existentes ...
  PRECON_ROOT_FOLDER_ID: '1YqVuKzEDJTAl503zYCoaHXW7-vovs7ja'
};
```

- [ ] **Step 6: Push e teste manual na planilha real**

```bash
clasp push
```

No editor do Apps Script, rodar 1x (apagar depois de confirmar):

```javascript
function _sub_debug_testOrganize() {
  var r = _sub_org_organizeItem({
    itemName: 'TESTE AUTOMACAO SUBMITTAL',
    discipline: 'TESTE',
    folder: 'TESTE',
    room: 'TESTE',
    sourceUrl: 'https://drive.google.com/file/d/ALGUM_ID_DE_PDF_REAL/view',
    existingFile: null
  });
  Logger.log(JSON.stringify({ created: r.created, url: r.file ? r.file.getUrl() : null }));
}
```

Expected: no Drive (pasta raiz `1YqVuKzEDJTAl503zYCoaHXW7-vovs7ja`), aparecem as subpastas `TESTE/TESTE/TESTE` com o arquivo `TESTE AUTOMACAO SUBMITTAL.pdf` dentro. Rodar a função de novo (mesmos parâmetros) e confirmar `created: false` na segunda vez (idempotência). Apagar a pasta de teste `TESTE` do Drive e a função `_sub_debug_testOrganize` depois de confirmar.

- [ ] **Step 7: Commit**

```bash
git add SubmittalDriveOrg.gs SubmittalDriveOrg.test.js Submittal.gs
git commit -m "feat: organizacao automatica de cut-sheets no Shared Drive (02-PRECON)"
```

---

## Task 6: Gatilho onEdit instalável (autofill + ingestão)

**Files:**
- Create: `C:\DEV\Sheets\Submittal-GAS\SubmittalTriggers.gs`
- Modify: `C:\DEV\Sheets\Submittal-GAS\Menu.gs` (adicionar o item "🔌 Ativar automação")

**Interfaces:**
- Consumes: `_sub_getColumnMap` (Task 3), `_sub_catalog_lookup`/`_sub_catalog_upsert` (Task 4), `_sub_org_organizeItem`/`_sub_org_resolveExistingFile` (Task 5)
- Produces: `_sub_trg_activate()` (chamada pelo menu), `_sub_trg_onEditInstalled(e)` (handler do gatilho)

- [ ] **Step 1: Validar a premissa "escrita por script não re-dispara onEdit" ANTES de construir o resto**

Esse é o risco técnico mais importante do plano (ver spec, seção de riscos). No editor do Apps Script, criar temporariamente:

```javascript
function _sub_debug_installLoopTest() {
  ScriptApp.newTrigger('_sub_debug_onEditLoopTest').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
}
function _sub_debug_onEditLoopTest(e) {
  var props = PropertiesService.getScriptProperties();
  var count = Number(props.getProperty('LOOP_TEST_COUNT') || 0) + 1;
  props.setProperty('LOOP_TEST_COUNT', String(count));
  Logger.log('onEdit disparou. count=' + count + ' cell=' + e.range.getA1Notation());
  if (count < 20) {
    e.range.offset(0, 1).setValue('escrita-do-script-' + count); // escreve na celula do lado
  }
}
```

Rodar `_sub_debug_installLoopTest` uma vez pelo editor (autoriza o gatilho). Depois, **editar manualmente 1 célula na planilha** (digitar qualquer coisa numa célula vazia de teste) e checar `Extensões > Apps Script > Execuções`. Expected: **exatamente 1 execução** do `_sub_debug_onEditLoopTest` (a escrita feita pelo próprio script na célula do lado NÃO deve gerar uma segunda execução). Se aparecer mais de uma execução (indício de que escrita por script realmente re-dispara onEdit), pare e avise o usuário antes de prosseguir — a lógica de guarda por coluna (Step 3 abaixo) continua válida de qualquer forma, mas seria ainda mais importante nesse cenário.

Depois do teste: apagar o gatilho de teste (`ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === '_sub_debug_onEditLoopTest') ScriptApp.deleteTrigger(t); })`, rodar 1x pelo editor), remover as 2 funções de debug e a propriedade `LOOP_TEST_COUNT`.

- [ ] **Step 2: Implementar SubmittalTriggers.gs**

```javascript
/**
 * @fileoverview Gatilho onEdit instalavel: autofill (coluna ITEM), ingestao (coluna PDF)
 * e move por mudanca de classificacao (DISCIPLINE/FOLDER/ROOM).
 * Instalavel (nao onEdit(e) simples) porque mover/criar arquivo no Drive exige autorizacao completa.
 * Premissa validada no Step 1: edicao feita por script NAO re-dispara onEdit — so edicao
 * humana. Mesmo se falhasse, todas as operacoes sao idempotentes (nao loopam, so gastam
 * uma execucao extra).
 */

function _sub_trg_activate() {
  _sub_trg_deactivate();
  ScriptApp.newTrigger('_sub_trg_onEditInstalled')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('Automação do Submittal ativada (autofill + organização de PDF).');
}

function _sub_trg_deactivate() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === '_sub_trg_onEditInstalled') ScriptApp.deleteTrigger(t);
  });
}

/** Handler instalado — nunca usar getUi().alert()/prompt() aqui (sem contexto de UI). */
function _sub_trg_onEditInstalled(e) {
  try {
    _sub_trg_handleEdit(e);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Erro na automação do Submittal: ' + err.message, '⚠️ Submittal', 10);
    console.error('[Submittal onEdit] ' + err.message + '\n' + (err.stack || ''));
  }
}

function _sub_trg_handleEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() === SUBMITTAL_CFG.CATALOG_SHEET_NAME) return; // nao reage a edicao na propria base

  var colMap = _sub_getColumnMap(sheet);
  var itemCol = colMap['ITEM'];
  var pdfCol = colMap['PDF'];
  if (!itemCol && !pdfCol) return; // aba sem esse layout — ignora

  var startRow = e.range.getRow();
  if (startRow < 2) return; // nao mexe no cabecalho

  var startCol = e.range.getColumn();
  var endCol = startCol + e.range.getNumColumns() - 1;
  var endRow = startRow + e.range.getNumRows() - 1;

  var inRange = function(col) { return !!col && col >= startCol && col <= endCol; };
  var touchesItem = inRange(itemCol);
  var touchesPdf = inRange(pdfCol);
  // Editar a classificacao move o arquivo (requisito: mover, nunca duplicar) — a
  // ingestao ja detecta pasta-destino diferente e faz o moveTo + upsert do catalogo.
  var touchesClassification = inRange(colMap['DISCIPLINE']) || inRange(colMap['FOLDER']) || inRange(colMap['ROOM']);
  if (!touchesItem && !touchesPdf && !touchesClassification) return;

  for (var row = startRow; row <= endRow; row++) {
    if (touchesItem) _sub_trg_runAutofill(sheet, colMap, row);
    if (touchesPdf || touchesClassification) _sub_trg_runIngest(sheet, colMap, row);
  }
}

function _sub_trg_runAutofill(sheet, colMap, row) {
  var itemName = String(sheet.getRange(row, colMap['ITEM']).getValue() || '').trim();
  if (!itemName) return;

  var match = _sub_catalog_lookup(itemName);
  if (!match) return; // item novo — nada a preencher ainda, fica pro fluxo de ingestao

  if (colMap['DISCIPLINE']) sheet.getRange(row, colMap['DISCIPLINE']).setValue(match.discipline);
  if (colMap['FOLDER'])     sheet.getRange(row, colMap['FOLDER']).setValue(match.folder);
  if (colMap['ROOM'])       sheet.getRange(row, colMap['ROOM']).setValue(match.room);
  if (colMap['LOCATION'])   sheet.getRange(row, colMap['LOCATION']).setValue(match.location);

  if (!colMap['PDF'] || !match.discipline || !match.folder || !match.room) return;

  // busca ao vivo — a base so guarda o NOME do arquivo, nao um link; pode nao existir
  // organizado ainda (itens legados cadastrados manualmente antes desta automacao existir)
  var found = _sub_org_organizeItem({
    itemName: itemName,
    discipline: match.discipline,
    folder: match.folder,
    room: match.room,
    sourceUrl: null,
    existingFile: null
  });
  if (found.file) {
    var rt = SpreadsheetApp.newRichTextValue()
      .setText(itemName + '.pdf')
      .setLinkUrl(found.file.getUrl())
      .build();
    sheet.getRange(row, colMap['PDF']).setRichTextValue(rt);
  }
  // se found.file for null: item conhecido mas nunca organizado de fato — PDF fica em
  // branco, e uma futura edicao na coluna PDF (usuario colando o link) cai no fluxo de ingestao.
}

function _sub_trg_runIngest(sheet, colMap, row) {
  if (!colMap['ITEM']) return;
  var itemName = String(sheet.getRange(row, colMap['ITEM']).getValue() || '').trim();
  if (!itemName) return; // sem item, nao da pra organizar nem catalogar

  var discipline = colMap['DISCIPLINE'] ? String(sheet.getRange(row, colMap['DISCIPLINE']).getValue() || '').trim() : '';
  var folder     = colMap['FOLDER']     ? String(sheet.getRange(row, colMap['FOLDER']).getValue() || '').trim()     : '';
  var room       = colMap['ROOM']       ? String(sheet.getRange(row, colMap['ROOM']).getValue() || '').trim()       : '';
  if (!discipline || !folder || !room) return; // sem os 3, nao da pra organizar ainda

  var pdfCell = sheet.getRange(row, colMap['PDF']);
  var richText = pdfCell.getRichTextValue();
  var existingFile = _sub_org_resolveExistingFile(richText);

  var plainText = String(pdfCell.getValue() || '').trim();
  var currentUrl = richText && richText.getLinkUrl ? richText.getLinkUrl() : null;
  var sourceUrl = null;
  if (!existingFile) {
    if (currentUrl) sourceUrl = currentUrl;
    else if (/^https?:\/\//i.test(plainText)) sourceUrl = plainText;
  }
  if (!existingFile && !sourceUrl) return; // celula sem link e sem texto de URL — nada a fazer

  var result = _sub_org_organizeItem({
    itemName: itemName,
    discipline: discipline,
    folder: folder,
    room: room,
    sourceUrl: sourceUrl,
    existingFile: existingFile
  });
  if (!result.file) return;

  var rt = SpreadsheetApp.newRichTextValue()
    .setText(itemName + '.pdf')
    .setLinkUrl(result.file.getUrl())
    .build();
  pdfCell.setRichTextValue(rt);

  // Upsert SEMPRE (nao so quando o item e novo): se a classificacao mudou e o arquivo
  // acabou de ser movido, o catalogo precisa acompanhar — senao um autofill futuro
  // apontaria pra pasta velha e uma nova colagem de link criaria duplicata.
  _sub_catalog_upsert({
    discipline: discipline,
    folder: folder,
    room: room,
    location: colMap['LOCATION'] ? String(sheet.getRange(row, colMap['LOCATION']).getValue() || '').trim() : '',
    item: itemName
  });
}
```

- [ ] **Step 3: Adicionar o item de ativação no Menu.gs**

Substituir o conteúdo de `Menu.gs` por:

```javascript
/**
 * @fileoverview Menu Principal — Submittal-GAS
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Submittal')
    .addItem('📦 Montar Submittal', 'openSubmittalSidebar')
    .addSeparator()
    .addItem('🔌 Ativar automação (autofill + organização)', '_sub_trg_activate')
    .addToUi();
}
```

- [ ] **Step 4: Push**

```bash
clasp push
```

- [ ] **Step 5: Teste manual — item NOVO (fluxo de ingestão)**

Numa linha vazia da aba de teste: digitar um `ITEM` que **não existe** na DATA BASE SUBMITTAL, preencher manualmente `DISCIPLINE`/`FOLDER`/`ROOM`, e colar um link de um PDF real (do Drive ou da web) na célula `PDF`. Expected: em poucos segundos, a célula `PDF` vira um hyperlink pro arquivo já dentro de `RAIZ/DISCIPLINE/FOLDER/ROOM/<ITEM>.pdf`, e uma linha nova aparece na `DATA BASE SUBMITTAL` com esses dados.

- [ ] **Step 6: Teste manual — item CONHECIDO (fluxo de autofill)**

Numa outra linha vazia: digitar um `ITEM` que **já existe** na DATA BASE SUBMITTAL (ex: `PLASTIC ACCESS PANEL 6 IN. X 6 IN.`, o mesmo teste da Task 4). Expected: `DISCIPLINE`/`FOLDER`/`ROOM`/`LOCATION` preenchem sozinhos; `PDF` preenche com link se esse item já tiver sido organizado por outro teste anterior, ou fica em branco se ainda não (comportamento esperado pros itens legados nunca organizados).

- [ ] **Step 7: Teste manual — mudança de classificação MOVE o arquivo (nunca duplica)**

Na linha do item criado no Step 5 (já organizado, PDF com link): editar a célula `ROOM` pra outro valor (ex: de `TESTE` pra `TESTE2`). Expected:
1. O arquivo **some** da pasta antiga (`RAIZ/DISCIPLINE/FOLDER/TESTE/`) e **aparece** na nova (`RAIZ/DISCIPLINE/FOLDER/TESTE2/`) — conferir no Drive que existe **uma única cópia** (movido, não duplicado).
2. A linha desse item na `DATA BASE SUBMITTAL` foi atualizada com o `ROOM` novo (upsert).
3. O link na célula `PDF` continua funcionando (o `fileId` não muda num move).

- [ ] **Step 8: Teste manual — paste de linha inteira**

Copiar uma linha inteira de uma aba (todas as colunas de uma vez) e colar numa linha vazia da aba de teste. Expected: mesmo comportamento das Steps 5/6 (autofill e/ou ingestão rodam pra essa linha), sem erro — confirma que o tratamento de `e.range` com várias colunas/linhas de uma vez funciona.

- [ ] **Step 9: Ativar e confirmar persistência do gatilho**

Pelo menu **📦 Submittal → 🔌 Ativar automação**, confirmar o alerta de ativação. Rodar `clasp push` de novo (simulando um deploy futuro) e confirmar, em `Extensões > Apps Script > Gatilhos`, que o gatilho `_sub_trg_onEditInstalled` continua listado (push não apaga gatilho).

- [ ] **Step 10: Commit**

```bash
git add SubmittalTriggers.gs Menu.gs
git commit -m "feat: gatilho onEdit instalavel (autofill, ingestao e move por classificacao)"
```

---

## Task 7: Verificação final de convivência com "Montar Submittal"

**Files:**
- Nenhum arquivo novo — só validação end-to-end.

- [ ] **Step 1: Rodar o fluxo completo de montagem numa aba com itens organizados pela automação**

Na aba de teste, com pelo menos 2-3 itens já processados pelas Tasks 5/6 (PDF apontando pro arquivo organizado), abrir **📦 Montar Submittal**, configurar a "Pasta pai no Drive" (config já existente, separada do 02-PRECON), carregar os itens do número de teste, e rodar a montagem.

Expected: o merge funciona normalmente (PDF final com capa + itens mesclados), confirmando que ler o link a partir da coluna `PDF` (resolvida por cabeçalho, Task 3) — que agora sempre aponta pro arquivo já organizado no 02-PRECON — não quebrou nada do fluxo existente.

- [ ] **Step 2: Conferir o relatório final**

Confirmar que o relatório da sidebar mostra a contagem certa de itens mesclados e nenhum item "sem link" pros que já foram organizados pela automação.

- [ ] **Step 3: Push final e commit de fechamento**

```bash
clasp push
git add -A
git commit -m "chore: validacao end-to-end — Montar Submittal convive com autofill/organizacao automatica" --allow-empty
git push
```
