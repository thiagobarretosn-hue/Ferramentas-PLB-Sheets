# Request Generator — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Criar uma ferramenta independente de geração de requisição de materiais no formato KOJO, com arredondamento automático configurável por tipo de material.

**Architecture:** `Request.gs` contém todo o backend (config própria via PropertiesService, processamento batch, escrita do output). `RequestSidebar.html` é a sidebar completa com 4 seções. `Menu.gs` recebe apenas um novo item de menu. Zero dependência de `BOM.gs`.

**Tech Stack:** Google Apps Script (V8/ES6+), SpreadsheetApp API, HtmlService, PropertiesService, setFormulas() para batch de fórmulas.

**Spec:** `docs/superpowers/specs/2026-05-07-request-generator-design.md`

**Nota sobre testes:** GAS não tem test runner. Os "testes" são funções manuais executadas no Script Editor — verificar output no Execution Log (View > Execution log).

---

## File Structure

| Arquivo | Ação | Responsabilidade |
|---------|------|-----------------|
| `Request.gs` | Criar | Config, helpers de coluna, queries de dados, processamento, entry points |
| `RequestSidebar.html` | Criar | UI completa da sidebar (4 seções + JS) |
| `Menu.gs` | Modificar | Adicionar item `📋 Gerador de Request` |

---

## Task 1: Request.gs — Constants, ConfigService e Column Helpers

**Files:**
- Create: `Request.gs`

- [ ] **Step 1: Criar `Request.gs` com constantes, ConfigService e helpers de coluna**

```javascript
/**
 * @fileoverview Request Generator — Gerador de Requisição de Materiais KOJO
 * @version 1.0.0
 */

// ============================================================================
// CONSTANTS
// ============================================================================
const REQUEST_CONFIG = {
  SETTINGS_KEY: 'REQUEST_SETTINGS_V1',
  CACHE_TTL: 180,
  HEADER_ROWS: 7,
  COL_HEADER_ROW: 9,
  DATA_START_ROW: 10,
  COLORS: {
    HEADER_BG: '#2c3e50',
    FONT_LIGHT: '#ffffff'
  },
  DEFAULT_RULES: [
    { pattern: 'FOAM CORE', roundUp: 20 },
    { pattern: 'CPVC', roundUp: 10 }
  ],
  COL_WIDTHS: { A: 120, B: 80, C: 550, D: 100, E: 100, F: 100 }
};

// ============================================================================
// CONFIG SERVICE
// ============================================================================
const RequestConfigService = {
  _defaults: function() {
    return {
      sourceSheet: '',
      colDesc: '',
      colUpc: '',
      colUom: '',
      colQty: '',
      groupL1: '',
      groupL2: '',
      groupL3: '',
      project: '',
      kojoPrefix: '',
      engineer: '',
      version: '01',
      roundingRules: REQUEST_CONFIG.DEFAULT_RULES.map(function(r) {
        return { pattern: r.pattern, roundUp: r.roundUp };
      })
    };
  },

  get: function() {
    try {
      const saved = PropertiesService.getDocumentProperties()
        .getProperty(REQUEST_CONFIG.SETTINGS_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        return Object.assign(RequestConfigService._defaults(), parsed);
      }
    } catch (e) {
      console.error('[RequestConfigService] Erro ao ler config:', e.message);
    }
    return RequestConfigService._defaults();
  },

  save: function(config) {
    try {
      PropertiesService.getDocumentProperties()
        .setProperty(REQUEST_CONFIG.SETTINGS_KEY, JSON.stringify(config));
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }
};

// ============================================================================
// COLUMN HELPERS
// ============================================================================

/**
 * "J - DESC" → 10 (1-indexed). Suporta colunas AA, AB, etc.
 */
function _req_getColumnIndex(colConfig) {
  if (!colConfig) return -1;
  const match = String(colConfig).match(/^([A-Z]+)\s*-/);
  if (!match) return -1;
  const letters = match[1];
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index;
}

/**
 * 1 → "A", 26 → "Z", 27 → "AA"
 */
function _req_numberToLetter(n) {
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

/**
 * Retorna array de strings "A - Header" para todos os headers da aba
 */
function _req_getColumnsFromSheet(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.map(function(h, i) {
    const letter = _req_numberToLetter(i + 1);
    return letter + ' - ' + (h || 'Col ' + letter);
  });
}

/**
 * Aplica a primeira regra que matcheia o DESC (case-insensitive).
 * Fallback: 1. Garante mínimo de 1 para evitar divisão por zero.
 */
function _req_applyRoundingRule(desc, rules) {
  if (!rules || rules.length === 0) return 1;
  const upper = String(desc).toUpperCase();
  const match = rules.find(function(r) {
    return r.pattern && upper.includes(r.pattern.toUpperCase());
  });
  return Math.max(match ? (Number(match.roundUp) || 1) : 1, 1);
}
```

- [ ] **Step 2: Criar função de teste manual e executar no Script Editor**

Adicionar ao final de `Request.gs`:

```javascript
// ============================================================================
// TESTES MANUAIS — executar no Script Editor, verificar Execution Log
// ============================================================================
function testRequestHelpers() {
  const rules = REQUEST_CONFIG.DEFAULT_RULES;
  const results = [];

  // _req_getColumnIndex
  results.push('J → ' + _req_getColumnIndex('J - DESC') + ' (expect 10)');
  results.push('A → ' + _req_getColumnIndex('A - ID') + ' (expect 1)');
  results.push('AA → ' + _req_getColumnIndex('AA - EXTRA') + ' (expect 27)');
  results.push('vazio → ' + _req_getColumnIndex('') + ' (expect -1)');

  // _req_numberToLetter
  results.push('1 → ' + _req_numberToLetter(1) + ' (expect A)');
  results.push('26 → ' + _req_numberToLetter(26) + ' (expect Z)');
  results.push('27 → ' + _req_numberToLetter(27) + ' (expect AA)');

  // _req_applyRoundingRule
  results.push('FOAM CORE → ' + _req_applyRoundingRule('PIPE 2 IN FOAM CORE', rules) + ' (expect 20)');
  results.push('CPVC → ' + _req_applyRoundingRule('PIPE 1 IN. X 10 FT. CPVC', rules) + ' (expect 10)');
  results.push('ELBOW → ' + _req_applyRoundingRule('ELBOW 90 DEGREES 2 IN.', rules) + ' (expect 1)');
  results.push('roundUp=0 → ' + _req_applyRoundingRule('PIPE FOAM CORE', [{pattern:'FOAM CORE', roundUp:0}]) + ' (expect 1)');

  // ConfigService defaults
  const defaults = RequestConfigService.get();
  results.push('defaults.roundingRules.length: ' + defaults.roundingRules.length + ' (expect 2)');
  results.push('defaults.version: ' + defaults.version + ' (expect 01)');

  console.log(results.join('\n'));
}
```

Executar `testRequestHelpers` no Script Editor.
Esperado no Execution Log: todos os valores com `(expect X)` batem.

- [ ] **Step 3: Commit**

```
git add Request.gs
git commit -m "feat(request): constants, ConfigService e column helpers"
```

---

## Task 2: Request.gs — Data Query Functions

**Files:**
- Modify: `Request.gs`

- [ ] **Step 1: Adicionar funções de query de dados**

Adicionar após os helpers em `Request.gs`:

```javascript
// ============================================================================
// DATA QUERIES — chamadas pelo HTML via google.script.run
// ============================================================================

/**
 * Dados iniciais para a sidebar ao carregar.
 * Retorna lista de abas, colunas da aba configurada e config salva.
 */
function getRequestInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = RequestConfigService.get();
  const allSheets = ss.getSheets().map(function(s) { return s.getName(); });

  let allColumns = [];
  if (config.sourceSheet) {
    const sourceSheet = ss.getSheetByName(config.sourceSheet);
    if (sourceSheet) allColumns = _req_getColumnsFromSheet(sourceSheet);
  }

  return { allSheets: allSheets, allColumns: allColumns, config: config };
}

/**
 * Retorna colunas de uma aba pelo nome.
 * Chamada quando usuário troca a aba fonte na sidebar.
 */
function getRequestSheetColumns(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet ? _req_getColumnsFromSheet(sheet) : [];
}

/**
 * Valores únicos de uma coluna para popular dropdowns de grupo.
 */
function getRequestUniqueValues(groupCol, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];

  const colIndex = _req_getColumnIndex(groupCol);
  if (colIndex < 0) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
  return [...new Set(values.filter(function(v) { return v !== '' && v !== null; }))]
    .map(function(v) { return String(v); })
    .sort(function(a, b) { return a.localeCompare(b, undefined, { numeric: true }); });
}

/**
 * Conta itens da fonte que batem com a combinação selecionada.
 * groupCols: ["I - FLOOR", "H - PHASE"]
 * groupVals: ["6th", "Job Site"]
 */
function getRequestItemCount(groupCols, groupVals, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const activeGroups = groupCols.filter(Boolean);
  if (activeGroups.length === 0) return 0;

  const groupIndices = activeGroups.map(_req_getColumnIndex);
  if (groupIndices.some(function(i) { return i < 0; })) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let count = 0;

  data.forEach(function(row) {
    const matches = groupIndices.every(function(colIdx, i) {
      return String(row[colIdx - 1] || '').trim() === String(groupVals[i] || '').trim();
    });
    if (matches) count++;
  });

  return count;
}

/**
 * Salva config do Request no PropertiesService.
 */
function saveRequestConfig(config) {
  return RequestConfigService.save(config);
}
```

- [ ] **Step 2: Teste manual**

```javascript
function testRequestDataQueries() {
  // Ajustar para a planilha ativa no momento do teste
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = ss.getActiveSheet().getName();
  console.log('Aba ativa:', sheetName);

  const cols = getRequestSheetColumns(sheetName);
  console.log('Colunas (primeiras 3):', cols.slice(0, 3).join(', '));
  console.log('getRequestInitData retorna allSheets:', getRequestInitData().allSheets.length > 0 ? 'OK' : 'VAZIO');
}
```

Executar `testRequestDataQueries`. Esperado: nome da aba ativa, lista de colunas, "OK" para allSheets.

- [ ] **Step 3: Commit**

```
git add Request.gs
git commit -m "feat(request): data query functions (initData, uniqueValues, itemCount)"
```

---

## Task 3: Request.gs — Output Writers

**Files:**
- Modify: `Request.gs`

- [ ] **Step 1: Adicionar funções de escrita do output**

```javascript
// ============================================================================
// OUTPUT WRITERS (privados)
// ============================================================================

function _req_writeHeader(sheet, settings, combination, groupCols) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const lastUpdate = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');
  const kojoFull = settings.kojoPrefix
    ? settings.kojoPrefix + '.' + (settings.kojoSuffix || '')
    : (settings.kojoSuffix || '');

  const generatedFrom = groupCols
    .filter(Boolean)
    .map(function(col, i) {
      const label = col.replace(/^[A-Z]+ - /, '');
      return label + ': ' + (combination.parts[i] || '');
    })
    .join(' | ');

  // Col A: labels
  const labels = [['PROJECT:'], ['REQUEST:'], ['BOM KOJO:'], ['ENG.:'], ['VERSION:'], ['LAST UPDATE:'], ['GENERATED FROM:']];
  sheet.getRange(1, 1, 7, 1).setValues(labels);

  // Col B: main values
  const mainVals = [
    [settings.project || ''],
    [settings.request || ''],
    [kojoFull],
    [settings.engineer || ''],
    [settings.version || '01'],
    [lastUpdate],
    [generatedFrom]
  ];
  sheet.getRange(1, 2, 7, 1).setValues(mainVals);

  // Col E e F: Requisition # e Need By (apenas nas linhas corretas)
  sheet.getRange(1, 5).setValue('Requisition #');
  sheet.getRange(1, 6).setValue('Need By');
  sheet.getRange(3, 5).setValue(settings.requisitionNum || '');
  sheet.getRange(3, 6).setValue(settings.needBy || '');

  // Merge B:D por linha (depois de setar valores)
  for (let r = 1; r <= 7; r++) {
    sheet.getRange(r, 2, 1, 3).merge();
  }

  sheet.getRange(1, 1, 7, 6).setNumberFormat('@STRING@');
  sheet.getRange(1, 1, 7, 6).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function _req_writeData(sheet, rows, dataStartRow) {
  if (rows.length === 0) return;

  // Colunas B-F: valores batch
  const values = rows.map(function(r) {
    return [r.uom, r.desc, r.upc, r.qty, r.roundUp];
  });
  sheet.getRange(dataStartRow, 2, rows.length, 5).setValues(values);

  // Coluna A: fórmulas batch via setFormulas (mais rápido que loop com setFormula)
  const formulas = rows.map(function(_, i) {
    const rowNum = dataStartRow + i;
    return ['=ROUNDUP(E' + rowNum + '/F' + rowNum + ')*F' + rowNum];
  });
  sheet.getRange(dataStartRow, 1, rows.length, 1).setFormulas(formulas);
}

function _req_formatSheet(sheet, dataRowCount) {
  const colHeaderRow = REQUEST_CONFIG.COL_HEADER_ROW;

  // Banding nas linhas de dados (inclui header de colunas)
  sheet.getRange(colHeaderRow, 1, dataRowCount + 1, 6)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  // Larguras
  sheet.setColumnWidth(1, REQUEST_CONFIG.COL_WIDTHS.A);
  sheet.setColumnWidth(2, REQUEST_CONFIG.COL_WIDTHS.B);
  sheet.setColumnWidth(3, REQUEST_CONFIG.COL_WIDTHS.C);
  sheet.setColumnWidth(4, REQUEST_CONFIG.COL_WIDTHS.D);
  sheet.setColumnWidth(5, REQUEST_CONFIG.COL_WIDTHS.E);
  sheet.setColumnWidth(6, REQUEST_CONFIG.COL_WIDTHS.F);
}
```

- [ ] **Step 2: Commit**

```
git add Request.gs
git commit -m "feat(request): output writer helpers (header, data, format)"
```

---

## Task 4: Request.gs — processRequestCore

**Files:**
- Modify: `Request.gs`

- [ ] **Step 1: Adicionar a função principal de processamento**

```javascript
// ============================================================================
// CORE PROCESSING
// ============================================================================

/**
 * Processa e gera a aba de request.
 *
 * @param {{parts: string[], label: string}} combination - Combinação selecionada
 * @param {Object} settings - Config completa + campos do header
 * @returns {{success: boolean, count?: number, message?: string}}
 */
function processRequestCore(combination, settings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(settings.sourceSheet);
    if (!sourceSheet) {
      return { success: false, message: 'Aba "' + settings.sourceSheet + '" não encontrada.' };
    }

    const descIdx  = _req_getColumnIndex(settings.colDesc)  - 1;
    const upcIdx   = _req_getColumnIndex(settings.colUpc)   - 1;
    const uomIdx   = _req_getColumnIndex(settings.colUom)   - 1;
    const qtyIdx   = _req_getColumnIndex(settings.colQty)   - 1;

    if ([descIdx, upcIdx, uomIdx, qtyIdx].some(function(i) { return i < 0; })) {
      return { success: false, message: 'Configuração de colunas inválida. Verifique a seção Configuração.' };
    }

    const groupCols = [settings.groupL1, settings.groupL2, settings.groupL3].filter(Boolean);
    const groupVals = combination.parts;
    const groupIndices = groupCols.map(_req_getColumnIndex);

    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Aba fonte está vazia.' };

    // 1. Leitura batch
    const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();

    // 2. Filtro + agrupamento (DESC + UPC + UOM como chave)
    const grouped = {};
    data.forEach(function(row) {
      const matches = groupIndices.every(function(colIdx, i) {
        return String(row[colIdx - 1] || '').trim() === String(groupVals[i] || '').trim();
      });
      if (!matches) return;

      const desc = String(row[descIdx] || '').trim();
      if (!desc) return;

      const upc  = String(row[upcIdx]  || '').trim();
      const uom  = String(row[uomIdx]  || '').trim();
      const qty  = Number(row[qtyIdx]) || 0;
      const key  = desc + '|||' + upc + '|||' + uom;

      if (grouped[key]) {
        grouped[key].qty += qty;
      } else {
        grouped[key] = { desc: desc, upc: upc, uom: uom, qty: qty };
      }
    });

    // 3. Ordenar por DESC
    const rows = Object.values(grouped).sort(function(a, b) {
      return a.desc.localeCompare(b.desc, undefined, { numeric: true });
    });

    if (rows.length === 0) {
      return { success: false, message: 'Nenhum item encontrado para a combinação selecionada.' };
    }

    // 4. Aplicar regras de arredondamento
    const rules = settings.roundingRules || REQUEST_CONFIG.DEFAULT_RULES;
    rows.forEach(function(row) {
      row.roundUp = _req_applyRoundingRule(row.desc, rules);
    });

    // 5. Criar ou limpar aba de destino
    const sheetName = String(settings.kojoSuffix || 'Request').trim().substring(0, 100);
    let targetSheet = ss.getSheetByName(sheetName);
    if (targetSheet) {
      targetSheet.clear();
    } else {
      targetSheet = ss.insertSheet(sheetName);
    }

    // 6. Escrever header
    _req_writeHeader(targetSheet, settings, combination, groupCols);

    // 7. Escrever cabeçalho de colunas (linha 9)
    const colHeaderRow = REQUEST_CONFIG.COL_HEADER_ROW;
    targetSheet.getRange(colHeaderRow, 1, 1, 6)
      .setValues([['QTY (ROUND UP)', 'UOM', 'DESC', 'UPC', 'QTY', 'ROUND UP']]);
    targetSheet.getRange(colHeaderRow, 1, 1, 6).setFontWeight('bold');

    // 8. Escrever dados (linha 10+)
    _req_writeData(targetSheet, rows, REQUEST_CONFIG.DATA_START_ROW);

    // 9. Formatar
    _req_formatSheet(targetSheet, rows.length);

    // 10. Proteger header
    try {
      const prot = targetSheet.getRange(1, 1, REQUEST_CONFIG.HEADER_ROWS, 6).protect();
      prot.setDescription('Header protegido').removeEditors(prot.getEditors());
    } catch (e) { /* ignora — pode não ter permissão em alguns contextos */ }

    return { success: true, count: rows.length };

  } catch (e) {
    console.error('[processRequestCore]', e.message, e.stack);
    return { success: false, message: e.message };
  }
}
```

- [ ] **Step 2: Teste manual de processamento (sem sidebar)**

```javascript
function testProcessRequestCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Ajustar esses valores para sua planilha de teste
  const TEST_SHEET = ss.getActiveSheet().getName();
  const TEST_COLS = _req_getColumnsFromSheet(ss.getActiveSheet());
  console.log('Colunas disponíveis:', TEST_COLS.slice(0, 5).join(', '));

  // Para rodar um teste real, preencher com valores reais da planilha:
  // const result = processRequestCore(
  //   { parts: ['6th', 'Job Site'], label: '6th | Job Site' },
  //   {
  //     sourceSheet: 'REVIT DES CRPLB RISER REQ',
  //     colDesc: 'J - DESC', colUpc: 'M - UPC', colUom: 'L - UOM', colQty: 'O - QTY',
  //     groupL1: 'I - FLOOR', groupL2: 'H - PHASE', groupL3: '',
  //     project: 'TEST', kojoPrefix: 'TEST', kojoSuffix: 'TEST.REQUEST',
  //     engineer: 'TEST', version: '01', request: 'Test Request',
  //     requisitionNum: 'REQ-001', needBy: '05/30/2026',
  //     roundingRules: [{ pattern: 'FOAM CORE', roundUp: 20 }, { pattern: 'CPVC', roundUp: 10 }]
  //   }
  // );
  // console.log('Result:', JSON.stringify(result));
  console.log('Descomentar o bloco acima com valores reais para testar o processamento completo.');
}
```

Executar `testProcessRequestCore` no Script Editor.

- [ ] **Step 3: Commit**

```
git add Request.gs
git commit -m "feat(request): processRequestCore — filtro, agrupamento, rounding, output"
```

---

## Task 5: Request.gs — Entry Points

**Files:**
- Modify: `Request.gs`

- [ ] **Step 1: Adicionar entry points públicos**

```javascript
// ============================================================================
// ENTRY POINTS — chamados pelo Menu e pelo HTML
// ============================================================================

/**
 * Abre a sidebar do Request Generator.
 * Chamado pelo menu.
 */
function openRequestSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('RequestSidebar')
    .setTitle('📋 Gerador de Request')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}
```

- [ ] **Step 2: Commit**

```
git add Request.gs
git commit -m "feat(request): openRequestSidebar entry point"
```

---

## Task 6: Menu.gs — Adicionar Item de Menu

**Files:**
- Modify: `Menu.gs` (linhas 47–53, bloco do menu `🔧 Relatórios Dinâmicos`)

- [ ] **Step 1: Localizar o bloco do menu e adicionar o item**

Localizar em `Menu.gs` o trecho:
```javascript
  ui.createMenu('🔧 Relatórios Dinâmicos')
    .addItem('📊 Gerador de BOM (Painel)', 'openBomSidebar')
    .addSeparator()
    .addItem('🔧 Fixadores → Fonte', 'abrirSeletorFixadores')
    .addSeparator()
    .addItem('📄 Exportar PDFs', 'exportPDFsWithFeedback')
    .addSeparator()
    .addItem('🗑️ Limpar Relatórios', 'clearOldReports')
    .addItem('🔄 Limpar Cache', 'forceRefreshCache')
    .addItem('🧪 Diagnóstico', 'testSystem')
    .addToUi();
```

Substituir por:
```javascript
  ui.createMenu('🔧 Relatórios Dinâmicos')
    .addItem('📊 Gerador de BOM (Painel)', 'openBomSidebar')
    .addSeparator()
    .addItem('📋 Gerador de Request', 'openRequestSidebar')
    .addSeparator()
    .addItem('🔧 Fixadores → Fonte', 'abrirSeletorFixadores')
    .addSeparator()
    .addItem('📄 Exportar PDFs', 'exportPDFsWithFeedback')
    .addSeparator()
    .addItem('🗑️ Limpar Relatórios', 'clearOldReports')
    .addItem('🔄 Limpar Cache', 'forceRefreshCache')
    .addItem('🧪 Diagnóstico', 'testSystem')
    .addToUi();
```

- [ ] **Step 2: Verificar**

Recarregar a planilha no Google Sheets (fechar e reabrir, ou Tools > Script editor > Run `onOpen`).
Verificar que `🔧 Relatórios Dinâmicos` > `📋 Gerador de Request` aparece no menu.

- [ ] **Step 3: Commit**

```
git add Menu.gs
git commit -m "feat(request): adicionar item de menu Gerador de Request"
```

---

## Task 7: RequestSidebar.html — Sidebar Completa

**Files:**
- Create: `RequestSidebar.html`

- [ ] **Step 1: Criar `RequestSidebar.html` completo**

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; font-size: 12px; color: #2c3e50; background: #f8f9fa; }
    .header { background: #2c3e50; color: #fff; padding: 10px 14px; font-weight: bold; font-size: 13px; }
    .section { padding: 10px 14px; border-bottom: 1px solid #e0e0e0; background: #fff; }
    .section-title { font-weight: bold; color: #2c3e50; margin-bottom: 8px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; }
    label { display: block; margin-bottom: 2px; color: #555; font-size: 10px; }
    select, input[type="text"] {
      width: 100%; padding: 4px 6px; border: 1px solid #bdc3c7; border-radius: 3px;
      font-size: 11px; margin-bottom: 6px; background: #fff;
    }
    select:focus, input:focus { outline: none; border-color: #3498db; }
    .row2 { display: flex; gap: 6px; }
    .row2 > div { flex: 1; }
    .btn { padding: 5px 10px; border: none; border-radius: 3px; cursor: pointer; font-size: 11px; }
    .btn-primary { background: #27ae60; color: #fff; width: 100%; padding: 10px; font-size: 13px; font-weight: bold; margin-top: 4px; }
    .btn-primary:disabled { background: #95a5a6; cursor: not-allowed; }
    .btn-secondary { background: #3498db; color: #fff; }
    .btn-small { padding: 2px 7px; font-size: 10px; }
    .btn-danger { background: #e74c3c; color: #fff; }
    .preview { background: #eaf4fb; border: 1px solid #aed6f1; padding: 5px 8px; border-radius: 3px; font-size: 11px; color: #2471a3; margin-top: 4px; }
    .kojo-full { font-size: 10px; color: #7f8c8d; margin-top: -4px; margin-bottom: 6px; font-style: italic; min-height: 14px; }
    table.rules { width: 100%; border-collapse: collapse; margin-bottom: 6px; }
    table.rules th { background: #ecf0f1; padding: 4px 5px; text-align: left; font-size: 10px; border: 1px solid #ddd; }
    table.rules td { padding: 2px 3px; border: 1px solid #ddd; }
    table.rules input[type="text"] { margin: 0; padding: 2px 4px; font-size: 11px; }
    .fallback-note { font-size: 10px; color: #7f8c8d; margin-top: 4px; }
    #status { padding: 8px 14px; font-size: 11px; min-height: 24px; }
    .status-ok { color: #27ae60; }
    .status-err { color: #e74c3c; }
    .loading { color: #7f8c8d; font-style: italic; }
    .warning { background: #fef9e7; border: 1px solid #f0b27a; padding: 5px 8px; border-radius: 3px; font-size: 10px; color: #935116; margin-bottom: 6px; }
  </style>
</head>
<body>

<div class="header">📋 Gerador de Request</div>

<!-- ==================== SECTION 1: CONFIG ==================== -->
<div class="section">
  <div class="section-title">⚙️ Configuração</div>

  <label>Aba Fonte</label>
  <select id="sel-source-sheet"></select>

  <div class="row2">
    <div><label>Coluna DESC</label><select id="sel-col-desc"></select></div>
    <div><label>Coluna UPC</label><select id="sel-col-upc"></select></div>
  </div>
  <div class="row2">
    <div><label>Coluna UOM</label><select id="sel-col-uom"></select></div>
    <div><label>Coluna QTY (raw)</label><select id="sel-col-qty"></select></div>
  </div>

  <label>Agrupar por Nível 1</label>
  <select id="sel-group-l1"></select>
  <label>Agrupar por Nível 2</label>
  <select id="sel-group-l2"></select>
  <label>Agrupar por Nível 3</label>
  <select id="sel-group-l3"></select>

  <button class="btn btn-secondary" id="btn-save-config" style="width:100%">💾 Salvar Configuração</button>
</div>

<!-- ==================== SECTION 2: HEADER ==================== -->
<div class="section">
  <div class="section-title">📋 Header do Request</div>

  <div class="row2">
    <div><label>PROJECT</label><input type="text" id="inp-project" /></div>
    <div><label>VERSION</label><input type="text" id="inp-version" value="01" /></div>
  </div>

  <label>REQUEST (descrição)</label>
  <input type="text" id="inp-request" placeholder="ex: RISERS 6th Floor" />

  <label>KOJO PREFIX</label>
  <input type="text" id="inp-kojo-prefix" placeholder="ex: MTN.PLB.RGH.JS.B1067" />

  <label>BOM KOJO Suffix</label>
  <input type="text" id="inp-kojo-suffix" placeholder="ex: CA.RSR.F6" />
  <div class="kojo-full" id="kojo-preview"></div>

  <label>ENG.</label>
  <input type="text" id="inp-engineer" />

  <div class="row2">
    <div><label>Requisition #</label><input type="text" id="inp-req-num" placeholder="REQ-A4799" /></div>
    <div><label>Need By</label><input type="text" id="inp-need-by" placeholder="MM/DD/YYYY" /></div>
  </div>
</div>

<!-- ==================== SECTION 3: COMBINATION ==================== -->
<div class="section">
  <div class="section-title">🔽 Seleção da Combinação</div>
  <div id="group-selectors">
    <p class="loading">Salve a configuração para carregar os grupos.</p>
  </div>
  <div class="preview" id="combo-preview" style="display:none"></div>
</div>

<!-- ==================== SECTION 4: ROUNDING ==================== -->
<div class="section">
  <div class="section-title">🔢 Regras de Arredondamento</div>
  <table class="rules">
    <thead>
      <tr>
        <th>Padrão (texto no DESC)</th>
        <th style="width:65px">Round Up</th>
        <th style="width:28px"></th>
      </tr>
    </thead>
    <tbody id="rules-body"></tbody>
  </table>
  <button class="btn btn-small btn-secondary" id="btn-add-rule">+ Regra</button>
  <p class="fallback-note">↳ Fallback: itens sem match → Round Up = 1</p>
</div>

<!-- ==================== GENERATE ==================== -->
<div style="padding:12px 14px;">
  <button class="btn btn-primary" id="btn-generate">📋 GERAR REQUEST</button>
</div>
<div id="status"></div>

<script>
  /* ============================================================
     STATE
  ============================================================ */
  var state = { config: {}, allSheets: [], allColumns: [] };

  /* ============================================================
     INIT
  ============================================================ */
  document.addEventListener('DOMContentLoaded', function() {
    bindEvents();
    google.script.run
      .withSuccessHandler(onInitData)
      .withFailureHandler(onError)
      .getRequestInitData();
    showStatus('Carregando...', 'loading');
  });

  function bindEvents() {
    document.getElementById('btn-save-config').addEventListener('click', onSaveConfig);
    document.getElementById('btn-add-rule').addEventListener('click', function() { addRuleRow('', ''); });
    document.getElementById('btn-generate').addEventListener('click', onGenerate);
    document.getElementById('inp-kojo-prefix').addEventListener('input', updateKojoPreview);
    document.getElementById('inp-kojo-suffix').addEventListener('input', updateKojoPreview);
    document.getElementById('sel-source-sheet').addEventListener('change', onSourceSheetChange);
  }

  function onInitData(data) {
    state.allSheets = data.allSheets || [];
    state.allColumns = data.allColumns || [];
    state.config = data.config || {};

    populateSheetDropdown(state.allSheets, state.config.sourceSheet);
    populateColumnDropdowns(state.allColumns, state.config);
    populateHeaderFields(state.config);
    populateRulesTable(state.config.roundingRules);

    if (state.config.groupL1 && state.config.sourceSheet) {
      loadGroupSelectors(state.config);
    }
    showStatus('', '');
  }

  function onError(err) {
    showStatus('Erro: ' + (err.message || err), 'err');
  }

  /* ============================================================
     POPULATE HELPERS
  ============================================================ */
  function populateSheetDropdown(sheets, selected) {
    var sel = document.getElementById('sel-source-sheet');
    sel.innerHTML = '<option value="">-- Selecione --</option>';
    sheets.forEach(function(s) {
      var opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      if (s === selected) opt.selected = true;
      sel.appendChild(opt);
    });
  }

  function populateColumnDropdowns(columns, config) {
    var map = {
      'sel-col-desc': config.colDesc || '',
      'sel-col-upc':  config.colUpc  || '',
      'sel-col-uom':  config.colUom  || '',
      'sel-col-qty':  config.colQty  || '',
      'sel-group-l1': config.groupL1 || '',
      'sel-group-l2': config.groupL2 || '',
      'sel-group-l3': config.groupL3 || ''
    };
    Object.keys(map).forEach(function(id) {
      var sel = document.getElementById(id);
      var isGroup = id.startsWith('sel-group');
      sel.innerHTML = isGroup
        ? '<option value="">-- Nenhum --</option>'
        : '<option value="">-- Selecione --</option>';
      columns.forEach(function(col) {
        var opt = document.createElement('option');
        opt.value = col; opt.textContent = col;
        if (col === map[id]) opt.selected = true;
        sel.appendChild(opt);
      });
    });
  }

  function populateHeaderFields(config) {
    document.getElementById('inp-project').value  = config.project    || '';
    document.getElementById('inp-version').value  = config.version    || '01';
    document.getElementById('inp-kojo-prefix').value = config.kojoPrefix || '';
    document.getElementById('inp-engineer').value = config.engineer   || '';
    updateKojoPreview();
  }

  function populateRulesTable(rules) {
    document.getElementById('rules-body').innerHTML = '';
    var defaultRules = (rules && rules.length > 0)
      ? rules
      : [{ pattern: 'FOAM CORE', roundUp: 20 }, { pattern: 'CPVC', roundUp: 10 }];
    defaultRules.forEach(function(r) { addRuleRow(r.pattern, r.roundUp); });
  }

  function addRuleRow(pattern, roundUp) {
    var tbody = document.getElementById('rules-body');
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td><input type="text" class="rule-pattern" value="' + (pattern || '') + '" placeholder="ex: FOAM CORE"></td>' +
      '<td><input type="text" class="rule-roundup" value="' + (roundUp || '') + '" style="width:55px"></td>' +
      '<td><button class="btn btn-small btn-danger btn-rm">✕</button></td>';
    tr.querySelector('.btn-rm').addEventListener('click', function() { tr.remove(); });
    tbody.appendChild(tr);
  }

  /* ============================================================
     SOURCE SHEET CHANGE
  ============================================================ */
  function onSourceSheetChange() {
    var sheetName = document.getElementById('sel-source-sheet').value;
    if (!sheetName) return;
    google.script.run
      .withSuccessHandler(function(cols) {
        state.allColumns = cols;
        populateColumnDropdowns(cols, {});
      })
      .withFailureHandler(onError)
      .getRequestSheetColumns(sheetName);
  }

  /* ============================================================
     SAVE CONFIG
  ============================================================ */
  function onSaveConfig() {
    var config = collectConfig();
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          state.config = config;
          showStatus('✅ Configuração salva!', 'ok');
          loadGroupSelectors(config);
        } else {
          showStatus('Erro ao salvar: ' + result.message, 'err');
        }
      })
      .withFailureHandler(onError)
      .saveRequestConfig(config);
  }

  function collectConfig() {
    return {
      sourceSheet: document.getElementById('sel-source-sheet').value,
      colDesc:  document.getElementById('sel-col-desc').value,
      colUpc:   document.getElementById('sel-col-upc').value,
      colUom:   document.getElementById('sel-col-uom').value,
      colQty:   document.getElementById('sel-col-qty').value,
      groupL1:  document.getElementById('sel-group-l1').value,
      groupL2:  document.getElementById('sel-group-l2').value,
      groupL3:  document.getElementById('sel-group-l3').value,
      project:  document.getElementById('inp-project').value,
      kojoPrefix: document.getElementById('inp-kojo-prefix').value,
      engineer: document.getElementById('inp-engineer').value,
      version:  document.getElementById('inp-version').value,
      roundingRules: collectRules()
    };
  }

  function collectRules() {
    var rules = [];
    document.querySelectorAll('#rules-body tr').forEach(function(tr) {
      var pattern = tr.querySelector('.rule-pattern').value.trim();
      var roundUp = Math.max(parseInt(tr.querySelector('.rule-roundup').value) || 1, 1);
      if (pattern) rules.push({ pattern: pattern, roundUp: roundUp });
    });
    return rules;
  }

  /* ============================================================
     GROUP SELECTORS
  ============================================================ */
  function loadGroupSelectors(config) {
    var groups = [config.groupL1, config.groupL2, config.groupL3].filter(Boolean);
    var container = document.getElementById('group-selectors');
    container.innerHTML = '';

    if (groups.length === 0) {
      container.innerHTML = '<p class="loading">Nenhum nível de grupo configurado.</p>';
      return;
    }

    groups.forEach(function(groupCol, i) {
      var label = groupCol.replace(/^[A-Z]+ - /, '');
      var selId = 'sel-gval-' + i;
      var div = document.createElement('div');
      div.innerHTML = '<label>' + label + '</label><select id="' + selId + '"><option>Carregando...</option></select>';
      container.appendChild(div);

      google.script.run
        .withSuccessHandler(function(vals) {
          var sel = document.getElementById(selId);
          sel.innerHTML = '<option value="">-- Selecione --</option>';
          vals.forEach(function(v) {
            var opt = document.createElement('option');
            opt.value = v; opt.textContent = v;
            sel.appendChild(opt);
          });
          sel.addEventListener('change', updateComboPreview);
        })
        .withFailureHandler(onError)
        .getRequestUniqueValues(groupCol, config.sourceSheet);
    });
  }

  function updateComboPreview() {
    var config = state.config;
    var groups = [config.groupL1, config.groupL2, config.groupL3].filter(Boolean);
    var vals = groups.map(function(_, i) {
      var sel = document.getElementById('sel-gval-' + i);
      return sel ? sel.value : '';
    });

    if (vals.some(function(v) { return !v; })) {
      document.getElementById('combo-preview').style.display = 'none';
      return;
    }

    google.script.run
      .withSuccessHandler(function(count) {
        var el = document.getElementById('combo-preview');
        el.style.display = 'block';
        el.textContent = '→ ' + count + ' itens encontrados para esta combinação';
      })
      .withFailureHandler(onError)
      .getRequestItemCount(groups, vals, config.sourceSheet);
  }

  /* ============================================================
     KOJO PREVIEW
  ============================================================ */
  function updateKojoPreview() {
    var prefix = document.getElementById('inp-kojo-prefix').value.trim();
    var suffix = document.getElementById('inp-kojo-suffix').value.trim();
    var full = prefix && suffix ? prefix + '.' + suffix : (prefix || suffix || '');
    document.getElementById('kojo-preview').textContent = full ? '→ ' + full : '';
  }

  /* ============================================================
     GENERATE
  ============================================================ */
  function onGenerate() {
    var config = state.config;
    var groups = [config.groupL1, config.groupL2, config.groupL3].filter(Boolean);
    var vals = groups.map(function(_, i) {
      var sel = document.getElementById('sel-gval-' + i);
      return sel ? sel.value : '';
    });

    if (!config.sourceSheet)
      return showStatus('⚠️ Configure e salve a Aba Fonte antes de gerar.', 'err');
    if (groups.length > 0 && vals.some(function(v) { return !v; }))
      return showStatus('⚠️ Selecione um valor para cada nível de grupo.', 'err');
    if (!document.getElementById('inp-kojo-suffix').value.trim())
      return showStatus('⚠️ Informe o BOM KOJO Suffix.', 'err');

    var combination = { parts: vals, label: vals.join(' | ') };
    var settings = Object.assign({}, config, {
      request:       document.getElementById('inp-request').value,
      kojoSuffix:    document.getElementById('inp-kojo-suffix').value.trim(),
      requisitionNum: document.getElementById('inp-req-num').value,
      needBy:        document.getElementById('inp-need-by').value,
      roundingRules: collectRules()
    });

    var btn = document.getElementById('btn-generate');
    btn.disabled = true;
    showStatus('⏳ Gerando request...', 'loading');

    google.script.run
      .withSuccessHandler(function(result) {
        btn.disabled = false;
        if (result.success) {
          showStatus('✅ Request gerado com ' + result.count + ' itens!', 'ok');
        } else {
          showStatus('❌ ' + result.message, 'err');
        }
      })
      .withFailureHandler(function(err) {
        btn.disabled = false;
        onError(err);
      })
      .processRequestCore(combination, settings);
  }

  /* ============================================================
     UTILS
  ============================================================ */
  function showStatus(msg, type) {
    var el = document.getElementById('status');
    el.textContent = msg;
    el.className = type === 'err' ? 'status-err' : type === 'loading' ? 'loading' : 'status-ok';
  }
</script>
</body>
</html>
```

- [ ] **Step 2: Verificar abertura da sidebar**

No Google Sheets, ir em `🔧 Relatórios Dinâmicos` > `📋 Gerador de Request`.
Verificar que a sidebar abre sem erros no console do browser.

- [ ] **Step 3: Commit**

```
git add RequestSidebar.html
git commit -m "feat(request): RequestSidebar.html completo (4 seções + JS)"
```

---

## Task 8: Teste de Integração End-to-End

**Files:** nenhum novo — verificação funcional

- [ ] **Step 1: Configurar via sidebar**
  1. Abrir `📋 Gerador de Request`
  2. Seção Config: selecionar aba fonte, mapear DESC/UPC/UOM/QTY, definir grupo L1 (FLOOR) e L2 (PHASE)
  3. Clicar "💾 Salvar Configuração" — verificar toast "✅ Configuração salva!"
  4. Verificar que dropdowns de grupo aparecem na Seção 3

- [ ] **Step 2: Preencher header e selecionar combinação**
  1. Preencher PROJECT, REQUEST, KOJO PREFIX, Suffix, ENG, VERSION, Requisition #, Need By
  2. Verificar preview do KOJO completo abaixo do Suffix
  3. Selecionar Floor e Phase nos dropdowns — verificar preview "→ N itens encontrados"

- [ ] **Step 3: Gerar e verificar output**
  1. Clicar "📋 GERAR REQUEST"
  2. Verificar status "✅ Request gerado com N itens!"
  3. Na planilha, verificar nova aba com o nome do KOJO Suffix
  4. Conferir:
     - 7 linhas de header com PROJECT, REQUEST, BOM KOJO, ENG, VERSION, LAST UPDATE, GENERATED FROM
     - Linha 1 col E = "Requisition #", col F = "Need By"
     - Linha 3 col E = valor do Requisition #, col F = Need By
     - Linha 9: cabeçalhos `QTY (ROUND UP) | UOM | DESC | UPC | QTY | ROUND UP`
     - Linha 10+: dados ordenados por DESC
     - Col A = fórmula `=ROUNDUP(E10/F10)*F10`
     - Col F = 20 para FOAM CORE pipes, 10 para CPVC pipes, 1 para demais
     - Editar col F manualmente → col A recalcula automaticamente

- [ ] **Step 4: Commit final**

```
git add .
git commit -m "feat(request): Request Generator v1.0 completo e testado"
```

---

## Checklist de Self-Review

- [x] Spec coberta: config própria ✅, sidebar 4 seções ✅, header 7 linhas ✅, Requisition #/Need By ✅, GENERATED FROM ✅, col A fórmula ✅, col F editável ✅, rounding rules tabela ✅, fallback=1 ✅, F=0 protegido via Math.max ✅
- [x] Sem placeholders: todo o código está completo em cada task
- [x] Nomes consistentes: `_req_*` para privados, `processRequestCore`, `getRequestInitData`, `saveRequestConfig` — usados consistentemente em .gs e .html
- [x] `setFormulas()` (batch 2D array) em vez de loop com `setFormula()` ✅
- [x] Todos `getValues()`/`setValues()` fora de loops ✅
