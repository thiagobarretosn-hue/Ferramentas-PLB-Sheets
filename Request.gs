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

  // Col E e F: Requisition # e Need By
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

// ============================================================================
// TESTES MANUAIS — executar no Script Editor, verificar Execution Log
// ============================================================================
function testRequestHelpers() {
  const rules = REQUEST_CONFIG.DEFAULT_RULES;
  const results = [];

  results.push('J → ' + _req_getColumnIndex('J - DESC') + ' (expect 10)');
  results.push('A → ' + _req_getColumnIndex('A - ID') + ' (expect 1)');
  results.push('AA → ' + _req_getColumnIndex('AA - EXTRA') + ' (expect 27)');
  results.push('vazio → ' + _req_getColumnIndex('') + ' (expect -1)');

  results.push('1 → ' + _req_numberToLetter(1) + ' (expect A)');
  results.push('26 → ' + _req_numberToLetter(26) + ' (expect Z)');
  results.push('27 → ' + _req_numberToLetter(27) + ' (expect AA)');

  results.push('FOAM CORE → ' + _req_applyRoundingRule('PIPE 2 IN FOAM CORE', rules) + ' (expect 20)');
  results.push('CPVC → ' + _req_applyRoundingRule('PIPE 1 IN. X 10 FT. CPVC', rules) + ' (expect 10)');
  results.push('ELBOW → ' + _req_applyRoundingRule('ELBOW 90 DEGREES 2 IN.', rules) + ' (expect 1)');
  results.push('roundUp=0 → ' + _req_applyRoundingRule('PIPE FOAM CORE', [{pattern:'FOAM CORE', roundUp:0}]) + ' (expect 1)');

  const defaults = RequestConfigService.get();
  results.push('defaults.roundingRules.length: ' + defaults.roundingRules.length + ' (expect 2)');
  results.push('defaults.version: ' + defaults.version + ' (expect 01)');

  console.log(results.join('\n'));
}
