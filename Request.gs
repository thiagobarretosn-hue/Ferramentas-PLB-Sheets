/**
 * @fileoverview Request Generator — Gerador de Requisição de Materiais KOJO
 * @version 1.2.0 - Autosave de config (sem botão Salvar); aba gerada com prefixo
 *                  "REQUEST-" e cor de aba amarela (#f1c232); Requisition #/Need By
 *                  saíram da sidebar (preenchidos depois, direto na aba);
 *                  config via factory compartilhado (lib/Shared/Config.gs)
 * @version 1.1.0
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
  SHEET_PREFIX: 'REQUEST-',
  TAB_COLOR: '#f1c232',
  COLORS: {
    HEADER_BG: '#2c3e50',
    FONT_LIGHT: '#ffffff'
  },
  // Regex patterns: first-match-wins, case-insensitive
  DEFAULT_RULES: [
    { pattern: '^PIPE.*FOAM CORE', roundUp: 20 },
    { pattern: '^PIPE.*CPVC',      roundUp: 10 }
  ],
  // Keywords used for auto-matching column headers
  AUTO_MATCH: [
    { field: 'colDesc',  keywords: ['DESC', 'DESCRIPTION', 'MATERIAL'] },
    { field: 'colUpc',   keywords: ['UPC', 'ITEM CODE', 'PART NUMBER', 'PART'] },
    { field: 'colUom',   keywords: ['UOM', 'UNIT', 'UM'] },
    { field: 'colQty',   keywords: ['QTY', 'QUANTITY'] },
    { field: 'groupL1',  keywords: ['FLOOR', 'LEVEL', 'ANDAR'] },
    { field: 'groupL2',  keywords: ['PHASE', 'FASE', 'SYSTEM'] },
    { field: 'groupL3',  keywords: ['ZONE', 'AREA', 'SECTION'] }
  ],
  COL_WIDTHS: { A: 120, B: 80, C: 550, D: 100, E: 100, F: 100 }
};

// ============================================================================
// CONFIG SERVICE
// ============================================================================
// V1.2: storage delegado ao factory compartilhado (lib/Shared/Config.gs).
// Só a migração de regras v1.0→v1.1 permanece aqui (específica do Request).
const RequestConfigService = {
  _svc: null,
  _store: function() {
    if (!this._svc) {
      this._svc = SharedConfig_createDocConfigService(
        REQUEST_CONFIG.SETTINGS_KEY,
        RequestConfigService._defaults
      );
    }
    return this._svc;
  },

  _defaults: function() {
    return {
      sourceSheet: '',
      colDesc: '', colUpc: '', colUom: '', colQty: '',
      groupL1: '', groupL2: '', groupL3: '',
      project: '', kojoPrefix: '', engineer: '', version: '01',
      // Header fields persisted across sessions
      // (requisitionNum/needBy: legado — mantidos p/ compat com configs salvas,
      //  não são mais editados na sidebar)
      request: '', kojoSuffix: '', requisitionNum: '', needBy: '',
      // Last combination used (restored in Section 3 dropdowns)
      lastGroupVals: [],
      roundingRules: REQUEST_CONFIG.DEFAULT_RULES.map(function(r) {
        return { pattern: r.pattern, roundUp: r.roundUp };
      })
    };
  },

  get: function() {
    const config = this._store().getAll();
    // Migrate v1.0 plain-text rules → v1.1 regex rules
    config.roundingRules = (config.roundingRules || []).map(function(r) {
      if (r.pattern === 'FOAM CORE') return { pattern: '^PIPE.*FOAM CORE', roundUp: r.roundUp };
      if (r.pattern === 'CPVC')      return { pattern: '^PIPE.*CPVC',      roundUp: r.roundUp };
      return r;
    });
    return config;
  },

  save: function(config) {
    return this._store().saveAll(config);
  }
};

// ============================================================================
// COLUMN HELPERS
// ============================================================================

// Wrappers finos sobre lib/Shared/Utils.gs — mantidos para não tocar os call sites.

/** "J - DESC" → 10 (1-indexed). Supports AA, AB, etc. Exige o formato "X - ..." */
function _req_getColumnIndex(colConfig) {
  if (!colConfig) return -1;
  const match = String(colConfig).match(/^([A-Za-z]+)\s*-/);
  if (!match) return -1;
  return SharedUtils_columnLetterToIndex(match[1]);
}

/** 1 → "A", 26 → "Z", 27 → "AA" */
function _req_numberToLetter(n) {
  return SharedUtils_numberToColumnLetter(n);
}

/** Returns ["A - Header", "B - Header", ...] for every column in the sheet */
function _req_getColumnsFromSheet(sheet) {
  return SharedUtils_getColumnLabelsFromSheet(sheet, 'Col');
}

/**
 * Scans column labels and returns the best-matching column per field.
 * Used to auto-fill dropdowns without a round-trip per column.
 */
function _req_autoMatch(columns) {
  const result = {};
  REQUEST_CONFIG.AUTO_MATCH.forEach(function(m) {
    const found = columns.find(function(col) {
      const upper = col.toUpperCase();
      return m.keywords.some(function(kw) { return upper.includes(kw); });
    });
    result[m.field] = found || '';
  });
  return result;
}

/**
 * Applies the first matching rounding rule (regex, case-insensitive).
 * Falls back to plain includes() if the pattern is not a valid regex.
 * Always returns >= 1 to prevent division-by-zero in the formula.
 */
function _req_applyRoundingRule(desc, rules) {
  if (!rules || rules.length === 0) return 1;
  const match = rules.find(function(r) {
    if (!r.pattern) return false;
    try {
      return new RegExp(r.pattern, 'i').test(String(desc));
    } catch (e) {
      return String(desc).toUpperCase().includes(r.pattern.toUpperCase());
    }
  });
  return Math.max(match ? (Number(match.roundUp) || 1) : 1, 1);
}

// ============================================================================
// DATA QUERIES — called by HTML via google.script.run
// ============================================================================

/**
 * Initial data for the sidebar.
 * Returns sheets, columns for the configured source, saved config, and
 * auto-match suggestions (only when column fields are not yet configured).
 */
function getRequestInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = RequestConfigService.get();
  const allSheets = ss.getSheets().map(function(s) { return s.getName(); });

  let allColumns = [];
  let autoMatch = {};
  if (config.sourceSheet) {
    const sourceSheet = ss.getSheetByName(config.sourceSheet);
    if (sourceSheet) {
      allColumns = _req_getColumnsFromSheet(sourceSheet);
      if (!config.colDesc && !config.colUpc) {
        autoMatch = _req_autoMatch(allColumns);
      }
    }
  }

  // ENG: se vazio, herda do BOM config (campo 'Engenheiro')
  if (!config.engineer) {
    try {
      const bomSaved = PropertiesService.getDocumentProperties().getProperty('BOM_SETTINGS_V3');
      if (bomSaved) config.engineer = (JSON.parse(bomSaved)['Engenheiro'] || '');
    } catch(e) {}
  }

  return { allSheets: allSheets, allColumns: allColumns, config: config, autoMatch: autoMatch };
}

/**
 * Returns columns for a sheet + auto-match suggestions.
 * Called when the user switches the source sheet.
 */
function getRequestSheetColumns(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { columns: [], autoMatch: {} };
  const columns = _req_getColumnsFromSheet(sheet);
  return { columns: columns, autoMatch: _req_autoMatch(columns) };
}

/** Unique sorted values for a column — populates group combination dropdowns. */
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
 * Counts source rows matching the selected combination.
 * groupCols: ["I - FLOOR", "H - PHASE"], groupVals: ["6th", "Job Site"]
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
      const rowVal = String(row[colIdx - 1] || '').trim();
      const accepted = Array.isArray(groupVals[i]) ? groupVals[i] : [groupVals[i]];
      return accepted.some(function(v) { return String(v || '').trim() === rowVal; });
    });
    if (matches) count++;
  });
  return count;
}

/** Persists the full config (column mapping + header fields + last group vals). */
function saveRequestConfig(config) {
  return RequestConfigService.save(config);
}

// ============================================================================
// OUTPUT WRITERS (private)
// ============================================================================

function _req_writeHeader(sheet, settings, combination, groupCols) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const lastUpdate = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');
  const kojoFull = settings.kojoSuffix || '';

  const generatedFrom = groupCols
    .filter(Boolean)
    .map(function(col, i) {
      const label = col.replace(/^[A-Z]+ - /, '');
      const part  = combination.parts[i];
      const val   = Array.isArray(part) ? part.join(', ') : (part || '');
      return label + ': ' + val;
    })
    .join(' | ');

  sheet.getRange(1, 1, 7, 6).setNumberFormat('@STRING@');

  const labels = [['PROJECT:'], ['REQUEST:'], ['BOM KOJO:'], ['ENG.:'], ['VERSION:'], ['LAST UPDATE:'], ['GENERATED FROM:']];
  sheet.getRange(1, 1, 7, 1).setValues(labels);

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

  sheet.getRange(1, 5).setValue('Requisition #');
  sheet.getRange(1, 6).setValue('Need By');
  sheet.getRange(3, 5).setValue(settings.requisitionNum || '');
  sheet.getRange(3, 6).setValue(settings.needBy || '');

  for (let r = 1; r <= 7; r++) {
    sheet.getRange(r, 2, 1, 3).merge();
  }
  sheet.getRange(1, 1, 7, 6).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function _req_writeData(sheet, rows, dataStartRow) {
  if (rows.length === 0) return;

  const values = rows.map(function(r) {
    return [r.uom, r.desc, r.upc, r.qty, r.roundUp];
  });
  sheet.getRange(dataStartRow, 2, rows.length, 5).setValues(values);

  // Batch formula insert — faster than per-row setFormula()
  const formulas = rows.map(function(_, i) {
    const rowNum = dataStartRow + i;
    return ['=ROUNDUP(E' + rowNum + '/F' + rowNum + ')*F' + rowNum];
  });
  sheet.getRange(dataStartRow, 1, rows.length, 1).setFormulas(formulas);
}

function _req_formatSheet(sheet, dataRowCount) {
  const colHeaderRow = REQUEST_CONFIG.COL_HEADER_ROW;
  sheet.getRange(colHeaderRow, 1, dataRowCount + 1, 6)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

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
 * Processes source data and writes the KOJO request sheet.
 * @param {{parts: string[], label: string}} combination
 * @param {Object} settings - full config + per-request header fields
 * @returns {{success: boolean, count?: number, message?: string}}
 */
function processRequestCore(combination, settings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(settings.sourceSheet);
    if (!sourceSheet) {
      return { success: false, message: 'Aba "' + settings.sourceSheet + '" não encontrada.' };
    }

    const descIdx = _req_getColumnIndex(settings.colDesc) - 1;
    const upcIdx  = _req_getColumnIndex(settings.colUpc)  - 1;
    const uomIdx  = _req_getColumnIndex(settings.colUom)  - 1;
    const qtyIdx  = _req_getColumnIndex(settings.colQty)  - 1;

    if ([descIdx, upcIdx, uomIdx, qtyIdx].some(function(i) { return i < 0; })) {
      return { success: false, message: 'Configuração de colunas inválida. Verifique a seção Configuração.' };
    }

    const groupCols    = [settings.groupL1, settings.groupL2, settings.groupL3].filter(Boolean);
    const groupVals    = combination.parts;
    const groupIndices = groupCols.map(_req_getColumnIndex);

    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Aba fonte está vazia.' };

    // 1. Batch read
    const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();

    // 2. Filter + group by DESC+UPC+UOM
    const grouped = {};
    data.forEach(function(row) {
      const matches = groupIndices.every(function(colIdx, i) {
        const rowVal = String(row[colIdx - 1] || '').trim();
        const accepted = Array.isArray(groupVals[i]) ? groupVals[i] : [groupVals[i]];
        return accepted.some(function(v) { return String(v || '').trim() === rowVal; });
      });
      if (!matches) return;

      const desc = String(row[descIdx] || '').trim();
      if (!desc) return;

      const upc = String(row[upcIdx] || '').trim();
      const uom = String(row[uomIdx] || '').trim();
      const qty = Number(row[qtyIdx]) || 0;
      const key = desc + '|||' + upc + '|||' + uom;

      if (grouped[key]) {
        grouped[key].qty += qty;
      } else {
        grouped[key] = { desc: desc, upc: upc, uom: uom, qty: qty };
      }
    });

    // 3. Sort by DESC
    const rows = Object.values(grouped).sort(function(a, b) {
      return a.desc.localeCompare(b.desc, undefined, { numeric: true });
    });

    if (rows.length === 0) {
      return { success: false, message: 'Nenhum item encontrado para a combinação selecionada.' };
    }

    // 4. Apply rounding rules (regex, first-match-wins)
    const rules = settings.roundingRules || REQUEST_CONFIG.DEFAULT_RULES;
    rows.forEach(function(row) {
      row.roundUp = _req_applyRoundingRule(row.desc, rules);
    });

    // 5. Create or clear target sheet — nome sempre com prefixo "REQUEST-"
    const baseName = String(settings.kojoSuffix || 'Request').trim();
    const prefixed = new RegExp('^' + REQUEST_CONFIG.SHEET_PREFIX, 'i').test(baseName)
      ? baseName
      : REQUEST_CONFIG.SHEET_PREFIX + baseName;
    const sheetName = SharedUtils_sanitizeSheetName(prefixed);
    let targetSheet = ss.getSheetByName(sheetName);
    if (targetSheet) {
      targetSheet.clear();
    } else {
      targetSheet = ss.insertSheet(sheetName);
    }
    targetSheet.setTabColor(REQUEST_CONFIG.TAB_COLOR);

    // 6. Write header (rows 1-7)
    _req_writeHeader(targetSheet, settings, combination, groupCols);

    // 7. Column header row (row 9)
    const colHeaderRow = REQUEST_CONFIG.COL_HEADER_ROW;
    targetSheet.getRange(colHeaderRow, 1, 1, 6)
      .setValues([['QTY (ROUND UP)', 'UOM', 'DESC', 'UPC', 'QTY', 'ROUND UP']]);
    targetSheet.getRange(colHeaderRow, 1, 1, 6).setFontWeight('bold');

    // 8. Data rows (row 10+)
    _req_writeData(targetSheet, rows, REQUEST_CONFIG.DATA_START_ROW);

    // 9. Format
    _req_formatSheet(targetSheet, rows.length);

    // 10. Protect header
    try {
      const prot = targetSheet.getRange(1, 1, REQUEST_CONFIG.HEADER_ROWS, 6).protect();
      prot.setDescription('Header protegido').removeEditors(prot.getEditors());
    } catch (e) { /* no permission in some contexts */ }

    return { success: true, count: rows.length };

  } catch (e) {
    console.error('[processRequestCore]', e.message, e.stack);
    return { success: false, message: e.message };
  }
}

// ============================================================================
// ENTRY POINTS
// ============================================================================

function openRequestSidebar() {
  // Template (não HtmlOutput direto): necessário para o include() de SharedStyles/SharedScripts
  const html = HtmlService.createTemplateFromFile('RequestSidebar')
    .evaluate()
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
  results.push('AA → ' + _req_getColumnIndex('AA - EXTRA') + ' (expect 27)');
  results.push('vazio → ' + _req_getColumnIndex('') + ' (expect -1)');

  results.push('1 → ' + _req_numberToLetter(1) + ' (expect A)');
  results.push('27 → ' + _req_numberToLetter(27) + ' (expect AA)');

  // Regex rules: only PIPE items match
  results.push('PIPE FOAM CORE → ' + _req_applyRoundingRule('PIPE 2 IN FOAM CORE', rules) + ' (expect 20)');
  results.push('PIPE CPVC → ' + _req_applyRoundingRule('PIPE 1 IN. X 10 FT. CPVC', rules) + ' (expect 10)');
  results.push('ELBOW CPVC → ' + _req_applyRoundingRule('ELBOW 90 CPVC', rules) + ' (expect 1 — not a pipe)');
  results.push('TEE FOAM CORE → ' + _req_applyRoundingRule('TEE 2 IN FOAM CORE', rules) + ' (expect 1 — not a pipe)');
  results.push('roundUp=0 → ' + _req_applyRoundingRule('PIPE FOAM CORE', [{pattern:'^PIPE.*FOAM CORE', roundUp:0}]) + ' (expect 1)');

  const defaults = RequestConfigService.get();
  results.push('defaults.roundingRules[0].pattern: ' + defaults.roundingRules[0].pattern);
  results.push('defaults.version: ' + defaults.version + ' (expect 01)');

  console.log(results.join('\n'));
}
