/**
 * @fileoverview Summary All — consolida e agrupa abas selecionadas
 * Config em PropertiesService. Suporta destino novo/existente e seleção de abas.
 * @version 3.0.0
 */

const SUMMARY_ALL_CFG = {
  PROPS_KEY:  'SUMMARY_ALL_CONFIG_V3',
  SKIP:       ['Claude Log'],
  HDR_COLOR:  '#1F4E79',
  HDR_FONT:   '#FFFFFF',
};

// ─── MENU ─────────────────────────────────────────────────────────────────
function openSummaryAllSidebar() {
  const data = _getSidebarData();
  const tpl  = HtmlService.createTemplateFromFile('SummaryAllSidebar');
  tpl.data   = JSON.stringify(data);
  const sidebar = tpl.evaluate().setTitle('📊 Summary All').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(sidebar);
}

// ─── APIs CHAMADAS PELO SIDEBAR ────────────────────────────────────────────

/** Retorna headers da primeira aba selecionada (chamado quando seleção muda) */
function getHeadersForFirstSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheetName) return { success: true, headers: [] };
    const s = ss.getSheetByName(sheetName);
    if (!s || s.getLastRow() < 1) return { success: true, headers: [] };
    const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String);
    return { success: true, headers };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/** Salva config completa (DocumentProperties: por planilha, não vaza entre docs/usuários) */
function saveSummaryAllConfig(fullConfig) {
  try {
    PropertiesService.getDocumentProperties()
      .setProperty(SUMMARY_ALL_CFG.PROPS_KEY, JSON.stringify(fullConfig));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/** Salva e executa */
function runSummaryAllFromSidebar(fullConfig) {
  try {
    saveSummaryAllConfig(fullConfig);
    const result = _runSummary(fullConfig);
    return { success: true, grupos: result.grupos, linhas: result.linhas, dest: result.dest };
  } catch(e) {
    console.error('[SummaryAll] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ─── ENTRY POINT DIRETO ───────────────────────────────────────────────────
function summarizeAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Gerando summary...', 'Summary All', 3);
  try {
    const cfg = _loadFullConfig();
    const r   = _runSummary(cfg);
    ss.toast(r.grupos + ' grupos | ' + r.linhas + ' linhas → ' + r.dest + ' ✅', 'Summary All', 6);
  } catch(e) {
    ss.toast('Erro: ' + e.message, 'Summary All', 8);
  }
}

// ─── CORE ─────────────────────────────────────────────────────────────────
function _runSummary(fullConfig) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const columns = fullConfig.columns || [];
  const config  = _parseColumnConfig(columns);

  if (config.groupCols.length === 0) throw new Error('Nenhuma coluna marcada como "Agrupar".');

  const selectedSheets = (fullConfig.selectedSheets || []).filter(Boolean);
  if (selectedSheets.length === 0) throw new Error('Selecione ao menos uma aba de origem.');

  const { headers, allRows } = _readSourceSheets(ss, selectedSheets);
  if (allRows.length === 0) throw new Error('Nenhuma linha de dados encontrada nas abas selecionadas.');

  const grouped = _groupAndAggregate(allRows, config.groupCols, config.sumCols, config.countCols);
  const sorted  = _sortRows(grouped, config.sortKeys);
  const dest    = _writeOutput(ss, headers, sorted, config.countCols, fullConfig, selectedSheets);

  return { grupos: sorted.length, linhas: allRows.length, dest };
}

// ─── SIDEBAR DATA ──────────────────────────────────────────────────────────
function _getSidebarData() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = _listSourceSheets(ss);
  const saved     = _loadFullConfig();

  // Defaults se não há config
  const selectedSheets = saved.selectedSheets && saved.selectedSheets.length > 0
    ? saved.selectedSheets
    : allSheets.map(s => s.name);

  // Headers da primeira aba selecionada
  const firstName  = selectedSheets.find(n => allSheets.some(s => s.name === n));
  const headers    = firstName ? _getSheetHeaders(ss, firstName) : [];
  const n          = Math.max(11, headers.length);

  // Expand colunas se necessário
  const columns = _buildDefaultConfig(n, saved.columns);

  return {
    headers,
    allSheets,
    config: {
      columns,
      selectedSheets,
      destMode:   saved.destMode   || 'new',
      destSuffix: saved.destSuffix !== undefined ? saved.destSuffix : '_SUMMARY',
      destSheet:  saved.destSheet  || '',
      destCell:   saved.destCell   || 'A1',
    }
  };
}

function _listSourceSheets(ss) {
  return ss.getSheets()
    .filter(s => SUMMARY_ALL_CFG.SKIP.indexOf(s.getName()) < 0)
    .map(s => ({
      name:    s.getName(),
      rows:    Math.max(0, s.getLastRow() - 1),
    }));
}

function _getSheetHeaders(ss, sheetName) {
  const s = ss.getSheetByName(sheetName);
  if (!s || s.getLastRow() < 1) return [];
  return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String);
}

// ─── CONFIG ───────────────────────────────────────────────────────────────
function _loadFullConfig() {
  // DocumentProperties é o local atual; ScriptProperties é fallback de migração
  // (config salva antes de 07/2026 era compartilhada entre todas as planilhas do script)
  const raw = PropertiesService.getDocumentProperties().getProperty(SUMMARY_ALL_CFG.PROPS_KEY)
    || PropertiesService.getScriptProperties().getProperty(SUMMARY_ALL_CFG.PROPS_KEY);
  if (!raw) return {};

  const parsed = JSON.parse(raw);
  // Migração do formato antigo (array de colunas)
  if (Array.isArray(parsed)) return { columns: parsed };
  return parsed;
}

function _parseColumnConfig(columns) {
  const groupCols = [], sumCols = [], countCols = [], sortKeys = [];

  columns.forEach(c => {
    const ci   = parseInt(c.ci);
    const acao = String(c.acao || '').trim().toUpperCase();
    const ord  = parseInt(c.ordem);
    const dir  = String(c.dir  || '').trim().toUpperCase();

    if      (acao === 'SOMAR VALORES')    sumCols.push(ci);
    else if (acao === 'CONTAR REGISTROS') countCols.push(ci);
    else if (acao === 'AGRUPAR')          groupCols.push(ci);

    if (!isNaN(ord) && ord > 0) {
      sortKeys.push({ colIndex: ci, ascending: dir !== 'DESC', priority: ord });
    }
  });

  sortKeys.sort((a, b) => a.priority - b.priority);
  return { groupCols, sumCols, countCols, sortKeys };
}

function _buildDefaultConfig(n, existing) {
  const DEFAULTS = [
    { acao: 'Copiar 1º registro', ordem: '', dir: 'ASC' },
    { acao: 'Agrupar',            ordem: '1', dir: 'ASC' },
    { acao: 'Copiar 1º registro', ordem: '', dir: 'ASC' },
    { acao: 'Agrupar',            ordem: '2', dir: 'ASC' },
    { acao: 'Agrupar',            ordem: '4', dir: 'ASC' },
    { acao: 'Agrupar',            ordem: '5', dir: 'ASC' },
    { acao: 'Copiar 1º registro', ordem: '', dir: 'ASC' },
    { acao: 'Agrupar',            ordem: '3', dir: 'ASC' },
    { acao: 'Copiar 1º registro', ordem: '', dir: 'ASC' },
    { acao: 'Agrupar',            ordem: '6', dir: 'ASC' },
    { acao: 'Somar valores',      ordem: '', dir: 'ASC' },
  ];

  const result = [];
  for (let i = 0; i < n; i++) {
    const saved = existing && existing[i];
    const def   = DEFAULTS[i] || { acao: 'Copiar 1º registro', ordem: '', dir: 'ASC' };
    result.push({
      ci:    i,
      acao:  saved ? saved.acao  : def.acao,
      ordem: saved ? saved.ordem : def.ordem,
      dir:   saved ? saved.dir   : def.dir,
    });
  }
  return result;
}

// ─── LER ABAS FONTE ────────────────────────────────────────────────────────
function _readSourceSheets(ss, selectedSheets) {
  let headers  = [];
  const allRows = [];

  // Processar na ordem da seleção (primeira = cabeçalho)
  selectedSheets.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const lastRow = s.getLastRow();
    if (lastRow < 2) return;

    const vals = s.getRange(1, 1, lastRow, s.getLastColumn()).getValues();
    if (headers.length === 0) headers = vals[0];

    for (let r = 1; r < vals.length; r++) {
      if (vals[r].some(c => c !== '' && c !== null && c !== undefined)) {
        allRows.push(vals[r]);
      }
    }
  });

  return { headers, allRows };
}

// ─── AGRUPAR E AGREGAR ─────────────────────────────────────────────────────
function _groupAndAggregate(allRows, groupCols, sumCols, countCols) {
  const groupMap = new Map();

  allRows.forEach(row => {
    let key = '';
    groupCols.forEach(ci => {
      key += String(row[ci] != null ? row[ci] : '') + '||';
    });

    if (!groupMap.has(key)) {
      const stored = row.slice();
      countCols.forEach(ci => { stored[ci] = 1; });
      groupMap.set(key, stored);
    } else {
      const stored = groupMap.get(key);
      sumCols.forEach(ci => {
        stored[ci] = (Number(stored[ci]) || 0) + (Number(row[ci]) || 0);
      });
      countCols.forEach(ci => {
        stored[ci] = (Number(stored[ci]) || 0) + 1;
      });
    }
  });

  const result = [];
  groupMap.forEach(row => result.push(row));
  return result;
}

// ─── ORDENAR ───────────────────────────────────────────────────────────────
function _sortRows(rows, sortKeys) {
  if (sortKeys.length === 0) return rows;

  return rows.slice().sort((a, b) => {
    for (let k = 0; k < sortKeys.length; k++) {
      const sk  = sortKeys[k];
      const va  = String(a[sk.colIndex] != null ? a[sk.colIndex] : '');
      const vb  = String(b[sk.colIndex] != null ? b[sk.colIndex] : '');
      const na  = Number(va), nb = Number(vb);
      const cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : va.localeCompare(vb, 'pt-BR');
      if (cmp !== 0) return sk.ascending ? cmp : -cmp;
    }
    return 0;
  });
}

// ─── ESCREVER SAÍDA ────────────────────────────────────────────────────────
function _writeOutput(ss, headers, resultRows, countCols, fullConfig, selectedSheets) {
  let out, startRow = 1, startCol = 1, destName;

  if (fullConfig.destMode === 'existing') {
    destName = fullConfig.destSheet;
    out = ss.getSheetByName(destName);
    if (!out) throw new Error('Aba de destino "' + destName + '" não encontrada.');

    const parsed = _parseCellRef(fullConfig.destCell || 'A1');
    startRow = parsed.row;
    startCol = parsed.col;

    // Limpar região a partir da célula de destino
    const clearRows = resultRows.length + 1;
    const clearCols = headers.length;
    out.getRange(startRow, startCol, clearRows, clearCols).clearContent().clearFormat();

  } else {
    // Nova aba: nome = primeira aba selecionada + sufixo
    const firstName = selectedSheets[0] || 'Summary';
    const suffix    = fullConfig.destSuffix !== undefined ? fullConfig.destSuffix : '_SUMMARY';
    destName = firstName + suffix;

    out = ss.getSheetByName(destName);
    if (out) { out.clearContents(); out.clearFormats(); }
    else      { out = ss.insertSheet(destName); }
  }

  const enrichedHeaders = headers.map((h, i) =>
    countCols.indexOf(i) >= 0 ? '# ' + String(h) : h
  );

  const nC = enrichedHeaders.length;
  const outputRows = [enrichedHeaders].concat(
    resultRows.map(row => {
      const uniform = [];
      for (let c = 0; c < nC; c++) uniform.push(row[c] != null ? row[c] : '');
      return uniform;
    })
  );

  out.getRange(startRow, startCol, outputRows.length, nC).setValues(outputRows);

  out.getRange(startRow, startCol, 1, nC)
    .setBackground(SUMMARY_ALL_CFG.HDR_COLOR)
    .setFontColor(SUMMARY_ALL_CFG.HDR_FONT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (countCols.length > 0) {
    countCols.forEach(ci => out.getRange(startRow, startCol + ci, 1, 1).setBackground('#2E7D32'));
  }

  for (let c = startCol; c < startCol + nC; c++) out.autoResizeColumn(c);
  out.activate();

  return destName;
}

// ─── HELPERS ──────────────────────────────────────────────────────────────
function _parseCellRef(ref) {
  const m = String(ref).toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return { row: 1, col: 1 };
  const col = m[1].split('').reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
  return { row: parseInt(m[2]), col };
}
