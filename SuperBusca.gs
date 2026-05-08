/**
 * SUPER BUSCA - FERRAMENTAS PLB SHEETS
 * Versão: 2.0 - sort, cache, persistência de estado, config dinâmica
 */

const SUPER_BUSCA_CONFIG = {
  get SOURCE_SHEET() {
    return AppConfig.get('SUPER_BUSCA_SOURCE_SHEET', '5.COST.LIST');
  },
  get DEFAULT_COL_INDEX() {
    return AppConfig.get('SUPER_BUSCA_DEFAULT_COL', 6);
  },
  MENU_TITLE: '🔍 Super Busca',
  CACHE_TTL: 300,
  USER_STATE_KEY: 'SB_USER_STATE_V1'
};

// Cache por aba+coluna para evitar re-leitura da planilha a cada reload
const SBCache = {
  _c: CacheService.getScriptCache(),
  get: (key) => {
    try { const v = SBCache._c.get(key); return v ? JSON.parse(v) : null; } catch(e) { return null; }
  },
  put: (key, value) => {
    try { SBCache._c.put(key, JSON.stringify(value), SUPER_BUSCA_CONFIG.CACHE_TTL); } catch(e) {}
  }
};

// ============================================================================
// FUNÇÕES PRINCIPAIS (PÚBLICAS)
// ============================================================================

/**
 * Abre a sidebar do Super Busca
 * @public
 */
function abrirSuperBuscaSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SuperBuscaSidebar.html')
    .setTitle('Busca de Materiais')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Retorna todos os dados de inicialização do sidebar em uma única chamada:
 * - defaults de config (AppConfig)
 * - lista de abas
 * - último estado salvo pelo usuário (aba + coluna)
 *
 * @public
 * @returns {{ defaultSheet: string, defaultCol: number, sheets: string[], userState: Object|null }}
 */
function getSuperBuscaInitData() {
  const defaultSheet = SUPER_BUSCA_CONFIG.SOURCE_SHEET;
  const defaultCol   = Number(SUPER_BUSCA_CONFIG.DEFAULT_COL_INDEX);
  const sheets       = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());

  let userState = null;
  try {
    const saved = PropertiesService.getUserProperties().getProperty(SUPER_BUSCA_CONFIG.USER_STATE_KEY);
    if (saved) userState = JSON.parse(saved);
  } catch(e) {}

  return { defaultSheet, defaultCol, sheets, userState };
}

/**
 * Persiste a última aba + coluna escolhida pelo usuário
 * Chamada automaticamente ao mudar aba ou coluna no sidebar
 *
 * @public
 * @param {string} sheetName
 * @param {number} colIndex
 */
function saveSuperBuscaState(sheetName, colIndex) {
  try {
    PropertiesService.getUserProperties().setProperty(
      SUPER_BUSCA_CONFIG.USER_STATE_KEY,
      JSON.stringify({ sheetName, colIndex: Number(colIndex) })
    );
    return { success: true };
  } catch(e) {
    return { success: false };
  }
}

/**
 * Retorna lista de nomes de todas as abas da planilha
 * @public
 * @returns {string[]}
 */
function getSheetList() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

/**
 * Retorna lista de colunas com nome e índice para o dropdown
 *
 * @public
 * @param {string} [sheetName]
 * @returns {Array<{name: string, index: number}>}
 */
function getColumnHeaders(sheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName || SUPER_BUSCA_CONFIG.SOURCE_SHEET);
  if (!sheet) return [];
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.map((h, i) => {
    const colIndex   = i + 1;
    const text       = String(h).trim();
    const letter     = indexToLetter(colIndex);
    const displayName = text ? text : `Coluna ${letter}`;
    return { name: displayName, index: colIndex };
  });
}

/**
 * Busca valores únicos de uma coluna, ordenados alfabeticamente (numeric-aware)
 * Usa cache de 5 minutos para evitar re-leitura em cada interação
 *
 * @public
 * @param {string} [sheetName]
 * @param {number} [colIndex]
 * @returns {string[]}
 */
function getDadosParaBusca(sheetName, colIndex) {
  try {
    const targetSheetName = sheetName || SUPER_BUSCA_CONFIG.SOURCE_SHEET;
    const targetCol       = SharedUtils_toPositiveInteger(colIndex, SUPER_BUSCA_CONFIG.DEFAULT_COL_INDEX);

    const cacheKey = `sb_${targetSheetName}_${targetCol}`;
    const cached   = SBCache.get(cacheKey);
    if (cached) return cached;

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) {
      console.error(`[SuperBusca] Aba "${targetSheetName}" nao encontrada`);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values    = sheet.getRange(2, targetCol, lastRow - 1, 1).getDisplayValues();
    const listaLimpa = [...new Set(values.flat().filter(item => item !== ''))];

    listaLimpa.sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

    SBCache.put(cacheKey, listaLimpa);
    return listaLimpa;

  } catch (error) {
    console.error(`[SuperBusca] Erro em getDadosParaBusca: ${error.message}`);
    return [];
  }
}

/**
 * Insere itens selecionados na célula ativa, verticalmente
 * Retorna resposta padronizada — sidebar DEVE verificar result.success
 *
 * @public
 * @param {string[]} itens
 * @returns {{ success: boolean, data?: { inserted: number }, error?: string }}
 */
function inserirItensSelecionados(itens) {
  try {
    if (!itens || itens.length === 0) {
      return SharedUtils_errorResponse('Nenhum item selecionado');
    }
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const sheet      = ss.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    const row        = activeCell.getRow();
    const col        = activeCell.getColumn();

    sheet.getRange(row, col, itens.length, 1).setValues(itens.map(item => [item]));
    sheet.getRange(row + itens.length, col).activate();

    return SharedUtils_successResponse({ inserted: itens.length });
  } catch (error) {
    console.error(`[SuperBusca] Erro em inserirItensSelecionados: ${error.message}`);
    return SharedUtils_errorResponse(error);
  }
}

/**
 * Auxiliar: Converte índice numérico em letra de coluna (1→A, 27→AA)
 */
function indexToLetter(column) {
  return SharedUtils_numberToColumnLetter(column);
}
