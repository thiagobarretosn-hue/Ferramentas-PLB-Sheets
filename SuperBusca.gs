/**
 * SUPER BUSCA - FERRAMENTAS PLB SHEETS
 * Versão: Dropdown Híbrido (Nome ou Letra da Coluna)
 */

/**
 * Configuracao do Super Busca
 * Valores dinamicos obtidos de AppConfig (lib/Shared/Config.gs)
 * Use showConfigDialog() para alterar via interface
 */
const SUPER_BUSCA_CONFIG = {
  get SOURCE_SHEET() {
    return AppConfig.get('SUPER_BUSCA_SOURCE_SHEET', '5.COST.LIST');
  },
  get DEFAULT_COL_INDEX() {
    return AppConfig.get('SUPER_BUSCA_DEFAULT_COL', 6);
  },
  MENU_TITLE: '🔍 Super Busca'
};

// ============================================================================
// FUNÇÕES PRINCIPAIS (PUBLICAS)
// Estas funções são chamadas pelo menu ou pela sidebar HTML
// ============================================================================

/**
 * Abre a sidebar do Super Busca
 * Chamada pelo menu: '🔍 Super Busca' > '🚀 Abrir Painel'
 *
 * @public
 * @returns {void}
 */
function abrirSuperBuscaSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SuperBuscaSidebar.html')
    .setTitle('Busca de Materiais')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Retorna lista de nomes de todas as abas da planilha
 * Chamada pela sidebar para popular dropdown de abas
 *
 * @public
 * @returns {string[]} Array com nomes das abas
 */
function getSheetList() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

/**
 * Retorna lista de colunas com nome/indice para dropdown
 * Se a coluna tiver cabeçalho na linha 1, mostra o nome
 * Caso contrário, mostra "Coluna X" (letra da coluna)
 *
 * @public
 * @param {string} [sheetName] - Nome da aba (usa SOURCE_SHEET se omitido)
 * @returns {Array<{name: string, index: number}>} Lista de colunas
 *
 * @example
 * getColumnHeaders('MinhaPlanilha')
 * // [{ name: 'Descrição', index: 1 }, { name: 'Coluna B', index: 2 }]
 */
function getColumnHeaders(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName || SUPER_BUSCA_CONFIG.SOURCE_SHEET);
  if (!sheet) return [];

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];

  // Pega valores da linha 1
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Mapeia TODAS as colunas
  const list = headers.map((h, i) => {
    const colIndex = i + 1;
    const text = String(h).trim();
    const letter = indexToLetter(colIndex);
    
    // Lógica: Se tem texto, usa o texto. Se vazio, usa "Coluna [Letra]"
    const displayName = text ? text : `Coluna ${letter}`;
    
    return { name: displayName, index: colIndex };
  });

  return list;
}

/**
 * Busca valores únicos de uma coluna para popular lista de busca
 * Ignora a linha 1 (cabeçalho) e remove valores vazios e duplicados
 *
 * @public
 * @param {string} [sheetName] - Nome da aba (usa SOURCE_SHEET se omitido)
 * @param {number} [colIndex] - Indice da coluna 1-indexed (usa DEFAULT_COL_INDEX se omitido)
 * @returns {string[]} Lista de valores únicos ordenados ou array vazio em caso de erro
 *
 * @example
 * getDadosParaBusca('Materiais', 3)
 * // ['Item A', 'Item B', 'Item C']
 */
function getDadosParaBusca(sheetName, colIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const targetSheetName = sheetName || SUPER_BUSCA_CONFIG.SOURCE_SHEET;
    const targetCol = SharedUtils_toPositiveInteger(colIndex, SUPER_BUSCA_CONFIG.DEFAULT_COL_INDEX);

    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) {
      console.error(`[SuperBusca] Aba "${targetSheetName}" nao encontrada`);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(2, targetCol, lastRow - 1, 1).getDisplayValues();

    const listaLimpa = values.flat().filter(item => item !== "");
    return [...new Set(listaLimpa)];

  } catch (error) {
    console.error(`[SuperBusca] Erro em getDadosParaBusca: ${error.message}`);
    return [];
  }
}

/**
 * Insere itens selecionados na célula ativa da planilha
 * Os itens são inseridos verticalmente, um por linha, a partir da célula ativa
 * Após inserção, a célula ativa é movida para a próxima linha disponível
 *
 * @public
 * @param {string[]} itens - Lista de itens a inserir
 * @returns {Object} Resposta padronizada { success: boolean, data?: { inserted: number }, error?: string }
 *
 * @example
 * // Na sidebar:
 * google.script.run.inserirItensSelecionados(['Item 1', 'Item 2', 'Item 3']);
 */
function inserirItensSelecionados(itens) {
  try {
    if (!itens || itens.length === 0) {
      return SharedUtils_errorResponse('Nenhum item selecionado');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const activeCell = sheet.getActiveCell();

    const row = activeCell.getRow();
    const col = activeCell.getColumn();

    const dadosParaInserir = itens.map(item => [item]);
    sheet.getRange(row, col, dadosParaInserir.length, 1).setValues(dadosParaInserir);

    const nextRow = row + dadosParaInserir.length;
    sheet.getRange(nextRow, col).activate();

    return SharedUtils_successResponse({ inserted: itens.length });

  } catch (error) {
    console.error(`[SuperBusca] Erro em inserirItensSelecionados: ${error.message}`);
    return SharedUtils_errorResponse(error);
  }
}

/**
 * Auxiliar: Converte 1 -> A, 2 -> B, 27 -> AA
 * Usa funcao centralizada de lib/Shared/Utils.gs
 */
function indexToLetter(column) {
  return SharedUtils_numberToColumnLetter(column);
}