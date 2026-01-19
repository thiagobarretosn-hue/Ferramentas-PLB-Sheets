/**
 * SUPER BUSCA - FERRAMENTAS PLB SHEETS
 *
 * Sistema de busca rápida de materiais em planilhas
 * Permite buscar e inserir itens de uma lista de custos
 */

// ============================================================================
// CONFIGURAÇÃO - SUPER BUSCA
// ============================================================================

const SUPER_BUSCA_CONFIG = {
  SOURCE_SHEET: '5.COST.LIST', // Nome exato da aba de origem
  SOURCE_COL: 6, // Coluna F = 6 (A=1, B=2...)
  MENU_TITLE: '🔍 Super Busca'
};

// ============================================================================
// FUNÇÕES DA SUPER BUSCA
// ============================================================================

function abrirSuperBuscaSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SuperBuscaSidebar.html')
    .setTitle('Busca de Materiais')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Pega os dados da lista de custos para a memória do Sidebar
 */
function getDadosParaBusca() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SUPER_BUSCA_CONFIG.SOURCE_SHEET);

  if (!sheet) return []; // Retorna vazio se a aba não existir

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Pega valores da coluna configurada
  const values = sheet.getRange(2, SUPER_BUSCA_CONFIG.SOURCE_COL, lastRow - 1, 1).getDisplayValues();

  // Limpa vazios e duplicatas para deixar a busca mais leve
  const listaLimpa = values.flat().filter(item => item !== "");
  return [...new Set(listaLimpa)]; // Remove duplicatas
}

/**
 * Insere os itens na célula ATIVA (onde o cursor do usuário está)
 */
function inserirItensSelecionados(itens) {
  if (!itens || itens.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeCell = sheet.getActiveCell();

  const row = activeCell.getRow();
  const col = activeCell.getColumn();

  // Prepara o array vertical
  const dadosParaInserir = itens.map(item => [item]);

  // Insere sobrescrevendo para baixo a partir da célula ativa
  sheet.getRange(row, col, dadosParaInserir.length, 1).setValues(dadosParaInserir);

  // Move a seleção para a próxima célula livre abaixo, facilitando inserções contínuas
  const nextRow = row + dadosParaInserir.length;
  sheet.getRange(nextRow, col).activate();
}
