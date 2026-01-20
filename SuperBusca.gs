/**
 * SUPER BUSCA - FERRAMENTAS PLB SHEETS
 * Versão: Dropdown Híbrido (Nome ou Letra da Coluna)
 */

const SUPER_BUSCA_CONFIG = {
  SOURCE_SHEET: '5.COST.LIST',
  DEFAULT_COL_INDEX: 6, // Padrão: Coluna F
  MENU_TITLE: '🔍 Super Busca'
};

// ============================================================================
// FUNÇÕES PRINCIPAIS
// ============================================================================

function abrirSuperBuscaSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SuperBuscaSidebar.html')
    .setTitle('Busca de Materiais')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetList() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

/**
 * Retorna lista de colunas.
 * Se tiver cabeçalho: Mostra o Nome.
 * Se não tiver: Mostra a Letra (ex: "Coluna F").
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
 * Busca dados da coluna selecionada
 */
function getDadosParaBusca(sheetName, colIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const targetSheetName = sheetName || SUPER_BUSCA_CONFIG.SOURCE_SHEET;
  const targetCol = colIndex ? Number(colIndex) : SUPER_BUSCA_CONFIG.DEFAULT_COL_INDEX;

  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error(`Aba não encontrada.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, targetCol, lastRow - 1, 1).getDisplayValues();
  
  const listaLimpa = values.flat().filter(item => item !== "");
  return [...new Set(listaLimpa)];
}

function inserirItensSelecionados(itens) {
  if (!itens || itens.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeCell = sheet.getActiveCell();
  
  const row = activeCell.getRow();
  const col = activeCell.getColumn();
  
  const dadosParaInserir = itens.map(item => [item]);
  sheet.getRange(row, col, dadosParaInserir.length, 1).setValues(dadosParaInserir);
  
  const nextRow = row + dadosParaInserir.length;
  sheet.getRange(nextRow, col).activate();
}

/**
 * Auxiliar: Converte 1 -> A, 2 -> B, 27 -> AA
 */
function indexToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}