/**
 * @fileoverview Utilitários para manipulação de planilhas
 * @version 1.0.0
 */

/**
 * Utilitários de planilha
 * @namespace
 */
const SheetUtils = {

  /**
   * Obtém ou cria uma aba
   * @param {Spreadsheet} ss - Planilha
   * @param {string} name - Nome da aba
   * @param {boolean} [clear=false] - Se deve limpar se existir
   * @returns {Sheet} Aba obtida ou criada
   */
  getOrCreateSheet(ss, name, clear = false) {
    let sheet = ss.getSheetByName(name);

    if (sheet && clear) {
      sheet.clear();
    } else if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    return sheet;
  },

  /**
   * Deleta aba se existir
   * @param {Spreadsheet} ss - Planilha
   * @param {string} name - Nome da aba
   * @returns {boolean} Se deletou
   */
  deleteSheetIfExists(ss, name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
      return true;
    }
    return false;
  },

  /**
   * Obtém última linha com dados em uma coluna específica
   * @param {Sheet} sheet - Aba
   * @param {number} col - Número da coluna (1-indexed)
   * @returns {number} Última linha com dados
   */
  getLastRowInColumn(sheet, col) {
    const data = sheet.getRange(1, col, sheet.getMaxRows(), 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] !== '') {
        return i + 1;
      }
    }
    return 0;
  },

  /**
   * Obtém dados como array de objetos
   * @param {Sheet} sheet - Aba
   * @param {number} [headerRow=1] - Linha do cabeçalho
   * @returns {Object[]} Array de objetos com dados
   */
  getDataAsObjects(sheet, headerRow = 1) {
    const data = sheet.getDataRange().getValues();
    if (data.length <= headerRow) return [];

    const headers = data[headerRow - 1];
    const result = [];

    for (let i = headerRow; i < data.length; i++) {
      const row = data[i];
      const obj = {};

      headers.forEach((header, j) => {
        obj[header] = row[j];
      });

      result.push(obj);
    }

    return result;
  },

  /**
   * Converte array de objetos para dados de planilha
   * @param {Object[]} objects - Array de objetos
   * @param {string[]} [headers] - Headers (se não fornecido, usa keys do primeiro objeto)
   * @returns {Array[]} Array 2D para setValues
   */
  objectsToSheetData(objects, headers = null) {
    if (objects.length === 0) return [];

    const keys = headers || Object.keys(objects[0]);
    const result = [keys];

    objects.forEach(obj => {
      result.push(keys.map(key => obj[key] !== undefined ? obj[key] : ''));
    });

    return result;
  },

  /**
   * Converte letra de coluna para índice
   * @param {string} letter - Letra(s) da coluna (A, B, AA, etc)
   * @returns {number} Índice 1-based
   */
  columnLetterToIndex(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
      index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index;
  },

  /**
   * Converte índice para letra de coluna
   * @param {number} index - Índice 1-based
   * @returns {string} Letra(s) da coluna
   */
  indexToColumnLetter(index) {
    let letter = '';
    while (index > 0) {
      const remainder = (index - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      index = Math.floor((index - 1) / 26);
    }
    return letter;
  },

  /**
   * Auto-redimensiona colunas para caber conteúdo
   * @param {Sheet} sheet - Aba
   * @param {number} [startCol=1] - Coluna inicial
   * @param {number} [numCols] - Número de colunas (padrão: todas)
   */
  autoResizeColumns(sheet, startCol = 1, numCols = null) {
    const cols = numCols || sheet.getLastColumn();
    for (let i = startCol; i <= startCol + cols - 1; i++) {
      sheet.autoResizeColumn(i);
    }
  },

  /**
   * Aplica formatação de cabeçalho
   * @param {Sheet} sheet - Aba
   * @param {number} [row=1] - Linha do cabeçalho
   */
  formatHeader(sheet, row = 1) {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    const range = sheet.getRange(row, 1, 1, lastCol);
    range
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
  },

  /**
   * Aplica zebra striping (linhas alternadas)
   * @param {Sheet} sheet - Aba
   * @param {number} [startRow=2] - Linha inicial
   * @param {string} [color='#f3f3f3'] - Cor das linhas pares
   */
  applyZebraStriping(sheet, startRow = 2, color = '#f3f3f3') {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < startRow || lastCol === 0) return;

    for (let i = startRow; i <= lastRow; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i, 1, 1, lastCol).setBackground(color);
      }
    }
  },

  /**
   * Sanitiza nome de aba (remove caracteres inválidos)
   * @param {string} name - Nome original
   * @returns {string} Nome sanitizado
   */
  sanitizeSheetName(name) {
    return name
      .replace(/[\\\/\?\*\[\]\']/g, '')
      .substring(0, 100);
  },

  /**
   * Gera nome único para aba
   * @param {Spreadsheet} ss - Planilha
   * @param {string} baseName - Nome base
   * @returns {string} Nome único
   */
  getUniqueSheetName(ss, baseName) {
    const sanitized = this.sanitizeSheetName(baseName);
    let name = sanitized;
    let counter = 1;

    while (ss.getSheetByName(name)) {
      name = `${sanitized} (${counter++})`;
    }

    return name;
  }
};

// Exportar para uso global
if (typeof module !== 'undefined') {
  module.exports = SheetUtils;
}
