/**
 * @fileoverview Utilitários para exportação de PDF
 * @version 1.0.0
 */

/**
 * Utilitários de exportação PDF
 * @namespace
 */
const PDFExport = {

  /**
   * Opções padrão para exportação PDF
   */
  DEFAULT_OPTIONS: {
    size: 'A4',           // A4, Letter, Legal
    orientation: 'portrait', // portrait, landscape
    fitToWidth: true,
    gridlines: false,
    headers: true,
    footers: true,
    margin: {
      top: 0.5,
      bottom: 0.5,
      left: 0.5,
      right: 0.5
    }
  },

  /**
   * Exporta planilha inteira como PDF
   * @param {Spreadsheet} ss - Planilha
   * @param {Folder} folder - Pasta de destino
   * @param {string} filename - Nome do arquivo (sem .pdf)
   * @returns {File} Arquivo PDF criado
   */
  exportSpreadsheet(ss, folder, filename) {
    const blob = ss.getAs('application/pdf');
    return folder.createFile(blob.setName(filename + '.pdf'));
  },

  /**
   * Exporta aba específica como PDF
   * @param {Spreadsheet} ss - Planilha
   * @param {Sheet} sheet - Aba a exportar
   * @param {Folder} folder - Pasta de destino
   * @param {string} filename - Nome do arquivo
   * @param {Object} [options] - Opções de exportação
   * @returns {File} Arquivo PDF criado
   */
  exportSheet(ss, sheet, folder, filename, options = {}) {
    const opts = { ...this.DEFAULT_OPTIONS, ...options };

    // Construir URL de exportação
    const url = this._buildExportUrl(ss, sheet, opts);

    // Fazer requisição autenticada
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Falha ao exportar PDF: ' + response.getContentText());
    }

    // Criar arquivo
    const blob = response.getBlob().setName(filename + '.pdf');
    return folder.createFile(blob);
  },

  /**
   * Exporta múltiplas abas como PDFs separados
   * @param {Spreadsheet} ss - Planilha
   * @param {Sheet[]} sheets - Array de abas
   * @param {Folder} folder - Pasta de destino
   * @param {Function} [nameGenerator] - Função para gerar nome (recebe sheet)
   * @param {Object} [options] - Opções de exportação
   * @returns {File[]} Array de arquivos criados
   */
  exportMultipleSheets(ss, sheets, folder, nameGenerator = null, options = {}) {
    const files = [];
    const getName = nameGenerator || (sheet => sheet.getName());

    sheets.forEach((sheet, index) => {
      try {
        const filename = getName(sheet, index);
        const file = this.exportSheet(ss, sheet, folder, filename, options);
        files.push(file);

        // Toast de progresso
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Exportado: ${filename}`,
          `Progresso: ${index + 1}/${sheets.length}`,
          2
        );

      } catch (error) {
        console.error(`Erro ao exportar ${sheet.getName()}: ${error.message}`);
      }
    });

    return files;
  },

  /**
   * Exporta range específico como PDF
   * @param {Spreadsheet} ss - Planilha
   * @param {Sheet} sheet - Aba
   * @param {string} rangeA1 - Range em notação A1
   * @param {Folder} folder - Pasta de destino
   * @param {string} filename - Nome do arquivo
   * @param {Object} [options] - Opções
   * @returns {File} Arquivo PDF
   */
  exportRange(ss, sheet, rangeA1, folder, filename, options = {}) {
    const opts = {
      ...this.DEFAULT_OPTIONS,
      ...options,
      range: rangeA1
    };

    const url = this._buildExportUrl(ss, sheet, opts);

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Falha ao exportar PDF: ' + response.getContentText());
    }

    const blob = response.getBlob().setName(filename + '.pdf');
    return folder.createFile(blob);
  },

  /**
   * Obtém ou cria pasta no Drive
   * @param {string} folderId - ID da pasta (opcional)
   * @param {string} folderName - Nome da pasta (se criar)
   * @returns {Folder} Pasta do Drive
   */
  getOrCreateFolder(folderId, folderName) {
    if (folderId) {
      try {
        return DriveApp.getFolderById(folderId);
      } catch (e) {
        console.warn('Pasta não encontrada pelo ID, criando nova...');
      }
    }

    // Procurar por nome ou criar
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      return folders.next();
    }

    return DriveApp.createFolder(folderName);
  },

  /**
   * Constrói URL de exportação
   * @private
   */
  _buildExportUrl(ss, sheet, options) {
    const baseUrl = ss.getUrl().replace(/\/edit.*$/, '/export?');

    const params = {
      format: 'pdf',
      gid: sheet.getSheetId(),
      size: options.size || 'A4',
      portrait: options.orientation !== 'landscape',
      fitw: options.fitToWidth,
      gridlines: options.gridlines,
      printheaders: options.headers,
      printfooters: options.footers,
      top_margin: options.margin?.top || 0.5,
      bottom_margin: options.margin?.bottom || 0.5,
      left_margin: options.margin?.left || 0.5,
      right_margin: options.margin?.right || 0.5
    };

    // Adicionar range se especificado
    if (options.range) {
      params.range = options.range;
    }

    // Construir query string
    const queryString = Object.entries(params)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join('&');

    return baseUrl + queryString;
  }
};

// Exportar para uso global
if (typeof module !== 'undefined') {
  module.exports = PDFExport;
}
