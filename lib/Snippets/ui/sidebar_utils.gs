/**
 * @fileoverview Utilitários para criação de sidebars e dialogs
 * @version 1.0.0
 */

/**
 * Utilitários de UI
 * @namespace
 */
const UIUtils = {

  /**
   * Mostra sidebar a partir de arquivo HTML
   * @param {string} filename - Nome do arquivo HTML (sem extensão)
   * @param {string} title - Título da sidebar
   * @param {number} [width=350] - Largura em pixels
   */
  showSidebar(filename, title, width = 350) {
    const html = HtmlService.createHtmlOutputFromFile(filename)
      .setTitle(title)
      .setWidth(width);

    SpreadsheetApp.getUi().showSidebar(html);
  },

  /**
   * Mostra dialog modal a partir de arquivo HTML
   * @param {string} filename - Nome do arquivo HTML
   * @param {string} title - Título do dialog
   * @param {number} [width=500] - Largura
   * @param {number} [height=400] - Altura
   */
  showModalDialog(filename, title, width = 500, height = 400) {
    const html = HtmlService.createHtmlOutputFromFile(filename)
      .setWidth(width)
      .setHeight(height);

    SpreadsheetApp.getUi().showModalDialog(html, title);
  },

  /**
   * Mostra dialog não-modal a partir de arquivo HTML
   * @param {string} filename - Nome do arquivo HTML
   * @param {string} title - Título do dialog
   * @param {number} [width=500] - Largura
   * @param {number} [height=400] - Altura
   */
  showModelessDialog(filename, title, width = 500, height = 400) {
    const html = HtmlService.createHtmlOutputFromFile(filename)
      .setWidth(width)
      .setHeight(height);

    SpreadsheetApp.getUi().showModelessDialog(html, title);
  },

  /**
   * Mostra sidebar com template e dados
   * @param {string} filename - Nome do arquivo HTML template
   * @param {string} title - Título
   * @param {Object} data - Dados para o template
   * @param {number} [width=350] - Largura
   */
  showSidebarWithData(filename, title, data, width = 350) {
    const template = HtmlService.createTemplateFromFile(filename);

    // Injeta dados no template
    Object.keys(data).forEach(key => {
      template[key] = data[key];
    });

    const html = template.evaluate()
      .setTitle(title)
      .setWidth(width);

    SpreadsheetApp.getUi().showSidebar(html);
  },

  /**
   * Mostra toast (notificação não-bloqueante)
   * @param {string} message - Mensagem
   * @param {string} [title=''] - Título opcional
   * @param {number} [duration=5] - Duração em segundos
   */
  toast(message, title = '', duration = 5) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, duration);
  },

  /**
   * Mostra alerta simples
   * @param {string} message - Mensagem
   */
  alert(message) {
    SpreadsheetApp.getUi().alert(message);
  },

  /**
   * Mostra alerta com título
   * @param {string} title - Título
   * @param {string} message - Mensagem
   */
  alertWithTitle(title, message) {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  },

  /**
   * Mostra confirmação Yes/No
   * @param {string} title - Título
   * @param {string} message - Mensagem
   * @returns {boolean} True se clicou Yes
   */
  confirm(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
  },

  /**
   * Mostra prompt para input de texto
   * @param {string} title - Título
   * @param {string} message - Mensagem/instrução
   * @returns {string|null} Texto digitado ou null se cancelou
   */
  prompt(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
      return response.getResponseText();
    }
    return null;
  },

  /**
   * Include helper para templates HTML
   * Permite incluir outros arquivos HTML
   * @param {string} filename - Nome do arquivo
   * @returns {string} Conteúdo do arquivo
   */
  include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  },

  /**
   * Cria menu personalizado
   * @param {string} name - Nome do menu
   * @param {Array<{name: string, function: string}|'separator'>} items - Itens do menu
   */
  createMenu(name, items) {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu(name);

    items.forEach(item => {
      if (item === 'separator') {
        menu.addSeparator();
      } else if (item.submenu) {
        const submenu = ui.createMenu(item.name);
        item.submenu.forEach(subitem => {
          submenu.addItem(subitem.name, subitem.function);
        });
        menu.addSubMenu(submenu);
      } else {
        menu.addItem(item.name, item.function);
      }
    });

    menu.addToUi();
  },

  /**
   * Mostra barra de progresso via toast
   * @param {number} current - Valor atual
   * @param {number} total - Valor total
   * @param {string} [message='Processando'] - Mensagem base
   */
  showProgress(current, total, message = 'Processando') {
    const percent = Math.round((current / total) * 100);
    const bar = '█'.repeat(Math.floor(percent / 5)) + '░'.repeat(20 - Math.floor(percent / 5));
    this.toast(`${message}: ${bar} ${percent}%`, 'Progresso', 30);
  }
};

// Exportar para uso global
if (typeof module !== 'undefined') {
  module.exports = UIUtils;
}
