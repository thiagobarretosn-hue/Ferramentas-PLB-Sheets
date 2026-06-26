/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 3.1.0 - Cores unificadas em um item; Gerenciador de Abas atualizado
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Sheet Tools')
    .addItem('📊 Gerador de BOM', 'openBomSidebar')
    .addItem('📋 Gerador de Request', 'openRequestSidebar')
    .addSeparator()
    .addItem('📑 Gerenciador de Abas', 'showSheetManager')
    .addItem('🎨 Cores das Abas', 'openColorConfig')
    .addSeparator()
    .addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar')
    .addSeparator()
    .addItem('📊 Summary All', 'openSummaryAllSidebar')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
}
