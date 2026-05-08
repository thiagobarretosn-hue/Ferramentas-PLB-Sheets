/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 3.0.0 - Menu único Sheet Tools
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Sheet Tools')
    .addItem('📊 Gerador de BOM', 'openBomSidebar')
    .addItem('📋 Gerador de Request', 'openRequestSidebar')
    .addSeparator()
    .addItem('Gerenciador de Abas', 'showSheetManager')
    .addItem('🎨 Configurar Cores', 'openColorConfig')
    .addItem('✨ Aplicar Cores', 'applyGroupColors')
    .addSeparator()
    .addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
}
