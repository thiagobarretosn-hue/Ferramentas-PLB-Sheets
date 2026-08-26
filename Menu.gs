/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 3.2.0 - Removido onEdit vazio (rodava a cada edição sem fazer nada)
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
    .addItem('📦 Montar Submittal', 'openSubmittalSidebar')
    .addToUi();
}
