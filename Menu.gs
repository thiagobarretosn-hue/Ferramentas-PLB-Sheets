/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 2.3.0 - Removidos itens legados migrados para sidebar
 *
 * Menus Disponíveis:
 * - 🔧 Relatórios Dinâmicos — BOM e Request Generator
 * - 📑 Gerenciar Abas — Organização e cores de abas
 * - 🔍 Super Busca — Busca rápida de materiais
 * - ⚙️ Configurações — Configurações gerais do sistema
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔧 Relatórios Dinâmicos')
    .addItem('📊 Gerador de BOM (Painel)', 'openBomSidebar')
    .addSeparator()
    .addItem('📋 Gerador de Request', 'openRequestSidebar')
    .addToUi();

  ui.createMenu('📑 Gerenciar Abas')
    .addItem('Gerenciador de Abas', 'showSheetManager')
    .addItem('🎨 Configurar Cores', 'openColorConfig')
    .addItem('✨ Aplicar Cores', 'applyGroupColors')
    .addToUi();

  ui.createMenu('🔍 Super Busca')
    .addItem('🚀 Abrir Painel', 'abrirSuperBuscaSidebar')
    .addToUi();

  ui.createMenu('⚙️ Configuracoes')
    .addItem('🔧 Configuracoes Gerais', 'showConfigDialog')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
}
