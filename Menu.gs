/**
 * @fileoverview Menu Principal - AMBAR TOOL (branch Cost)
 * @version 1.0.0
 *
 * Aplicações ativas neste branch:
 * - Gerenciador de Abas (SheetManager.gs)
 * - Super Busca (SuperBusca.gs)
 */

// ============================================================================
// FUNÇÃO PRINCIPAL - onOpen
// ============================================================================

/**
 * Cria o menu AMBAR TOOL ao abrir a planilha
 * @trigger onOpen
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AMBAR TOOL')
    .addItem('📑 Gerenciador de Abas', 'showSheetManager')
    .addSeparator()
    .addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar')
    .addToUi();
}

// ============================================================================
// TRIGGER DE EDIÇÃO
// ============================================================================

/**
 * @trigger onEdit
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) return;
}
