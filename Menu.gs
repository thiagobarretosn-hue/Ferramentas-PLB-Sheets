/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 2.0.0 - V3.0 Arquitetura independente (sem aba Config)
 *
 * Este arquivo centraliza todos os menus e funções onOpen do sistema.
 * Facilita a adição ou remoção de funcionalidades do menu principal.
 *
 * IMPORTANTE: Este é o único arquivo que deve conter funções onOpen e onEdit.
 * Todas as outras funcionalidades devem ser chamadas a partir daqui.
 *
 * Menus Disponíveis:
 * - 🔧 Relatórios Dinâmicos (BOM) - Geração e exportação de relatórios
 * - 🏗️ PLB Templates - Sistema de templates
 * - 📑 Gerenciar Abas - Organização e cores de abas
 * - 🔍 Super Busca - Busca rápida de materiais
 * - ⚙️ Configurações - Configurações gerais do sistema
 */

// ============================================================================
// FUNÇÃO PRINCIPAL - onOpen (TRIGGER SIMPLES)
// ============================================================================

/**
 * Cria todos os menus quando a planilha é aberta
 * Esta é a ÚNICA função onOpen que deve existir no projeto
 *
 * IMPORTANTE: Simple triggers têm limitações:
 * - Não podem acessar serviços que requerem autorização (não se aplica a createMenu)
 * - Tempo máximo de execução de 30 segundos
 * - Não podem fazer alterações que afetem outros usuários
 *
 * @trigger onOpen - Executada automaticamente ao abrir a planilha
 * @returns {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ========================================
  // MENU: RELATÓRIOS DINÂMICOS (BOM)
  // ========================================
  ui.createMenu('🔧 Relatórios Dinâmicos')
    .addItem('📊 Gerador de BOM (Painel)', 'openBomSidebar')
    .addSeparator()
    .addItem('🔧 Fixadores → Fonte', 'abrirSeletorFixadores')
    .addSeparator()
    .addItem('📄 Exportar PDFs', 'exportPDFsWithFeedback')
    .addSeparator()
    .addItem('🗑️ Limpar Relatórios', 'clearOldReports')
    .addItem('🔄 Limpar Cache', 'forceRefreshCache')
    .addItem('🧪 Diagnóstico', 'testSystem')
    .addToUi();

  // ========================================
  // MENU: PLB TEMPLATES
  // ========================================
  ui.createMenu('🏗️ PLB Templates')
    .addItem('📋 Abrir Sidebar', 'openTemplateSidebar')
    .addItem('🔄 Atualizar Templates', 'refreshTemplates')
    .addSeparator()
    .addItem('➕ Criar Template da Seleção', 'createTemplateFromSelection')
    .addItem('⚙️ Configurar Sistema', 'openSystemConfig')
    .addSeparator()
    .addItem('📂 Abrir Base de Dados', 'openCentralDatabase')
    .addItem('🧪 Testar Sistema', 'testSystemTemplate')
    .addItem('Substituir SHELL em FIRESTOP', 'substituirShellFirestop')
    .addToUi();

  // ========================================
  // MENU: GERENCIAR ABAS
  // ========================================
  ui.createMenu('📑 Gerenciar Abas')
    .addItem('Gerenciador de Abas', 'showSheetManager')
    .addItem('🎨 Configurar Cores', 'openColorConfig')
    .addItem('✨ Aplicar Cores', 'applyGroupColors')
    .addToUi();

  // ========================================
  // MENU: SUPER BUSCA
  // ========================================
  ui.createMenu('🔍 Super Busca')
    .addItem('🚀 Abrir Painel', 'abrirSuperBuscaSidebar')
    .addToUi();

  // ========================================
  // MENU: CONFIGURACOES DO SISTEMA
  // ========================================
  ui.createMenu('⚙️ Configuracoes')
    .addItem('🔧 Configuracoes Gerais', 'showConfigDialog')
    .addToUi();

}

// ============================================================================
// FUNCAO onEdit - TRIGGER DE EDIÇÃO
// ============================================================================

/**
 * Gerencia todos os triggers de edição do sistema
 * Redireciona para os handlers apropriados baseado na aba editada
 *
 * IMPORTANTE: Simple trigger - limite de 30 segundos
 * - Não pode acessar PropertiesService
 * - Não pode enviar emails
 * - Não pode criar arquivos no Drive
 *
 * Handlers registrados:
 * - onEditColorTrigger() - Coloração automática por grupo
 *
 * @trigger onEdit - Executada automaticamente ao editar uma célula
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - Evento de edição
 * @returns {void}
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) return;

  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    // ========================================
    // TRIGGER: Coloracao automatica
    // ========================================
    if (typeof onEditColorTrigger === 'function') {
      const allConfigs = getAllColorConfigs();
      const config = allConfigs[sheetName];

      if (config && config.automaticColoring) {
        const editedCol = e.range.getColumn();
        if (editedCol === config.groupCol) {
          onEditColorTrigger(e);
        }
      }
    }

  } catch (error) {
    console.error(`[onEdit] Erro: ${error.message}`);
  }
}
