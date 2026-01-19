/**
 * MENU PRINCIPAL - FERRAMENTAS PLB SHEETS
 *
 * Este arquivo centraliza todos os menus e funções onOpen do sistema.
 * Facilita a adição ou remoção de funcionalidades do menu principal.
 *
 * Estrutura:
 * - onOpen(): Função principal que cria todos os menus
 * - onEdit(): Gerencia todos os triggers de edição
 */

// ============================================================================
// FUNÇÃO PRINCIPAL - onOpen
// ============================================================================

/**
 * Cria todos os menus quando a planilha é aberta.
 * Esta é a única função onOpen que deve existir no projeto.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ========================================
  // MENU: RELATÓRIOS DINÂMICOS (BOM)
  // ========================================
  ui.createMenu('🔧 Relatórios Dinâmicos')
    .addItem('⚙️ Painel de Controle (Sidebar)', 'openConfigSidebar')
    .addItem('📊 Gerador de BOM (Painel)', 'openBomSidebar')
    .addSeparator()
    .addItem('🔧 Fixadores → Fonte', 'abrirSeletorFixadores')
    .addSeparator()
    .addItem('📄 Exportar PDFs (da Aba Config)', 'exportPDFsWithFeedback')
    .addSeparator()
    .addItem('🗑️ Limpar Relatórios', 'clearOldReports')
    .addItem('🔄 Limpar Cache', 'forceRefreshCache')
    .addItem('🧪 Diagnóstico', 'testSystem')
    .addItem('🔧 Recriar Config', 'forceCreateConfig')
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
  // INICIALIZAÇÃO
  // ========================================
  // Garante que a aba Config existe (BOM)
  ensureConfigExists();
}

// ============================================================================
// FUNÇÃO onEdit - GERENCIADOR DE TRIGGERS
// ============================================================================

/**
 * Gerencia todos os triggers de edição do sistema.
 * Chama as funções apropriadas dependendo da aba editada.
 */
function onEdit(e) {
  // Chama o onEdit do BOM (para aba Config)
  if (typeof onEditBom === 'function') {
    onEditBom(e);
  }

  // Chama o onEdit de cores (para coloração automática)
  if (typeof onEditColorTrigger === 'function') {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const allConfigs = getAllColorConfigs();
    const config = allConfigs[sheetName];

    if (config && config.automaticColoring && e.range.getColumn() === config.groupCol) {
      onEditColorTrigger(e);
    }
  }
}
