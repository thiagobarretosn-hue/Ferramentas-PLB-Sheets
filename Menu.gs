/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 3.1.0 - Cores unificadas em um item; Gerenciador de Abas atualizado
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Sheet Tools')
    .addItem('📊 Gerador de BOM', 'openBomSidebar')
    .addItem('📋 Gerador de Request', 'openRequestSidebar')
    .addSeparator()
    .addItem('📑 Gerenciador de Abas', 'showSheetManager')
    .addItem('🎨 Cores das Abas', 'openColorConfig')
    .addSeparator()
    .addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('🏆 Bolão Copa 2026')
      .addItem('⚙️ Setup Inicial do Bolão', 'setupBolao')
      .addItem('➕ Adicionar Participante', 'adicionarJogadorBolao')
      .addSeparator()
      .addItem('🔄 Buscar / Atualizar Resultados', 'buscarEPopularJogos')
      .addItem('📊 Recalcular Pontuação', 'recalcularPontuacao')
      .addSeparator()
      .addItem('🔑 Configurar Chave de API', 'configurarChaveAPI')
      .addItem('📋 Guia de Pontuação', 'mostrarGuiaPontuacao')
    )
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  onEditBolao(e);
}
