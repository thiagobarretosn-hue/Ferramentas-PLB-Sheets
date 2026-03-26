/**
 * @fileoverview Gerenciador de Abas - AMBAR TOOL
 * @version 1.0.0
 *
 * Funcionalidades:
 * - Renomear abas (prefixo, sufixo, nome completo, localizar/substituir)
 * - Duplicar abas (simples ou com renomeação)
 */

// ============================================================================
// SIDEBAR
// ============================================================================

/**
 * Abre o gerenciador de abas
 * @public
 */
function showSheetManager() {
  SpreadsheetApp.getUi()
    .showModalDialog(
      HtmlService.createTemplateFromFile('SheetManager.html')
        .evaluate()
        .setWidth(600)
        .setHeight(500),
      'Gerenciador de Abas'
    );
}

// ============================================================================
// FUNÇÕES CHAMADAS PELO HTML
// ============================================================================

function getAllSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName());
}

function renameCompleteSelected(selectedSheets, newBaseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return selectedSheets.map((name, index) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    const finalName = selectedSheets.length === 1
      ? getUniqueName(ss, newBaseName)
      : getUniqueName(ss, `${newBaseName} ${index + 1}`);
    sh.setName(finalName);
    return `"${name}" → "${finalName}"`;
  });
}

function findAndReplaceSelected(selectedSheets, findText, replaceText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!findText) throw new Error('Texto para localizar não pode estar vazio');
  return selectedSheets.map(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    if (!name.includes(findText)) return `"${name}" - texto não encontrado`;
    const newName = name.split(findText).join(replaceText);
    const finalName = getUniqueName(ss, newName);
    sh.setName(finalName);
    return `"${name}" → "${finalName}"`;
  });
}

function renameSelected(selectedSheets, prefix, suffix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return selectedSheets.map(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    const newName = (prefix || '') + name + (suffix || '');
    const finalName = getUniqueName(ss, newName);
    sh.setName(finalName);
    return `"${name}" → "${finalName}"`;
  });
}

function duplicateSelected(selectedSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return selectedSheets.map(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    const c = sh.copyTo(ss);
    const novo = getUniqueName(ss, `${name} (cópia)`);
    c.setName(novo);
    return `"${name}" duplicada como "${novo}"`;
  });
}

function duplicateAndRename(selectedSheets, prefix, suffix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return selectedSheets.map(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    const c = sh.copyTo(ss);
    const alvo = getUniqueName(ss, `${prefix || ''}${name}${suffix || ''}`);
    c.setName(alvo);
    return `"${name}" duplicada como "${alvo}"`;
  });
}

function duplicateAndFindReplace(selectedSheets, findText, replaceText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!findText) throw new Error('Texto para localizar não pode estar vazio');
  return selectedSheets.map(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    if (!name.includes(findText)) return `"${name}" - texto não encontrado (sem duplicação)`;
    const c = sh.copyTo(ss);
    const novoNome = name.split(findText).join(replaceText);
    const alvo = getUniqueName(ss, novoNome);
    c.setName(alvo);
    return `"${name}" duplicada como "${alvo}"`;
  });
}

function duplicateAndRenameComplete(selectedSheets, newBaseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return selectedSheets.map((name, index) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return `Erro: "${name}" não encontrada`;
    const c = sh.copyTo(ss);
    const alvo = selectedSheets.length === 1
      ? getUniqueName(ss, newBaseName)
      : getUniqueName(ss, `${newBaseName} ${index + 1}`);
    c.setName(alvo);
    return `"${name}" duplicada como "${alvo}"`;
  });
}

// ============================================================================
// UTILITÁRIOS
// ============================================================================

function getUniqueName(ss, base) {
  let nome = base, i = 2;
  while (ss.getSheetByName(nome)) nome = `${base} (${i++})`;
  return nome;
}
