/**
 * @fileoverview SheetManager — Backend do Gerenciador de Abas
 * @version 1.1.0 — hide/unhide + status
 */

function showSheetManager() {
  const html = HtmlService.createHtmlOutputFromFile('SheetManager')
    .setTitle('Gerenciador de Abas')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAllSheetNamesWithStatus() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => ({
    name: s.getName(),
    isHidden: s.isSheetHidden()
  }));
}

function getAllSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

function hideSelected(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visibleCount = ss.getSheets().filter(s => !s.isSheetHidden()).length;
  let willHide = 0;
  const results = [];
  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { results.push(name + ': não encontrada'); return; }
    if (sheet.isSheetHidden()) { results.push(name + ': já estava oculta'); return; }
    if (visibleCount - willHide <= 1) {
      results.push(name + ': única aba visível — não pode ocultar');
      return;
    }
    sheet.hideSheet();
    willHide++;
    results.push(name + ': ocultada ✓');
  });
  return results;
}

function unhideSelected(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { results.push(name + ': não encontrada'); return; }
    if (!sheet.isSheetHidden()) { results.push(name + ': já estava visível'); return; }
    sheet.showSheet();
    results.push(name + ': exibida ✓');
  });
  return results;
}

function renameSelected(names, prefix, suffix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return names.map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const newName = (prefix || '') + name + (suffix || '');
    sheet.setName(newName);
    return name + ' → ' + newName;
  });
}

function renameCompleteSelected(names, baseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return names.map((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const newName = names.length === 1 ? baseName : baseName + ' ' + (i + 1);
    sheet.setName(newName);
    return name + ' → ' + newName;
  });
}

function findAndReplaceSelected(names, find, replace) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return names.map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const newName = name.split(find).join(replace || '');
    if (newName === name) return name + ': sem correspondência';
    sheet.setName(newName);
    return name + ' → ' + newName;
  });
}

function duplicateSelected(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const lastIndex = sheets.length;
  return names.map((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const copy = sheet.copyTo(ss);
    ss.moveActiveSheet(lastIndex + i + 1);
    copy.setName('Cópia de ' + name);
    return name + ': duplicada';
  });
}

function duplicateAndRename(names, prefix, suffix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastIndex = ss.getSheets().length;
  return names.map((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const copy = sheet.copyTo(ss);
    ss.moveActiveSheet(lastIndex + i + 1);
    const newName = (prefix || '') + name + (suffix || '');
    copy.setName(newName);
    return name + ' → ' + newName;
  });
}

function duplicateAndRenameComplete(names, baseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastIndex = ss.getSheets().length;
  return names.map((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const copy = sheet.copyTo(ss);
    ss.moveActiveSheet(lastIndex + i + 1);
    const newName = names.length === 1 ? baseName : baseName + ' ' + (i + 1);
    copy.setName(newName);
    return name + ' → ' + newName;
  });
}

function duplicateAndFindReplace(names, find, replace) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastIndex = ss.getSheets().length;
  return names.map((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const copy = sheet.copyTo(ss);
    ss.moveActiveSheet(lastIndex + i + 1);
    const newName = name.split(find).join(replace || '');
    copy.setName(newName || ('Cópia de ' + name));
    return name + ' → ' + copy.getName();
  });
}
