/**
 * @fileoverview SheetManager — Backend do Gerenciador de Abas
 * @version 1.3.0 — 📌 fixar posição de abas (pins persistidos em DocumentProperties,
 *                  respeitados por ordenar/mover); log com limpeza
 * @version 1.2.0 — organizar em lote: ordenar A→Z/Z→A, cor de aba, mover posição;
 *                  fix duplicar (copy.activate antes de moveActiveSheet)
 * @version 1.1.0 — hide/unhide + status
 */

// ============================================================================
// PINS — abas com posição fixa (v1.3)
// ============================================================================

const SHEETMGR_PINNED_KEY = 'SHEETMGR_PINNED_V1';

/** Lista crua salva (pode conter nomes de abas já apagadas/renomeadas) */
function _getPinnedRaw() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(SHEETMGR_PINNED_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

/** Lista de abas fixadas, já sem nomes que não existem mais */
function getPinnedSheets() {
  const existing = new Set(SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()));
  return _getPinnedRaw().filter(n => existing.has(n));
}

function setPinnedSheets(names) {
  PropertiesService.getDocumentProperties()
    .setProperty(SHEETMGR_PINNED_KEY, JSON.stringify(names || []));
  return { success: true };
}

/** Alterna o pin de uma aba. Retorna a lista atualizada (para a sidebar). */
function togglePinnedSheet(name) {
  const pinned = getPinnedSheets(); // já poda nomes órfãos
  const idx = pinned.indexOf(name);
  if (idx >= 0) pinned.splice(idx, 1);
  else pinned.push(name);
  setPinnedSheets(pinned);
  return pinned;
}

/**
 * Fixa ou solta várias abas de uma vez (v1.4 — hold em lote).
 * @param {string[]} names - Abas selecionadas
 * @param {boolean} pinned - true = fixar, false = soltar
 * @returns {string[]} Log por aba
 */
function setPinnedForSheets(names, pinned) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = new Set(ss.getSheets().map(s => s.getName()));
  const cur = new Set(getPinnedSheets());
  const results = [];

  (names || []).forEach(name => {
    if (!existing.has(name)) { results.push(name + ': não encontrada'); return; }
    if (pinned) {
      if (cur.has(name)) { results.push('📌 ' + name + ': já estava fixada'); }
      else { cur.add(name); results.push('📌 ' + name + ': fixada ✓'); }
    } else {
      if (cur.has(name)) { cur.delete(name); results.push(name + ': solta ✓'); }
      else { results.push(name + ': não estava fixada'); }
    }
  });

  setPinnedSheets([...cur]);
  return results;
}

/** Mantém o pin ao renomear uma aba fixada (pins são por nome) */
function _syncPinnedRename(oldName, newName) {
  const raw = _getPinnedRaw();
  const idx = raw.indexOf(oldName);
  if (idx >= 0) {
    raw[idx] = newName;
    setPinnedSheets(raw);
  }
}

function showSheetManager() {
  // Template (não HtmlOutput direto): necessário para o include() de SharedScripts
  const html = HtmlService.createTemplateFromFile('SheetManager')
    .evaluate()
    .setTitle('Gerenciador de Abas')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAllSheetNamesWithStatus() {
  const pinned = new Set(getPinnedSheets());
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => ({
    name: s.getName(),
    isHidden: s.isSheetHidden(),
    isPinned: pinned.has(s.getName())
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
    _syncPinnedRename(name, newName);
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
    _syncPinnedRename(name, newName);
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
    _syncPinnedRename(name, newName);
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
    // moveActiveSheet move a aba ATIVA — sem activate() moveria a aba que o usuário estava vendo
    copy.activate();
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
    copy.activate();
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
    copy.activate();
    ss.moveActiveSheet(lastIndex + i + 1);
    const newName = names.length === 1 ? baseName : baseName + ' ' + (i + 1);
    copy.setName(newName);
    return name + ' → ' + newName;
  });
}

// ============================================================================
// ORGANIZAR EM LOTE (v1.2/v1.3) — ordenar, colorir, posicionar (com pins)
// ============================================================================

/**
 * Calcula a ordem final de TODAS as abas:
 * - abas fixadas (📌) permanecem nos seus índices absolutos atuais;
 * - as demais formam uma sequência onde o bloco `movingOrdered` é inserido
 *   no slot não-fixado correspondente a `startPos`.
 *
 * @param {string[]} allNames - Ordem atual de todas as abas
 * @param {Set<string>} pinnedSet - Abas fixadas
 * @param {string[]} movingOrdered - Abas a mover, já na ordem final desejada
 * @param {number} startPos - Posição-alvo (1-based) do início do bloco
 * @returns {string[]} Ordem final completa
 */
function _finalOrderWithPins(allNames, pinnedSet, movingOrdered, startPos) {
  const movingSet = new Set(movingOrdered);

  // Índices absolutos (0-based) das fixadas — não mudam
  const pinnedAt = {};
  allNames.forEach((n, i) => { if (pinnedSet.has(n)) pinnedAt[i] = n; });

  // Sequência não-fixada sem as que vão se mover
  const rest = allNames.filter(n => !pinnedSet.has(n) && !movingSet.has(n));

  // Slot de inserção na sequência não-fixada = qtde de slots não-fixados antes de startPos
  let k = 0;
  for (let i = 0; i < Math.min(startPos - 1, allNames.length); i++) {
    if (!Object.prototype.hasOwnProperty.call(pinnedAt, i)) k++;
  }
  k = Math.min(k, rest.length);
  const seq = rest.slice(0, k).concat(movingOrdered, rest.slice(k));

  // Monta a ordem final: fixadas nos seus índices, o resto preenche em sequência
  const final = new Array(allNames.length);
  Object.keys(pinnedAt).forEach(i => { final[i] = pinnedAt[i]; });
  let j = 0;
  for (let i = 0; i < final.length; i++) {
    if (final[i] === undefined) final[i] = seq[j++];
  }
  return final;
}

/**
 * Aplica uma ordem completa de abas, da esquerda para a direita.
 * Abas ocultas são exibidas temporariamente (moveActiveSheet exige aba ativa,
 * e aba oculta não pode ser ativada) e re-ocultadas em seguida.
 */
function _applySheetOrder(ss, desiredOrder) {
  desiredOrder.forEach((name, i) => {
    const current = ss.getSheets();
    if (current[i] && current[i].getName() === name) return; // já no lugar
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const wasHidden = sheet.isSheetHidden();
    if (wasHidden) sheet.showSheet();
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1);
    if (wasHidden) sheet.hideSheet();
  });
}

/** Move `movingOrdered` para um bloco a partir de startPos, respeitando pins. */
function _moveWithPins(ss, allNames, pinnedSet, movingOrdered, startPos) {
  const final = _finalOrderWithPins(allNames, pinnedSet, movingOrdered, startPos);
  _applySheetOrder(ss, final);
  return movingOrdered.map(n => n + ' → posição ' + (final.indexOf(n) + 1));
}

/**
 * Reordena as abas selecionadas alfabeticamente (natural sort), como bloco
 * começando na posição da primeira selecionada. Abas 📌 fixadas não se movem
 * e mantêm o índice absoluto (o bloco flui ao redor delas).
 * @param {string[]} names
 * @param {string} direction - 'asc' (A→Z) ou 'desc' (Z→A)
 */
function sortSelectedSheets(names, direction) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = ss.getSheets().map(s => s.getName());
  const pinnedSet = new Set(getPinnedSheets());

  const valid = (names || []).filter(n => all.indexOf(n) >= 0);
  const skipped = valid.filter(n => pinnedSet.has(n));
  const movable = valid.filter(n => !pinnedSet.has(n));
  const skippedMsgs = skipped.map(n => '📌 ' + n + ': fixada — não movida');

  if (movable.length < 2) {
    return ['Selecione pelo menos 2 abas não fixadas para ordenar'].concat(skippedMsgs);
  }

  const sorted = movable.slice().sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  if (direction === 'desc') sorted.reverse();

  const startPos = Math.min(...movable.map(n => all.indexOf(n))) + 1;
  return _moveWithPins(ss, all, pinnedSet, sorted, startPos).concat(skippedMsgs);
}

/**
 * Move as abas selecionadas (mantendo a ordem atual entre elas) para a posição
 * indicada (1 = primeira). Abas 📌 fixadas não se movem e mantêm o índice
 * absoluto. Posições fora do intervalo são ajustadas.
 */
function moveSelectedToPosition(names, position) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let pos = parseInt(position, 10);
  if (isNaN(pos) || pos < 1) return ['Posição inválida (1 = primeira)'];

  const all = ss.getSheets().map(s => s.getName());
  const pinnedSet = new Set(getPinnedSheets());

  const selectedSet = new Set(names || []);
  const skipped = all.filter(n => selectedSet.has(n) && pinnedSet.has(n));
  const ordered = all.filter(n => selectedSet.has(n) && !pinnedSet.has(n));
  const skippedMsgs = skipped.map(n => '📌 ' + n + ': fixada — não movida');

  if (ordered.length === 0) {
    return ['Selecione pelo menos uma aba não fixada'].concat(skippedMsgs);
  }

  pos = Math.min(pos, all.length);
  return _moveWithPins(ss, all, pinnedSet, ordered, pos).concat(skippedMsgs);
}

/**
 * Aplica cor de aba em lote. Cor vazia/null remove a cor.
 * @param {string[]} names
 * @param {string} color - Hex (ex: '#f1c232') ou '' para remover
 */
function setTabColorSelected(names, color) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hex = color && String(color).trim() ? String(color).trim() : null;
  return (names || []).map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    sheet.setTabColor(hex);
    return name + (hex ? ': cor ' + hex + ' ✓' : ': cor removida ✓');
  });
}

function duplicateAndFindReplace(names, find, replace) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastIndex = ss.getSheets().length;
  return names.map((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return name + ': não encontrada';
    const copy = sheet.copyTo(ss);
    copy.activate();
    ss.moveActiveSheet(lastIndex + i + 1);
    const newName = name.split(find).join(replace || '');
    copy.setName(newName || ('Cópia de ' + name));
    return name + ' → ' + copy.getName();
  });
}
