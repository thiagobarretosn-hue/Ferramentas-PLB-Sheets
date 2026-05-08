/**
 * @fileoverview Configuração de Cores por Grupo — Sheet Tools
 * @version 1.0.0
 *
 * Gerencia coloração automática de abas por valor de coluna de grupo.
 * Configurações persistidas em DocumentProperties (por planilha).
 *
 * Funções públicas (menu / sidebar):
 *   openColorConfig()
 *
 * Funções chamadas pela color-config-sidebar.html:
 *   getInitialDataForSidebar()
 *   salvarConfiguracaoCores(data, sheetName)
 *   applyGroupColors(sheetName, configToApply)
 *   getRefreshedGroups(sheetName, startRow, groupColLetter)
 *   NOME_ABA()
 */

// ============================================================================
// CONSTANTES
// ============================================================================

const DEFAULT_COLOR_CONFIG = {
  startRow: 2,
  startCol: 1,
  endCol: 15,
  groupCol: 1,
  groupColors: {},
  automaticColoring: false,
  saturation: 80,
  luminosity: 95
};

// ============================================================================
// PERSISTÊNCIA
// ============================================================================

function getAllColorConfigs() {
  const props = PropertiesService.getDocumentProperties();
  const saved = props.getProperty('SHEET_COLOR_CONFIGS');
  return saved ? JSON.parse(saved) : {};
}

function saveAllColorConfigs(configs) {
  PropertiesService.getDocumentProperties()
    .setProperty('SHEET_COLOR_CONFIGS', JSON.stringify(configs));
}

// ============================================================================
// COLORAÇÃO
// ============================================================================

/**
 * Aplica cores de fundo baseadas nos valores da coluna de grupo.
 * @param {string|null} sheetName - Nome da aba (usa ativa se null)
 * @param {Object|null} configToApply - Config de cores (usa salva se null)
 */
function applyGroupColors(sheetName, configToApply) {
  sheetName = sheetName || null;
  configToApply = configToApply || null;

  const sheet = sheetName
    ? SpreadsheetApp.getActive().getSheetByName(sheetName)
    : SpreadsheetApp.getActive().getActiveSheet();
  if (!sheet) return;

  const config = configToApply
    || Object.assign({}, DEFAULT_COLOR_CONFIG, getAllColorConfigs()[sheet.getName()] || {});

  const lastRow = sheet.getLastRow();
  if (lastRow < config.startRow) return;

  const range = sheet.getRange(
    config.startRow,
    config.startCol,
    lastRow - config.startRow + 1,
    config.endCol - config.startCol + 1
  );

  const groupValues = sheet.getRange(
    config.startRow,
    config.groupCol,
    lastRow - config.startRow + 1,
    1
  ).getValues();

  const backgrounds = range.getBackgrounds();

  groupValues.forEach(function (row, rowIndex) {
    const groupValue = row[0];
    const color = groupValue ? (config.groupColors[String(groupValue)] || null) : null;
    for (let c = 0; c < backgrounds[rowIndex].length; c++) {
      backgrounds[rowIndex][c] = color;
    }
  });

  range.setBackgrounds(backgrounds);
}

// ============================================================================
// SALVAR CONFIGURAÇÃO
// ============================================================================

function saveColorConfiguration(configData, sheetName) {
  try {
    if (!sheetName) throw new Error('Nome da aba não fornecido.');

    const allConfigs = getAllColorConfigs();

    const newConfig = {
      startRow:  SharedUtils_validateInteger(configData.startRow, 1),
      startCol:  SharedUtils_columnLetterToIndex(configData.startCol),
      endCol:    SharedUtils_columnLetterToIndex(configData.endCol),
      groupCol:  SharedUtils_columnLetterToIndex(configData.groupCol),
      groupColors: configData.groupColors || {},
      automaticColoring: !!configData.automaticColoring,
      saturation: configData.saturation,
      luminosity: configData.luminosity
    };

    allConfigs[sheetName] = newConfig;
    saveAllColorConfigs(allConfigs);

    // Gerencia trigger de coloração automática
    const triggers = ScriptApp.getProjectTriggers();
    let triggerExists = false;
    triggers.forEach(function (t) {
      if (t.getHandlerFunction() === 'onEditColorTrigger') triggerExists = true;
    });

    const anyAutoColoring = Object.values(allConfigs).some(function (c) { return c.automaticColoring; });

    if (anyAutoColoring && !triggerExists) {
      ScriptApp.newTrigger('onEditColorTrigger')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
    } else if (!anyAutoColoring && triggerExists) {
      triggers.forEach(function (t) {
        if (t.getHandlerFunction() === 'onEditColorTrigger') ScriptApp.deleteTrigger(t);
      });
    }

    applyGroupColors(sheetName, newConfig);

    return { success: true, message: 'Configurações salvas para a aba "' + sheetName + '"!' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

/** Wrapper chamado pela color-config-sidebar.html */
function salvarConfiguracaoCores(data, sheetName) {
  return saveColorConfiguration(data, sheetName);
}

// ============================================================================
// TRIGGER AUTOMÁTICO (installable)
// ============================================================================

function onEditColorTrigger(e) {
  if (!e || !e.source || !e.range) return;
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const config = getAllColorConfigs()[sheetName];

    if (!config || !config.automaticColoring) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (row >= config.startRow && col >= config.startCol && col <= config.endCol) {
      applyGroupColors(sheetName, config);
    }
  } catch (error) {
    console.error('[onEditColorTrigger] ' + error.message);
  }
}

// ============================================================================
// SIDEBAR
// ============================================================================

function openColorConfig() {
  const sidebar = HtmlService.createTemplateFromFile('color-config-sidebar.html')
    .evaluate()
    .setTitle('🎨 Configuração de Cores')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(sidebar);
}

function getInitialDataForSidebar() {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const sheetName = sheet.getName();

  const savedConfig = getAllColorConfigs()[sheetName] || {};
  const config = Object.assign({}, DEFAULT_COLOR_CONFIG, savedConfig);

  config.startColLetter = SharedUtils_numberToColumnLetter(config.startCol);
  config.endColLetter   = SharedUtils_numberToColumnLetter(config.endCol);
  config.groupColLetter = SharedUtils_numberToColumnLetter(config.groupCol);

  const currentGroups = _getGroupValues(sheet, config.startRow, config.groupCol);

  const lastCol = sheet.getLastColumn();
  const headers = [];
  if (lastCol > 0) {
    sheet.getRange(1, 1, 1, lastCol).getValues()[0].forEach(function (h, i) {
      const letter = SharedUtils_numberToColumnLetter(i + 1);
      headers.push({ letter: letter, name: h ? String(h) : '', label: letter + (h ? ' - ' + h : '') });
    });
  }

  return { config: config, currentGroups: currentGroups, sheetName: sheetName, headers: headers };
}

function getCurrentGroups(sheet, startRow, groupColumn) {
  return _getGroupValues(sheet, startRow, groupColumn);
}

function _getGroupValues(sheet, startRow, groupColumn) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return [];
    const values = sheet.getRange(startRow, groupColumn, lastRow - startRow + 1, 1).getValues();
    const unique = new Set();
    values.forEach(function (row) {
      if (row[0] !== null && row[0] !== '') unique.add(String(row[0]));
    });
    return Array.from(unique).sort();
  } catch (e) {
    return ['Erro ao ler grupos: ' + e.message];
  }
}

function NOME_ABA() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function getRefreshedGroups(sheetName, startRow, groupColLetter) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return ['Erro: Aba não encontrada'];
  return _getGroupValues(sheet, startRow, SharedUtils_columnLetterToIndex(groupColLetter));
}
