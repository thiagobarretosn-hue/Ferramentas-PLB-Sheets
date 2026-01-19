/**
 * SISTEMA DE TEMPLATES PLB
 * Ferramentas PLB Sheets
 *
 * Este arquivo contém todas as funções relacionadas ao sistema de Templates:
 * - Gerenciamento de templates de tarefas
 * - Inserção de templates em planilhas
 * - Configuração de cores automáticas
 * - Gerenciamento de abas (renomear, duplicar, etc.)
 * - Criação e atualização de templates
 */

// ============================================================================
// CONFIGURAÇÃO GLOBAL - TEMPLATE
// ============================================================================

const CENTRAL_SPREADSHEET_ID = "1KD_NXdtEjORJJFYwiFc9uzrZeO57TPPSXWh7bokZMdU";
const CENTRAL_SHEET_NAME = "DATA BASE";

const COLUMN_MAPPING = {
  DESTINATION: { TASK: 4, 'SUB-TASK': 5, 'SUB-TRADE': 6, LOCAL: 9, DESC: 11, QTY: 12 },
  SOURCE: { TASK: 15, 'SUB-TASK': 16, 'SUB-TRADE': 17, LOCAL: 18, DESC: 19, QTY: 20 }
};

const TASK_COLOR_PALETTE = {
  "UNITS": "#ffe5a0",
  "COMMON AREAS": "#473822",
  "CONTINGENCY": "#d4edbc",
  "CONSUMABLES": "#11734b",
  "SITE": "#bfe1f6",
  "NA": "#0a53a8",
  "TEMPORARY FOR CONSTRUCTION": "#3d3d3d",
  "UNDERGROUND": "#ffc8aa",
  "SHELL": "#ffe5a0",
  "ROUGH": "#d4edbc",
  "FINISH": "#bfe1f6",
  "FINAL CONNECTION": "#e6cff2"
};

const SUBTRADE_COLOR_PALETTE = {
  "SEWER FACT.UND.": "#473822",
  "SEWER FACT.": "#ffc8aa",
  "WS FACT.": "#e6e6e6",
  "STORM DRAIN": "#3d3d3d",
  "WH DRAIN": "#ffcfc9",
  "AC DRAIN": "#bfe1f6",
  "SEWER": "#753800",
  "WS": "#0a53a8",
  "METER": "#e6cff2",
  "GAS": "#215a6c"
};

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
// CACHE - TEMPLATE
// ============================================================================

class TemplateCache {
  constructor(ttlMinutes = 5) {
    this.clear();
    this.ttl = ttlMinutes * 60 * 1000;
  }

  isValid() {
    return this.timestamp && (Date.now() - this.timestamp < this.ttl);
  }

  setData(data) {
    this.data = data;
    this.timestamp = Date.now();
  }

  getData() {
    return this.isValid() ? this.data : null;
  }

  clear() {
    this.data = null;
    this.timestamp = null;
  }
}

const templateCache = new TemplateCache();

// ============================================================================
// CONFIGURAÇÃO DE CORES - TEMPLATE
// ============================================================================

function getAllColorConfigs() {
  const properties = PropertiesService.getDocumentProperties();
  const savedConfigs = properties.getProperty('SHEET_COLOR_CONFIGS');
  return savedConfigs ? JSON.parse(savedConfigs) : {};
}

function saveAllColorConfigs(configs) {
  const properties = PropertiesService.getDocumentProperties();
  properties.setProperty('SHEET_COLOR_CONFIGS', JSON.stringify(configs));
}

function applyGroupColors(sheetName = null, configToApply = null) {
  const sheet = sheetName ? SpreadsheetApp.getActive().getSheetByName(sheetName) : SpreadsheetApp.getActive().getActiveSheet();
  if (!sheet) return;

  const config = configToApply || Object.assign({}, DEFAULT_COLOR_CONFIG, getAllColorConfigs()[sheet.getName()] || {});

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

  groupValues.forEach((row, rowIndex) => {
    const groupValue = row[0];
    let color = null;

    if (groupValue) {
      const groupKey = String(groupValue);
      color = config.groupColors[groupKey] || null;
    }

    for (let colIndex = 0; colIndex < backgrounds[rowIndex].length; colIndex++) {
      backgrounds[rowIndex][colIndex] = color;
    }
  });

  range.setBackgrounds(backgrounds);
}

function saveColorConfiguration(configData, sheetName) {
  try {
    if (!sheetName) {
      throw new Error("Nome da aba não fornecido.");
    }
    const allConfigs = getAllColorConfigs();

    const newConfig = {
      startRow: validateInteger(configData.startRow, 1),
      startCol: columnLetterToIndex(configData.startCol),
      endCol: columnLetterToIndex(configData.endCol),
      groupCol: columnLetterToIndex(configData.groupCol),
      groupColors: configData.groupColors || {},
      automaticColoring: !!configData.automaticColoring,
      saturation: configData.saturation,
      luminosity: configData.luminosity
    };

    allConfigs[sheetName] = newConfig;
    saveAllColorConfigs(allConfigs);

    const triggers = ScriptApp.getProjectTriggers();
    let triggerExists = false;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEditColorTrigger') {
        triggerExists = true;
      }
    });
    const anyAutoColoring = Object.values(allConfigs).some(c => c.automaticColoring);
    if (anyAutoColoring && !triggerExists) {
      ScriptApp.newTrigger('onEditColorTrigger')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
    } else if (!anyAutoColoring && triggerExists) {
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onEditColorTrigger') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
    }

    applyGroupColors(sheetName, newConfig);

    return {
      success: true,
      message: `Configurações salvas para a aba "${sheetName}"!`
    };

  } catch (error) {
    return {
      success: false,
      message: `Erro: ${error.message}`
    };
  }
}

function onEditColorTrigger(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const allConfigs = getAllColorConfigs();
  const config = allConfigs[sheetName];

  if (config && config.automaticColoring && e.range.getColumn() === config.groupCol) {
    applyGroupColors(sheetName, config);
  }
}

// ============================================================================
// UI E SIDEBARS - TEMPLATE
// ============================================================================

function openTemplateSidebar() {
  const html = HtmlService.createTemplateFromFile('template-sidebar.html');
  html.templates = JSON.stringify(loadTemplatesWithCache());

  const sidebar = html.evaluate()
    .setTitle('🏗️ PLB Templates')
    .setWidth(450);

  SpreadsheetApp.getUi().showSidebar(sidebar);
}

function openSystemConfig() {
  const config = getSystemConfiguration();
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '⚙️ Configuração do Sistema',
    `Linha padrão para inserção (atual: ${config.defaultInsertRow}):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const newRow = parseInt(response.getResponseText(), 10);
    if (!isNaN(newRow) && newRow > 0) {
      config.defaultInsertRow = newRow;
      saveSystemConfiguration(config);
      ui.alert('✅ Configuração salva com sucesso!');
    } else {
      ui.alert('❌ Valor inválido. Digite um número maior que 0.');
    }
  }
}

function testSystemTemplate() {
  const templates = loadTemplatesWithCache();
  let formulaCount = 0;
  Object.values(templates.tasks).forEach(task => {
    Object.values(task.subTrades).forEach(subTrade => {
      Object.values(subTrade.locals).forEach(local => {
        local.templates.forEach(template => {
          if (template.formula) formulaCount++;
        });
      });
    });
  });
  const message = `Sistema OK!\n\n` +
    `Tasks: ${Object.keys(templates.tasks).length}\n` +
    `Templates: ${templates.total}\n` +
    `Com fórmulas: ${formulaCount}`;
  SpreadsheetApp.getUi().alert('🧪 Teste do Sistema', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function getSelectedRowForClient() {
  return getSelectedRowForUI();
}

function inserirLocal(taskName, subTradeName, localName) {
  return insertLocalTemplates(taskName, subTradeName, localName);
}

function inserirMultiplosLocals(selections) {
  return insertMultipleLocals(selections);
}

function atualizarTemplates() {
  return refreshTemplates();
}

function salvarConfiguracaoCores(data, sheetName) {
  return saveColorConfiguration(data, sheetName);
}

// ============================================================================
// UTILITÁRIOS - TEMPLATE
// ============================================================================

function numberToColumnLetter(columnNumber) {
  let result = '';
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}

function columnLetterToIndex(columnLetter) {
  let index = 0;
  const letters = String(columnLetter || '');
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index;
}

function validateInteger(value, min = 1, max = 10000) {
  if (value == null || value === '') return min;
  const number = parseInt(value, 10);
  if (isNaN(number) || number < min || number > max) {
    throw new Error(`Valor inválido: ${value}. Deve estar entre ${min} e ${max}`);
  }
  return number;
}

function createSafeId(text) {
  return String(text || '')
    .replace(/[^A-Za-z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

function substituirShellFirestop() {
  const planilha = SpreadsheetApp.getActiveSheet();
  const dados = planilha.getDataRange().getValues();
  let contador = 0;
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][10] && dados[i][10].toString().toUpperCase().includes('FIRESTOP')) {
      planilha.getRange(i + 1, 4).setValue('SHELL');
      planilha.getRange(i + 1, 7).setValue('C.A.');
      contador++;
    }
  }

  SpreadsheetApp.getUi().alert(`Operação concluída!\n${contador} linhas alteradas.`);
}

// ============================================================================
// SHEET MANAGER - TEMPLATE
// ============================================================================

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

function getUniqueName(ss, base) {
  let nome = base, i = 2;
  while (ss.getSheetByName(nome)) nome = `${base} (${i++})`;
  return nome;
}

// ============================================================================
// CONFIGURAÇÃO DO SISTEMA - TEMPLATE
// ============================================================================

function getSystemConfiguration() {
  const savedConfig = PropertiesService.getScriptProperties().getProperty('SYSTEM_CONFIG');
  const config = savedConfig ?
    JSON.parse(savedConfig) : {
    defaultInsertRow: 57,
    centralSpreadsheetId: CENTRAL_SPREADSHEET_ID,
    centralSheetName: CENTRAL_SHEET_NAME
  };
  config.defaultInsertRow = validateInteger(config.defaultInsertRow, 1, 5000);
  return config;
}

function saveSystemConfiguration(config) {
  PropertiesService.getScriptProperties()
    .setProperty('SYSTEM_CONFIG', JSON.stringify(config));
}

// ============================================================================
// TEMPLATES - CARREGAMENTO E CACHE
// ============================================================================

function loadTemplatesWithCache(forceRefresh = false) {
  if (!forceRefresh) {
    const cachedData = templateCache.getData();
    if (cachedData) return cachedData;
  }

  const freshData = loadTemplatesFromCentralSheet();
  templateCache.setData(freshData);
  return freshData;
}

function loadTemplatesFromCentralSheet() {
  try {
    const config = getSystemConfiguration();
    const centralSheet = SpreadsheetApp.openById(config.centralSpreadsheetId);
    const dataSheet = centralSheet.getSheetByName(config.centralSheetName);

    if (!dataSheet) {
      return { tasks: {}, total: 0, error: 'Planilha central não encontrada' };
    }

    const lastRow = dataSheet.getLastRow();
    if (lastRow < 4) {
      return { tasks: {}, total: 0 };
    }

    const taskColorCodes = loadTaskColorCodes(centralSheet);
    const subTradeColorCodes = loadSubTradeColorCodes(centralSheet);
    const templateData = extractTemplateData(dataSheet, lastRow, taskColorCodes, subTradeColorCodes);
    return templateData;

  } catch (error) {
    console.error('Erro ao carregar templates:', error);
    return {
      tasks: {},
      total: 0,
      error: error.message
    };
  }
}

function loadTaskColorCodes(spreadsheet) {
  const taskColorCodes = {};
  const taskSheet = spreadsheet.getSheetByName('Task');
  if (taskSheet && taskSheet.getLastRow() > 1) {
    const taskData = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 3).getValues();
    taskData.forEach(([taskName, , colorCode]) => {
      if (taskName) {
        const cleanTaskName = String(taskName).trim();
        taskColorCodes[cleanTaskName] = colorCode || TASK_COLOR_PALETTE[cleanTaskName] || '';
      }
    });
  }

  return taskColorCodes;
}

function loadSubTradeColorCodes(spreadsheet) {
  const subTradeColorCodes = {};
  const subTradeSheet = spreadsheet.getSheetByName('SubTrade');
  if (subTradeSheet && subTradeSheet.getLastRow() > 1) {
    const subTradeData = subTradeSheet.getRange(2, 1, subTradeSheet.getLastRow() - 1, 3).getValues();
    subTradeData.forEach(([subTradeName, , colorCode]) => {
      if (subTradeName) {
        const cleanSubTradeName = String(subTradeName).trim();
        subTradeColorCodes[cleanSubTradeName] = colorCode || SUBTRADE_COLOR_PALETTE[cleanSubTradeName] || '';
      }
    });
  }

  return subTradeColorCodes;
}

function extractTemplateData(sheet, lastRow, taskColorCodes, subTradeColorCodes) {
  const srcCol = COLUMN_MAPPING.SOURCE;
  const dataRange = sheet.getRange(4, srcCol.TASK, lastRow - 3, 6);
  const values = dataRange.getValues();
  const formulas = sheet.getRange(4, srcCol.QTY, lastRow - 3, 1).getFormulas();

  const templates = { tasks: {}, total: 0 };
  values.forEach((row, index) => {
    const [task, subTask, subTrade, local, description, quantity] = row;
    const formula = formulas[index][0] || '';

    if (!task || !subTrade || !local || !description) return;

    const taskName = String(task).trim();
    const subTradeName = String(subTrade).trim();
    const localName = String(local).trim();

    if (!templates.tasks[taskName]) {
      templates.tasks[taskName] = {
        name: taskName,
        colorCode: taskColorCodes[taskName] || TASK_COLOR_PALETTE[taskName] || '#667eea',
        safeId: createSafeId(taskName),
        subTrades: {},
        totalTemplates: 0
      };
    }

    const taskRef = templates.tasks[taskName];

    if (!taskRef.subTrades[subTradeName]) {
      taskRef.subTrades[subTradeName] = {
        name: subTradeName,
        colorCode: subTradeColorCodes[subTradeName] || SUBTRADE_COLOR_PALETTE[subTradeName] || taskRef.colorCode,
        safeId: createSafeId(subTradeName),
        locals: {},
        totalTemplates: 0
      };
    }

    const subTradeRef = taskRef.subTrades[subTradeName];
    if (!subTradeRef.locals[localName]) {
      subTradeRef.locals[localName] = {
        name: localName,
        safeId: createSafeId(localName),
        templates: []
      };
    }

    subTradeRef.locals[localName].templates.push({
      task: taskName,
      subTask: String(subTask || '').trim(),
      subTrade: subTradeName,
      local: localName,
      description: String(description || '').trim(),
      quantity: quantity || 0,
      formula: formula,
      originalRow: index + 4
    });
    subTradeRef.totalTemplates++;
    taskRef.totalTemplates++;
    templates.total++;
  });

  return templates;
}

function refreshTemplates() {
  templateCache.clear();
  return loadTemplatesWithCache(true);
}

// ============================================================================
// INSERÇÃO DE TEMPLATES
// ============================================================================

function getRequiredSelectionRow() {
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) {
    throw new Error('Selecione uma célula na planilha antes de inserir templates');
  }
  return activeRange.getRow();
}

function getSelectedRowForUI() {
  try {
    const activeRange = SpreadsheetApp.getActiveRange();
    return activeRange ?
      activeRange.getRow() : null;
  } catch (error) {
    return null;
  }
}

function adjustFormula(formula, originalRow, targetRow) {
  if (!formula) return '';
  try {
    const rowDifference = (targetRow || 0) - (originalRow || targetRow || 0);
    return String(formula).replace(/([A-Z]+)(\d+)/g, (match, column, row) => {
      const currentRowNum = parseInt(row, 10);
      const newRowNum = currentRowNum + rowDifference;

      if (column === 'T') {
        return numberToColumnLetter(COLUMN_MAPPING.DESTINATION.QTY) + newRowNum;
      }

      return column + newRowNum;
    });
  } catch (error) {
    console.error('Erro ao ajustar fórmula:', error);
    return formula;
  }
}

function insertLocalTemplates(taskName, subTradeName, localName) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);
    if (!taskName || !subTradeName || !localName) {
      return { success: false, message: 'Parâmetros obrigatórios faltando' };
    }

    const templates = loadTemplatesWithCache();
    const task = templates.tasks[taskName];
    if (!task) {
      return { success: false, message: 'Task não encontrada' };
    }

    const subTrade = task.subTrades[subTradeName];
    if (!subTrade) {
      return { success: false, message: 'Sub-trade não encontrada' };
    }

    const local = subTrade.locals[localName];
    if (!local || !local.templates || local.templates.length === 0) {
      return { success: false, message: 'Local não possui templates' };
    }

    const startRow = getRequiredSelectionRow();
    const insertResult = pasteTemplateData(local.templates, startRow);
    return {
      success: true,
      totalInserted: insertResult.count,
      startRow: startRow
    };
  } catch (error) {
    return {
      success: false,
      message: error.message ||
        String(error)
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function insertMultipleLocals(selections) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);
    if (!selections || selections.length === 0) {
      return { success: false, message: 'Nenhuma seleção fornecida' };
    }

    const templates = loadTemplatesWithCache();
    const allTemplates = [];
    selections.forEach(selection => {
      const { task: taskName, subTrade: subTradeName } = selection;
      const localName = selection.nomeLocal || selection.local || selection.nome || selection.localName;

      if (!taskName || !subTradeName || !localName) return;

      const task = templates.tasks[taskName];
      if (!task) return;

      const subTrade = task.subTrades[subTradeName];
      if (!subTrade) return;

      const local = subTrade.locals[localName];
      if (!local || !local.templates) return;

      allTemplates.push(...local.templates);
    });
    if (allTemplates.length === 0) {
      return { success: false, message: 'Nenhum template encontrado nas seleções' };
    }

    const startRow = getRequiredSelectionRow();
    const insertResult = pasteTemplateData(allTemplates, startRow);
    return {
      success: true,
      totalInserted: insertResult.count,
      startRow: startRow
    };
  } catch (error) {
    return {
      success: false,
      message: error.message ||
        String(error)
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function pasteTemplateData(templates, startRow) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const destCol = COLUMN_MAPPING.DESTINATION;

  const endRow = startRow + templates.length - 1;
  const lastSheetRow = sheet.getLastRow();
  if (endRow > lastSheetRow) {
    const rowsToAdd = endRow - lastSheetRow;
    sheet.insertRowsAfter(lastSheetRow, rowsToAdd);
  }

  templates.forEach((template, index) => {
    const currentRow = startRow + index;

    sheet.getRange(currentRow, destCol.TASK).setValue(template.task || '');
    sheet.getRange(currentRow, destCol['SUB-TASK']).setValue(template.subTask || '');
    sheet.getRange(currentRow, destCol['SUB-TRADE']).setValue(template.subTrade || '');
    sheet.getRange(currentRow, destCol.LOCAL).setValue(template.local || '');
    sheet.getRange(currentRow, destCol.DESC).setValue(template.description || '');

    if (template.formula) {
      const adjustedFormula = adjustFormula(
        template.formula,
        template.originalRow || currentRow,
        currentRow
      );
      try {
        sheet.getRange(currentRow, destCol.QTY).setFormula(adjustedFormula);
      } catch (error) {
        console.error('Erro ao aplicar fórmula:', error);
        sheet.getRange(currentRow, destCol.QTY).setValue(template.quantity || '');
      }
    } else {
      sheet.getRange(currentRow, destCol.QTY).setValue(template.quantity ||
        '');
    }
  });

  return { count: templates.length };
}

// ============================================================================
// CRIAÇÃO DE TEMPLATES
// ============================================================================

function isTemplateIdentical(existing, newTemplate) {
  if (existing.length !== newTemplate.length) return false;
  for (let i = 0; i < existing.length; i++) {
    const e = existing[i];
    const n = newTemplate[i];

    if (e.task !== n.task ||
        e.subTrade !== n.subTrade ||
        e.description !== n.description ||
        e.local !== n.local) {
      return false;
    }
  }

  return true;
}

function createTemplateFromSelection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = spreadsheet.getActiveSheet();
    const config = getSystemConfiguration();
    if (sheet.getName() === config.centralSheetName) {
      throw new Error('Selecione uma aba de orçamento, não a planilha central');
    }

    const range = spreadsheet.getActiveRange();
    if (!range) {
      throw new Error('Selecione as linhas que compõem o template');
    }

    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const templateData = extractSelectionData(sheet, startRow, numRows);

    if (!templateData || templateData.length === 0) {
      throw new Error('Nenhum dado válido encontrado na seleção');
    }

    const existingTemplate = findExistingTemplate(templateData[0].local);
    if (!existingTemplate) {
      saveTemplateToDatabase(templateData);
      refreshTemplates();
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Template "${templateData[0].local}" salvo na base central com ${templateData.length} item(ns).`,
        '✅ Template Salvo',
        5
      );
    } else {
      if (isTemplateIdentical(existingTemplate.templates, templateData)) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Template "${templateData[0].local}" já existe com configuração idêntica.`,
          'ℹ️ Template Existente',
          4
        );
      } else {
        showDuplicateTemplateDialog(existingTemplate, templateData);
      }
    }

    return { success: true };
  } catch (error) {
    ui.alert('❌ Erro', error.message || String(error), ui.ButtonSet.OK);
    return { success: false, error: error.message };
  }
}

function extractSelectionData(sheet, startRow, numRows) {
  const destCol = COLUMN_MAPPING.DESTINATION;
  const columnData = {
    task: sheet.getRange(startRow, destCol.TASK, numRows, 1).getValues(),
    subTask: sheet.getRange(startRow, destCol['SUB-TASK'], numRows, 1).getValues(),
    subTrade: sheet.getRange(startRow, destCol['SUB-TRADE'], numRows, 1).getValues(),
    local: sheet.getRange(startRow, destCol.LOCAL, numRows, 1).getValues(),
    description: sheet.getRange(startRow, destCol.DESC, numRows, 1).getValues(),
    quantity: sheet.getRange(startRow, destCol.QTY, numRows, 1).getValues()
  };
  const formulas = sheet.getRange(startRow, destCol.QTY, numRows, 1).getFormulas();

  const extractedData = [];
  let firstLocal = null;
  for (let i = 0; i < numRows; i++) {
    const task = String(columnData.task[i][0] || '').trim();
    const subTrade = String(columnData.subTrade[i][0] || '').trim();
    const local = String(columnData.local[i][0] || '').trim();
    const description = String(columnData.description[i][0] || '').trim();
    if (!firstLocal && local) firstLocal = local;
    if (local && firstLocal && local !== firstLocal) {
      throw new Error(
        'A seleção deve conter apenas um único LOCAL. ' +
        `Encontrados: "${firstLocal}" e "${local}"`
      );
    }

    if (task && subTrade && description && firstLocal) {
      extractedData.push({
        task: task,
        subTask: String(columnData.subTask[i][0] || '').trim(),
        subTrade: subTrade,
        local: firstLocal,
        description: description,
        quantity: columnData.quantity[i][0] || '',
        formula: formulas[i][0] || '',
        originalRow:
          startRow + i
      });
    }
  }

  return extractedData;
}

function findExistingTemplate(localName) {
  try {
    const config = getSystemConfiguration();
    const centralSheet = SpreadsheetApp.openById(config.centralSpreadsheetId);
    const dataSheet = centralSheet.getSheetByName(config.centralSheetName);

    if (!dataSheet || dataSheet.getLastRow() < 4) {
      return null;
    }

    const srcCol = COLUMN_MAPPING.SOURCE;
    const lastRow = dataSheet.getLastRow();
    const values = dataSheet.getRange(4, srcCol.TASK, lastRow - 3, 6).getValues();
    const formulas = dataSheet.getRange(4, srcCol.QTY, lastRow - 3, 1).getFormulas();
    const groupedByLocal = {};

    values.forEach((row, index) => {
      const [task, subTask, subTrade, local, description, quantity] = row;
      const cleanLocal = String(local || '').trim();

      if (!cleanLocal) return;

      if (!groupedByLocal[cleanLocal]) {
        groupedByLocal[cleanLocal] = {
          startRow: index + 4,
          templates: []
        };
      }

      groupedByLocal[cleanLocal].templates.push({
        task: task,
        subTask: subTask,
        subTrade: subTrade,
        local: local,
        description: description,
        quantity: quantity,
        formula: formulas[index][0] || '',
        originalRow: index + 4
      });
    });

    return groupedByLocal[localName] || null;
  } catch (error) {
    console.error('Erro ao procurar template existente:', error);
    return null;
  }
}

function saveTemplateToDatabase(templateData) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const config = getSystemConfiguration();
    const centralSheet = SpreadsheetApp.openById(config.centralSpreadsheetId);
    const dataSheet = centralSheet.getSheetByName(config.centralSheetName);

    if (!dataSheet) {
      throw new Error('Planilha central não encontrada');
    }

    const srcCol = COLUMN_MAPPING.SOURCE;
    const startRow = dataSheet.getLastRow() < 4 ?
      4 : dataSheet.getLastRow() + 1;

    const rowsData = templateData.map(item => [
      item.task,
      item.subTask || '',
      item.subTrade || '',
      item.local || '',
      item.description || '',
      item.formula ? '' : (item.quantity || '')
    ]);
    dataSheet.getRange(startRow, srcCol.TASK, rowsData.length, 6).setValues(rowsData);

    templateData.forEach((item, index) => {
      if (item.formula) {
        try {
          const adjustedFormula = adjustFormulaForDatabase(
            item.formula,
            item.originalRow,
            startRow + index
          );
          dataSheet.getRange(startRow + index,
            srcCol.QTY).setFormula(adjustedFormula);
        } catch (error) {
          console.error('Erro ao aplicar fórmula na base:', error);
        }
      }
    });
    return {
      success: true,
      startRow: startRow,
      totalRows: rowsData.length
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function adjustFormulaForDatabase(formula, originalRow, targetRow) {
  if (!formula) return '';
  try {
    const rowDifference = targetRow - originalRow;
    return formula.replace(/([A-Z]+)(\d+)/g, (match, column, row) => {
      const newRowNumber = parseInt(row, 10) + rowDifference;
      if (column === numberToColumnLetter(COLUMN_MAPPING.DESTINATION.QTY)) {
        return `T${newRowNumber}`;
      }
      return `${column}${newRowNumber}`;
    });
  } catch (error) {
    console.error('Erro ao ajustar fórmula para base:', error);
    return formula;
  }
}

function updateExistingTemplate(oldTemplates, newTemplates) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const config = getSystemConfiguration();
    const centralSheet = SpreadsheetApp.openById(config.centralSpreadsheetId);
    const dataSheet = centralSheet.getSheetByName(config.centralSheetName);

    if (!dataSheet) {
      throw new Error('Planilha central não encontrada');
    }

    const firstOldRow = oldTemplates[0].originalRow;
    dataSheet.deleteRows(firstOldRow, oldTemplates.length);

    return saveTemplateToDatabase(newTemplates);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function showDuplicateTemplateDialog(existing, newTemplate) {
  const html = HtmlService.createTemplateFromFile('duplicate-dialog.html');
  html.existingTemplates = JSON.stringify(existing.templates);
  html.newTemplates = JSON.stringify(newTemplate);
  html.localName = newTemplate[0].local;

  const dialog = html.evaluate()
    .setWidth(650)
    .setHeight(480)
    .setTitle('⚠️ Conflito de Template');
  SpreadsheetApp.getUi().showModalDialog(dialog, '⚠️ Conflito de Template');
}

// ============================================================================
// COLOR CONFIG SIDEBAR
// ============================================================================

function openColorConfig() {
  const html = HtmlService.createTemplateFromFile('color-config-sidebar.html');
  const sidebar = html.evaluate()
    .setTitle('🎨 Configuração de Cores')
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(sidebar);
}

function getInitialDataForSidebar() {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const sheetName = sheet.getName();

  const allConfigs = getAllColorConfigs();
  const savedConfig = allConfigs[sheetName] || {};

  const config = Object.assign({}, DEFAULT_COLOR_CONFIG, savedConfig);
  config.startColLetter = numberToColumnLetter(config.startCol);
  config.endColLetter = numberToColumnLetter(config.endCol);
  config.groupColLetter = numberToColumnLetter(config.groupCol);

  const currentGroups = getCurrentGroups(sheet, config.startRow, config.groupCol);

  return {
    config: config,
    currentGroups: currentGroups,
    sheetName: sheetName
  };
}

function getCurrentGroups(sheet, startRow, groupColumn) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return [];

    const range = sheet.getRange(startRow, groupColumn, lastRow - startRow + 1, 1);
    const values = range.getValues();

    const uniqueGroups = new Set();
    values.forEach(row => {
      const value = row[0];
      if (value !== null && value !== '') uniqueGroups.add(String(value));
    });

    return Array.from(uniqueGroups).sort();
  } catch (e) {
    return [`Erro ao ler grupos: ${e.message}`];
  }
}

function NOME_ABA() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function getRefreshedGroups(sheetName, startRow, groupColLetter) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) {
    return ['Erro: Aba não encontrada'];
  }
  const groupCol = columnLetterToIndex(groupColLetter);
  return getCurrentGroups(sheet, startRow, groupCol);
}

function openCentralDatabase() {
  const url = `https://docs.google.com/spreadsheets/d/${CENTRAL_SPREADSHEET_ID}/edit`;
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial, sans-serif; font-size:13px; padding:6px;">
       📂 Abrindo Base...
       <script>
         window.open("${url}", "_blank");
         google.script.host.close();
       </script>
     </div>`
  ).setWidth(180).setHeight(50);

  SpreadsheetApp.getUi().showModelessDialog(html, "📂 Abrindo Base... ");
}
