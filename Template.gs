/**
 * @fileoverview Sistema de Templates PLB - Ferramentas PLB Sheets
 * @version 1.2.0
 *
 * FUNCIONALIDADES:
 * - Gerenciamento de templates de tarefas (carregar, inserir, criar)
 * - Inserção de templates na planilha ativa
 * - Criação e atualização de templates na base central
 *
 * DEPENDÊNCIAS:
 * - lib/Shared/Config.gs (AppConfig)
 * - lib/Shared/Utils.gs (SharedUtils)
 *
 * SheetManager e ColorConfig foram movidos para seus próprios arquivos:
 *   SheetManager.gs — backend do gerenciador de abas
 *   ColorConfig.gs  — configuração de cores por grupo
 */

// ============================================================================
// CONFIGURACAO GLOBAL
// ============================================================================

const _AppConfigFallback = {
  _defaults: {
    CENTRAL_SPREADSHEET_ID: '1IE_NTWtwB9PHlrFsM853SkkVwWttiZxVZPcBDE6qjKk',
    CENTRAL_SHEET_NAME: 'DATA BASE'
  },
  get: function(key, defaultValue) {
    try {
      const props = PropertiesService.getScriptProperties();
      const value = props.getProperty('PLB_CONFIG_' + key);
      if (value) {
        try { return JSON.parse(value); } catch(e) { return value; }
      }
    } catch(e) {}
    return this._defaults[key] || defaultValue;
  }
};

const _Config = (typeof AppConfig !== 'undefined') ? AppConfig : _AppConfigFallback;

function getCentralSpreadsheetId() {
  const config = TemplateConfigService.getAll();
  return config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID] || _Config.get('CENTRAL_SPREADSHEET_ID');
}

function getCentralSheetName() {
  const config = TemplateConfigService.getAll();
  return config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || _Config.get('CENTRAL_SHEET_NAME') || 'DATA BASE';
}

// DEPRECATED: Mantidas apenas para retrocompatibilidade com código legado
const CENTRAL_SPREADSHEET_ID = _Config.get('CENTRAL_SPREADSHEET_ID');
const CENTRAL_SHEET_NAME = _Config.get('CENTRAL_SHEET_NAME');

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

// ============================================================================
// CACHE
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
// UI E SIDEBARS
// ============================================================================

function openTemplateSidebar() {
  const html = HtmlService.createTemplateFromFile('template-sidebar.html');
  const initData = getTemplateInitData();
  html.initData = JSON.stringify(initData);
  const templates = loadTemplatesWithDynamicConfig();
  html.templates = JSON.stringify(templates);

  const sidebar = html.evaluate()
    .setTitle('🏗️ PLB Templates')
    .setWidth(500);

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
    const newRow = SharedUtils_toPositiveInteger(response.getResponseText(), 0);
    if (newRow > 0) {
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

function getSelectedRowForClient() { return getSelectedRowForUI(); }
function inserirLocal(taskName, subTradeName, localName) { return insertLocalTemplates(taskName, subTradeName, localName); }
function inserirMultiplosLocals(selections) { return insertMultipleLocals(selections); }
function atualizarTemplates() { return refreshTemplates(); }

// ============================================================================
// UTILITÁRIOS — wrappers para lib/Shared/Utils.gs
// ============================================================================

function numberToColumnLetter(columnNumber) { return SharedUtils_numberToColumnLetter(columnNumber); }
function columnLetterToIndex(columnLetter) { return SharedUtils_columnLetterToIndex(columnLetter); }
function validateInteger(value, min = 1, max = 10000) { return SharedUtils_validateInteger(value, min, max); }
function createSafeId(text) { return SharedUtils_createSafeId(text); }

/**
 * Substitui TASK/SUB-TRADE em linhas com FIRESTOP na descrição.
 */
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
// CONFIGURAÇÃO DO SISTEMA
// ============================================================================

function getSystemConfiguration() {
  const templateConfig = TemplateConfigService.getAll();
  const dynamicCentralId = templateConfig[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
  const dynamicSheetName = templateConfig[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET];

  const savedConfig = PropertiesService.getScriptProperties().getProperty('SYSTEM_CONFIG');
  const config = savedConfig ? JSON.parse(savedConfig) : { defaultInsertRow: 57 };

  config.centralSpreadsheetId = dynamicCentralId || _Config.get('CENTRAL_SPREADSHEET_ID');
  config.centralSheetName = dynamicSheetName || _Config.get('CENTRAL_SHEET_NAME') || 'DATA BASE';
  config.defaultInsertRow = validateInteger(config.defaultInsertRow, 1, 5000);

  return config;
}

function saveSystemConfiguration(config) {
  PropertiesService.getScriptProperties()
    .setProperty('SYSTEM_CONFIG', JSON.stringify(config));
}

// ============================================================================
// TEMPLATES — CARREGAMENTO E CACHE
// ============================================================================

function loadTemplatesWithCache(forceRefresh = false) {
  if (!forceRefresh) {
    const cachedData = templateCache.getData();
    if (cachedData) return cachedData;
  }
  const freshData = loadTemplatesWithDynamicConfig();
  templateCache.setData(freshData);
  return freshData;
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
  if (!activeRange) throw new Error('Selecione uma célula na planilha antes de inserir templates');
  return activeRange.getRow();
}

function getSelectedRowForUI() {
  try {
    const activeRange = SpreadsheetApp.getActiveRange();
    return activeRange ? activeRange.getRow() : null;
  } catch (error) {
    return null;
  }
}

function adjustFormula(formula, originalRow, targetRow) {
  if (!formula) return '';
  try {
    const rowDifference = (targetRow || 0) - (originalRow || targetRow || 0);
    const destQtyCol = getDynamicColumnMapping().DESTINATION.QTY;

    return String(formula).replace(/([A-Z]+)(\d+)/g, (match, column, row) => {
      const currentRowNum = SharedUtils_toInteger(row);
      const newRowNum = currentRowNum + rowDifference;
      if (column === 'T') return numberToColumnLetter(destQtyCol) + newRowNum;
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
    if (!task) return { success: false, message: 'Task não encontrada' };

    const subTrade = task.subTrades[subTradeName];
    if (!subTrade) return { success: false, message: 'Sub-trade não encontrada' };

    const local = subTrade.locals[localName];
    if (!local || !local.templates || local.templates.length === 0) {
      return { success: false, message: 'Local não possui templates' };
    }

    const startRow = getRequiredSelectionRow();
    const insertResult = pasteTemplateData(local.templates, startRow);
    return { success: true, totalInserted: insertResult.count, startRow: startRow };
  } catch (error) {
    return { success: false, message: error.message || String(error) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
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
    return { success: true, totalInserted: insertResult.count, startRow: startRow };
  } catch (error) {
    return { success: false, message: error.message || String(error) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function pasteTemplateData(templates, startRow) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const destCol = getDynamicColumnMapping().DESTINATION;

  const endRow = startRow + templates.length - 1;
  const lastSheetRow = sheet.getLastRow();
  if (endRow > lastSheetRow) {
    sheet.insertRowsAfter(lastSheetRow, endRow - lastSheetRow);
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
      sheet.getRange(currentRow, destCol.QTY).setValue(template.quantity || '');
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
    const e = existing[i], n = newTemplate[i];
    if (e.task !== n.task || e.subTrade !== n.subTrade || e.description !== n.description || e.local !== n.local) {
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
    const config = TemplateConfigService.getAll();
    const centralSheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';
    if (sheet.getName() === centralSheetName) {
      throw new Error('Selecione uma aba de orçamento, não a planilha central');
    }

    const range = spreadsheet.getActiveRange();
    if (!range) throw new Error('Selecione as linhas que compõem o template');

    const templateData = extractSelectionData(sheet, range.getRow(), range.getNumRows());

    if (!templateData || templateData.length === 0) {
      throw new Error('Nenhum dado válido encontrado na seleção');
    }

    const existingTemplate = findExistingTemplate(templateData[0].local);
    if (!existingTemplate) {
      saveTemplateToDatabase(templateData);
      refreshTemplates();
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Template "${templateData[0].local}" salvo na base central com ${templateData.length} item(ns).`,
        '✅ Template Salvo', 5
      );
    } else {
      if (isTemplateIdentical(existingTemplate.templates, templateData)) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Template "${templateData[0].local}" já existe com configuração idêntica.`,
          'ℹ️ Template Existente', 4
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
  const destCol = getDynamicColumnMapping().DESTINATION;
  const columnData = {
    task:        sheet.getRange(startRow, destCol.TASK,         numRows, 1).getValues(),
    subTask:     sheet.getRange(startRow, destCol['SUB-TASK'],  numRows, 1).getValues(),
    subTrade:    sheet.getRange(startRow, destCol['SUB-TRADE'], numRows, 1).getValues(),
    local:       sheet.getRange(startRow, destCol.LOCAL,        numRows, 1).getValues(),
    description: sheet.getRange(startRow, destCol.DESC,         numRows, 1).getValues(),
    quantity:    sheet.getRange(startRow, destCol.QTY,          numRows, 1).getValues()
  };
  const formulas = sheet.getRange(startRow, destCol.QTY, numRows, 1).getFormulas();

  const extractedData = [];
  let firstLocal = null;
  for (let i = 0; i < numRows; i++) {
    const task        = String(columnData.task[i][0] || '').trim();
    const subTrade    = String(columnData.subTrade[i][0] || '').trim();
    const local       = String(columnData.local[i][0] || '').trim();
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
        task,
        subTask:     String(columnData.subTask[i][0] || '').trim(),
        subTrade,
        local:       firstLocal,
        description,
        quantity:    columnData.quantity[i][0] || '',
        formula:     formulas[i][0] || '',
        originalRow: startRow + i
      });
    }
  }
  return extractedData;
}

function findExistingTemplate(localName) {
  try {
    const config = TemplateConfigService.getAll();
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) { console.warn('ID da planilha central não configurado'); return null; }

    const centralSheet = SpreadsheetApp.openById(centralId);
    const dataSheet = centralSheet.getSheetByName(sheetName);
    if (!dataSheet || dataSheet.getLastRow() < 4) return null;

    const srcCol = COLUMN_MAPPING.SOURCE;
    const lastRow = dataSheet.getLastRow();
    const values = dataSheet.getRange(4, srcCol.TASK, lastRow - 3, 6).getValues();
    const formulas = dataSheet.getRange(4, srcCol.QTY, lastRow - 3, 1).getFormulas();
    const groupedByLocal = {};

    values.forEach((row, index) => {
      const [task, subTask, subTrade, local, description, quantity] = row;
      const cleanLocal = String(local || '').trim();
      if (!cleanLocal) return;
      if (!groupedByLocal[cleanLocal]) groupedByLocal[cleanLocal] = { startRow: index + 4, templates: [] };
      groupedByLocal[cleanLocal].templates.push({
        task, subTask, subTrade, local, description, quantity,
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

    const config = TemplateConfigService.getAll();
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) throw new Error('ID da planilha central não configurado. Configure na sidebar.');

    const centralSheet = SpreadsheetApp.openById(centralId);
    const dataSheet = centralSheet.getSheetByName(sheetName);
    if (!dataSheet) throw new Error('Planilha central não encontrada');

    const srcCol = COLUMN_MAPPING.SOURCE;
    const startRow = dataSheet.getLastRow() < 4 ? 4 : dataSheet.getLastRow() + 1;

    const rowsData = templateData.map(item => [
      item.task, item.subTask || '', item.subTrade || '',
      item.local || '', item.description || '',
      item.formula ? '' : (item.quantity || '')
    ]);
    dataSheet.getRange(startRow, srcCol.TASK, rowsData.length, 6).setValues(rowsData);

    templateData.forEach((item, index) => {
      if (item.formula) {
        try {
          const adjustedFormula = adjustFormulaForDatabase(item.formula, item.originalRow, startRow + index);
          dataSheet.getRange(startRow + index, srcCol.QTY).setFormula(adjustedFormula);
        } catch (error) {
          console.error('Erro ao aplicar fórmula na base:', error);
        }
      }
    });
    return { success: true, startRow, totalRows: rowsData.length };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function adjustFormulaForDatabase(formula, originalRow, targetRow) {
  if (!formula) return '';
  try {
    const rowDifference = targetRow - originalRow;
    const destQtyColLetter = numberToColumnLetter(getDynamicColumnMapping().DESTINATION.QTY);
    return formula.replace(/([A-Z]+)(\d+)/g, (match, column, row) => {
      const newRowNumber = SharedUtils_toInteger(row) + rowDifference;
      if (column === destQtyColLetter) return `T${newRowNumber}`;
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
    const config = TemplateConfigService.getAll();
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';
    if (!centralId) throw new Error('ID da planilha central não configurado. Configure na sidebar.');
    const centralSheet = SpreadsheetApp.openById(centralId);
    const dataSheet = centralSheet.getSheetByName(sheetName);
    if (!dataSheet) throw new Error('Planilha central não encontrada');
    dataSheet.deleteRows(oldTemplates[0].originalRow, oldTemplates.length);
    return saveTemplateToDatabase(newTemplates);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function showDuplicateTemplateDialog(existing, newTemplate) {
  const html = HtmlService.createTemplateFromFile('duplicate-dialog.html');
  html.existingTemplates = JSON.stringify(existing.templates);
  html.newTemplates = JSON.stringify(newTemplate);
  html.localName = newTemplate[0].local;
  const dialog = html.evaluate().setWidth(650).setHeight(480).setTitle('⚠️ Conflito de Template');
  SpreadsheetApp.getUi().showModalDialog(dialog, '⚠️ Conflito de Template');
}

function openCentralDatabase() {
  const config = TemplateConfigService.getAll();
  const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
  if (!centralId) {
    SpreadsheetApp.getUi().alert('Configure o ID da planilha central na sidebar de Templates primeiro.');
    return;
  }
  const url = `https://docs.google.com/spreadsheets/d/${centralId}/edit`;
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial, sans-serif; font-size:13px; padding:6px;">
       📂 Abrindo Base...
       <script>window.open("${url}", "_blank"); google.script.host.close();</script>
     </div>`
  ).setWidth(180).setHeight(50);
  SpreadsheetApp.getUi().showModelessDialog(html, "📂 Abrindo Base... ");
}

// ============================================================================
// CONFIGURAÇÃO DINÂMICA — TemplateConfigService
// ============================================================================

const TEMPLATE_CONFIG_KEYS = {
  CENTRAL_ID:         'TEMPLATE_CENTRAL_SPREADSHEET_ID',
  CENTRAL_SHEET:      'TEMPLATE_CENTRAL_SHEET_NAME',
  DEST_TASK:          'TEMPLATE_DEST_COL_TASK',
  DEST_SUBTASK:       'TEMPLATE_DEST_COL_SUBTASK',
  DEST_SUBTRADE:      'TEMPLATE_DEST_COL_SUBTRADE',
  DEST_LOCAL:         'TEMPLATE_DEST_COL_LOCAL',
  DEST_DESC:          'TEMPLATE_DEST_COL_DESC',
  DEST_QTY:           'TEMPLATE_DEST_COL_QTY',
  DEFAULT_INSERT_ROW: 'TEMPLATE_DEFAULT_INSERT_ROW',
  TASK_COLORS:        'TEMPLATE_TASK_COLORS'
};

const TEMPLATE_CONFIG_DEFAULTS = {
  [TEMPLATE_CONFIG_KEYS.CENTRAL_ID]:         '',
  [TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET]:      'DATA BASE',
  [TEMPLATE_CONFIG_KEYS.DEST_TASK]:          'D',
  [TEMPLATE_CONFIG_KEYS.DEST_SUBTASK]:       'E',
  [TEMPLATE_CONFIG_KEYS.DEST_SUBTRADE]:      'F',
  [TEMPLATE_CONFIG_KEYS.DEST_LOCAL]:         'I',
  [TEMPLATE_CONFIG_KEYS.DEST_DESC]:          'K',
  [TEMPLATE_CONFIG_KEYS.DEST_QTY]:           'L',
  [TEMPLATE_CONFIG_KEYS.DEFAULT_INSERT_ROW]: 57
};

const TemplateConfigService = {
  _cache: null,
  _cacheTime: null,
  _cacheTTL: 180000,

  getAll: function() {
    const now = Date.now();
    if (this._cache && this._cacheTime && (now - this._cacheTime < this._cacheTTL)) return this._cache;

    const props = PropertiesService.getDocumentProperties();
    const config = {};

    Object.entries(TEMPLATE_CONFIG_KEYS).forEach(([key, propKey]) => {
      const value = props.getProperty(propKey);
      config[propKey] = value !== null ? value : (TEMPLATE_CONFIG_DEFAULTS[propKey] || '');
    });

    const hasAnyConfig = props.getProperty(TEMPLATE_CONFIG_KEYS.CENTRAL_ID) !== null;
    if (!hasAnyConfig && !config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID]) {
      config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID] = _Config.get('CENTRAL_SPREADSHEET_ID');
    }
    if (!hasAnyConfig && !config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET]) {
      config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] = _Config.get('CENTRAL_SHEET_NAME') || 'DATA BASE';
    }

    this._cache = config;
    this._cacheTime = now;
    return config;
  },

  get: function(key, defaultValue) {
    defaultValue = defaultValue || '';
    const all = this.getAll();
    return all[key] !== undefined && all[key] !== '' ? all[key] : defaultValue;
  },

  set: function(key, value) {
    PropertiesService.getDocumentProperties().setProperty(key, String(value));
    this._cache = null;
  },

  setAll: function(settings) {
    const props = PropertiesService.getDocumentProperties();
    Object.entries(settings).forEach(([key, value]) => {
      if (value !== undefined && value !== null) props.setProperty(key, String(value));
    });
    this._cache = null;
  },

  clearCache: function() {
    this._cache = null;
    this._cacheTime = null;
  },

  debug: function() {
    const props = PropertiesService.getDocumentProperties();
    const all = props.getProperties();
    const templateProps = {};
    Object.keys(all).forEach(key => { if (key.startsWith('TEMPLATE_')) templateProps[key] = all[key]; });
    console.log('=== DEBUG TemplateConfigService ===');
    console.log('DocumentProperties (TEMPLATE_*):', JSON.stringify(templateProps, null, 2));
    console.log('getAll() retorna:', JSON.stringify(this.getAll(), null, 2));
    return { raw: templateProps, processed: this.getAll() };
  }
};

function debugTemplateConfig() {
  const result = TemplateConfigService.debug();
  const msg = `=== CONFIG DEBUG ===\n\nRAW (DocumentProperties):\n${JSON.stringify(result.raw, null, 2)}\n\nPROCESSED (getAll):\nCENTRAL_ID: ${result.processed[TEMPLATE_CONFIG_KEYS.CENTRAL_ID] || '(vazio)'}\nCENTRAL_SHEET: ${result.processed[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || '(vazio)'}\n`;
  console.log(msg);
  SpreadsheetApp.getUi().alert('Debug Config', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return result;
}

function forceSetCentralSpreadsheet(spreadsheetId, sheetName) {
  if (!spreadsheetId) {
    SpreadsheetApp.getUi().alert('Erro', 'Forneça o ID da planilha como parâmetro', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(TEMPLATE_CONFIG_KEYS.CENTRAL_ID, spreadsheetId);
  props.setProperty(TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET, sheetName || 'DATA BASE');
  TemplateConfigService.clearCache();
  templateCache.clear();
  const savedId = props.getProperty(TEMPLATE_CONFIG_KEYS.CENTRAL_ID);
  const savedSheet = props.getProperty(TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET);
  SpreadsheetApp.getUi().alert('✅ Configuração Forçada',
    `Configuração salva!\n\nID: ${savedId}\nAba: ${savedSheet}\n\nReabra a sidebar para aplicar.`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function configurarPlanilhaCentral() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '⚙️ Configurar Planilha Central',
    'Cole o ID da planilha central de templates:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const id = response.getResponseText().trim();
    if (id) forceSetCentralSpreadsheet(id, 'DATA BASE');
  }
}

function getTemplateInitData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();

    let config = {};
    try { config = TemplateConfigService.getAll(); } catch (e) { console.warn('Erro ao carregar config:', e.message); }

    let localSheets = [];
    try {
      localSheets = ss.getSheets()
        .map(s => s.getName())
        .filter(n => !['Config', 'Template Config'].includes(n));
    } catch (e) { console.warn('Erro ao listar abas:', e.message); }

    let destColumns = [];
    try {
      if (activeSheet) {
        const lastCol = activeSheet.getLastColumn();
        if (lastCol > 0) {
          const headers = activeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
          destColumns = headers.map((h, i) => {
            const letter = numberToColumnLetter(i + 1);
            return `${letter} - ${h || 'Coluna ' + letter}`;
          });
        }
      }
    } catch (e) { console.warn('Erro ao obter colunas:', e.message); }

    let externalSheets = [];
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    if (centralId) {
      try {
        externalSheets = SpreadsheetApp.openById(centralId).getSheets().map(s => s.getName());
      } catch (e) { console.warn('Não foi possível acessar planilha central:', e.message); }
    }

    let taskColors = TASK_COLOR_PALETTE || {};
    const savedColors = config[TEMPLATE_CONFIG_KEYS.TASK_COLORS];
    if (savedColors) {
      try { taskColors = JSON.parse(savedColors); } catch (e) { taskColors = TASK_COLOR_PALETTE || {}; }
    }

    let savedState = null;
    try { savedState = getUserTemplateSidebarState(); } catch (e) { console.warn('Erro ao carregar estado:', e.message); }

    return {
      config, configKeys: TEMPLATE_CONFIG_KEYS,
      localSheets, externalSheets, destColumns,
      taskColors, defaultColors: TASK_COLOR_PALETTE || {},
      subTradeColors: SUBTRADE_COLOR_PALETTE || {},
      savedState, activeSheetName: activeSheet ? activeSheet.getName() : ''
    };
  } catch (error) {
    console.error('Erro em getTemplateInitData:', error);
    return {
      config: {}, configKeys: TEMPLATE_CONFIG_KEYS,
      localSheets: [], externalSheets: [], destColumns: [],
      taskColors: {}, defaultColors: {}, subTradeColors: {},
      savedState: null, activeSheetName: ''
    };
  }
}

function saveTemplateConfig(settings) {
  try {
    TemplateConfigService.setAll(settings);
    templateCache.clear();
    TemplateConfigService.clearCache();
    return { success: true, message: 'Configurações salvas!' };
  } catch (error) {
    console.error('Erro ao salvar configurações:', error);
    return { success: false, message: error.message };
  }
}

function saveUserTemplateSidebarState(stateJson) {
  try { PropertiesService.getUserProperties().setProperty('TEMPLATE_SIDEBAR_STATE', stateJson); }
  catch (e) { console.warn('Erro ao salvar estado:', e.message); }
}

function getUserTemplateSidebarState() {
  try {
    const stateJson = PropertiesService.getUserProperties().getProperty('TEMPLATE_SIDEBAR_STATE');
    return stateJson ? JSON.parse(stateJson) : null;
  } catch (e) { return null; }
}

function getExternalSpreadsheetSheets(spreadsheetId) {
  try {
    if (!spreadsheetId || spreadsheetId.trim() === '') {
      return { success: false, sheets: [], message: 'ID não fornecido' };
    }
    const ss = SpreadsheetApp.openById(spreadsheetId.trim());
    return { success: true, sheets: ss.getSheets().map(s => s.getName()), name: ss.getName() };
  } catch (error) {
    return { success: false, sheets: [], message: error.message };
  }
}

function saveTaskColors(colors) {
  try {
    TemplateConfigService.set(TEMPLATE_CONFIG_KEYS.TASK_COLORS, JSON.stringify(colors));
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function loadTemplatesWithDynamicConfig() {
  try {
    let config = {};
    try {
      config = TemplateConfigService.getAll();
    } catch (e) {
      console.warn('Erro ao carregar config:', e.message);
      config = {
        [TEMPLATE_CONFIG_KEYS.CENTRAL_ID]:    _Config.get('CENTRAL_SPREADSHEET_ID'),
        [TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET]: _Config.get('CENTRAL_SHEET_NAME') || 'DATA BASE'
      };
    }

    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) {
      return { tasks: {}, total: 0, error: 'Configure o ID da planilha central na aba "1. Config"' };
    }

    let centralSheet;
    try {
      centralSheet = SpreadsheetApp.openById(centralId);
    } catch (e) {
      if (e.message && e.message.includes('permiss')) {
        return { tasks: {}, total: 0, error: 'Sem permissao para acessar a planilha. Execute qualquer funcao do menu para autorizar o script.' };
      }
      return { tasks: {}, total: 0, error: `Erro ao acessar planilha central: ${e.message}` };
    }

    const dataSheet = centralSheet.getSheetByName(sheetName);
    if (!dataSheet) {
      return { tasks: {}, total: 0, error: `Aba "${sheetName}" não encontrada na planilha central` };
    }

    const lastRow = dataSheet.getLastRow();
    if (lastRow < 4) {
      return { tasks: {}, total: 0, message: 'Planilha central vazia ou sem dados suficientes' };
    }

    let taskColorCodes = {}, subTradeColorCodes = {};
    try { taskColorCodes = loadTaskColorCodes(centralSheet); } catch (e) { console.warn('Erro ao carregar cores de tasks:', e.message); }
    try { subTradeColorCodes = loadSubTradeColorCodes(centralSheet); } catch (e) { console.warn('Erro ao carregar cores de subtrades:', e.message); }

    const savedColors = config[TEMPLATE_CONFIG_KEYS.TASK_COLORS];
    if (savedColors) {
      try { Object.assign(taskColorCodes, JSON.parse(savedColors)); } catch (e) {}
    }

    return extractTemplateData(dataSheet, lastRow, taskColorCodes, subTradeColorCodes);
  } catch (error) {
    console.error('Erro ao carregar templates:', error);
    return { tasks: {}, total: 0, error: `Erro ao carregar: ${error.message}` };
  }
}

function getDynamicColumnMapping() {
  const config = TemplateConfigService.getAll();
  return {
    DESTINATION: {
      TASK:        columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_TASK]?.split(' - ')[0] || 'D'),
      'SUB-TASK':  columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_SUBTASK]?.split(' - ')[0] || 'E'),
      'SUB-TRADE': columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_SUBTRADE]?.split(' - ')[0] || 'F'),
      LOCAL:       columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_LOCAL]?.split(' - ')[0] || 'I'),
      DESC:        columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_DESC]?.split(' - ')[0] || 'K'),
      QTY:         columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_QTY]?.split(' - ')[0] || 'L')
    },
    SOURCE: COLUMN_MAPPING.SOURCE
  };
}
