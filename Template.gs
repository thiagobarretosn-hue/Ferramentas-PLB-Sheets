/**
 * @fileoverview Sistema de Templates PLB - Ferramentas PLB Sheets
 * @version 1.1.0
 *
 * Este arquivo contém todas as funções relacionadas ao sistema de Templates:
 *
 * FUNCIONALIDADES:
 * - Gerenciamento de templates de tarefas (carregar, inserir, criar)
 * - Inserção de templates na planilha ativa
 * - Configuração de cores automáticas por grupo/tarefa
 * - Gerenciamento de abas (renomear, duplicar, buscar/substituir)
 * - Criação e atualização de templates na base central
 *
 * DEPENDÊNCIAS:
 * - lib/Shared/Config.gs (AppConfig) - Configurações centralizadas
 * - lib/Shared/Utils.gs (SharedUtils) - Funções utilitárias
 *
 * MENUS:
 * - '🏗️ PLB Templates' - Menu principal de templates
 * - '📑 Gerenciar Abas' - Menu de gerenciamento de abas
 */

// ============================================================================
// CONFIGURACAO GLOBAL - TEMPLATE
// ============================================================================

// Fallback para AppConfig se não estiver definido (lib/Shared/Config.gs)
const _AppConfigFallback = {
  _defaults: {
    CENTRAL_SPREADSHEET_ID: '1KD_NXdtEjORJJFYwiFc9uzrZeO57TPPSXWh7bokZMdU',
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

// Usa AppConfig se existir, senão usa fallback
const _Config = (typeof AppConfig !== 'undefined') ? AppConfig : _AppConfigFallback;

/**
 * Obtem ID da planilha central
 * @returns {string} ID da planilha
 */
function getCentralSpreadsheetId() {
  return _Config.get('CENTRAL_SPREADSHEET_ID');
}

/**
 * Obtem nome da aba central
 * @returns {string} Nome da aba
 */
function getCentralSheetName() {
  return _Config.get('CENTRAL_SHEET_NAME');
}

// DEPRECATED: Constantes antigas - NÃO USAR DIRETAMENTE
// Use TemplateConfigService.get(TEMPLATE_CONFIG_KEYS.CENTRAL_ID) em vez disso
// Mantidas apenas para retrocompatibilidade com código legado
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

/**
 * Aplica cores de fundo baseadas nos valores da coluna de grupo
 * Usa a paleta de cores configurada para cada valor de grupo
 *
 * @public
 * @menuitem '📑 Gerenciar Abas' > '✨ Aplicar Cores'
 * @param {string} [sheetName=null] - Nome da aba (usa aba ativa se omitido)
 * @param {Object} [configToApply=null] - Configuração de cores (usa salva se omitido)
 * @returns {void}
 */
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

/**
 * Abre a sidebar principal de templates
 * Carrega templates da base central e exibe interface de seleção
 *
 * @public
 * @menuitem '🏗️ PLB Templates' > '📋 Abrir Sidebar'
 * @returns {void}
 */
function openTemplateSidebar() {
  const html = HtmlService.createTemplateFromFile('template-sidebar.html');

  // Carrega dados iniciais para a sidebar
  const initData = getTemplateInitData();
  html.initData = JSON.stringify(initData);

  // Carrega templates usando configuração dinâmica
  const templates = loadTemplatesWithDynamicConfig();
  html.templates = JSON.stringify(templates);

  const sidebar = html.evaluate()
    .setTitle('🏗️ PLB Templates')
    .setWidth(500);

  SpreadsheetApp.getUi().showSidebar(sidebar);
}

/**
 * Abre diálogo de configuração do sistema de templates
 * Permite configurar a linha padrão de inserção
 *
 * @public
 * @menuitem '🏗️ PLB Templates' > '⚙️ Configurar Sistema'
 * @returns {void}
 */
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
// Wrappers para funções centralizadas em lib/Shared/Utils.gs
// Mantidos para retrocompatibilidade com código existente
// ============================================================================

/**
 * Converte número de coluna para letra(s)
 * @param {number} columnNumber - Número da coluna (1-indexed)
 * @returns {string} Letra(s) da coluna
 */
function numberToColumnLetter(columnNumber) {
  return SharedUtils_numberToColumnLetter(columnNumber);
}

/**
 * Converte letra(s) de coluna para número
 * @param {string} columnLetter - Letra(s) da coluna
 * @returns {number} Número da coluna (1-indexed)
 */
function columnLetterToIndex(columnLetter) {
  return SharedUtils_columnLetterToIndex(columnLetter);
}

/**
 * Valida e converte para inteiro dentro de um range
 * @param {*} value - Valor a validar
 * @param {number} [min=1] - Valor mínimo
 * @param {number} [max=10000] - Valor máximo
 * @returns {number} Valor validado
 */
function validateInteger(value, min = 1, max = 10000) {
  return SharedUtils_validateInteger(value, min, max);
}

/**
 * Cria ID seguro para uso em HTML/CSS
 * @param {string} text - Texto original
 * @returns {string} ID seguro
 */
function createSafeId(text) {
  return SharedUtils_createSafeId(text);
}

/**
 * Substitui TASK/SUB-TRADE em linhas com FIRESTOP na descrição
 * Define TASK como 'SHELL' e SUB-TRADE como 'C.A.'
 *
 * @public
 * @menuitem '🏗️ PLB Templates' > 'Substituir SHELL em FIRESTOP'
 * @returns {void}
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
// SHEET MANAGER - TEMPLATE
// ============================================================================

/**
 * Abre o gerenciador de abas da planilha
 * Permite renomear, duplicar e buscar/substituir em múltiplas abas
 *
 * @public
 * @menuitem '📑 Gerenciar Abas' > 'Gerenciador de Abas'
 * @returns {void}
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
  // PRIORIZA CONFIGURACAO DINAMICA do TemplateConfigService
  const templateConfig = TemplateConfigService.getAll();
  const dynamicCentralId = templateConfig[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
  const dynamicSheetName = templateConfig[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET];

  const savedConfig = PropertiesService.getScriptProperties().getProperty('SYSTEM_CONFIG');
  const config = savedConfig ?
    JSON.parse(savedConfig) : {
    defaultInsertRow: 57
  };

  // Usa config dinâmica se disponível, senão fallback para _Config
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
// TEMPLATES - CARREGAMENTO E CACHE
// ============================================================================

/**
 * Carrega templates com cache - USA CONFIGURACAO DINAMICA
 * @param {boolean} forceRefresh - Forçar recarregamento
 * @returns {Object} Dados dos templates
 */
function loadTemplatesWithCache(forceRefresh = false) {
  if (!forceRefresh) {
    const cachedData = templateCache.getData();
    if (cachedData) return cachedData;
  }

  // USA CONFIGURACAO DINAMICA em vez de getSystemConfiguration()
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
      const currentRowNum = SharedUtils_toInteger(row);
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
  // Usa mapeamento dinâmico em vez do fixo
  const destCol = getDynamicColumnMapping().DESTINATION;

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
    // USA CONFIGURACAO DINAMICA
    const config = TemplateConfigService.getAll();
    const centralSheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';
    if (sheet.getName() === centralSheetName) {
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
    // USA CONFIGURACAO DINAMICA
    const config = TemplateConfigService.getAll();
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) {
      console.warn('ID da planilha central não configurado');
      return null;
    }

    const centralSheet = SpreadsheetApp.openById(centralId);
    const dataSheet = centralSheet.getSheetByName(sheetName);

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

    // USA CONFIGURACAO DINAMICA
    const config = TemplateConfigService.getAll();
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) {
      throw new Error('ID da planilha central não configurado. Configure na sidebar.');
    }

    const centralSheet = SpreadsheetApp.openById(centralId);
    const dataSheet = centralSheet.getSheetByName(sheetName);

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
      const newRowNumber = SharedUtils_toInteger(row) + rowDifference;
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

    // USA CONFIGURACAO DINAMICA
    const config = TemplateConfigService.getAll();
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) {
      throw new Error('ID da planilha central não configurado. Configure na sidebar.');
    }

    const centralSheet = SpreadsheetApp.openById(centralId);
    const dataSheet = centralSheet.getSheetByName(sheetName);

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
  // USA CONFIGURACAO DINAMICA
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
       <script>
         window.open("${url}", "_blank");
         google.script.host.close();
       </script>
     </div>`
  ).setWidth(180).setHeight(50);

  SpreadsheetApp.getUi().showModelessDialog(html, "📂 Abrindo Base... ");
}

// ============================================================================
// CONFIGURAÇÃO DINÂMICA DA SIDEBAR - TEMPLATE
// ============================================================================

/**
 * Chaves de configuração do Template
 * Armazenadas em DocumentProperties para persistir por documento
 */
const TEMPLATE_CONFIG_KEYS = {
  CENTRAL_ID: 'TEMPLATE_CENTRAL_SPREADSHEET_ID',
  CENTRAL_SHEET: 'TEMPLATE_CENTRAL_SHEET_NAME',
  DEST_TASK: 'TEMPLATE_DEST_COL_TASK',
  DEST_SUBTASK: 'TEMPLATE_DEST_COL_SUBTASK',
  DEST_SUBTRADE: 'TEMPLATE_DEST_COL_SUBTRADE',
  DEST_LOCAL: 'TEMPLATE_DEST_COL_LOCAL',
  DEST_DESC: 'TEMPLATE_DEST_COL_DESC',
  DEST_QTY: 'TEMPLATE_DEST_COL_QTY',
  DEFAULT_INSERT_ROW: 'TEMPLATE_DEFAULT_INSERT_ROW',
  TASK_COLORS: 'TEMPLATE_TASK_COLORS'
};

/**
 * Valores padrão de configuração
 */
const TEMPLATE_CONFIG_DEFAULTS = {
  [TEMPLATE_CONFIG_KEYS.CENTRAL_ID]: '',
  [TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET]: 'DATA BASE',
  [TEMPLATE_CONFIG_KEYS.DEST_TASK]: 'D',
  [TEMPLATE_CONFIG_KEYS.DEST_SUBTASK]: 'E',
  [TEMPLATE_CONFIG_KEYS.DEST_SUBTRADE]: 'F',
  [TEMPLATE_CONFIG_KEYS.DEST_LOCAL]: 'I',
  [TEMPLATE_CONFIG_KEYS.DEST_DESC]: 'K',
  [TEMPLATE_CONFIG_KEYS.DEST_QTY]: 'L',
  [TEMPLATE_CONFIG_KEYS.DEFAULT_INSERT_ROW]: 57
};

/**
 * Serviço de configuração do Template
 * Usa DocumentProperties para persistir configurações por documento
 */
const TemplateConfigService = {
  _cache: null,
  _cacheTime: null,
  _cacheTTL: 180000, // 3 minutos

  /**
   * Obtém todas as configurações
   * @returns {Object} Configurações do template
   */
  getAll: function() {
    const now = Date.now();
    if (this._cache && this._cacheTime && (now - this._cacheTime < this._cacheTTL)) {
      return this._cache;
    }

    const props = PropertiesService.getDocumentProperties();
    const config = {};

    Object.entries(TEMPLATE_CONFIG_KEYS).forEach(([key, propKey]) => {
      const value = props.getProperty(propKey);
      config[propKey] = value !== null ? value : (TEMPLATE_CONFIG_DEFAULTS[propKey] || '');
    });

    // Se não tem ID central configurado, tenta usar o AppConfig
    if (!config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID]) {
      config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID] = _Config.get('CENTRAL_SPREADSHEET_ID');
    }
    if (!config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET]) {
      config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] = _Config.get('CENTRAL_SHEET_NAME') || 'DATA BASE';
    }

    this._cache = config;
    this._cacheTime = now;
    return config;
  },

  /**
   * Obtém uma configuração específica
   * @param {string} key - Chave da configuração
   * @param {*} defaultValue - Valor padrão
   * @returns {*} Valor da configuração
   */
  get: function(key, defaultValue = '') {
    const all = this.getAll();
    return all[key] !== undefined && all[key] !== '' ? all[key] : defaultValue;
  },

  /**
   * Define uma configuração
   * @param {string} key - Chave
   * @param {*} value - Valor
   */
  set: function(key, value) {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(key, String(value));
    this._cache = null; // Invalida cache
  },

  /**
   * Define múltiplas configurações
   * @param {Object} settings - Objeto com chave/valor
   */
  setAll: function(settings) {
    const props = PropertiesService.getDocumentProperties();
    Object.entries(settings).forEach(([key, value]) => {
      if (value !== undefined && value !== null) {
        props.setProperty(key, String(value));
      }
    });
    this._cache = null;
  },

  /**
   * Limpa o cache
   */
  clearCache: function() {
    this._cache = null;
    this._cacheTime = null;
  }
};

/**
 * Obtém dados iniciais para a sidebar do Template
 * @returns {Object} Dados para inicialização
 */
function getTemplateInitData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();

    // Configurações atuais (com fallback)
    let config = {};
    try {
      config = TemplateConfigService.getAll();
    } catch (e) {
      console.warn('Erro ao carregar config:', e.message);
    }

    // Lista de abas do documento atual
    let localSheets = [];
    try {
      localSheets = ss.getSheets()
        .map(s => s.getName())
        .filter(n => !['Config', 'Template Config'].includes(n));
    } catch (e) {
      console.warn('Erro ao listar abas:', e.message);
    }

    // Colunas da aba ativa (para mapeamento de destino)
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
    } catch (e) {
      console.warn('Erro ao obter colunas:', e.message);
    }

    // Lista de planilhas externas recentes (se houver)
    let externalSheets = [];
    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    if (centralId) {
      try {
        const centralSS = SpreadsheetApp.openById(centralId);
        externalSheets = centralSS.getSheets().map(s => s.getName());
      } catch (e) {
        console.warn('Não foi possível acessar planilha central:', e.message);
      }
    }

    // Cores das tasks (salvas ou paleta padrão)
    let taskColors = TASK_COLOR_PALETTE || {};
    const savedColors = config[TEMPLATE_CONFIG_KEYS.TASK_COLORS];
    if (savedColors) {
      try {
        taskColors = JSON.parse(savedColors);
      } catch (e) {
        taskColors = TASK_COLOR_PALETTE || {};
      }
    }

    // Estado salvo do usuário
    let savedState = null;
    try {
      savedState = getUserTemplateSidebarState();
    } catch (e) {
      console.warn('Erro ao carregar estado:', e.message);
    }

    return {
      config: config,
      configKeys: TEMPLATE_CONFIG_KEYS,
      localSheets: localSheets,
      externalSheets: externalSheets,
      destColumns: destColumns,
      taskColors: taskColors,
      defaultColors: TASK_COLOR_PALETTE || {},
      subTradeColors: SUBTRADE_COLOR_PALETTE || {},
      savedState: savedState,
      activeSheetName: activeSheet ? activeSheet.getName() : ''
    };
  } catch (error) {
    console.error('Erro em getTemplateInitData:', error);
    // Retorna objeto mínimo para não quebrar a sidebar
    return {
      config: {},
      configKeys: TEMPLATE_CONFIG_KEYS,
      localSheets: [],
      externalSheets: [],
      destColumns: [],
      taskColors: {},
      defaultColors: {},
      subTradeColors: {},
      savedState: null,
      activeSheetName: ''
    };
  }
}

/**
 * Salva configurações do Template
 * @param {Object} settings - Configurações a salvar
 * @returns {Object} Resultado da operação
 */
function saveTemplateConfig(settings) {
  try {
    TemplateConfigService.setAll(settings);

    // Atualiza COLUMN_MAPPING dinâmicamente para a sessão
    if (settings[TEMPLATE_CONFIG_KEYS.DEST_TASK]) {
      COLUMN_MAPPING.DESTINATION.TASK = columnLetterToIndex(settings[TEMPLATE_CONFIG_KEYS.DEST_TASK].split(' - ')[0]);
    }
    if (settings[TEMPLATE_CONFIG_KEYS.DEST_SUBTASK]) {
      COLUMN_MAPPING.DESTINATION['SUB-TASK'] = columnLetterToIndex(settings[TEMPLATE_CONFIG_KEYS.DEST_SUBTASK].split(' - ')[0]);
    }
    if (settings[TEMPLATE_CONFIG_KEYS.DEST_SUBTRADE]) {
      COLUMN_MAPPING.DESTINATION['SUB-TRADE'] = columnLetterToIndex(settings[TEMPLATE_CONFIG_KEYS.DEST_SUBTRADE].split(' - ')[0]);
    }
    if (settings[TEMPLATE_CONFIG_KEYS.DEST_LOCAL]) {
      COLUMN_MAPPING.DESTINATION.LOCAL = columnLetterToIndex(settings[TEMPLATE_CONFIG_KEYS.DEST_LOCAL].split(' - ')[0]);
    }
    if (settings[TEMPLATE_CONFIG_KEYS.DEST_DESC]) {
      COLUMN_MAPPING.DESTINATION.DESC = columnLetterToIndex(settings[TEMPLATE_CONFIG_KEYS.DEST_DESC].split(' - ')[0]);
    }
    if (settings[TEMPLATE_CONFIG_KEYS.DEST_QTY]) {
      COLUMN_MAPPING.DESTINATION.QTY = columnLetterToIndex(settings[TEMPLATE_CONFIG_KEYS.DEST_QTY].split(' - ')[0]);
    }

    // Limpa cache de templates para recarregar
    templateCache.clear();

    return { success: true, message: 'Configurações salvas!' };
  } catch (error) {
    console.error('Erro ao salvar configurações:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Salva estado da sidebar do Template
 * @param {string} stateJson - Estado em JSON
 */
function saveUserTemplateSidebarState(stateJson) {
  try {
    PropertiesService.getUserProperties().setProperty('TEMPLATE_SIDEBAR_STATE', stateJson);
  } catch (e) {
    console.warn('Erro ao salvar estado:', e.message);
  }
}

/**
 * Carrega estado da sidebar do Template
 * @returns {Object|null} Estado salvo
 */
function getUserTemplateSidebarState() {
  try {
    const stateJson = PropertiesService.getUserProperties().getProperty('TEMPLATE_SIDEBAR_STATE');
    return stateJson ? JSON.parse(stateJson) : null;
  } catch (e) {
    return null;
  }
}

/**
 * Obtém abas de uma planilha externa pelo ID
 * @param {string} spreadsheetId - ID da planilha
 * @returns {Object} Lista de abas ou erro
 */
function getExternalSpreadsheetSheets(spreadsheetId) {
  try {
    if (!spreadsheetId || spreadsheetId.trim() === '') {
      return { success: false, sheets: [], message: 'ID não fornecido' };
    }

    const ss = SpreadsheetApp.openById(spreadsheetId.trim());
    const sheets = ss.getSheets().map(s => s.getName());
    const name = ss.getName();

    return { success: true, sheets: sheets, name: name };
  } catch (error) {
    return { success: false, sheets: [], message: error.message };
  }
}

/**
 * Salva cores customizadas das tasks
 * @param {Object} colors - Objeto com nome da task e cor hex
 * @returns {Object} Resultado
 */
function saveTaskColors(colors) {
  try {
    TemplateConfigService.set(TEMPLATE_CONFIG_KEYS.TASK_COLORS, JSON.stringify(colors));
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Carrega templates usando configuração dinâmica
 * Substitui loadTemplatesFromCentralSheet para usar config do documento
 * @returns {Object} Dados dos templates
 */
function loadTemplatesWithDynamicConfig() {
  try {
    let config = {};
    try {
      config = TemplateConfigService.getAll();
    } catch (e) {
      console.warn('Erro ao carregar config:', e.message);
      // Tenta fallback para AppConfig
      config = {
        [TEMPLATE_CONFIG_KEYS.CENTRAL_ID]: _Config.get('CENTRAL_SPREADSHEET_ID'),
        [TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET]: _Config.get('CENTRAL_SHEET_NAME') || 'DATA BASE'
      };
    }

    const centralId = config[TEMPLATE_CONFIG_KEYS.CENTRAL_ID];
    const sheetName = config[TEMPLATE_CONFIG_KEYS.CENTRAL_SHEET] || 'DATA BASE';

    if (!centralId) {
      return {
        tasks: {},
        total: 0,
        error: 'Configure o ID da planilha central na aba "1. Config"'
      };
    }

    let centralSheet, dataSheet;
    try {
      centralSheet = SpreadsheetApp.openById(centralId);
    } catch (e) {
      if (e.message && e.message.includes('permiss')) {
        return {
          tasks: {},
          total: 0,
          error: 'Sem permissao para acessar a planilha. Execute qualquer funcao do menu para autorizar o script.'
        };
      }
      return {
        tasks: {},
        total: 0,
        error: `Erro ao acessar planilha central: ${e.message}`
      };
    }

    dataSheet = centralSheet.getSheetByName(sheetName);

    if (!dataSheet) {
      return {
        tasks: {},
        total: 0,
        error: `Aba "${sheetName}" não encontrada na planilha central`
      };
    }

    const lastRow = dataSheet.getLastRow();
    if (lastRow < 4) {
      return { tasks: {}, total: 0, message: 'Planilha central vazia ou sem dados suficientes' };
    }

    let taskColorCodes = {};
    let subTradeColorCodes = {};

    try {
      taskColorCodes = loadTaskColorCodes(centralSheet);
    } catch (e) {
      console.warn('Erro ao carregar cores de tasks:', e.message);
    }

    try {
      subTradeColorCodes = loadSubTradeColorCodes(centralSheet);
    } catch (e) {
      console.warn('Erro ao carregar cores de subtrades:', e.message);
    }

    // Mescla cores customizadas salvas
    const savedColors = config[TEMPLATE_CONFIG_KEYS.TASK_COLORS];
    if (savedColors) {
      try {
        const customColors = JSON.parse(savedColors);
        Object.assign(taskColorCodes, customColors);
      } catch (e) {}
    }

    const templateData = extractTemplateData(dataSheet, lastRow, taskColorCodes, subTradeColorCodes);
    return templateData;

  } catch (error) {
    console.error('Erro ao carregar templates:', error);
    return {
      tasks: {},
      total: 0,
      error: `Erro ao carregar: ${error.message}`
    };
  }
}

/**
 * Obtém mapeamento de colunas de destino atualizado
 * @returns {Object} Mapeamento de colunas
 */
function getDynamicColumnMapping() {
  const config = TemplateConfigService.getAll();

  return {
    DESTINATION: {
      TASK: columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_TASK]?.split(' - ')[0] || 'D'),
      'SUB-TASK': columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_SUBTASK]?.split(' - ')[0] || 'E'),
      'SUB-TRADE': columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_SUBTRADE]?.split(' - ')[0] || 'F'),
      LOCAL: columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_LOCAL]?.split(' - ')[0] || 'I'),
      DESC: columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_DESC]?.split(' - ')[0] || 'K'),
      QTY: columnLetterToIndex(config[TEMPLATE_CONFIG_KEYS.DEST_QTY]?.split(' - ')[0] || 'L')
    },
    SOURCE: COLUMN_MAPPING.SOURCE
  };
}
