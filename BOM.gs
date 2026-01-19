/**
 * @OnlyCurrentDoc
 * SISTEMA UNIFICADO DE RELATÓRIOS DINÂMICOS + FIXADORES (BOM)
 * Versão 2.13 - Ferramentas PLB Sheets
 *
 * Este arquivo contém todas as funções relacionadas ao sistema BOM:
 * - Configuração e gestão da aba Config
 * - Processamento de relatórios dinâmicos
 * - Exportação de PDFs
 * - Gestão de fixadores
 * - Painéis de agrupamento e pré-visualização
 */

// ============================================================================
// CONFIGURAÇÃO GLOBAL - BOM
// ============================================================================

const CONFIG = {
  SHEETS: {
    CONFIG: 'Config'
  },
  KEYS: {
    // Chaves do BOM
    SOURCE_SHEET: 'Aba Origem',
    GROUP_L1: 'Agrupar por Nível 1',
    GROUP_L2: 'Agrupar por Nível 2',
    GROUP_L3: 'Agrupar por Nível 3',
    COL_1: 'Coluna 1',
    COL_2: 'Coluna 2',
    COL_3: 'Coluna 3',
    COL_4: 'Coluna 4',
    COL_5: 'Coluna 5',
    PROJECT: 'Project',
    BOM: 'BOM',
    KOJO_PREFIX: 'KOJO Prefixo',
    ENGINEER: 'Engenheiro',
    VERSION: 'Versão',
    DRIVE_FOLDER_ID: 'Pasta Drive ID',
    DRIVE_FOLDER_NAME: 'Pasta Nome',
    PDF_PREFIX: 'PDF Prefixo',
    SORT_BY: 'CLASSIFICAR POR',
    SORT_ORDER: 'ORDEM',

    // Chaves Fixadores (V2.7)
    FIX_SECTION: 'Fixador: Coluna Seção',
    FIX_DESC: 'Fixador: Coluna Descrição',
    FIX_QTY: 'Fixador: Coluna Quantidade',
    FIX_UOM: 'Fixador: Coluna UOM',
    FIX_TRADE: 'Fixador: Coluna Trade (FIX)',
  },
  COLORS: {
    HEADER_BG: '#2c3e50',
    SECTION_BG: '#3498db',
    INPUT_BG: '#ecf0f1',
    FONT_LIGHT: '#ffffff',
    FONT_DARK: '#2c3e50',
    FONT_SUBTLE: '#7f8c8d',
    BORDER: '#bdc3c7',
    FONT_FAMILY: 'Inter',
    PANEL_EMPTY_BG: '#95a5a6',
    PANEL_ERROR_BG: '#c0392b'
  },
  CACHE_TTL: 180,
  DELIMITER: '|||',
  FIXADORES: {
    RISER: {
      interval: 10,
      clamps: {
        '1/2': 'RISER CLAMP 1 IN. METAL', '3/4': 'RISER CLAMP 1 IN. METAL', '1': 'RISER CLAMP 1 IN. METAL',
        '1-1/4': 'RISER CLAMP 1-1/4 IN. METAL', '1-1/2': 'RISER CLAMP 1-1/2 IN. METAL', '2': 'RISER CLAMP 2 IN. METAL',
        '2-1/2': 'RISER CLAMP 2 IN. METAL', '3': 'RISER CLAMP 3 IN. METAL', '4': 'RISER CLAMP 4 IN. METAL',
        '6': 'RISER CLAMP 6 IN. METAL', '8': 'RISER CLAMP 8 IN. METAL', '10': 'RISER CLAMP 12 IN. METAL', '12': 'RISER CLAMP 12 IN. METAL'
      },
      materials: [
        { desc: 'NUT 3/8 IN. METAL', factor: 4 }, { desc: 'FENDER WASHER 3/8 X 1-1/2', factor: 4 },
        { desc: 'ANCHOR DROP-IN 3/8 IN. X 3/4 IN. LONG HDI-P (W/ AUTO SET TOOL) [HILTI 409499]', factor: 2 },
        { desc: 'PLTD STEEL ALL THREAD ROD 3/8 IN. X 6 FT.', factor: 2 }
      ]
    },
    LOOP: {
      interval: 3,
      hangs: {
        '1/2': 'LOOP HANG 1/2 IN. HANGER', '3/4': 'LOOP HANG 3/4 IN. HANGER', '1': 'LOOP HANG 1 IN. METAL',
        '1-1/4': 'LOOP HANG 1-1/4 IN. METAL', '1-1/2': 'LOOP HANG 1-1/2 IN. HANGER', '2': 'LOOP HANG 2 IN. METAL',
        '2-1/2': 'LOOP HANG 2-1/2 IN. METAL', '3': 'LOOP HANG 3 IN. METAL', '4': 'LOOP HANG 4 IN. METAL',
        '6': 'LOOP HANG 6 IN. METAL', '8': 'LOOP HANG 6 IN. METAL', '10': 'LOOP HANG 12 IN. METAL', '12': 'LOOP HANG 12 IN. METAL'
      },
      materials: [
        { desc: 'NUT 3/8 IN. METAL', factor: 2 }, { desc: 'FENDER WASHER 3/8 X 1-1/2', factor: 2 },
        { desc: 'ANCHOR DROP-IN 3/8 IN. X 3/4 IN. LONG HDI-P (W/ AUTO SET TOOL) [HILTI 409499]', factor: 1 },
        { desc: 'PLTD STEEL ALL THREAD ROD 3/8 IN. X 6 FT.', factor: 1 }
      ]
    }
  }
};

// ============================================================================
// UTILITÁRIOS - BOM
// ============================================================================

const Utils = {
  getColumnIndex: (colConfig) => {
    if (!colConfig || typeof colConfig !== 'string') return -1;
    return colConfig.charAt(0).toUpperCase().charCodeAt(0) - 64;
  },
  getColumnHeader: (colConfig) => {
    if (!colConfig || typeof colConfig !== 'string') return '';
    const parts = colConfig.split(' - ');
    return parts.length > 1 ? parts[1].trim() : '';
  },
  formatVersion: (input) => {
    if (!input) return '';
    const str = String(input).trim();
    if (!str) return '';
    const num = parseInt(str, 10);
    return (isNaN(num) || num < 1) ? str : (num < 10 ? `0${num}` : `${num}`);
  },
  sanitizeSheetName: (name) => {
    return String(name).replace(/[\/\\\?\*\[\]:]/g, '_').substring(0, 100);
  },
  extractDiameter: (desc) => {
    const patterns = [
      /PIPE\s+(\d+\/\d+)\s+IN/i, /PIPE\s+(\d+)\s+IN/i, /PIPE\s+(\d+-\d+\/\d+)\s+IN/i
    ];
    for (const pattern of patterns) {
      const match = desc.match(pattern);
      if (match) {
        const diam = match[1].replace(/\s+/g, '');
        const mapping = {
          '1/2': '1/2', '3/4': '3/4', '1': '1', '1-1/4': '1-1/4',
          '1-1/2': '1-1/2', '2': '2', '2-1/2': '2-1/2', '3': '3',
          '4': '4', '6': '6', '8': '8', '10': '10', '12': '12'
        };
        return mapping[diam] || diam;
      }
    }
    return null;
  }
};

// ============================================================================
// CACHE - BOM
// ============================================================================

const CacheManager = {
  _cache: CacheService.getScriptCache(),
  get: (key) => {
    const cached = CacheManager._cache.get(key);
    return cached ? JSON.parse(cached) : null;
  },
  put: (key, value, ttl = CONFIG.CACHE_TTL) => {
    CacheManager._cache.put(key, JSON.stringify(value), ttl);
  },
  invalidateAll: () => {
    CacheManager._cache.removeAll([ 'all_config_values', 'unique_values' ]);
    const config = ConfigService.getAll();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
    if (sourceSheet) {
        const sourceSheetId = sourceSheet.getSheetId();
        const groupConfigs = [
            config[CONFIG.KEYS.GROUP_L1], config[CONFIG.KEYS.GROUP_L2], config[CONFIG.KEYS.GROUP_L3]
        ].filter(Boolean);
        const cacheKeysToRemove = groupConfigs.map(conf => {
            const colIndex = Utils.getColumnIndex(conf);
            if (colIndex !== -1) return `unique_${sourceSheetId}_${colIndex}`;
            return null;
        }).filter(Boolean);
        if (cacheKeysToRemove.length > 0) {
            CacheManager._cache.removeAll(cacheKeysToRemove);
        }
    }
  }
};

// ============================================================================
// CONFIGURAÇÃO - BOM
// ============================================================================

const ConfigService = {
  getAll: () => {
    const cached = CacheManager.get('all_config_values');
    if (cached) return cached;
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
    if (!configSheet) return {};
    const values = configSheet.getRange("A1:B" + configSheet.getLastRow()).getDisplayValues();
    const config = {};
    values.forEach(row => {
      const key = String(row[0]).trim();
      if (key && !key.startsWith(' ')) config[key] = row[1];
    });
    CacheManager.put('all_config_values', config);
    return config;
  },
  get: (key, defaultValue = '') => {
    return ConfigService.getAll()[key] || defaultValue;
  }
};

// ============================================================================
// CRIAÇÃO E ATUALIZAÇÃO DA ABA CONFIG (V2.7)
// ============================================================================

function ensureConfigExists() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG)) {
    forceCreateConfig();
  }
}

function forceCreateConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (configSheet) ss.deleteSheet(configSheet);

  configSheet = ss.insertSheet(CONFIG.SHEETS.CONFIG, 0);
  configSheet.clear();

  const K = CONFIG.KEYS;
  const configData = [
    ['CONFIGURAÇÃO', 'VALOR', 'DESCRIÇÃO'],
    ['', '', ''],
    [' 📊 AGRUPAMENTO (BOM)', '', 'Defina as colunas para agrupar os relatórios BOM'],
    [K.SOURCE_SHEET, '', 'Aba com dados brutos'],
    [K.GROUP_L1, '', 'Coluna principal de agrupamento'],
    [K.GROUP_L2, '', 'Segundo nível (opcional)'],
    [K.GROUP_L3, '', 'Terceiro nível (opcional)'],
    ['', '', ''],
    [' 📋 DADOS BOMS', '', 'Mapeamento das colunas para os relatórios BOM'],
    [K.COL_1, 'D - UNIT ID', 'ID ou identificador (ex: D - UNIT ID)'],
    [K.COL_2, 'J - DESC', 'Descrição do item (ex: J - DESC)'],
    [K.COL_3, 'M - UPC', 'Código de barras (ex: M - UPC)'],
    [K.COL_4, 'L - UOM', 'Unidade de medida (ex: L - UOM)'],
    [K.COL_5, 'O - PROJECT', 'Quantidade (ex: O - PROJECT)'],
    ['', '', ''],
    [' 🏷️ CABEÇALHO (BOM)', '', 'Informações do relatório BOM'],
    [K.PROJECT, 'HG1 BE', 'Nome do projeto'],
    [K.BOM, 'RISERS JS', 'Nome da BOM'],
    [K.KOJO_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo KOJO'],
    [K.ENGINEER, 'WANDERSON', 'Engenheiro'],
    [K.VERSION, '', 'Versão (ex: 01, 02)'],
    ['', '', ''],
    [' ⚙️ OPÇÕES (BOM)', '', 'Classificação dos dados do relatório BOM'],
    [K.SORT_BY, 'J - DESC', 'Coluna da Aba Origem para classificar'],
    [K.SORT_ORDER, 'Ascendente (A-Z, 0-9)', 'Ordem de classificação'],
    ['', '', ''],
    [' 💾 SALVAMENTO (BOM)', '', 'Exportação de PDFs dos relatórios BOM'],
    [K.DRIVE_FOLDER_ID, '', 'Link ou ID da pasta'],
    [K.DRIVE_FOLDER_NAME, '', 'Nome da pasta'],
    [K.PDF_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo dos PDFs'],
    ['', '', ''],
    [' 🔩 FIXADORES', '', 'Mapeamento de colunas para a ferramenta "Incluir Fixadores"'],
    [K.FIX_SECTION, 'B - SECTION', 'Coluna que contém "RISER" ou "COLGANTE"'],
    [K.FIX_DESC, 'J - DESC', 'Coluna com a descrição do tubo (para "PIPE" e diâmetro)'],
    [K.FIX_QTY, 'K - QTT', 'Coluna com a quantidade do tubo (ex: 100 FT)'],
    [K.FIX_UOM, 'L - UOM', 'Coluna com a unidade de medida do tubo (ex: FT)'],
    [K.FIX_TRADE, 'G - TRADE', 'Coluna onde o script escreverá "FIX"'],
  ];

  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);

  // Formatação
  const s = CONFIG.COLORS;
  configSheet.getRange("A1:C" + configData.length)
    .setFontFamily(s.FONT_FAMILY)
    .setVerticalAlignment('middle')
    .setFontColor(s.FONT_DARK);

  configSheet.setColumnWidth(1, 260).setColumnWidth(2, 300).setColumnWidth(3, 350);
  configSheet.getRange('A1:C1').merge()
    .setValue('⚙️ PAINEL DE CONFIGURAÇÃO | RELATÓRIOS DINÂMICOS')
    .setBackground(s.HEADER_BG)
    .setFontColor(s.FONT_LIGHT)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const sections = {
    3: { endRow: 7 },   // Agrupamento
    9: { endRow: 14 },  // Dados BOMS
    16: { endRow: 21 }, // Cabeçalho
    23: { endRow: 25 }, // Opções
    27: { endRow: 30 }, // Salvamento
    32: { endRow: 37 }  // Fixadores
  };

  for (const startRow in sections) {
    const start = parseInt(startRow);
    const endRow = sections[startRow].endRow;
    configSheet.getRange(start, 1, 1, 3).merge()
      .setBackground(s.SECTION_BG)
      .setFontColor(s.FONT_LIGHT)
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    configSheet.setRowHeight(start, 30);

    for (let r = start + 1; r <= endRow; r++) {
      if (configSheet.getRange(r, 1).getValue() === '') {
        configSheet.setRowHeight(r, 20);
        continue;
      }
      configSheet.getRange(r, 2).setBackground(s.INPUT_BG).setHorizontalAlignment('center');
      configSheet.getRange(r, 1).setFontWeight('500');
      configSheet.getRange(r, 3).setFontStyle('italic').setFontColor(s.FONT_SUBTLE);
      configSheet.setRowHeight(r, 28);
    }

    configSheet.getRange(start, 1, endRow - start + 1, 3)
      .setBorder(true, true, true, true, null, null, s.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  configSheet.getRange(21, 2).setNumberFormat('@STRING@'); // Linha da Versão
  configSheet.setFrozenRows(1);

  configSheet.getRange('E1:M1').merge()
    .setValue('PAINEL UNIFICADO DE AGRUPAMENTO (BOM)')
    .setBackground(s.HEADER_BG)
    .setFontColor(s.FONT_LIGHT)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontFamily(s.FONT_FAMILY);
  configSheet.setRowHeight(1, 40);

  SpreadsheetApp.flush();
  Utilities.sleep(100);
  updateConfigDropdowns();
}

function updateConfigDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!configSheet) return;
  const config = ConfigService.getAll();

  const sourceSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n !== CONFIG.SHEETS.CONFIG);
  if (sourceSheets.length > 0) {
    configSheet.getRange('B4').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(sourceSheets, true).setAllowInvalid(false).build()
    );
  }

  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet || sourceSheet.getLastColumn() === 0) return;

  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const columnOptions = headers.map((h, i) =>
    `${String.fromCharCode(65 + i)} - ${h || `Coluna ${String.fromCharCode(65 + i)}`}`
  );
  const columnRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(columnOptions, true).setAllowInvalid(false).build();

  const rowsToValidate = [
    5, 6, 7, // Agrupamento BOM
    10, 11, 12, 13, 14, // Dados BOM
    24, // Sort By BOM
    33, 34, 35, 36, 37 // Colunas Fixadores
  ];
  rowsToValidate.forEach(row =>
    configSheet.getRange(row, 2).setDataValidation(columnRule)
  );

  const sortOrderOptions = ['Ascendente (A-Z, 0-9)', 'Descendente (Z-A, 9-0)'];
  configSheet.getRange(25, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(sortOrderOptions, true).setAllowInvalid(false).build()
  );
}

// ============================================================================
// PAINÉIS DE AGRUPAMENTO (BOM - Aba Config)
// ============================================================================

function updateGroupingPanel() {
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  for (let level = 1; level <= 3; level++) {
    updateSingleLevelPanel(level, config, sourceSheet);
  }
}

function updateSingleLevelPanel(level, config, sourceSheet) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
  const startCol = 5 + (level - 1) * 2;
  const levelId = `NÍVEL ${level}`;
  const groupConfig = config[`Agrupar por Nível ${level}`];

  configSheet.getRange(3, startCol, configSheet.getMaxRows() - 2, 2).clear();
  configSheet.setColumnWidth(startCol, 150).setColumnWidth(startCol + 1, 50);

  const headerRange = configSheet.getRange(3, startCol, 1, 2).merge();
  if (sourceSheet && groupConfig && groupConfig.trim() !== '') {
    const colIndex = Utils.getColumnIndex(groupConfig);
    if (colIndex === -1) {
      headerRange.setValue(`${levelId}: ERRO`).setBackground(CONFIG.COLORS.PANEL_ERROR_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
      configSheet.getRange(4, startCol).setValue('Config. inválida.').setFontColor(CONFIG.COLORS.FONT_SUBTLE).setFontStyle('italic');
    } else {
      const colHeader = Utils.getColumnHeader(groupConfig);
      headerRange.setValue(`${levelId}: ${colHeader.toUpperCase()}`).setBackground(CONFIG.COLORS.SECTION_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
      const uniqueValues = getUniqueColumnValues(sourceSheet, colIndex);
      if (uniqueValues.length > 0) {
        const valuesFormatted = uniqueValues.map(v => [v ?? '']);
        configSheet.getRange(4, startCol, valuesFormatted.length, 1).setValues(valuesFormatted);
        configSheet.getRange(4, startCol + 1, valuesFormatted.length, 1).insertCheckboxes();
      }
    }
  } else {
    headerRange.setValue(levelId).setBackground(CONFIG.COLORS.PANEL_EMPTY_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
    const message = sourceSheet ? '(Não configurado)' : '(Selecione Aba Origem)';
    configSheet.getRange(4, startCol).setValue(message).setFontColor(CONFIG.COLORS.FONT_SUBTLE).setFontStyle('italic');
  }

  const lastRowInPanel = configSheet.getRange(configSheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastRowInPanel);
  configSheet.getRange(3, startCol, effectiveLastRow - 2, 2).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

function getUniqueColumnValues(sheet, columnIndex) {
  const cacheKey = `unique_${sheet.getSheetId()}_${columnIndex}`;
  const cached = CacheManager.get(cacheKey);
  if (cached) return cached;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues().flat();
  const uniqueValues = [...new Set(values.filter(v => v))]
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

  CacheManager.put(cacheKey, uniqueValues);
  return uniqueValues;
}

function getPanelSelections(sheet) {
  const selections = {};
  for (let i = 0; i < 3; i++) {
    const startCol = 5 + (i * 2);
    const headerRange = sheet.getRange(3, startCol);
    if (headerRange.isBlank()) continue;
    const levelId = headerRange.getValue().split(':')[0].trim();
    const lastDataRow = sheet.getRange(sheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    if (lastDataRow >= 4) {
      const numItems = lastDataRow - 3;
      const values = sheet.getRange(4, startCol, numItems, 1).getValues().flat();
      const checkboxes = sheet.getRange(4, startCol + 1, numItems, 1).getValues().flat();
      const selectedValues = values
        .map((v, index) => ({ value: v, checked: checkboxes[index] }))
        .filter(item => item.checked === true)
        .map(item => item.value ?? '');
      selections[levelId] = new Set(selectedValues);
    } else {
      selections[levelId] = new Set();
    }
  }
  return selections;
}

// ============================================================================
// PRÉ-VISUALIZAÇÃO (BOM - Aba Config)
// ============================================================================

function updatePreviewPanel(previewStartCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);

  const combinations = getSelectedCombinations();
  const existingSuffixes = new Map();
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  if (lastPreviewRow >= 4) {
    configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 3).getValues()
      .forEach(row => {
        if (row[0]) existingSuffixes.set(row[0], row[2]);
      });
  }

  configSheet.getRange(3, previewStartCol, configSheet.getMaxRows() - 2, 4).clear();
  configSheet.getRange(3, previewStartCol, 1, 3)
    .setValues([['COMBINAÇÃO GERADA', 'CRIAR?', 'SUFIXO KOJO FINAL']])
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold');

  if (combinations.length > 0) {
    const tableData = combinations.map(combo => {
      const displayCombo = combo.replace(/\|\|\|/g, '.');
      const kojoSuffix = existingSuffixes.get(combo) || displayCombo;
      return [displayCombo, true, kojoSuffix, combo];
    });
    configSheet.getRange(4, previewStartCol, tableData.length, 4).setValues(tableData);
    configSheet.getRange(4, previewStartCol + 1, tableData.length, 1).insertCheckboxes();
    configSheet.getRange(4, previewStartCol + 2, tableData.length, 1).setBackground(CONFIG.COLORS.INPUT_BG);
  }

  configSheet.setColumnWidth(previewStartCol, 250).setColumnWidth(previewStartCol + 1, 80).setColumnWidth(previewStartCol + 2, 250);
  configSheet.hideColumns(previewStartCol + 3);

  const lastCombinationRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastCombinationRow);
  configSheet.getRange(3, previewStartCol, effectiveLastRow - 2, 3).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  ss.toast('Pré-visualização atualizada!', 'Sucesso', 3);
}

function getSelectedCombinations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const config = ConfigService.getAll();
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return [];

  const groupConfigs = [
    config[CONFIG.KEYS.GROUP_L1], config[CONFIG.KEYS.GROUP_L2], config[CONFIG.KEYS.GROUP_L3]
  ].filter(Boolean);
  if (groupConfigs.length === 0) return [];

  const groupIndices = groupConfigs.map(Utils.getColumnIndex);
  const panelSelections = getPanelSelections(configSheet);
  const activeSelectionLevels = Object.keys(panelSelections).filter(level => panelSelections[level].size > 0);
  if (activeSelectionLevels.length === 0) return [];

  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  const existingCombinations = new Set();

  allData.forEach(row => {
    const combinationParts = groupIndices.map(index => row[index - 1] ?? '');
    if (combinationParts.every(part => String(part).trim() !== '')) {
      existingCombinations.add(combinationParts.join(CONFIG.DELIMITER));
    }
  });

  const finalCombinations = [...existingCombinations].filter(combo => {
    const parts = combo.split(CONFIG.DELIMITER);
    return parts.every((part, i) => {
      const levelId = `NÍVEL ${i + 1}`;
      if (!panelSelections[levelId]) return false;
      return Array.from(panelSelections[levelId]).some(selected => String(selected).trim() === String(part).trim());
    });
  });

  return finalCombinations.sort((a, b) => {
    const aParts = a.split(CONFIG.DELIMITER);
    const bParts = b.split(CONFIG.DELIMITER);
    for (let i = 0; i < Math.min(aParts.length, bParts.length); i++) {
      const aNum = parseFloat(aParts[i]);
      const bNum = parseFloat(bParts[i]);
      if (!isNaN(aNum) && !isNaN(bNum)) {
        if (aNum !== bNum) return aNum - bNum;
      } else {
         const comparison = aParts[i].localeCompare(bParts[i], undefined, { numeric: true });
        if (comparison !== 0) return comparison;
      }
    }
    return aParts.length - bParts.length;
  });
}

// ============================================================================
// PROCESSAMENTO DE RELATÓRIOS (BOM)
// ============================================================================

/**
 * (V2.8) Esta é a função "antiga" que lê da aba 'Config'.
 */
function runProcessing() {
  CacheManager.invalidateAll();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll(); // Lê da aba 'Config'
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);

  const previewStartCol = 11;
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (lastPreviewRow < 4) {
    return { success: false, message: 'Nenhuma combinação na pré-visualização' };
  }

  const previewData = configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 4).getValues();
  const combinationsToProcess = previewData
    .filter(row => row[1] === true)
    .map(row => ({
      displayCombination: String(row[0]),
      combination: String(row[3]),
      kojoSuffix: row[2]
    }));

  if (combinationsToProcess.length === 0) {
    return { success: false, message: 'Nenhum relatório selecionado' };
  }

  return processBomCore(combinationsToProcess, config);
}

/**
 * (V2.10) Processa relatórios com base em dados do HTML.
 */
function runProcessingFromHtml(selections, reportSettings) {
  CacheManager.invalidateAll();
  const baseConfig = ConfigService.getAll();
  const settings = { ...baseConfig, ...reportSettings };

  if (!selections || selections.length === 0) {
    return { success: false, message: 'Nenhuma combinação recebida do painel' };
  }

  return processBomCore(selections, settings);
}


/**
 * ✅ ATUALIZADO (V2.12): Nomeia a aba pelo kojoSuffix
 */
function processBomCore(combinationsToProcess, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const K = CONFIG.KEYS;

  const sourceSheet = ss.getSheetByName(settings[K.SOURCE_SHEET]);
  if (!sourceSheet) {
    return { success: false, message: `Aba de origem "${settings[K.SOURCE_SHEET]}" não encontrada` };
  }

  const groupConfigs = [
    settings[K.GROUP_L1], settings[K.GROUP_L2], settings[K.GROUP_L3]
  ].filter(Boolean);
  const groupIndices = groupConfigs.map(Utils.getColumnIndex);

  const bomCols = {
    c1: Utils.getColumnIndex(settings[K.COL_1]),
    c2: Utils.getColumnIndex(settings[K.COL_2]),
    c3: Utils.getColumnIndex(settings[K.COL_3]),
    c4: Utils.getColumnIndex(settings[K.COL_4]),
    c5: Utils.getColumnIndex(settings[K.COL_5])
  };

  const sortColumnConfig = settings[K.SORT_BY] || settings[K.COL_2];
  const sortColumnIndex = Utils.getColumnIndex(sortColumnConfig) - 1;
  const sortOrder = settings[K.SORT_ORDER] === 'Descendente (Z-A, 9-0)' ? 'desc' : 'asc';

  if(sortColumnIndex < 0) {
    return { success: false, message: `Coluna de classificação "${sortColumnConfig}" inválida.` };
  }

  const dataMap = new Map();
  combinationsToProcess.forEach(item => dataMap.set(item.combination, []));
  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();

  for (const row of allData) {
    const rowCombination = groupIndices
      .map(index => String(row[index - 1] ?? '').trim())
      .join(CONFIG.DELIMITER);

    if (dataMap.has(rowCombination)) {
      dataMap.get(rowCombination).push([
        row[bomCols.c1 - 1], row[bomCols.c2 - 1], row[bomCols.c3 - 1],
        row[bomCols.c4 - 1], parseFloat(row[bomCols.c5 - 1]) || 0,
        row[sortColumnIndex] // Valor para classificação
      ]);
    }
  }

  let createdCount = 0;
  combinationsToProcess.forEach(item => {
    const { combination, kojoSuffix } = item;
    const rawData = dataMap.get(combination);
    if (!rawData || rawData.length === 0) return;

    const processedData = groupAndSumData(rawData, sortOrder); // Corrigido

    // ✅ MUDANÇA (V2.12): Nomeia a aba usando o kojoSuffix
    const sanitizedName = Utils.sanitizeSheetName(kojoSuffix);

    let targetSheet = ss.getSheetByName(sanitizedName);
    if (targetSheet) {
      targetSheet.clear();
    } else {
      targetSheet = ss.insertSheet(sanitizedName);
    }

    createAndFormatReport(targetSheet, kojoSuffix, processedData, settings);
    createdCount++;
  });

  return { success: true, created: createdCount };
}


function groupAndSumData(data, sortOrder) {
  const grouped = {};

  data.forEach(row => {
    const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
    if (grouped[key]) {
      grouped[key][4] += row[4];
    } else {
      grouped[key] = [...row];
    }
  });

  const groupedData = Object.values(grouped);
  const direction = sortOrder === 'asc' ? 1 : -1;

  groupedData.sort((a, b) => {
    const valA = a[5]; // O valor de ordenação está sempre no índice 5
    const valB = b[5];
    if (valA === null || valA === undefined || valA === '') return 1;
    if (valB === null || valB === undefined || valB === '') return -1;
    const aIsNum = !isNaN(parseFloat(valA)) && isFinite(valA);
    const bIsNum = !isNaN(parseFloat(valB)) && isFinite(valB);
    if (aIsNum && bIsNum) {
      return (parseFloat(valA) - parseFloat(valB)) * direction;
    }
    return String(valA).localeCompare(String(valB), undefined, { numeric: true }) * direction;
  });

  return groupedData.map(row => row.slice(0, 5));
}


function createAndFormatReport(sheet, kojoSuffix, data, settings) {
  const K = CONFIG.KEYS;
  const reportConfig = {
    project: settings[K.PROJECT],
    bom: settings[K.BOM],
    kojoPrefix: settings[K.KOJO_PREFIX],
    engineer: settings[K.ENGINEER],
    version: settings[K.VERSION]
  };
  const headers = {
    h1: Utils.getColumnHeader(settings[K.COL_1]),
    h2: Utils.getColumnHeader(settings[K.COL_2]),
    h3: Utils.getColumnHeader(settings[K.COL_3]),
    h4: Utils.getColumnHeader(settings[K.COL_4]),
    h5: 'QTY'
  };

  const lastUpdate = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd/yyyy');
  const bomKojoComplete = `${reportConfig.kojoPrefix}.${kojoSuffix}`;

  const headerValues = [
    ['PROJECT:', reportConfig.project], ['BOM:', reportConfig.bom], ['BOM KOJO:', bomKojoComplete],
    ['ENG.:', reportConfig.engineer], ['VERSION:', reportConfig.version], ['LAST UPDATE:', lastUpdate]
  ];

  sheet.getRange(1, 1, headerValues.length, 2).setValues(headerValues);
  sheet.getRange(1, 2, headerValues.length, 1).setNumberFormat('@STRING@');
  for (let r = 1; r <= headerValues.length; r++) {
    sheet.getRange(r, 2, 1, 4).merge();
  }
  sheet.getRange(1, 1, headerValues.length, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  const dataStartRow = headerValues.length + 2;
  const finalData = [[headers.h1, headers.h2, headers.h3, headers.h4, headers.h5]].concat(data);
  sheet.getRange(dataStartRow, 1, finalData.length, 5).setValues(finalData);
  sheet.getRange(dataStartRow, 1, finalData.length, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  sheet.getRange(dataStartRow, 1, 1, 5).setFontWeight('bold');

  sheet.setColumnWidth(1, 105).setColumnWidth(2, 570).setColumnWidth(3, 105).setColumnWidth(4, 105).setColumnWidth(5, 105);

  try {
    const protection = sheet.getRange(1, 1, dataStartRow - 1, 5).protect();
    protection.setDescription('Cabeçalho protegido').removeEditors(protection.getEditors());
  } catch (e) {
    Logger.log(`Aviso: ${e.message}`);
  }
}

function clearOldReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmação', 'Apagar TODAS as abas de relatório?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll();
  const protectedSheets = [CONFIG.SHEETS.CONFIG, config[CONFIG.KEYS.SOURCE_SHEET]].filter(Boolean);
  let deletedCount = 0;
  ss.getSheets().forEach(sheet => {
    if (!protectedSheets.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });
  ui.alert('Limpeza Concluída', `${deletedCount} abas removidas.`, ui.ButtonSet.OK);
}

// ============================================================================
// FIXADORES (SEÇÃO V2.7 - CÓPIA INTELIGENTE)
// ============================================================================

function abrirSeletorFixadores() {
  const html = HtmlService.createHtmlOutputFromFile('FixadoresSidebar.html')
    .setTitle('Seletor de Fixadores')
    .setWidth(900).setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Seletor de Fixadores');
}

function getPipesElegiveis() {
  const config = ConfigService.getAll();
  const K = CONFIG.KEYS;
  const fixIdx = {
    section: Utils.getColumnIndex(config[K.FIX_SECTION]) - 1,
    desc: Utils.getColumnIndex(config[K.FIX_DESC]) - 1,
    qty: Utils.getColumnIndex(config[K.FIX_QTY]) - 1,
    uom: Utils.getColumnIndex(config[K.FIX_UOM]) - 1,
    trade: Utils.getColumnIndex(config[K.FIX_TRADE]) - 1
  };

  if ([fixIdx.section, fixIdx.desc, fixIdx.qty, fixIdx.uom, fixIdx.trade].some(idx => idx < 0)) {
    Logger.log('Erro de Configuração de Fixadores: Pelo menos uma coluna-chave não está definida.');
    return [];
  }

  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[K.SOURCE_SHEET]);
  if (!sourceSheet) return [];
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  const pipes = [];

  data.forEach((row, idx) => {
    const section = String(row[fixIdx.section] || '');
    const desc = String(row[fixIdx.desc] || '');
    const qty = parseFloat(row[fixIdx.qty]) || 0;
    const uom = String(row[fixIdx.uom] || '');

    if (validarTipoFixacao(section) && desc.toUpperCase().includes('PIPE') && qty > 0) {
      const diameter = Utils.extractDiameter(desc);
      const isRiser = section.toUpperCase().includes('RISER');
      const fixConfig = isRiser ? CONFIG.FIXADORES.RISER : CONFIG.FIXADORES.LOOP;
      const itemMap = isRiser ? fixConfig.clamps : fixConfig.hangs;

      if (diameter && itemMap[diameter]) {
        let jaTemFixador = false;
        if (idx + 1 < data.length) {
          const nextRow = data[idx + 1];
          const nextRowTrade = String(nextRow[fixIdx.trade] || '').toUpperCase();
          if (nextRowTrade === 'FIX') jaTemFixador = true;
        }

        pipes.push({
          rowIndex: idx + 2,
          section: section, desc: desc, qty: qty, uom: uom,
          diameter: diameter, isRiser: isRiser, jaTemFixador: jaTemFixador,
          originalRow: [...row],
          trade: String(row[fixIdx.trade] || ''),
          floor: String(row[8] || ''),
          unitType: String(row[4] || ''),
          phase: String(row[7])
        });
      }
    }
  });
  return pipes;
}

function validarTipoFixacao(section) {
  const s = String(section).toUpperCase();
  return s.includes('RISER') || s.includes('COLGANTE');
}

function processarFixadoresSelecionados(selectedPipes) {
  const config = ConfigService.getAll();
  const K = CONFIG.KEYS;
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[K.SOURCE_SHEET]);
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }

  const fixIdx = {
    desc: Utils.getColumnIndex(config[K.FIX_DESC]) - 1,
    qty: Utils.getColumnIndex(config[K.FIX_QTY]) - 1,
    trade: Utils.getColumnIndex(config[K.FIX_TRADE]) - 1
  };

  if (fixIdx.desc < 0 || fixIdx.qty < 0 || fixIdx.trade < 0) {
     return { success: false, message: 'Configuração de colunas "Fixadores" inválida na aba "Config".' };
  }

  selectedPipes.sort((a, b) => b.rowIndex - a.rowIndex);
  let totalAdded = 0;
  const maxCol = sourceSheet.getLastColumn();

  selectedPipes.forEach(pipe => {
    const fixConfig = pipe.isRiser ? CONFIG.FIXADORES.RISER : CONFIG.FIXADORES.LOOP;
    const itemMap = pipe.isRiser ? fixConfig.clamps : fixConfig.hangs;
    const fixadorItem = itemMap[pipe.diameter];
    if (!fixadorItem) return;

    const originalFormulas = sourceSheet.getRange(pipe.rowIndex, 1, 1, maxCol).getFormulasR1C1()[0];
    const linhasParaInserir = [];
    const insertRow = pipe.rowIndex + 1;

    // Linha do fixador
    const linhaFixador = [...pipe.originalRow];
    linhaFixador[fixIdx.trade] = 'FIX';
    linhaFixador[fixIdx.desc] = fixadorItem;
    linhaFixador[fixIdx.qty] = `=ROUNDUP(R[-1]C/${fixConfig.interval})`;
    linhasParaInserir.push(linhaFixador);
    const fixadorRow = insertRow;

    // Materiais
    fixConfig.materials.forEach(mat => {
      const linhaMat = [...pipe.originalRow];
      linhaMat[fixIdx.trade] = 'FIX';
      linhaMat[fixIdx.desc] = mat.desc;
      linhaMat[fixIdx.qty] = `=R${fixadorRow}C*${mat.factor}`;
      linhasParaInserir.push(linhaMat);
    });

    sourceSheet.insertRowsAfter(pipe.rowIndex, linhasParaInserir.length);
    const formatoOrigem = sourceSheet.getRange(pipe.rowIndex, 1, 1, maxCol);
    formatoOrigem.copyFormatToRange(sourceSheet, 1, maxCol, insertRow, insertRow + linhasParaInserir.length - 1);
    const rangeDestino = sourceSheet.getRange(insertRow, 1, linhasParaInserir.length, maxCol);
    rangeDestino.setValues(linhasParaInserir);

    // Restaura fórmulas
    linhasParaInserir.forEach((row, idx) => {
      const currentRow = insertRow + idx;
      for (let col = 0; col < maxCol; col++) {
        if (originalFormulas[col] && col !== fixIdx.trade && col !== fixIdx.desc && col !== fixIdx.qty) {
          sourceSheet.getRange(currentRow, col + 1).setFormulaR1C1(originalFormulas[col]);
        }
      }
      sourceSheet.getRange(currentRow, fixIdx.qty + 1).setFormulaR1C1(row[fixIdx.qty]);
    });
    totalAdded += linhasParaInserir.length;
  });

  return { success: true, added: totalAdded };
}

function removerFixadoresSelecionados(selectedPipes) {
  const config = ConfigService.getAll();
  const K = CONFIG.KEYS;
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[K.SOURCE_SHEET]);
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }

  const tradeColIndex = Utils.getColumnIndex(config[K.FIX_TRADE]) - 1;
  if (tradeColIndex < 0) {
    return { success: false, message: 'Configuração de "Fixador: Coluna Trade" inválida.' };
  }

  selectedPipes.sort((a, b) => b.rowIndex - a.rowIndex);
  let totalRemoved = 0;
  const allData = sourceSheet.getDataRange().getValues();

  try {
    selectedPipes.forEach(pipe => {
      const rowIndex = pipe.rowIndex;
      let rowsToDelete = 0;
      for (let i = rowIndex; i < allData.length; i++) {
        const rowData = allData[i];
        const trade = String(rowData[tradeColIndex] || '').toUpperCase();
        if (trade === 'FIX') rowsToDelete++;
        else break;
      }
      if (rowsToDelete > 0) {
        sourceSheet.deleteRows(rowIndex + 1, rowsToDelete);
        totalRemoved += rowsToDelete;
      }
    });
    return { success: true, removed: totalRemoved };
  } catch (e) {
    Logger.log(`Erro ao remover fixadores: ${e.message}`);
    return { success: false, message: `Erro ao apagar linhas: ${e.message}` };
  }
}


// ============================================================================
// EXPORTAÇÃO PDF (V2.12)
// ============================================================================

function runPdfExportFromHtml(sheetNames, folderInput, prefix) {
  if (!sheetNames || sheetNames.length === 0) {
    return { success: false, message: 'Nenhuma aba selecionada' };
  }
  const folder = getFolderFromInput(folderInput, folderInput); // Usa o input como ID e fallback de nome
  if (!folder) return { success: false, message: `Pasta não encontrada ou inválida: ${folderInput}` };

  sheetNames.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      // Lê o nome da BOM KOJO da célula B3 (onde está "BOM KOJO:")
      const bomKojoName = getBomKojoNameFromSheet(sheet);
      const fileName = bomKojoName || sheetName; // Fallback para o nome da sheet se não encontrar
      exportSheetToPdf(sheet, fileName, folder);
    }
  });
  return { success: true, exported: sheetNames.length, folder: folder.getName() };
}

function exportPDFsWithFeedback() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Exportando PDFs...', 'Aguarde', -1);
  const config = ConfigService.getAll();
  const K = CONFIG.KEYS;

  const sheetNames = getReportSheetNames();
  if (!sheetNames || sheetNames.length === 0) {
     SpreadsheetApp.getUi().alert('Erro', 'Nenhum relatório gerado para exportar.', SpreadsheetApp.getUi().ButtonSet.OK);
     return;
  }

  const folder = getFolderFromInput(config[K.DRIVE_FOLDER_ID], config[K.DRIVE_FOLDER_NAME]);
  if (!folder) {
    SpreadsheetApp.getUi().alert('Erro', 'Pasta de destino não configurada ou inválida na aba "Config".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  sheetNames.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      // Lê o nome da BOM KOJO da célula B3 (onde está "BOM KOJO:")
      const bomKojoName = getBomKojoNameFromSheet(sheet);
      const fileName = bomKojoName || sheetName; // Fallback para o nome da sheet se não encontrar
      exportSheetToPdf(sheet, fileName, folder);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ ${sheetNames.length} PDFs exportados para "${folder.getName()}"!`,
      'Sucesso', 5
    );
}

/**
 * Extrai o nome da BOM KOJO da célula B3 de uma sheet de relatório.
 * Retorna null se não encontrar o valor.
 */
function getBomKojoNameFromSheet(sheet) {
  try {
    if (!sheet) return null;
    // O valor da BOM KOJO está na célula B3 (linha 3, coluna 2)
    const bomKojoValue = sheet.getRange(3, 2).getValue();
    if (bomKojoValue && String(bomKojoValue).trim() !== '') {
      return String(bomKojoValue).trim();
    }
  } catch (error) {
    Logger.log(`Erro ao ler BOM KOJO da sheet "${sheet.getName()}": ${error.message}`);
  }
  return null;
}

function getFolderFromInput(folderInput, folderName) {
  try {
    if (folderInput) {
      try {
        const folderById = DriveApp.getFolderById(folderInput);
        if (folderById) return folderById;
      } catch (e) { /* Não é um ID */ }

      const match = folderInput.match(/folders\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
         try {
           const folderById = DriveApp.getFolderById(match[1]);
           if (folderById) return folderById;
         } catch(e) { /* Link inválido */ }
      }
    }

    const nameToSearch = folderInput || folderName;
    if (nameToSearch) {
      const folders = DriveApp.getFoldersByName(nameToSearch);
      if (folders.hasNext()) return folders.next();
      return DriveApp.createFolder(nameToSearch);
    }

    const defaultName = `${SpreadsheetApp.getActiveSpreadsheet().getName()} - PDFs`;
    const folders = DriveApp.getFoldersByName(defaultName);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(defaultName);

  } catch (error) {
    Logger.log(`Erro ao acessar pasta: ${error.message}`);
    return null;
  }
}

function exportSheetToPdf(sheet, pdfName, folder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    `gid=${sheet.getSheetId()}&format=pdf&size=A4&portrait=true&fitw=true&` +
    `sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
  for (let attempt = 0, delay = 1000; attempt < 5; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200) {
        folder.createFile(response.getBlob().setName(`${pdfName}.pdf`));
        return;
      }
      if (response.getResponseCode() === 429) {
        Utilities.sleep(delay + Math.floor(Math.random() * 500));
        delay *= 2;
        continue;
      }
      throw new Error(`Código HTTP ${response.getResponseCode()}`);
    } catch (error) {
      if (attempt === 4) throw error;
      Utilities.sleep(delay);
    }
  }
}

// ============================================================================
// (V2.8) Backend para Painel BOM
// ============================================================================

/**
 * ✅ ATUALIZADO (V2.12): Envia o estado salvo para o HTML
 */
function getBomHtmlInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll();

  const allSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n !== CONFIG.SHEETS.CONFIG);

  let allColumns = [];
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (sourceSheet && sourceSheet.getLastColumn() > 0) {
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    allColumns = headers.map((h, i) =>
      `${String.fromCharCode(65 + i)} - ${h || `Coluna ${String.fromCharCode(65 + i)}`}`
    );
  }

  // Pega o estado salvo
  const savedState = getUserSidebarState();

  return {
    allSheets: allSheets,
    allColumns: allColumns,
    currentConfig: config, // Configurações padrão
    savedState: savedState // Configurações salvas do usuário
  };
}

function getUniqueValuesForColumn(colName) {
  if (!colName) return [];
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return [];
  const colIndex = Utils.getColumnIndex(colName);
  if (colIndex === -1) return [];
  return getUniqueColumnValues(sourceSheet, colIndex);
}

function getCombinationsForPreview(selectedGroups, groupConfigs) {
  const config = ConfigService.getAll();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return [];

  const activeGroupConfigs = groupConfigs.filter(Boolean);
  if (activeGroupConfigs.length === 0) return [];

  const groupIndices = activeGroupConfigs.map(Utils.getColumnIndex);

  const panelSelections = {};
  for (const level in selectedGroups) {
    panelSelections[level] = new Set(selectedGroups[level]);
  }

  const activeSelectionLevels = Object.keys(panelSelections).filter(level => panelSelections[level].size > 0);
  if (activeSelectionLevels.length === 0) return [];

  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  const existingCombinations = new Set();

  allData.forEach(row => {
    const combinationParts = groupIndices.map(index => row[index - 1] ?? '');
    if (combinationParts.every(part => String(part).trim() !== '')) {
      existingCombinations.add(combinationParts.join(CONFIG.DELIMITER));
    }
  });

  const finalCombinations = [...existingCombinations].filter(combo => {
    const parts = combo.split(CONFIG.DELIMITER);
    return parts.every((part, i) => {
      const levelId = `NÍVEL ${i + 1}`;
      if (!panelSelections[levelId]) return false;
      return Array.from(panelSelections[levelId]).some(selected => String(selected).trim() === String(part).trim());
    });
  });

  return finalCombinations.sort((a, b) => {
    const aParts = a.split(CONFIG.DELIMITER);
    const bParts = b.split(CONFIG.DELIMITER);
    for (let i = 0; i < Math.min(aParts.length, bParts.length); i++) {
      const aNum = parseFloat(aParts[i]);
      const bNum = parseFloat(bParts[i]);
      if (!isNaN(aNum) && !isNaN(bNum)) {
        if (aNum !== bNum) return aNum - bNum;
      } else {
         const comparison = aParts[i].localeCompare(bParts[i], undefined, { numeric: true });
        if (comparison !== 0) return comparison;
      }
    }
    return aParts.length - bParts.length;
  });
}

// (V2.12) Salvar e Carregar Estado
const STATE_PROPERTY_KEY = 'BOM_SIDEBAR_STATE_V2_12';

function saveUserSidebarState(stateJson) {
  try {
    PropertiesService.getUserProperties().setProperty(STATE_PROPERTY_KEY, stateJson);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getUserSidebarState() {
  try {
    const state = PropertiesService.getUserProperties().getProperty(STATE_PROPERTY_KEY);
    return state ? JSON.parse(state) : null; // Des-serializa
  } catch (e) {
    return null;
  }
}

// ============================================================================
// TRIGGERS E UI - BOM
// ============================================================================

/**
 * ✅ CORRIGIDO (V2.11): Adicionado setWidth e setHeight
 */
function openBomSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('BomSidebar.html')
    .setTitle('Painel Gerador de BOM')
    .setWidth(1100) // ✅ Tamanho grande
    .setHeight(750);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Painel Gerador de BOM');
}

function openConfigSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigSidebar.html')
    .setTitle('Painel de Controle')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Funções de Feedback (Wrapper)
function runProcessingWithFeedback() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Processando (via Aba Config)...', 'Aguarde', -1);
  const result = runProcessing();
  if (result.success) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${result.created} relatórios criados!`, 'Sucesso', 5);
  } else {
    SpreadsheetApp.getUi().alert('Erro (Aba Config)', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  return result;
}

function testSystem() {
  try {
    const config = ConfigService.getAll();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
    const msg = [
      `Versão do Script: 2.13 (Correção Exportar PDF)`, // Atualizado
      `Aba de Origem: ${config[CONFIG.KEYS.SOURCE_SHEET] || 'Não configurada'}`,
      `Linhas na Origem: ${sourceSheet ? sourceSheet.getLastRow() - 1 : 0}`,
      `Relatórios Gerados: ${getReportSheetNames().length}`,
      `Fixador Trade: ${config[CONFIG.KEYS.FIX_TRADE] || 'Não configurado'}`
    ].join('\n');
    SpreadsheetApp.getUi().alert('🧪 Diagnóstico do Sistema', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    return {
      "Modo de Filtro": "Classificação Avançada",
      "Tipologias Selecionadas": sourceSheet ? sourceSheet.getLastRow() - 1 : 0
    };
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro no Diagnóstico', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ✅ CORREÇÃO (V2.13): Esta função estava em falta na V2.12
 */
function getReportSheetNames() {
  const config = ConfigService.getAll();
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(s => s.getName())
    .filter(name => name !== CONFIG.SHEETS.CONFIG && name !== config[CONFIG.KEYS.SOURCE_SHEET]);
}

/**
 * (V2.12) Wrapper para a função `getReportSheetNames` para o HTML.
 */
function getReportSheetNamesForHtml() {
  return getReportSheetNames();
}

function forceRefreshCache() {
  CacheManager.invalidateAll();
  SpreadsheetApp.getActiveSpreadsheet().toast('Cache limpo!', 'Sucesso', 3);
}

// Trigger onEdit para BOM
function onEditBom(e) {
  if (!e || !e.source) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName === CONFIG.SHEETS.CONFIG) {
    CacheManager.invalidateAll();
    const range = e.range;
    const col = range.getColumn();
    const row = range.getRow();

    if (col === 2 && row === 21) { // Versão
      const formatted = Utils.formatVersion(range.getValue());
      if (formatted !== String(range.getValue())) range.setValue(formatted);
    }

    if (col === 2 && row >= 4 && row <= 7) { // Painéis de Agrupamento
      Utilities.sleep(200);
      const config = ConfigService.getAll();
      const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
      if (row === 4) {
        updateConfigDropdowns();
        updateGroupingPanel();
      } else {
        updateSingleLevelPanel(row - 4, config, sourceSheet);
      }
      updatePreviewPanel(11);
    }

    if ((col === 6 || col === 8 || col === 10) && row >= 4) { // Checkboxes
      updatePreviewPanel(11);
    }
  } else {
    try {
      const config = ConfigService.getAll();
      if (sheetName === config[CONFIG.KEYS.SOURCE_SHEET]) {
        CacheManager.invalidateAll();
      }
    } catch (error) {}
  }
}
