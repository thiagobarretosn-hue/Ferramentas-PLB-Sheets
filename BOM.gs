/**
 * @fileoverview SISTEMA UNIFICADO DE RELATÓRIOS DINÂMICOS + FIXADORES (BOM)
 * @version 3.1.0 - Correções: prefixo PDF, versão zero-pad, cache registry, empty-sheet guards
 *
 * V3.0: Cada ferramenta opera de forma independente:
 * - BomSidebar gerencia suas próprias configurações via PropertiesService
 * - FixadoresSidebar gerencia mapeamento de colunas via PropertiesService
 * - Todas as funções usam a aba ativa como fonte de dados
 * - Não existe mais dependência de uma aba "Config"
 */

// ============================================================================
// CONFIGURAÇÃO GLOBAL - BOM
// ============================================================================

const BOM_CONFIG = {
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
    PDF_BLOCKS: 'PDF Blocos Nome',
    SORT_BY: 'CLASSIFICAR POR',
    SORT_ORDER: 'ORDEM',

    // Chaves Fixadores
    FIX_SECTION: 'Fixador: Coluna Seção',
    FIX_DESC: 'Fixador: Coluna Descrição',
    FIX_QTY: 'Fixador: Coluna Quantidade',
    FIX_UOM: 'Fixador: Coluna UOM',
    FIX_TRADE: 'Fixador: Coluna Trade (FIX)',
  },
  // Valores padrão para quando nenhuma configuração existe
  DEFAULTS: {
    'Coluna 1': 'D - UNIT ID',
    'Coluna 2': 'J - DESC',
    'Coluna 3': 'M - UPC',
    'Coluna 4': 'L - UOM',
    'Coluna 5': 'O - PROJECT',
    'CLASSIFICAR POR': 'J - DESC',
    'ORDEM': 'Ascendente (A-Z, 0-9)',
    'Fixador: Coluna Seção': 'B - SECTION',
    'Fixador: Coluna Descrição': 'K - DESC',
    'Fixador: Coluna Quantidade': 'L - QTT',
    'Fixador: Coluna UOM': 'M - UOM',
    'Fixador: Coluna Trade (FIX)': 'G - TRADE',
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
// Usa funções centralizadas de lib/Shared/Utils.gs
// ============================================================================

const Utils = {
  /**
   * Extrai índice de coluna de config "A - Nome" ou "AA - Nome"
   * Usa SharedUtils para suporte a colunas AA, AB, etc.
   */
  getColumnIndex: (colConfig) => {
    return SharedUtils_getColumnIndexFromConfig(colConfig);
  },

  /**
   * Extrai nome do cabeçalho de config "A - Nome"
   */
  getColumnHeader: (colConfig) => {
    return SharedUtils_getColumnHeaderFromConfig(colConfig);
  },

  /**
   * Formata versão com zeros à esquerda: "1" → "01"
   */
  formatVersion: (input) => {
    return SharedUtils_formatVersion(input);
  },

  /**
   * Remove caracteres inválidos de nome de aba
   */
  sanitizeSheetName: (name) => {
    return SharedUtils_sanitizeSheetName(name);
  },

  /**
   * Extrai diâmetro de descrição de tubo (PIPE X IN)
   */
  extractDiameter: (desc) => {
    return SharedUtils_extractPipeDiameter(desc);
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
  put: (key, value, ttl = BOM_CONFIG.CACHE_TTL) => {
    CacheManager._cache.put(key, JSON.stringify(value), ttl);
    // Registra a chave para que invalidateAll possa removê-la
    try {
      const regJson = CacheManager._cache.get('bom_cache_keys');
      const keys = regJson ? JSON.parse(regJson) : [];
      if (!keys.includes(key)) {
        keys.push(key);
        CacheManager._cache.put('bom_cache_keys', JSON.stringify(keys), 3600);
      }
    } catch (e) {}
  },
  invalidateAll: () => {
    try {
      const regJson = CacheManager._cache.get('bom_cache_keys');
      const keys = regJson ? JSON.parse(regJson) : [];
      if (keys.length) CacheManager._cache.removeAll(keys);
      CacheManager._cache.remove('bom_cache_keys');
    } catch (e) {}
  }
};

// ============================================================================
// CONFIGURAÇÃO - BOM (V3.0 - PropertiesService)
// Configurações salvas via PropertiesService, sem dependência de aba Config
// ============================================================================

const BOM_SETTINGS_KEY = 'BOM_SETTINGS_V3';
const FIX_SETTINGS_KEY = 'FIX_SETTINGS_V3';

const ConfigService = {
  /**
   * Retorna todas as configurações BOM do PropertiesService
   * Se não existirem, retorna os defaults
   */
  getAll: () => {
    try {
      const saved = PropertiesService.getDocumentProperties().getProperty(BOM_SETTINGS_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        // Merge com defaults para garantir que novas chaves existam
        return { ...BOM_CONFIG.DEFAULTS, ...parsed };
      }
    } catch (e) {
      console.error('[ConfigService] Erro ao ler configurações:', e.message);
    }
    return { ...BOM_CONFIG.DEFAULTS };
  },

  /**
   * Obtém uma configuração específica
   */
  get: (key, defaultValue = '') => {
    const all = ConfigService.getAll();
    return all[key] || defaultValue;
  },

  /**
   * Salva todas as configurações BOM
   */
  saveAll: (config) => {
    try {
      PropertiesService.getDocumentProperties().setProperty(BOM_SETTINGS_KEY, JSON.stringify(config));
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  },

  /**
   * Salva configurações de fixadores
   */
  getFixConfig: () => {
    try {
      const saved = PropertiesService.getDocumentProperties().getProperty(FIX_SETTINGS_KEY);
      if (saved) return JSON.parse(saved);
    } catch (e) { /* ignora */ }
    // Defaults de fixadores
    const K = BOM_CONFIG.KEYS;
    return {
      [K.FIX_SECTION]: BOM_CONFIG.DEFAULTS[K.FIX_SECTION],
      [K.FIX_DESC]: BOM_CONFIG.DEFAULTS[K.FIX_DESC],
      [K.FIX_QTY]: BOM_CONFIG.DEFAULTS[K.FIX_QTY],
      [K.FIX_UOM]: BOM_CONFIG.DEFAULTS[K.FIX_UOM],
      [K.FIX_TRADE]: BOM_CONFIG.DEFAULTS[K.FIX_TRADE],
    };
  },

  saveFixConfig: (config) => {
    try {
      PropertiesService.getDocumentProperties().setProperty(FIX_SETTINGS_KEY, JSON.stringify(config));
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }
};

// ============================================================================
// FUNÇÕES DE SUPORTE - BOM
// ============================================================================

/**
 * Retorna valores únicos de uma coluna para uso em filtros e painéis
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba de dados
 * @param {number} columnIndex - Índice da coluna (1-indexed)
 * @returns {Array} Valores únicos ordenados
 */
function getUniqueColumnValues(sheet, columnIndex) {
  if (!sheet) return [];
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

// ============================================================================
// PROCESSAMENTO DE RELATÓRIOS (BOM)
// ============================================================================

/**
 * Processa relatórios BOM com configurações passadas pela sidebar HTML
 * Esta é a ÚNICA função de processamento - recebe TUDO do frontend
 *
 * @public
 * @param {Array<{combination: string, kojoSuffix: string}>} selections - Combinações a processar
 * @param {Object} reportSettings - Todas as configurações da sidebar
 * @returns {{success: boolean, message: string}}
 */
function runProcessingFromHtml(selections, reportSettings) {
  if (!selections || selections.length === 0) {
    return { success: false, message: 'Nenhuma combinação recebida do painel' };
  }

  // Salva configurações para próximo uso
  ConfigService.saveAll(reportSettings);

  return processBomCore(selections, reportSettings);
}



/**
 * Processa e gera relatórios BOM (Bill of Materials)
 * Função principal do sistema de relatórios
 *
 * Para cada combinação selecionada:
 * 1. Filtra dados da aba fonte pelo grupo
 * 2. Agrupa e soma quantidades por item
 * 3. Cria/limpa aba de destino
 * 4. Formata relatório com cabeçalhos e totais
 *
 * @public
 * @param {Array<{combination: string, kojoSuffix: string}>} combinationsToProcess - Combinações a processar
 * @param {Object} settings - Configurações da sidebar
 * @param {string} settings.Aba_Fonte - Nome da aba com dados brutos
 * @param {string} settings.Grupo_Nivel_1 - Configuração do primeiro nível de grupo
 * @param {string} [settings.Grupo_Nivel_2] - Segundo nível de grupo (opcional)
 * @param {string} [settings.Grupo_Nivel_3] - Terceiro nível de grupo (opcional)
 * @param {string} settings.Ordenar_Por - Coluna para ordenação
 * @param {string} settings.Ordem_Classificacao - 'asc' ou 'desc'
 * @returns {{success: boolean, message: string}} Resultado da operação
 */
function processBomCore(combinationsToProcess, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const K = BOM_CONFIG.KEYS;

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

  const lastDataRow = sourceSheet.getLastRow();
  if (lastDataRow < 2) {
    return { success: false, message: `Aba "${settings[K.SOURCE_SHEET]}" não tem dados.` };
  }
  const allData = sourceSheet.getRange(2, 1, lastDataRow - 1, sourceSheet.getLastColumn()).getValues();

  for (const row of allData) {
    const rowCombination = groupIndices
      .map(index => String(row[index - 1] ?? '').trim())
      .join(BOM_CONFIG.DELIMITER);

    if (dataMap.has(rowCombination)) {
      dataMap.get(rowCombination).push([
        row[bomCols.c1 - 1], row[bomCols.c2 - 1], row[bomCols.c3 - 1],
        row[bomCols.c4 - 1], SharedUtils_toNumber(row[bomCols.c5 - 1]),
        row[sortColumnIndex] // Valor para classificação
      ]);
    }
  }

  let createdCount = 0;
  combinationsToProcess.forEach(item => {
    const { combination, kojoSuffix } = item;
    const rawData = dataMap.get(combination);
    if (!rawData || rawData.length === 0) return;

    const processedData = groupAndSumData(rawData, sortOrder);

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
    // Usa BOM_CONFIG.DELIMITER para consistência com as combinações de grupo
    const key = [row[0], row[1], row[2], row[3]].join(BOM_CONFIG.DELIMITER);
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
    if (SharedUtils_isEmpty(valA)) return 1;
    if (SharedUtils_isEmpty(valB)) return -1;
    const aIsNum = SharedUtils_isValidNumber(valA);
    const bIsNum = SharedUtils_isValidNumber(valB);
    if (aIsNum && bIsNum) {
      return (SharedUtils_toNumber(valA) - SharedUtils_toNumber(valB)) * direction;
    }
    return String(valA).localeCompare(String(valB), undefined, { numeric: true }) * direction;
  });

  return groupedData.map(row => row.slice(0, 5));
}


function createAndFormatReport(sheet, kojoSuffix, data, settings) {
  const K = BOM_CONFIG.KEYS;
  const reportConfig = {
    project: settings[K.PROJECT],
    bom: settings[K.BOM],
    kojoPrefix: settings[K.KOJO_PREFIX],
    engineer: settings[K.ENGINEER],
    version: Utils.formatVersion(settings[K.VERSION])  // ✅ FIX: aplica zero-pad na escrita
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

  sheet.getRange(1, 2, headerValues.length, 1).setNumberFormat('@STRING@');
  sheet.getRange(1, 1, headerValues.length, 2).setValues(headerValues);
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

/**
 * Remove abas de relatório geradas
 * V3.1: Usa aba fonte configurada (não aba ativa) como proteção
 *
 * @public
 * @menuitem '🔧 Relatórios Dinâmicos' > '🗑️ Limpar Relatórios'
 * @returns {void}
 */
function clearOldReports() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ✅ FIX: usa aba fonte configurada, não a aba ativa
  const config = ConfigService.getAll();
  const sourceSheetName = config[BOM_CONFIG.KEYS.SOURCE_SHEET] || ss.getActiveSheet().getName();

  const response = ui.alert(
    'Confirmação',
    `Apagar TODAS as abas exceto "${sourceSheetName}" (aba fonte)?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let deletedCount = 0;
  ss.getSheets().forEach(sheet => {
    if (sheet.getName() !== sourceSheetName) {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });
  ui.alert('Limpeza Concluída', `${deletedCount} abas removidas.`, ui.ButtonSet.OK);
}

// ============================================================================
// FIXADORES (SEÇÃO V2.7 - CÓPIA INTELIGENTE)
// ============================================================================

/**
 * Abre a sidebar para seleção de fixadores
 * Permite adicionar automaticamente fixadores (clamps, hangers) para tubulações
 *
 * @public
 * @menuitem '🔧 Relatórios Dinâmicos' > '🔧 Fixadores → Fonte'
 * @returns {void}
 */
function abrirSeletorFixadores() {
  const html = HtmlService.createHtmlOutputFromFile('FixadoresSidebar.html')
    .setTitle('Seletor de Fixadores')
    .setWidth(900).setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Seletor de Fixadores');
}

/**
 * Retorna configuração de colunas dos fixadores
 * Chamada pela sidebar para popular a UI de configuração
 */
function getFixadorConfig() {
  return ConfigService.getFixConfig();
}

/**
 * Salva configuração de colunas dos fixadores
 * Chamada pela sidebar ao alterar mapeamento de colunas
 */
function saveFixadorConfig(config) {
  return ConfigService.saveFixConfig(config);
}

/**
 * Retorna colunas disponíveis na aba ativa para configurar fixadores
 */
function getActiveSheetColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return _getColumnsFromSheet(sheet);
}

/**
 * Retorna colunas de uma aba pelo nome (ou aba ativa se não informado)
 */
function getSheetColumnsByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  return _getColumnsFromSheet(sheet);
}

function _getColumnsFromSheet(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.map((h, i) => {
    const letter = SharedUtils_numberToColumnLetter(i + 1);
    return `${letter} - ${h || 'Coluna ' + letter}`;
  });
}

function getPipesElegiveis() {
  const fixConfig = ConfigService.getFixConfig();
  const K = BOM_CONFIG.KEYS;
  const fixIdx = {
    section: Utils.getColumnIndex(fixConfig[K.FIX_SECTION]) - 1,
    desc: Utils.getColumnIndex(fixConfig[K.FIX_DESC]) - 1,
    qty: Utils.getColumnIndex(fixConfig[K.FIX_QTY]) - 1,
    uom: Utils.getColumnIndex(fixConfig[K.FIX_UOM]) - 1,
    trade: Utils.getColumnIndex(fixConfig[K.FIX_TRADE]) - 1
  };

  if ([fixIdx.section, fixIdx.desc, fixIdx.qty, fixIdx.uom, fixIdx.trade].some(idx => idx < 0)) {
    Logger.log('Erro de Configuração de Fixadores: Pelo menos uma coluna-chave não está definida.');
    return [];
  }

  // Usa a aba ativa como fonte de dados
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sourceSheet) return [];
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  const pipes = [];

  data.forEach((row, idx) => {
    const section = String(row[fixIdx.section] || '');
    const desc = String(row[fixIdx.desc] || '');
    const qty = SharedUtils_toNumber(row[fixIdx.qty]);
    const uom = String(row[fixIdx.uom] || '');

    if (validarTipoFixacao(section) && desc.toUpperCase().includes('PIPE') && qty > 0) {
      const diameter = Utils.extractDiameter(desc);
      const isRiser = section.toUpperCase().includes('RISER');
      const fixConfig = isRiser ? BOM_CONFIG.FIXADORES.RISER : BOM_CONFIG.FIXADORES.LOOP;
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
  const fixConfig = ConfigService.getFixConfig();
  const K = BOM_CONFIG.KEYS;
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }

  const fixIdx = {
    desc: Utils.getColumnIndex(fixConfig[K.FIX_DESC]) - 1,
    qty: Utils.getColumnIndex(fixConfig[K.FIX_QTY]) - 1,
    trade: Utils.getColumnIndex(fixConfig[K.FIX_TRADE]) - 1
  };

  if (fixIdx.desc < 0 || fixIdx.qty < 0 || fixIdx.trade < 0) {
     return { success: false, message: 'Configuração de colunas "Fixadores" inválida. Verifique no painel de configuração.' };
  }

  selectedPipes.sort((a, b) => b.rowIndex - a.rowIndex);
  let totalAdded = 0;
  const maxCol = sourceSheet.getLastColumn();

  selectedPipes.forEach(pipe => {
    const pipeFixConfig = pipe.isRiser ? BOM_CONFIG.FIXADORES.RISER : BOM_CONFIG.FIXADORES.LOOP;
    const itemMap = pipe.isRiser ? pipeFixConfig.clamps : pipeFixConfig.hangs;
    const fixadorItem = itemMap[pipe.diameter];
    if (!fixadorItem) return;

    const originalFormulas = sourceSheet.getRange(pipe.rowIndex, 1, 1, maxCol).getFormulasR1C1()[0];
    const linhasParaInserir = [];
    const insertRow = pipe.rowIndex + 1;

    // Linha do fixador
    const linhaFixador = [...pipe.originalRow];
    linhaFixador[fixIdx.trade] = 'FIX';
    linhaFixador[fixIdx.desc] = fixadorItem;
    linhaFixador[fixIdx.qty] = `=ROUNDUP(R[-1]C/${pipeFixConfig.interval})`;
    linhasParaInserir.push(linhaFixador);
    const fixadorRow = insertRow;

    // Materiais
    pipeFixConfig.materials.forEach(mat => {
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
    // Limpa validações de dados copiadas para evitar conflito ao setar valores
    rangeDestino.clearDataValidations();
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
  const fixConfig = ConfigService.getFixConfig();
  const K = BOM_CONFIG.KEYS;
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }

  const tradeColIndex = Utils.getColumnIndex(fixConfig[K.FIX_TRADE]) - 1;
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
// EXPORTAÇÃO PDF (V3.1)
// ============================================================================

/**
 * Monta o nome do arquivo PDF a partir de blocos configurados no sidebar.
 * @param {Sheet} sheet - Aba do relatório
 * @param {string} blocksConfigJson - JSON com {blocks, separator, find, replace}
 * @returns {string} Nome do arquivo (sem .pdf)
 */
function _assemblePdfFilename(sheet, blocksConfigJson) {
  try {
    const config = blocksConfigJson ? JSON.parse(blocksConfigJson) : null;
    if (!config || !config.blocks || config.blocks.length === 0) {
      return getBomKojoNameFromSheet(sheet) || sheet.getName();
    }
    const vals = sheet.getRange(1, 2, 6, 1).getValues();
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const fields = {
      project:    String(vals[0][0] || ''),
      bom:        String(vals[1][0] || ''),
      kojo:       String(vals[2][0] || ''),
      engineer:   String(vals[3][0] || ''),
      version:    String(vals[4][0] || ''),
      sheet_name: sheet.getName(),
      today:      todayStr,
    };
    const globalSep = config.separator || '-';
    const segments = [];
    config.blocks.forEach(b => {
      const val = b.type === 'text' ? (b.value || '') : (fields[b.type] || '');
      if (!val) return;
      if (segments.length > 0) {
        segments.push((b.sep !== '' && b.sep != null) ? b.sep : globalSep);
      }
      segments.push(val);
    });
    let name = segments.join('');
    // Suporte a formato antigo (find/replace único) e novo (array de regras)
    const rules = config.findReplaceRules || (config.find ? [{ find: config.find, replace: config.replace || '' }] : []);
    rules.forEach(r => { if (r.find) name = name.split(r.find).join(r.replace || ''); });
    return name || sheet.getName();
  } catch (e) {
    Logger.log('_assemblePdfFilename error: ' + e.message);
    return getBomKojoNameFromSheet(sheet) || sheet.getName();
  }
}

/**
 * Exporta abas selecionadas para PDF via sidebar HTML
 * Chamada pela sidebar BomSidebar.html
 *
 * @public
 * @param {string[]} sheetNames - Nomes das abas a exportar
 * @param {string} folderInput - ID da pasta Drive ou nome para criar
 * @param {string} blocksConfigJson - JSON com blocos do nome ({blocks, separator, find, replace})
 * @returns {{success: boolean, message?: string, exported?: number, folder?: string}}
 */
function runPdfExportFromHtml(sheetNames, folderInput, blocksConfigJson) {
  if (!sheetNames || sheetNames.length === 0) {
    return { success: false, message: 'Nenhuma aba selecionada' };
  }
  const folder = getFolderFromInput(folderInput, folderInput);
  if (!folder) return { success: false, message: `Pasta não encontrada ou inválida: ${folderInput}` };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const fileName = _assemblePdfFilename(sheet, blocksConfigJson);
      exportSheetToPdf(sheet, fileName, folder);
    }
  });
  return { success: true, exported: sheetNames.length, folder: folder.getName() };
}

/**
 * Exporta PDFs via menu (usa configurações salvas)
 * V3.0: Lê pasta Drive das configurações salvas em PropertiesService
 *
 * @public
 * @menuitem '🔧 Relatórios Dinâmicos' > '📄 Exportar PDFs'
 * @returns {void}
 */
function exportPDFsWithFeedback() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Exportando PDFs...', 'Aguarde', -1);
  const config = ConfigService.getAll();
  const K = BOM_CONFIG.KEYS;

  const sheetNames = getReportSheetNames();
  if (!sheetNames || sheetNames.length === 0) {
     SpreadsheetApp.getUi().alert('Erro', 'Nenhum relatório gerado para exportar.', SpreadsheetApp.getUi().ButtonSet.OK);
     return;
  }

  const folder = getFolderFromInput(config[K.DRIVE_FOLDER_ID], config[K.DRIVE_FOLDER_NAME]);
  if (!folder) {
    SpreadsheetApp.getUi().alert('Erro', 'Pasta de destino não configurada. Configure pelo Painel BOM.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const prefix = config[K.PDF_PREFIX] || '';
  sheetNames.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      const bomKojoName = getBomKojoNameFromSheet(sheet);
      const baseName = bomKojoName || sheetName;
      const fileName = prefix ? `${prefix}${baseName}` : baseName;
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
 * Retorna dados iniciais para a BomSidebar
 * V3.0: Usa aba ativa como fonte e PropertiesService para config
 */
function getBomHtmlInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const savedConfig = ConfigService.getAll();
  const activeSheet = ss.getActiveSheet();

  const allSheets = ss.getSheets().map(s => s.getName());

  // Pega colunas da aba ativa como referência inicial
  let allColumns = [];
  const targetSheet = savedConfig[BOM_CONFIG.KEYS.SOURCE_SHEET]
    ? ss.getSheetByName(savedConfig[BOM_CONFIG.KEYS.SOURCE_SHEET])
    : activeSheet;

  if (targetSheet && targetSheet.getLastColumn() > 0) {
    const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    allColumns = headers.map((h, i) => {
      const letter = SharedUtils_numberToColumnLetter(i + 1);
      return `${letter} - ${h || 'Coluna ' + letter}`;
    });
  }

  const savedState = getUserSidebarState();

  return {
    allSheets: allSheets,
    allColumns: allColumns,
    currentConfig: savedConfig,
    savedState: savedState,
    activeSheetName: activeSheet ? activeSheet.getName() : ''
  };
}

/**
 * Retorna valores únicos de uma coluna
 * V3.0: Aceita sheetName como parâmetro direto
 */
function getUniqueValuesForColumn(colName, sheetName) {
  if (!colName) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  if (!sourceSheet) return [];
  const colIndex = Utils.getColumnIndex(colName);
  if (colIndex === -1) return [];
  return getUniqueColumnValues(sourceSheet, colIndex);
}

/**
 * Gera combinações para pré-visualização
 * V3.0: Aceita sheetName como parâmetro direto
 */
function getCombinationsForPreview(selectedGroups, groupConfigs, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
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

  // ✅ FIX: guard contra aba vazia
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return [];
  const allData = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  const existingCombinations = new Set();

  allData.forEach(row => {
    const combinationParts = groupIndices.map(index => row[index - 1] ?? '');
    if (combinationParts.every(part => String(part).trim() !== '')) {
      existingCombinations.add(combinationParts.join(BOM_CONFIG.DELIMITER));
    }
  });

  const finalCombinations = [...existingCombinations].filter(combo => {
    const parts = combo.split(BOM_CONFIG.DELIMITER);
    return parts.every((part, i) => {
      const levelId = `NÍVEL ${i + 1}`;
      if (!panelSelections[levelId]) return false;
      return Array.from(panelSelections[levelId]).some(selected => String(selected).trim() === String(part).trim());
    });
  });

  return finalCombinations.sort((a, b) => {
    const aParts = a.split(BOM_CONFIG.DELIMITER);
    const bParts = b.split(BOM_CONFIG.DELIMITER);
    for (let i = 0; i < Math.min(aParts.length, bParts.length); i++) {
      const aIsNum = SharedUtils_isValidNumber(aParts[i]);
      const bIsNum = SharedUtils_isValidNumber(bParts[i]);
      if (aIsNum && bIsNum) {
        const aNum = SharedUtils_toNumber(aParts[i]);
        const bNum = SharedUtils_toNumber(bParts[i]);
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
    return state ? JSON.parse(state) : null;
  } catch (e) {
    return null;
  }
}

// ============================================================================
// TRIGGERS E UI - BOM
// ============================================================================

/**
 * Abre o painel modal do Gerador de BOM
 * Dialog maior (1100x750) com interface completa de geração de relatórios
 *
 * @public
 * @menuitem '🔧 Relatórios Dinâmicos' > '📊 Gerador de BOM (Painel)'
 * @returns {void}
 */
function openBomSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('BomSidebar.html')
    .setTitle('Painel Gerador de BOM')
    .setWidth(1100)
    .setHeight(750);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Painel Gerador de BOM');
}

/**
 * Executa diagnóstico do sistema BOM
 * V3.0: Mostra info da aba ativa
 *
 * @public
 * @menuitem '🔧 Relatórios Dinâmicos' > '🧪 Diagnóstico'
 */
function testSystem() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const config = ConfigService.getAll();
    const fixConfig = ConfigService.getFixConfig();
    const msg = [
      `Versão do Script: 3.1 (Correções PDF/Versão/Cache)`,
      `Aba Ativa: ${activeSheet.getName()}`,
      `Linhas na Aba Ativa: ${activeSheet.getLastRow() - 1}`,
      `Total de Abas: ${ss.getSheets().length}`,
      `Fixador Trade: ${fixConfig[BOM_CONFIG.KEYS.FIX_TRADE] || 'Não configurado'}`,
      `Config BOM salva: ${config[BOM_CONFIG.KEYS.SOURCE_SHEET] ? 'Sim' : 'Usando defaults'}`
    ].join('\n');
    SpreadsheetApp.getUi().alert('🧪 Diagnóstico do Sistema', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro no Diagnóstico', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * V3.0: Retorna abas que são relatórios gerados (todas exceto a fonte)
 */
function getReportSheetNames() {
  const config = ConfigService.getAll();
  const sourceSheetName = config[BOM_CONFIG.KEYS.SOURCE_SHEET] || '';
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(s => s.getName())
    .filter(name => name !== sourceSheetName);
}

function getReportSheetNamesForHtml() {
  return getReportSheetNames();
}

/**
 * Retorna dados de header de cada aba de relatório para montar preview de nome PDF.
 * @public
 * @returns {Array<{name, project, bom, kojo, engineer, version}>}
 */
function getReportSheetDataForHtml() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getReportSheetNames().map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return { name, project: '', bom: '', kojo: '', engineer: '', version: '' };
    try {
      const vals = sheet.getRange(1, 2, 6, 1).getValues();
      return {
        name,
        project:  String(vals[0][0] || ''),
        bom:      String(vals[1][0] || ''),
        kojo:     String(vals[2][0] || ''),
        engineer: String(vals[3][0] || ''),
        version:  String(vals[4][0] || ''),
      };
    } catch (e) {
      return { name, project: '', bom: '', kojo: '', engineer: '', version: '' };
    }
  });
}

/**
 * Força limpeza completa do cache do sistema BOM
 * Útil quando os dados parecem desatualizados
 *
 * @public
 * @menuitem '🔧 Relatórios Dinâmicos' > '🔄 Limpar Cache'
 * @returns {void}
 */
function forceRefreshCache() {
  CacheManager.invalidateAll();
  SpreadsheetApp.getActiveSpreadsheet().toast('Cache limpo!', 'Sucesso', 3);
}
