/**
 * @fileoverview SISTEMA UNIFICADO DE RELATÓRIOS DINÂMICOS (BOM)
 * @version 3.4.0 - Higiene 07/2026: ferramenta Fixadores removida (sem ponto de entrada;
 *                  código preservado em C:\DEV\_OBSOLETO\Sheets\); órfãos removidos
 *                  (testSystem, exportPDFsWithFeedback, getReportSheetNamesForHtml);
 *                  fix ConfigService.get com valores falsy
 * @version 3.3.0 - Cache de valores únicos removido (sempre fresco) + botão Atualizar na sidebar;
 *                  lista de exportação = todas as abas exceto a fonte, em ordem alfabética (natural sort);
 *                  clearOldReports protege a fonte (assinatura de cabeçalho);
 *                  banding idempotente na regeneração; autodetect col5 PROJECT→QTY; debounce de estado
 *
 * V3.0: Cada ferramenta opera de forma independente:
 * - BomSidebar gerencia suas próprias configurações via PropertiesService
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

    EXCL_FILTERS: 'Filtros de Exclusão',
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
    'Filtros de Exclusão': [],
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
  DELIMITER: '|||'
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
  }
};

// ============================================================================
// CONFIGURAÇÃO - BOM (V3.0 - PropertiesService)
// Configurações salvas via PropertiesService, sem dependência de aba Config
// ============================================================================

const BOM_SETTINGS_KEY = 'BOM_SETTINGS_V3';

// V3.4: delega ao factory compartilhado (lib/Shared/Config.gs).
// Inicialização lazy: a ordem de avaliação dos arquivos no GAS não é garantida,
// então a chamada cross-file só acontece em runtime, nunca no load.
const ConfigService = {
  _svc: null,
  _store() {
    if (!this._svc) {
      this._svc = SharedConfig_createDocConfigService(
        BOM_SETTINGS_KEY,
        () => ({ ...BOM_CONFIG.DEFAULTS })
      );
    }
    return this._svc;
  },
  getAll() { return this._store().getAll(); },
  get(key, defaultValue = '') { return this._store().get(key, defaultValue); },
  saveAll(config) { return this._store().saveAll(config); }
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

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Sem cache: os valores únicos são lidos sempre frescos da aba fonte.
  // O cache anterior (180s, chave sheetId+coluna) não detectava edição
  // in-place dos dados e deixava os painéis desatualizados.
  const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues().flat();
  return [...new Set(values.filter(v => v))]
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
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

  // Filtros de exclusão: vals = valores PERMITIDOS (check = manter, uncheck = excluir)
  const exclRules = (settings[K.EXCL_FILTERS] || [])
    .filter(f => f.col && Array.isArray(f.vals) && f.vals.length > 0)
    .map(f => ({
      colIdx: Utils.getColumnIndex(f.col) - 1,
      allowed: new Set(f.vals.map(v => String(v).trim().toLowerCase()))
    }))
    .filter(r => r.colIdx >= 0);

  for (const row of allData) {
    if (exclRules.length > 0 && exclRules.some(r =>
      !r.allowed.has(String(row[r.colIdx] ?? '').trim().toLowerCase())
    )) continue;

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

    // Persiste os valores individuais de cada nível para uso no nome do PDF
    const levelValues = combination.split(BOM_CONFIG.DELIMITER);
    PropertiesService.getDocumentProperties().setProperty(
      'BOM_META_' + sanitizedName,
      JSON.stringify({ l1: levelValues[0] || '', l2: levelValues[1] || '', l3: levelValues[2] || '' })
    );

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

  // ✅ FIX (B4): clear() não remove banding (é objeto do Sheet, não do Range).
  // Sem isto, regerar a mesma aba lança "conflicting banding" em applyRowBanding.
  sheet.getBandings().forEach(b => b.remove());

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

  // ✅ FIX (B1): apaga SOMENTE abas com assinatura de relatório BOM gerado.
  // A aba de dados (fonte) e abas auxiliares nunca são tocadas.
  const reportSheets = ss.getSheets().filter(_isGeneratedReportSheet);
  if (reportSheets.length === 0) {
    ui.alert('Nada a limpar', 'Nenhuma aba de relatório BOM gerado foi encontrada.', ui.ButtonSet.OK);
    return;
  }

  const names = reportSheets.map(s => s.getName());
  const response = ui.alert(
    'Confirmação',
    `Apagar ${names.length} aba(s) de relatório BOM?\n\n${names.join(', ')}\n\n` +
    `A aba de dados e quaisquer abas auxiliares NÃO serão tocadas.`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  // ✅ FIX (B8): remove também a meta (BOM_META_*) de cada relatório apagado
  const docProps = PropertiesService.getDocumentProperties();
  let deletedCount = 0;
  reportSheets.forEach(sheet => {
    docProps.deleteProperty('BOM_META_' + sheet.getName());
    ss.deleteSheet(sheet);
    deletedCount++;
  });
  ui.alert('Limpeza Concluída', `${deletedCount} aba(s) de relatório removida(s).`, ui.ButtonSet.OK);
}

// ============================================================================
// COLUNAS — helpers usados pelo BomSidebar
// ============================================================================

/**
 * Retorna colunas de uma aba pelo nome (ou aba ativa se não informado)
 */
function getSheetColumnsByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  return _getColumnsFromSheet(sheet);
}

function _getColumnsFromSheet(sheet) {
  return SharedUtils_getColumnLabelsFromSheet(sheet);
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
    let name;
    if (!config || !config.blocks || config.blocks.length === 0) {
      name = getBomKojoNameFromSheet(sheet) || sheet.getName();
    } else {
      const vals = sheet.getRange(1, 2, 6, 1).getValues();
      const metaJson = PropertiesService.getDocumentProperties().getProperty('BOM_META_' + sheet.getName());
      const meta = metaJson ? JSON.parse(metaJson) : {};
      const fields = {
        project:    String(vals[0][0] || ''),
        bom:        String(vals[1][0] || ''),
        kojo:       String(vals[2][0] || ''),
        engineer:   String(vals[3][0] || ''),
        version:    String(vals[4][0] || ''),
        sheet_name: sheet.getName(),
        l1:         meta.l1 || '',
        l2:         meta.l2 || '',
        l3:         meta.l3 || '',
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
      name = segments.join('');
    }
    const rules = config ? (config.findReplaceRules || (config.find ? [{ find: config.find, replace: config.replace || '' }] : [])) : [];
    rules.forEach(r => {
      if (!r.find) return;
      if (r.regex) {
        try { name = name.replace(new RegExp(r.find, 'g'), r.replace || ''); } catch(e) {}
      } else {
        name = name.split(r.find).join(r.replace || '');
      }
    });
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
const _EXPORT_PROGRESS_KEY = 'PDF_EXPORT_PROGRESS_V1';

function getExportProgress() {
  const raw = CacheService.getUserCache().get(_EXPORT_PROGRESS_KEY);
  return raw ? JSON.parse(raw) : { status: 'idle' };
}

function _setExportProgress(done, total, current, status) {
  CacheService.getUserCache().put(
    _EXPORT_PROGRESS_KEY,
    JSON.stringify({ status: status || 'running', done, total, current: current || '' }),
    600
  );
}

function runPdfExportFromHtml(sheetNames, folderInput, blocksConfigJson) {
  if (!sheetNames || sheetNames.length === 0) {
    return { success: false, message: 'Nenhuma aba selecionada' };
  }
  const folder = getFolderFromInput(folderInput, folderInput);
  if (!folder) return { success: false, message: `Pasta não encontrada ou inválida: ${folderInput}` };

  const total = sheetNames.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  let exported = 0;

  _setExportProgress(0, total, '', 'running');

  sheetNames.forEach((sheetName, i) => {
    _setExportProgress(i, total, sheetName, 'running');
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      try {
        const fileName = _assemblePdfFilename(sheet, blocksConfigJson);
        exportSheetToPdf(sheet, fileName, folder);
        exported++;
      } catch(e) {
        errors.push(sheetName);
        Logger.log(`Erro ao exportar "${sheetName}": ${e.message}`);
      }
    }
  });

  _setExportProgress(total, total, '', 'done');
  return { success: true, exported, folder: folder.getName(), errors };
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

function _extractDriveFolderId(input) {
  if (!input || typeof input !== 'string') return null;
  // ?id= ou &id= (ex: https://drive.google.com/open?id=XXX)
  const idParam = input.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParam) return idParam[1];
  // /folders/ID
  const folderPath = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderPath) return folderPath[1];
  // ID puro (só alfanum + - _)
  if (/^[a-zA-Z0-9_-]{10,}$/.test(input.trim())) return input.trim();
  return null;
}

function getFolderFromInput(folderInput, folderName) {
  try {
    if (folderInput) {
      const extractedId = _extractDriveFolderId(folderInput);
      if (extractedId) {
        try {
          const f = DriveApp.getFolderById(extractedId);
          if (f) return f;
        } catch(e) { /* ID inválido */ }
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
        const existing = folder.getFilesByName(`${pdfName}.pdf`);
        while (existing.hasNext()) existing.next().setTrashed(true);
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
/**
 * Remove properties BOM_META_* sem aba correspondente (relatório apagado/renomeado).
 * Garbage collection barato (só chaves + nomes de abas, sem ler células), ao abrir o painel.
 */
function _syncReportMeta() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = new Set(ss.getSheets().map(s => s.getName()));
    const docProps = PropertiesService.getDocumentProperties();
    docProps.getKeys().forEach(k => {
      if (k.indexOf('BOM_META_') === 0 && !existing.has(k.substring(9))) {
        docProps.deleteProperty(k);
      }
    });
  } catch (e) { /* não crítico */ }
}

function getBomHtmlInitData() {
  _syncReportMeta();  // sincroniza o registro de relatórios (prune + adoção de legados)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const savedConfig = ConfigService.getAll();
  const activeSheet = ss.getActiveSheet();

  const allSheets = ss.getSheets().map(s => s.getName());

  // Pega colunas da aba ativa como referência inicial
  const targetSheet = savedConfig[BOM_CONFIG.KEYS.SOURCE_SHEET]
    ? ss.getSheetByName(savedConfig[BOM_CONFIG.KEYS.SOURCE_SHEET])
    : activeSheet;
  const allColumns = _getColumnsFromSheet(targetSheet);

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
 * Detecta se uma aba é um relatório BOM gerado por esta ferramenta.
 * Usa a assinatura determinística do cabeçalho escrito por createAndFormatReport
 * (A1 = "PROJECT:", A3 = "BOM KOJO:..."), em vez de "tudo que não é a fonte".
 * Protege a aba de dados e quaisquer abas auxiliares.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {boolean}
 */
function _isGeneratedReportSheet(sheet) {
  try {
    if (!sheet || sheet.getLastRow() < 3) return false;
    const colA = sheet.getRange(1, 1, 3, 1).getValues();
    const a1 = String(colA[0][0] || '').trim();
    const a3 = String(colA[2][0] || '').trim();
    return a1 === 'PROJECT:' && a3.indexOf('BOM KOJO') === 0;
  } catch (e) {
    return false;
  }
}

/**
 * V3.3: Retorna os nomes das abas de relatório, em ordem alfabética (natural sort).
 * Estratégia robusta: TODAS as abas exceto a fonte de dados — não usa assinatura de
 * cabeçalho (que variava entre versões e escondia relatórios legados).
 * Rápido: só nomes de abas, sem leitura de células.
 * @param {string} [sourceSheetName] Aba fonte a excluir (vinda do painel). Fallback: config salva.
 */
function getReportSheetNames(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = sourceSheetName || ConfigService.getAll()[BOM_CONFIG.KEYS.SOURCE_SHEET] || '';
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => name !== source)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
}

/**
 * Retorna dados de header de cada aba de relatório para montar preview de nome PDF.
 * @public
 * @returns {Array<{name, project, bom, kojo, engineer, version, l1, l2, l3}>}
 */
function getReportSheetDataForHtml(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docProps = PropertiesService.getDocumentProperties();
  return getReportSheetNames(sourceSheetName).map(name => {
    const sheet = ss.getSheetByName(name);
    const metaJson = docProps.getProperty('BOM_META_' + name);
    if (!sheet) {
      const meta = metaJson ? JSON.parse(metaJson) : {};
      return { name, project: '', bom: '', kojo: '', engineer: '', version: '', l1: meta.l1 || '', l2: meta.l2 || '', l3: meta.l3 || '' };
    }
    try {
      const vals = sheet.getRange(1, 2, 6, 1).getValues();
      const kojo = String(vals[2][0] || '');
      let l1 = '', l2 = '', l3 = '';
      if (metaJson) {
        const meta = JSON.parse(metaJson);
        l1 = meta.l1 || ''; l2 = meta.l2 || ''; l3 = meta.l3 || '';
      } else {
        // Fallback para BOMs existentes: infere do kojo (funciona se não foi editado manualmente)
        const parts = kojo.split('.');
        l1 = (parts[0] || '').trim(); l2 = (parts[1] || '').trim(); l3 = (parts[2] || '').trim();
      }
      return {
        name,
        project:  String(vals[0][0] || ''),
        bom:      String(vals[1][0] || ''),
        kojo,
        engineer: String(vals[3][0] || ''),
        version:  String(vals[4][0] || ''),
        l1, l2, l3,
      };
    } catch (e) {
      return { name, project: '', bom: '', kojo: '', engineer: '', version: '', l1: '', l2: '', l3: '' };
    }
  });
}

// Cache removido (V3.2): getUniqueColumnValues lê sempre fresco; não há mais
// cache a limpar. O botão "🔄 Atualizar" na sidebar relê os dados sob demanda.
