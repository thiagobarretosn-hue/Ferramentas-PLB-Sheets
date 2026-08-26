/**
 * @fileoverview Montar Submittal — monta o pacote de submittal para o cliente
 * @version 1.0.0
 *
 * Fluxo:
 *   1. Filtra linhas com STATUS (col I) = "NOT COMMITTED" do número escolhido (col B)
 *   2. Cria pasta "SUBMITTAL <num> - MM/DD/YYYY" dentro da pasta pai configurada
 *   3. Gera a capa a partir do template Google Doc (placeholders {{...}})
 *   4. Mescla capa + PDFs dos links (col L) em um único PDF via pdf-lib
 *   5. Arquivos não-PDF/protegidos são copiados soltos para a pasta
 *   6. Escreve o link do PDF final na col M (LINK SUBMITTAL) de cada linha
 *
 * Links da col L: suporta smart chip de arquivo (Sheets API v4 avançada),
 * hyperlink normal (RichTextValue) e URL em texto puro.
 *
 * Requer no appsscript.json: enabledAdvancedServices → Sheets v4.
 */

// ============================================================================
// CONSTANTES
// ============================================================================

const SUBMITTAL_CFG = {
  SETTINGS_KEY: 'SUBMITTAL_SETTINGS_V1',
  PROGRESS_KEY: 'SUBMITTAL_PROGRESS_V1',
  STATUS_FILTER: 'NOT COMMITTED',

  // ⚠️ ID FIXO do template da capa (Google Doc com placeholders):
  // {{PROJECT_NAME}} {{PROJECT_ADDRESS}} {{TO}} {{TITLE}} {{DATE}} {{PAGES}}
  TEMPLATE_DOC_ID: '1NUhDb-rHw-cdCrrIKo1rFuWVv7k_flB3Fa8bMDFIkKo',

  PDF_LIB_URL: 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js',

  // Colunas fixas da planilha Submittal Items LOG (1-indexed)
  COLS: {
    NUM: 2,        // B - #SUBMITTAL AMBAR
    DATE: 3,       // C - SUBMITTAL DATE
    DISCIPLINE: 4, // D - DISCIPLINE
    ITEM: 7,       // G - ITEM
    STATUS: 9,     // I - STATUS
    LINK: 12,      // L - LINK
    LINK_OUT: 13   // M - LINK SUBMITTAL
  }
};

// ============================================================================
// CONFIG (factory compartilhado — DocumentProperties)
// ============================================================================

const SubmittalConfigService = {
  _svc: null,
  _store: function() {
    if (!this._svc) {
      this._svc = SharedConfig_createDocConfigService(
        SUBMITTAL_CFG.SETTINGS_KEY,
        SubmittalConfigService._defaults
      );
    }
    return this._svc;
  },
  _defaults: function() {
    return {
      sourceSheet: '',
      parentFolderId: '',
      projectName: '',
      projectAddress: '',
      toName: '',
      submittalTitle: ''
    };
  },
  get: function() { return this._store().getAll(); },
  save: function(config) { return this._store().saveAll(config); }
};

function saveSubmittalConfig(config) {
  return SubmittalConfigService.save(config);
}

// ============================================================================
// LEITURA — números, itens e links
// ============================================================================

function getSubmittalInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = SubmittalConfigService.get();
  const allSheets = ss.getSheets().map(function(s) { return s.getName(); });
  return { allSheets: allSheets, config: config, templateOk: SUBMITTAL_CFG.TEMPLATE_DOC_ID.indexOf('COLE_AQUI') < 0 };
}

/** Números de submittal (col B) com itens NOT COMMITTED, com contagem */
function getSubmittalNumbers(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const C = SUBMITTAL_CFG.COLS;
  const data = sheet.getRange(2, 1, lastRow - 1, C.STATUS).getValues();
  const counts = {};
  data.forEach(function(row) {
    const status = String(row[C.STATUS - 1] || '').trim().toUpperCase();
    if (status !== SUBMITTAL_CFG.STATUS_FILTER) return;
    const num = String(row[C.NUM - 1] || '').trim();
    if (!num) return;
    counts[num] = (counts[num] || 0) + 1;
  });

  return Object.keys(counts)
    .sort(function(a, b) { return a.localeCompare(b, undefined, { numeric: true }); })
    .map(function(num) { return { num: num, count: counts[num] }; });
}

/** Extrai ID do Drive de uma URL (file/d/, /d/, open?id=, uc?id=) */
function _sub_extractDriveId(url) {
  if (!url) return null;
  let m = String(url).match(/\/(?:file\/)?d\/([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  m = String(url).match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  return null;
}

/**
 * Lê os links da coluna L em lote, cobrindo os 3 formatos:
 *   1. smart chip de arquivo → Sheets API v4 (chipRuns.richLinkProperties.uri)
 *   2. hyperlink normal      → campo hyperlink / textFormatRuns.link
 *   3. URL em texto puro     → formattedValue
 * Retorna mapa { numeroDaLinha: { url, label } }
 */
function _sub_getLinksMap(sheetName, startRow, endRow) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const colLetter = SharedUtils_numberToColumnLetter(SUBMITTAL_CFG.COLS.LINK);
  const range = "'" + sheetName + "'!" + colLetter + startRow + ':' + colLetter + endRow;

  const resp = Sheets.Spreadsheets.get(ssId, {
    ranges: [range],
    fields: 'sheets(data(rowData(values(hyperlink,formattedValue,chipRuns,textFormatRuns))))'
  });

  const map = {};
  try {
    const rows = resp.sheets[0].data[0].rowData || [];
    rows.forEach(function(rd, i) {
      const rowNum = startRow + i;
      const cell = (rd && rd.values && rd.values[0]) ? rd.values[0] : null;
      if (!cell) return;

      const label = String(cell.formattedValue || '').trim();

      // 1) smart chip de arquivo
      if (cell.chipRuns && cell.chipRuns.length > 0) {
        for (let c = 0; c < cell.chipRuns.length; c++) {
          const chip = cell.chipRuns[c].chip;
          if (chip && chip.richLinkProperties && chip.richLinkProperties.uri) {
            map[rowNum] = { url: chip.richLinkProperties.uri, label: label || 'arquivo' };
            return;
          }
        }
      }

      // 2) hyperlink da célula inteira
      if (cell.hyperlink) {
        map[rowNum] = { url: cell.hyperlink, label: label || cell.hyperlink };
        return;
      }

      // 2b) hyperlink em trecho do texto (textFormatRuns)
      if (cell.textFormatRuns && cell.textFormatRuns.length > 0) {
        for (let r = 0; r < cell.textFormatRuns.length; r++) {
          const fmt = cell.textFormatRuns[r].format;
          if (fmt && fmt.link && fmt.link.uri) {
            map[rowNum] = { url: fmt.link.uri, label: label || fmt.link.uri };
            return;
          }
        }
      }

      // 3) texto puro que é URL
      if (/^https?:\/\//i.test(label)) {
        map[rowNum] = { url: label, label: label };
      }
    });
  } catch (e) {
    console.error('[Submittal] Erro ao ler links: ' + e.message);
  }
  return map;
}

/**
 * Itens NOT COMMITTED do número escolhido, com o link resolvido de cada linha.
 * Chamado pela sidebar para montar a lista.
 */
function getSubmittalItems(num, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'Aba "' + sheetName + '" não encontrada.' };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Aba sem dados.' };

  const C = SUBMITTAL_CFG.COLS;
  const data = sheet.getRange(2, 1, lastRow - 1, C.LINK).getValues();
  const linksMap = _sub_getLinksMap(sheetName, 2, lastRow);

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const items = [];
  data.forEach(function(row, i) {
    const rowNum = i + 2;
    const status = String(row[C.STATUS - 1] || '').trim().toUpperCase();
    const rowSubNum = String(row[C.NUM - 1] || '').trim();
    if (status !== SUBMITTAL_CFG.STATUS_FILTER || rowSubNum !== String(num).trim()) return;

    let dateVal = row[C.DATE - 1];
    if (dateVal instanceof Date) dateVal = Utilities.formatDate(dateVal, tz, 'MM/dd/yyyy');

    items.push({
      row: rowNum,
      item: String(row[C.ITEM - 1] || '').trim(),
      discipline: String(row[C.DISCIPLINE - 1] || '').trim(),
      date: String(dateVal || '').trim(),
      link: linksMap[rowNum] || null
    });
  });

  return { success: true, items: items };
}

/** Grava um link (URL) na coluna L de uma linha — para itens sem link */
function saveSubmittalItemLink(rowNum, url, label, sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Aba não encontrada.' };
    const cleanUrl = String(url || '').trim();
    if (!/^https?:\/\//i.test(cleanUrl)) {
      return { success: false, message: 'URL inválida (deve começar com http/https).' };
    }
    const text = String(label || '').trim() || cleanUrl;
    const rt = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(cleanUrl).build();
    sheet.getRange(rowNum, SUBMITTAL_CFG.COLS.LINK).setRichTextValue(rt);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// PROGRESSO (sidebar faz polling — cobre execuções longas e async)
// ============================================================================

function _sub_setProgress(obj) {
  CacheService.getUserCache().put(SUBMITTAL_CFG.PROGRESS_KEY, JSON.stringify(obj), 900);
}

function getSubmittalProgress() {
  const raw = CacheService.getUserCache().get(SUBMITTAL_CFG.PROGRESS_KEY);
  return raw ? JSON.parse(raw) : { status: 'idle' };
}

// ============================================================================
// MONTAGEM — pasta, coleta, capa, merge
// ============================================================================

/**
 * Polyfill de timers para o pdf-lib: o V8 do GAS não tem setTimeout/setInterval
 * (erro "setTimeout is not defined" ao rodar o merge). Execução síncrona via
 * Utilities.sleep é suficiente para o uso que o pdf-lib faz.
 */
function _sub_installTimerPolyfill() {
  if (typeof globalThis.setTimeout === 'undefined') {
    globalThis.setTimeout = function(fn, ms) {
      Utilities.sleep(Math.min(ms || 0, 100));
      if (typeof fn === 'function') fn();
      return 0;
    };
    globalThis.clearTimeout = function() {};
  }
  if (typeof globalThis.setInterval === 'undefined') {
    globalThis.setInterval = function() { return 0; };
    globalThis.clearInterval = function() {};
  }
}

/** Resolve um link em blob. kind: 'pdf' | 'other'. Lança erro se inacessível. */
function _sub_resolveBlob(url) {
  const driveId = _sub_extractDriveId(url);
  if (driveId) {
    const file = DriveApp.getFileById(driveId);
    const mime = file.getMimeType();
    if (mime === MimeType.PDF) return { blob: file.getBlob(), kind: 'pdf', name: file.getName() };
    if (mime.indexOf('application/vnd.google-apps') === 0) {
      // Google Doc/Sheet/Slide → converte para PDF
      return { blob: file.getAs('application/pdf'), kind: 'pdf', name: file.getName() + '.pdf' };
    }
    return { blob: file.getBlob(), kind: 'other', name: file.getName() };
  }

  // Web: precisa ser download direto
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
  const blob = resp.getBlob();
  const ct = String(blob.getContentType() || '').toLowerCase();
  const isPdf = ct.indexOf('pdf') >= 0;
  let name = blob.getName() || url.split('/').pop().split('?')[0] || 'arquivo';
  if (isPdf && !/\.pdf$/i.test(name)) name += '.pdf';
  return { blob: blob, kind: isPdf ? 'pdf' : 'other', name: name };
}

/** Gera o PDF da capa a partir do template Doc, com placeholders substituídos */
function _sub_makeCoverPdf(fields, totalPages) {
  const tplId = SUBMITTAL_CFG.TEMPLATE_DOC_ID;
  if (tplId.indexOf('COLE_AQUI') === 0) {
    throw new Error('TEMPLATE_DOC_ID não configurado no Submittal.gs (constante SUBMITTAL_CFG.TEMPLATE_DOC_ID).');
  }
  const copy = DriveApp.getFileById(tplId).makeCopy('tmp-submittal-cover');
  try {
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    const map = {
      'PROJECT_NAME': fields.projectName || '',
      'PROJECT_ADDRESS': fields.projectAddress || '',
      'TO': fields.toName || '',
      'TITLE': fields.title || '',
      'DATE': fields.date || '',
      'PAGES': String(totalPages)
    };
    Object.keys(map).forEach(function(key) {
      body.replaceText('\\{\\{' + key + '\\}\\}', map[key]);
    });
    doc.saveAndClose();
    return copy.getAs('application/pdf');
  } finally {
    copy.setTrashed(true);
  }
}

/**
 * Monta o submittal. ASYNC (pdf-lib é promise-based) — a sidebar NÃO usa o
 * retorno: acompanha via getSubmittalProgress() (status 'done' traz o report).
 *
 * @param {string} num - Número do submittal (col B)
 * @param {Object} coverFields - { title, date } (projectName/address/to vêm da config)
 * @param {string} sheetName - Aba fonte
 */
async function montarSubmittal(num, coverFields, sheetName) {
  try {
    _sub_setProgress({ status: 'running', step: 'Lendo itens...', done: 0, total: 0 });

    const config = SubmittalConfigService.get();
    if (!config.parentFolderId) throw new Error('Pasta pai não configurada. Preencha na seção Configuração.');

    const itemsResp = getSubmittalItems(num, sheetName);
    if (!itemsResp.success) throw new Error(itemsResp.message);
    const items = itemsResp.items;
    if (items.length === 0) throw new Error('Nenhum item "' + SUBMITTAL_CFG.STATUS_FILTER + '" para o submittal ' + num + '.');

    // Pasta: SUBMITTAL <num> - MM/DD/YYYY (data da montagem, padrão americano)
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');
    const folderName = 'SUBMITTAL ' + num + ' - ' + today;

    let parent;
    try {
      parent = DriveApp.getFolderById(String(config.parentFolderId).trim());
    } catch (e) {
      throw new Error('Pasta pai inacessível. Confira o ID na Configuração.');
    }
    // Reusa pasta homônima se já existir (re-montagem no mesmo dia)
    const existing = parent.getFoldersByName(folderName);
    const folder = existing.hasNext() ? existing.next() : parent.createFolder(folderName);

    // ── Coleta dos arquivos ──
    const pdfEntries = [];   // { blob, name, rows: [rowNum] }
    const looseFiles = [];   // não-PDF ou erro de merge → copiados soltos
    const noLink = [];       // itens sem link
    const failed = [];       // link presente mas inacessível

    items.forEach(function(it, i) {
      _sub_setProgress({ status: 'running', step: 'Coletando: ' + it.item, done: i, total: items.length });
      if (!it.link) { noLink.push(it.item); return; }
      try {
        const r = _sub_resolveBlob(it.link.url);
        if (r.kind === 'pdf') {
          pdfEntries.push({ blob: r.blob, name: r.name, item: it.item, row: it.row });
        } else {
          folder.createFile(r.blob.setName(r.name));
          looseFiles.push(it.item + ' (' + r.name + ')');
        }
      } catch (e) {
        failed.push(it.item + ' — ' + e.message);
      }
    });

    if (pdfEntries.length === 0 && looseFiles.length === 0) {
      throw new Error('Nenhum arquivo pôde ser coletado. Sem link: ' + noLink.length + ' | Falhas: ' + failed.length);
    }

    // ── Merge com pdf-lib ──
    _sub_setProgress({ status: 'running', step: 'Carregando pdf-lib...', done: items.length, total: items.length });
    _sub_installTimerPolyfill(); // pdf-lib exige setTimeout (inexistente no GAS)
    eval(UrlFetchApp.fetch(SUBMITTAL_CFG.PDF_LIB_URL).getContentText());

    // 1ª passada: carregar e contar páginas (necessário para {{PAGES}} na capa)
    const loadedDocs = [];
    let totalItemPages = 0;
    for (let i = 0; i < pdfEntries.length; i++) {
      _sub_setProgress({ status: 'running', step: 'Lendo PDF: ' + pdfEntries[i].item, done: i, total: pdfEntries.length });
      try {
        const src = await PDFLib.PDFDocument.load(
          new Uint8Array(pdfEntries[i].blob.getBytes()), { ignoreEncryption: true });
        loadedDocs.push({ doc: src, entry: pdfEntries[i] });
        totalItemPages += src.getPageCount();
      } catch (e) {
        // PDF corrompido/criptografado: vai solto para a pasta
        folder.createFile(pdfEntries[i].blob.setName(pdfEntries[i].name));
        looseFiles.push(pdfEntries[i].item + ' (não mesclável: ' + pdfEntries[i].name + ')');
      }
    }

    // 2ª passada: capa (agora com o total de páginas) + merge
    _sub_setProgress({ status: 'running', step: 'Gerando capa...', done: 0, total: loadedDocs.length + 1 });
    const coverBlob = _sub_makeCoverPdf({
      projectName: config.projectName,
      projectAddress: config.projectAddress,
      toName: config.toName,
      title: (coverFields && coverFields.title) || config.submittalTitle || '',
      date: (coverFields && coverFields.date) || today
    }, totalItemPages + 1); // +1 = a própria capa

    const merged = await PDFLib.PDFDocument.create();
    const coverDoc = await PDFLib.PDFDocument.load(new Uint8Array(coverBlob.getBytes()));
    let pages = await merged.copyPages(coverDoc, coverDoc.getPageIndices());
    pages.forEach(function(p) { merged.addPage(p); });

    for (let i = 0; i < loadedDocs.length; i++) {
      _sub_setProgress({ status: 'running', step: 'Mesclando: ' + loadedDocs[i].entry.item, done: i + 1, total: loadedDocs.length + 1 });
      pages = await merged.copyPages(loadedDocs[i].doc, loadedDocs[i].doc.getPageIndices());
      pages.forEach(function(p) { merged.addPage(p); });
    }

    const pdfName = folderName + '.pdf';
    const bytes = await merged.save();
    const mergedBlob = Utilities.newBlob([].slice.call(new Int8Array(bytes)), MimeType.PDF, pdfName);

    // Substitui PDF homônimo (re-montagem)
    const oldPdfs = folder.getFilesByName(pdfName);
    while (oldPdfs.hasNext()) oldPdfs.next().setTrashed(true);
    const pdfFile = folder.createFile(mergedBlob);

    // ── Coluna M (LINK SUBMITTAL) em todas as linhas do submittal ──
    _sub_setProgress({ status: 'running', step: 'Gravando links na planilha...', done: loadedDocs.length + 1, total: loadedDocs.length + 1 });
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    // O grid pode terminar na coluna L — garante que a coluna M exista
    // (getRange além do grid lança "coordenadas fora das dimensões da página")
    const outCol = SUBMITTAL_CFG.COLS.LINK_OUT;
    if (sheet.getMaxColumns() < outCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), outCol - sheet.getMaxColumns());
    }
    const headerCell = sheet.getRange(1, outCol);
    if (!String(headerCell.getValue()).trim()) {
      headerCell.setValue('LINK SUBMITTAL').setFontWeight('bold');
    }

    const pdfUrl = pdfFile.getUrl();
    const linkRt = SpreadsheetApp.newRichTextValue().setText(pdfName).setLinkUrl(pdfUrl).build();
    items.forEach(function(it) {
      sheet.getRange(it.row, SUBMITTAL_CFG.COLS.LINK_OUT).setRichTextValue(linkRt);
    });

    const report = {
      total: items.length,
      merged: loadedDocs.length,
      loose: looseFiles,
      noLink: noLink,
      failed: failed,
      pages: totalItemPages + 1,
      folderUrl: folder.getUrl(),
      pdfUrl: pdfUrl,
      pdfName: pdfName
    };
    _sub_setProgress({ status: 'done', report: report });
    return report;

  } catch (e) {
    console.error('[Submittal] ' + e.message + '\n' + (e.stack || ''));
    _sub_setProgress({ status: 'error', message: e.message });
    return { success: false, message: e.message };
  }
}

// ============================================================================
// ENTRY POINT
// ============================================================================

function openSubmittalSidebar() {
  const html = HtmlService.createTemplateFromFile('SubmittalSidebar')
    .evaluate()
    .setTitle('📦 Montar Submittal')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}
