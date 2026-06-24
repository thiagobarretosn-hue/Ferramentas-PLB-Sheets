/**
 * ExportProReceiver.gs
 * Recebe dados de schedules do Revit (ExportPro) via HTTP POST JSON
 * e escreve cada schedule em uma aba separada da planilha ativa.
 *
 * Deploy manual (necessario uma vez):
 *   Extensions > Apps Script > Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone
 *   Copiar a URL gerada (formato: https://script.google.com/macros/s/XXXXX/exec)
 *   Colar essa URL na UI do ExportPro em Revit (Advanced > GAS Config).
 *
 * Payload esperado (JSON):
 * {
 *   "projectName": "NomeDoProjeto",
 *   "schedules": [
 *     {
 *       "name": "Pipe Schedule",
 *       "headers": ["Sistema", "Comprimento"],
 *       "rows": [["Supply", "12.5"], ["Return", "8.0"]]
 *     }
 *   ]
 * }
 *
 * Resposta em sucesso: {"status":"ok","sheets":N}
 * Resposta em erro:    {"status":"error","message":"..."}
 */

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var count   = 0;

    (payload.schedules || []).forEach(function(sched) {
      var tabName = String(sched.name || 'Sheet').substring(0, 100);
      var sheet   = ss.getSheetByName(tabName);

      if (sheet) {
        sheet.clearContents();
        sheet.clearFormats();
      } else {
        sheet = ss.insertSheet(tabName);
      }

      var allRows = [];
      if (sched.headers && sched.headers.length > 0) {
        allRows.push(sched.headers);
      }
      (sched.rows || []).forEach(function(row) { allRows.push(row); });

      if (allRows.length > 0) {
        var nCols = Math.max.apply(null, allRows.map(function(r) { return r.length; }));
        var uniform = allRows.map(function(row) {
          var r = row.slice();
          while (r.length < nCols) r.push('');
          return r;
        });
        sheet.getRange(1, 1, uniform.length, nCols).setValues(uniform);

        if (sched.headers && sched.headers.length > 0) {
          sheet.getRange(1, 1, 1, sched.headers.length)
            .setBackground('#1F4E79')
            .setFontColor('#FFFFFF')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
          for (var c = 1; c <= nCols; c++) sheet.autoResizeColumn(c);
        }
      }
      count++;
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', sheets: count }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
