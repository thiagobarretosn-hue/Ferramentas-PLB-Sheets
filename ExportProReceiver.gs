// ExportPro — Google Apps Script doPost
// Deploy como web app standalone em script.google.com
// Execute as: Me | Who has access: Anyone
// Versao 3.0 — token obrigatorio + allowlist de pasta + auditoria
//
// POR QUE "Anyone" E NAO "Anyone within domain":
//   Trocar para "Anyone within domain" NAO funciona com este cliente. O GAS
//   passa a exigir OAuth do usuario do dominio; o cliente e um
//   System.Net.WebClient anonimo (sem cookie, sem Authorization: Bearer),
//   entao o Google devolve a PAGINA DE LOGIN em HTML e a chamada nunca chega
//   ao doPost. A restricao de acesso, aqui, e feita no payload (token) e no
//   destino (allowlist de pasta) — nao na configuracao do deployment.
//
// O QUE MUDOU NA v3.0 e por que:
//   Ate a v2.3 o doPost aceitava QUALQUER chamada e fazia openById() do que
//   viesse no payload, rodando como o dono do script. Na pratica: quem tivesse
//   a URL podia escrever em qualquer planilha que o dono pode editar, bastando
//   saber o ID. Agora exige token e so grava dentro das pastas permitidas.
//
// SETUP (rodar UMA VEZ no editor, depois apagar os valores do historico):
//   configurarAcesso('<token-secreto>', ['<idPasta1>', '<idPasta2>'], '<idPlanilhaDeLog>')
//   Depois: Implantar > Gerenciar implantacoes > nova versao.
//   Os valores ficam em PropertiesService — nunca no codigo, nunca no git.

var PROP_TOKEN = 'EXPORTPRO_TOKEN';
var PROP_PASTAS = 'EXPORTPRO_PASTAS';       // JSON array de folder IDs
var PROP_LOG_ID = 'EXPORTPRO_LOG_ID';       // planilha de auditoria (opcional)
var PROP_TRANSICAO = 'EXPORTPRO_TRANSICAO'; // '1' = aceita sem token, so audita
var PROFUNDIDADE_MAX = 5;                   // niveis de subpasta aceitos
var CACHE_DESTINO_S = 600;                  // 10 min

/**
 * Modo de transicao: aceita chamada SEM token, mas registra na auditoria.
 * Serve para o intervalo entre implantar esta versao e todos os engenheiros
 * receberem o ExportPro atualizado — sem isso, quem ainda esta na versao
 * antiga para de exportar no minuto do deploy.
 * A allowlist de pasta continua valendo em modo de transicao: o buraco grave
 * (escrever em qualquer planilha do dono) fecha imediatamente.
 * DESLIGAR assim que a distribuicao terminar: modoTransicao(false)
 */
function modoTransicao(ligado) {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_TRANSICAO, ligado ? '1' : '0');
  return ligado
    ? 'TRANSICAO LIGADA — aceitando envios sem token (auditados). Desligar depois.'
    : 'Transicao desligada — token obrigatorio.';
}

function _emTransicao() {
  return PropertiesService.getScriptProperties().getProperty(PROP_TRANSICAO) === '1';
}

/**
 * Setup unico. Guarda token, pastas permitidas e planilha de log.
 * @param {string} token Segredo compartilhado com o cliente pyRevit.
 * @param {string[]} pastasIds IDs das pastas do Drive onde e permitido gravar.
 * @param {string} logSpreadsheetId Planilha de auditoria (pode ser '').
 */
function configurarAcesso(token, pastasIds, logSpreadsheetId) {
  if (!token || String(token).length < 16) {
    throw new Error('Token muito curto — use pelo menos 16 caracteres aleatorios.');
  }
  if (!pastasIds || !pastasIds.length) {
    throw new Error('Informe ao menos uma pasta permitida.');
  }
  PropertiesService.getScriptProperties().setProperties({
    EXPORTPRO_TOKEN: String(token),
    EXPORTPRO_PASTAS: JSON.stringify(pastasIds),
    EXPORTPRO_LOG_ID: String(logSpreadsheetId || '')
  });
  return 'Configurado: ' + pastasIds.length + ' pasta(s) permitida(s).';
}

/** Mostra a configuracao atual sem revelar o token. */
function verificarConfiguracao() {
  var p = PropertiesService.getScriptProperties();
  var token = p.getProperty(PROP_TOKEN);
  return {
    tokenDefinido: !!token,
    tokenTamanho: token ? token.length : 0,
    pastas: JSON.parse(p.getProperty(PROP_PASTAS) || '[]'),
    logDefinido: !!p.getProperty(PROP_LOG_ID)
  };
}

function _tokenValido(recebido) {
  var esperado = PropertiesService.getScriptProperties().getProperty(PROP_TOKEN);
  if (!esperado) return false;          // sem setup, ninguem entra
  if (!recebido) return false;
  if (String(recebido).length !== esperado.length) return false;
  // comparacao de tempo constante: nao vaza o tamanho do prefixo correto
  var diff = 0;
  for (var i = 0; i < esperado.length; i++) {
    diff |= esperado.charCodeAt(i) ^ String(recebido).charCodeAt(i);
  }
  return diff === 0;
}

/**
 * A planilha esta dentro de alguma pasta permitida (ou subpasta dela)?
 * E esta checagem — nao o token — que impede usar o web app como proxy de
 * escrita para qualquer planilha que o dono do script consegue editar.
 */
function _destinoPermitido(spreadsheetId) {
  var permitidas = JSON.parse(
    PropertiesService.getScriptProperties().getProperty(PROP_PASTAS) || '[]');
  if (!permitidas.length) return false;

  var cache = CacheService.getScriptCache();
  var chave = 'destino_' + spreadsheetId;
  var memo = cache.get(chave);
  if (memo !== null) return memo === '1';

  var alvo = {};
  for (var i = 0; i < permitidas.length; i++) alvo[permitidas[i]] = true;

  var ok = false;
  try {
    var nivel = [DriveApp.getFileById(spreadsheetId)];
    for (var d = 0; d < PROFUNDIDADE_MAX && !ok && nivel.length; d++) {
      var acima = [];
      for (var n = 0; n < nivel.length; n++) {
        var pais = nivel[n].getParents();
        while (pais.hasNext()) {
          var pasta = pais.next();
          if (alvo[pasta.getId()]) { ok = true; break; }
          acima.push(pasta);
        }
        if (ok) break;
      }
      nivel = acima;
    }
  } catch (err) {
    ok = false;   // sem acesso ao arquivo = destino nao permitido
  }

  cache.put(chave, ok ? '1' : '0', CACHE_DESTINO_S);
  return ok;
}

/** Registra a chamada na planilha de auditoria, se houver. */
function _auditar(dados) {
  var logId = PropertiesService.getScriptProperties().getProperty(PROP_LOG_ID);
  if (!logId) return;
  try {
    var ss = SpreadsheetApp.openById(logId);
    var aba = ss.getSheetByName('ExportPro-Log');
    if (!aba) {
      aba = ss.insertSheet('ExportPro-Log');
      aba.appendRow(['quando', 'projeto', 'spreadsheetId', 'planilha',
                     'modo', 'linhas', 'resultado']);
    }
    aba.appendRow([new Date(), dados.projeto || '', dados.spreadsheetId || '',
                   dados.planilha || '', dados.modo || '', dados.linhas || 0,
                   dados.resultado || '']);
  } catch (err) {
    // auditoria nunca pode derrubar a operacao principal
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _uniqueTabName(ss, baseName) {
  if (!ss.getSheetByName(baseName)) return baseName;
  var n = 2;
  while (ss.getSheetByName(baseName + '-' + n)) n++;
  return baseName + '-' + n;
}

function doPost(e) {
  var spreadsheetId = '';
  var projeto = '';
  try {
    var payload = JSON.parse(e.postData.contents);
    projeto = payload.projectName || '';
    spreadsheetId = payload.spreadsheetId || '';

    // --- portao 1: token -------------------------------------------------
    var comToken = _tokenValido(payload.token);
    if (!comToken) {
      if (!_emTransicao()) {
        _auditar({ projeto: projeto, spreadsheetId: spreadsheetId,
                   resultado: 'recusado: token' });
        return _json({ status: 'error', code: 'token',
                       message: 'Token ausente ou invalido. Configure o token do ExportPro.' });
      }
      // transicao: deixa passar, mas fica no log quem ainda nao atualizou
      _auditar({ projeto: projeto, spreadsheetId: spreadsheetId,
                 resultado: 'TRANSICAO: aceito sem token' });
    }

    if (!spreadsheetId) throw new Error('spreadsheetId ausente no payload.');

    // --- portao 2: destino ------------------------------------------------
    if (!_destinoPermitido(spreadsheetId)) {
      _auditar({ projeto: projeto, spreadsheetId: spreadsheetId,
                 resultado: 'recusado: fora da pasta permitida' });
      return _json({ status: 'error', code: 'destino',
                     message: 'A planilha de destino nao esta em uma pasta autorizada. ' +
                              'Mova a planilha para a pasta do ExportPro ou peca a inclusao da pasta.' });
    }

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var mode = payload.mode || 'separate';
    var schedules = payload.schedules || [];

    if (mode === 'stacked') {
      // Todos os schedules empilhados em uma unica aba
      var tabName = payload.projectName || 'Dados';
      tabName = tabName.substring(0, 90).replace(/[\[\]:*?\/\\]/g, '_');
      tabName = _uniqueTabName(ss, tabName);
      var sheet = ss.insertSheet(tabName);
      var allRows = [];
      for (var i = 0; i < schedules.length; i++) {
        if (i > 0) allRows.push(['']);  // linha em branco entre schedules
        var s = schedules[i];
        var incHeader = s.includeHeader !== false;
        if (incHeader && s.headers && s.headers.length > 0) {
          allRows.push(s.headers);
        }
        for (var r = 0; r < s.rows.length; r++) {
          allRows.push(s.rows[r]);
        }
      }
      if (allRows.length > 0) {
        var maxCols = allRows.reduce(function(m, row) { return Math.max(m, row.length); }, 0);
        var padded = allRows.map(function(row) {
          var copy = row.slice();
          while (copy.length < maxCols) copy.push('');
          return copy;
        });
        sheet.getRange(1, 1, padded.length, maxCols).setValues(padded);
      }
      _auditar({ projeto: projeto, spreadsheetId: spreadsheetId, planilha: ss.getName(),
                 modo: 'stacked', linhas: allRows.length, resultado: 'ok' });
      return _json({ status: 'ok', mode: 'stacked', rows: allRows.length,
                     spreadsheet: ss.getName(), tabName: tabName });
    }

    // mode === 'append': insert below existing content in a named tab
    if (mode === 'append') {
      var appendTabName = payload.tabName;
      if (!appendTabName) throw new Error('tabName missing for mode=append.');
      var appendSheet = ss.getSheetByName(appendTabName);
      if (!appendSheet) throw new Error('Tab "' + appendTabName + '" not found in spreadsheet.');
      var appendRows = [];
      for (var ai = 0; ai < schedules.length; ai++) {
        if (ai > 0) appendRows.push(['']);
        var as = schedules[ai];
        var aIncHeader = as.includeHeader !== false;
        if (aIncHeader && as.headers && as.headers.length > 0) {
          appendRows.push(as.headers);
        }
        for (var ar = 0; ar < as.rows.length; ar++) {
          appendRows.push(as.rows[ar]);
        }
      }
      if (appendRows.length > 0) {
        var maxColsA = appendRows.reduce(function(m, row) { return Math.max(m, row.length); }, 0);
        if (appendSheet.getMaxColumns() < maxColsA) {
          appendSheet.insertColumnsAfter(appendSheet.getMaxColumns(),
                                         maxColsA - appendSheet.getMaxColumns());
        }

        // Encontra o ultimo row com dado SOMENTE nas colunas que serao preenchidas (A:maxCols).
        // Ignora formulas/dados em outras colunas para nao inflar a posicao de insercao.
        var totalRows = appendSheet.getLastRow();
        var startRow = 1;
        if (totalRows > 0) {
          var colValues = appendSheet.getRange(1, 1, totalRows, maxColsA).getValues();
          for (var ri = colValues.length - 1; ri >= 0; ri--) {
            var hasData = false;
            for (var ci = 0; ci < colValues[ri].length; ci++) {
              if (colValues[ri][ci] !== '' && colValues[ri][ci] !== null) {
                hasData = true;
                break;
              }
            }
            if (hasData) {
              startRow = ri + 2; // ri e 0-based; +1 para 1-based, +1 para proxima linha
              break;
            }
          }
        }

        var paddedA = appendRows.map(function(row) {
          var copy = row.slice();
          while (copy.length < maxColsA) copy.push('');
          return copy;
        });
        appendSheet.getRange(startRow, 1, paddedA.length, maxColsA).setValues(paddedA);
      }
      _auditar({ projeto: projeto, spreadsheetId: spreadsheetId, planilha: ss.getName(),
                 modo: 'append', linhas: appendRows.length, resultado: 'ok' });
      return _json({ status: 'ok', mode: 'append', tabName: appendTabName,
                     rows: appendRows.length, spreadsheet: ss.getName() });
    }

    // mode === 'separate': uma aba por schedule
    var sheetsWritten = 0;
    var linhasTotal = 0;
    for (var j = 0; j < schedules.length; j++) {
      var sj = schedules[j];
      var tabNameJ = (sj.name || ('Sheet' + (j + 1))).substring(0, 90);
      tabNameJ = tabNameJ.replace(/[\[\]:*?\/\\]/g, '_');
      tabNameJ = _uniqueTabName(ss, tabNameJ);
      var sheetJ = ss.insertSheet(tabNameJ);
      var rows = [];
      var incHeaderJ = sj.includeHeader !== false;
      if (incHeaderJ && sj.headers && sj.headers.length > 0) {
        rows.push(sj.headers);
      }
      for (var rj = 0; rj < sj.rows.length; rj++) {
        rows.push(sj.rows[rj]);
      }
      if (rows.length > 0) {
        var maxColsJ = rows.reduce(function(m, row) { return Math.max(m, row.length); }, 0);
        var paddedJ = rows.map(function(row) {
          var copy = row.slice();
          while (copy.length < maxColsJ) copy.push('');
          return copy;
        });
        sheetJ.getRange(1, 1, paddedJ.length, maxColsJ).setValues(paddedJ);
      }
      linhasTotal += rows.length;
      sheetsWritten++;
    }
    _auditar({ projeto: projeto, spreadsheetId: spreadsheetId, planilha: ss.getName(),
               modo: 'separate', linhas: linhasTotal, resultado: 'ok' });
    return _json({ status: 'ok', sheets: sheetsWritten, spreadsheet: ss.getName() });

  } catch (err) {
    _auditar({ projeto: projeto, spreadsheetId: spreadsheetId,
               resultado: 'erro: ' + err.toString() });
    return _json({ status: 'error', message: err.toString() });
  }
}
