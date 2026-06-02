/**
 * @fileoverview Bolão Copa 2026 — Integração com football-data.org
 * @version 1.0.0
 *
 * API gratuita: https://www.football-data.org/client/register
 * Retorna dados do Mundial em tempo real (placares, status, escalações).
 * Trigger horário chama `buscarResultadosAutomatico()` automaticamente.
 */

// Mapeamento: stage da API → fase interna do Bolão
const API_STAGE_MAP = {
  GROUP_STAGE:    'GRUPOS',
  ROUND_OF_32:    'RODADA_32',
  LAST_32:        'RODADA_32',  // nome alternativo na API
  ROUND_OF_16:    'OITAVAS',
  LAST_16:        'OITAVAS',
  QUARTER_FINALS: 'QUARTAS',
  SEMI_FINALS:    'SEMI',
  THIRD_PLACE:    'TERCEIRO_LUGAR',
  FINAL:          'FINAL',
};

// Mapeamento: status da API → status interno do Bolão
const API_STATUS_MAP = {
  TIMED:      'AGENDADO',
  SCHEDULED:  'AGENDADO',
  LIVE:       'EM ANDAMENTO',
  IN_PLAY:    'EM ANDAMENTO',
  PAUSED:     'EM ANDAMENTO',
  HALFTIME:   'EM ANDAMENTO',
  EXTRA_TIME: 'EM ANDAMENTO',
  PENALTY:    'EM ANDAMENTO',
  FINISHED:   'ENCERRADO',
  AWARDED:    'ENCERRADO',
  POSTPONED:  'ADIADO',
  CANCELLED:  'ADIADO',
  SUSPENDED:  'ADIADO',
};

// ============================================================
// BUSCA E POPULAÇÃO INICIAL DOS JOGOS
// ============================================================

/**
 * Busca todos os jogos da Copa 2026 na API e popula a aba JOGOS.
 * Chamada no setup e quando o admin aciona manualmente.
 *
 * @returns {boolean} true se os dados foram carregados com sucesso.
 */
function buscarEPopularJogos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jogosSheet = ss.getSheetByName(BOLAO.ABAS.JOGOS);
  if (!jogosSheet) return false;

  const partidas = _fetchPartidasAPI();
  if (!partidas || !partidas.length) {
    console.warn('[BOLAO] buscarEPopularJogos: nenhuma partida retornada pela API.');
    return false;
  }

  // Limpa dados anteriores (mantém cabeçalho)
  if (jogosSheet.getLastRow() > 1) {
    jogosSheet.getRange(2, 1, jogosSheet.getLastRow() - 1, 10).clearContent().setBackground(null);
  }

  // Ordena por data/fase
  partidas.sort((a, b) => {
    if (!a.dataHora) return 1;
    if (!b.dataHora) return -1;
    return a.dataHora - b.dataHora;
  });

  const rows = partidas.map(p => [
    p.id,
    p.fase,
    p.grupo,
    p.dataHora,   // Date object → exibido como dd/mm/yyyy pela formatação da coluna
    p.dataHora,   // Date object → exibido como HH:mm pela formatação da coluna
    p.timeA,
    p.timeB,
    p.golsA !== null ? p.golsA : '',
    p.golsB !== null ? p.golsB : '',
    p.status,
  ]);

  jogosSheet.getRange(2, 1, rows.length, 10).setValues(rows);
  jogosSheet.getRange(2, BOLAO.CJ.DATA, rows.length, 1).setNumberFormat('dd/mm/yyyy');
  jogosSheet.getRange(2, BOLAO.CJ.HORA, rows.length, 1).setNumberFormat('HH:mm');

  _colorirFasesJogos(jogosSheet, rows.length);
  SpreadsheetApp.flush();

  console.log(`[BOLAO] ${rows.length} partidas carregadas da API.`);
  return true;
}

// ============================================================
// ATUALIZAÇÃO AUTOMÁTICA HORÁRIA (TRIGGER)
// ============================================================

/**
 * Chamada pelo trigger horário.
 * Atualiza placares/status na aba JOGOS e recalcula pontuação.
 */
function buscarResultadosAutomatico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jogosSheet = ss.getSheetByName(BOLAO.ABAS.JOGOS);
  if (!jogosSheet || jogosSheet.getLastRow() < 2) return;

  try {
    const partidas = _fetchPartidasAPI();
    if (!partidas || !partidas.length) return;

    // Indexa partidas da API por ID
    const partidaMap = {};
    partidas.forEach(p => { partidaMap[p.id] = p; });

    const data = jogosSheet.getRange(2, 1, jogosSheet.getLastRow() - 1, 10).getValues();
    let houveMudanca = false;

    data.forEach((row, i) => {
      const apiId  = row[BOLAO.CJ.ID - 1];
      const p      = partidaMap[apiId];
      if (!p) return;

      const sheetRow  = i + 2; // linha na planilha (base 1, +1 pelo header)
      const statusAtual = row[BOLAO.CJ.STATUS - 1];
      const golsAAtual  = row[BOLAO.CJ.GOLS_A - 1];
      const golsBAtual  = row[BOLAO.CJ.GOLS_B - 1];

      const novoStatus = p.status;
      const novoGolsA  = p.golsA !== null ? p.golsA : golsAAtual;
      const novoGolsB  = p.golsB !== null ? p.golsB : golsBAtual;

      const mudouStatus = statusAtual !== novoStatus;
      const mudouPlacar = golsAAtual !== novoGolsA || golsBAtual !== novoGolsB;

      if (mudouStatus || mudouPlacar) {
        jogosSheet.getRange(sheetRow, BOLAO.CJ.GOLS_A, 1, 3).setValues([[novoGolsA, novoGolsB, novoStatus]]);
        houveMudanca = true;
      }

      // Atualiza nomes de times nas fases eliminatórias (TBD → time real)
      const timeAAtual = row[BOLAO.CJ.TIME_A - 1];
      const timeBAtual = row[BOLAO.CJ.TIME_B - 1];
      if (p.fase !== 'GRUPOS' && (timeAAtual !== p.timeA || timeBAtual !== p.timeB)
          && p.timeA !== 'TBD' && p.timeB !== 'TBD') {
        jogosSheet.getRange(sheetRow, BOLAO.CJ.TIME_A, 1, 2).setValues([[p.timeA, p.timeB]]);
        houveMudanca = true;
      }
    });

    if (houveMudanca) {
      SpreadsheetApp.flush();
      _recalcularTudo(ss);
      console.log('[BOLAO] Resultados atualizados e pontuação recalculada automaticamente.');
    }
  } catch (err) {
    console.error(`[BOLAO] buscarResultadosAutomatico: ${err.message}\n${err.stack}`);
  }
}

// ============================================================
// CONFIGURAÇÃO DE CHAVE DE API
// ============================================================

function configurarChaveAPI() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '🔑 Chave da API football-data.org',
    'Cadastre-se gratuitamente em:\nhttps://www.football-data.org/client/register\n\nCole sua chave abaixo:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const key = resp.getResponseText().trim();
  if (!key) return;

  PropertiesService.getScriptProperties().setProperty(BOLAO.API.PROP_KEY, key);

  const fetch = ui.alert('✅ Chave salva!', 'Deseja buscar os jogos agora?', ui.ButtonSet.YES_NO);
  if (fetch === ui.Button.YES) {
    const ok = buscarEPopularJogos();
    if (ok) {
      _configurarTriggerHorario();
      ui.alert(
        '✅ Jogos carregados!',
        'Trigger horário ativado.\nResultados serão atualizados automaticamente a cada hora.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Falha ao buscar dados', 'Verifique sua chave de API e tente novamente.', ui.ButtonSet.OK);
    }
  }
}

// ============================================================
// FETCH E MAPEAMENTO DA API
// ============================================================

function _fetchPartidasAPI() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(BOLAO.API.PROP_KEY);
  if (!apiKey) {
    console.warn('[BOLAO] Chave de API não configurada.');
    return null;
  }

  const url = `${BOLAO.API.BASE_URL}/competitions/${BOLAO.API.COMPETITION}/matches`;
  let response;

  try {
    response = UrlFetchApp.fetch(url, {
      headers: { 'X-Auth-Token': apiKey },
      muteHttpExceptions: true,
    });
  } catch (err) {
    console.error(`[BOLAO] Erro de rede ao chamar API: ${err.message}`);
    return null;
  }

  const code = response.getResponseCode();
  if (code !== 200) {
    console.error(`[BOLAO] API retornou HTTP ${code}: ${response.getContentText().substring(0, 400)}`);
    return null;
  }

  let json;
  try {
    json = JSON.parse(response.getContentText());
  } catch (err) {
    console.error('[BOLAO] Erro ao parsear resposta JSON da API.');
    return null;
  }

  const matches = json.matches || json.value || [];
  return matches.map(_mapearPartida).filter(Boolean);
}

function _mapearPartida(match) {
  try {
    const fase = API_STAGE_MAP[match.stage] || 'GRUPOS';

    // "GROUP_A" → "Grupo A", "SEMI_FINALS" → "Semifinais", etc.
    let grupo = match.group || '';
    if (grupo.startsWith('GROUP_')) {
      grupo = `Grupo ${grupo.replace('GROUP_', '')}`;
    } else if (!grupo) {
      const labelMap = {
        RODADA_32: 'Rodada de 32',
        OITAVAS:   'Oitavas de Final',
        QUARTAS:   'Quartas de Final',
        SEMI:      'Semifinal',
        TERCEIRO_LUGAR: '3º Lugar',
        FINAL:     'Final',
      };
      grupo = labelMap[fase] || fase;
    }

    // Data/hora UTC → Date object (o fuso da planilha cuida da exibição)
    const dataHora = match.utcDate ? new Date(match.utcDate) : null;

    // Placar final (suporta formatos v4 e legado)
    const score    = match.score || {};
    const fullTime = score.fullTime || score.fulltime || {};
    const golsA    = fullTime.home  ?? fullTime.homeTeam  ?? null;
    const golsB    = fullTime.away  ?? fullTime.awayTeam  ?? null;

    return {
      id:      match.id,
      fase,
      grupo,
      dataHora,
      timeA:   match.homeTeam?.shortName || match.homeTeam?.name || 'TBD',
      timeB:   match.awayTeam?.shortName || match.awayTeam?.name || 'TBD',
      golsA:   golsA !== null && golsA !== undefined ? Number(golsA) : null,
      golsB:   golsB !== null && golsB !== undefined ? Number(golsB) : null,
      status:  API_STATUS_MAP[match.status] || 'AGENDADO',
    };
  } catch (err) {
    console.error(`[BOLAO] _mapearPartida: ${err.message} — match.id=${match?.id}`);
    return null;
  }
}
