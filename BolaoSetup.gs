/**
 * @fileoverview Bolão Copa do Mundo 2026 — Configuração e Setup
 * @version 1.0.0
 *
 * Changelog:
 * - v1.0.0: Versão inicial com setup, gestão de jogadores e integração com API
 */

// ============================================================
// CONFIGURAÇÃO GLOBAL DO BOLÃO
// ============================================================
const BOLAO = {
  ABAS: {
    JOGOS: '📋 JOGOS',
    CLASSIFICACAO: '🏆 CLASSIFICAÇÃO',
    CONFIG: '_BOLAO_CONFIG',
  },
  // Pontuação flat em todas as fases
  PONTUACAO: { EXATO: 3, VENCEDOR: 1 },
  // Índices de coluna (base 1) — aba JOGOS
  CJ: { ID: 1, FASE: 2, GRUPO: 3, DATA: 4, HORA: 5, TIME_A: 6, TIME_B: 7, GOLS_A: 8, GOLS_B: 9, STATUS: 10 },
  // Índices de coluna (base 1) — aba do jogador
  CP: { ID_JOGO: 1, FASE: 2, DATA: 3, HORA: 4, PARTIDA: 5, PAL_A: 6, PAL_B: 7, LOCK: 8, RESULTADO: 9, ACERTO: 10, PONTOS: 11 },
  // Índices de coluna (base 1) — aba CLASSIFICAÇÃO
  CC: { POS: 1, NOME: 2, EMAIL: 3, PONTOS: 4, EXATOS: 5, VENCEDORES: 6, PALPITES: 7, PERC: 8 },
  STATUS: {
    AGENDADO: 'AGENDADO',
    EM_ANDAMENTO: 'EM ANDAMENTO',
    ENCERRADO: 'ENCERRADO',
    ADIADO: 'ADIADO',
  },
  API: {
    BASE_URL: 'https://api.football-data.org/v4',
    COMPETITION: 'WC',
    PROP_KEY: 'BOLAO_API_KEY',
  },
  CORES: {
    HEADER_JOGOS: '#1565c0',
    HEADER_CLASS: '#e65100',
    HEADER_JOGADOR: '#1b5e20',
    FASE: {
      GRUPOS:         '#e3f2fd',
      RODADA_32:      '#fff3e0',
      OITAVAS:        '#fce4ec',
      QUARTAS:        '#f3e5f5',
      SEMI:           '#e8f5e9',
      TERCEIRO_LUGAR: '#efebe9',
      FINAL:          '#fff9c4',
    },
  },
};

// ============================================================
// SETUP PRINCIPAL
// ============================================================

function setupBolao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    '🏆 Setup — Bolão Copa do Mundo 2026',
    'Isso criará as abas:\n• 📋 JOGOS (todos os 104 jogos)\n• 🏆 CLASSIFICAÇÃO\n\n⚠️ Abas existentes com esses nomes serão recriadas.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    _criarAbaConfig(ss);
    _criarAbaJogos(ss);
    _criarAbaClassificacao(ss);
    SpreadsheetApp.flush();

    const apiResp = ui.prompt(
      '🔑 Chave de API — football-data.org (gratuita)',
      'Para busca automática de resultados, cadastre-se grátis em:\nhttps://www.football-data.org/client/register\n\nCole sua chave abaixo (ou deixe em branco para modo manual):',
      ui.ButtonSet.OK_CANCEL
    );

    let modoApi = false;
    if (apiResp.getSelectedButton() === ui.Button.OK && apiResp.getResponseText().trim()) {
      const apiKey = apiResp.getResponseText().trim();
      PropertiesService.getScriptProperties().setProperty(BOLAO.API.PROP_KEY, apiKey);

      const ok = buscarEPopularJogos();
      if (ok) {
        _configurarTriggerHorario();
        modoApi = true;
      } else {
        ui.alert(
          '⚠️ API não retornou dados',
          'A estrutura foi criada com jogos placeholder.\nVerifique sua chave e tente "Bolão > 🔄 Buscar/Atualizar Resultados".',
          ui.ButtonSet.OK
        );
        _popularJogosPlaceholder(ss.getSheetByName(BOLAO.ABAS.JOGOS));
      }
    } else {
      _popularJogosPlaceholder(ss.getSheetByName(BOLAO.ABAS.JOGOS));
    }

    ss.setActiveSheet(ss.getSheetByName(BOLAO.ABAS.JOGOS));
    ui.alert(
      '✅ Bolão criado!',
      (modoApi
        ? '✅ Jogos carregados via API.\n🕐 Atualização automática a cada hora.\n\n'
        : '⚠️ Modo manual — preencha times, datas e resultados na aba JOGOS.\n\n') +
      'Próximo passo: use\n"Bolão > ➕ Adicionar Jogador"\npara cadastrar cada participante.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    console.error(`[BOLAO] setupBolao: ${err.message}\n${err.stack}`);
    ui.alert('❌ Erro no setup', err.message, ui.ButtonSet.OK);
  }
}

// ============================================================
// CRIAÇÃO DAS ABAS
// ============================================================

function _criarAbaJogos(ss) {
  let sheet = ss.getSheetByName(BOLAO.ABAS.JOGOS);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(BOLAO.ABAS.JOGOS, 0);

  const headers = ['ID', 'FASE', 'GRUPO / RODADA', 'DATA', 'HORA (BRT)', 'TIME A', 'TIME B', 'GOLS A', 'GOLS B', 'STATUS'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(BOLAO.CORES.HEADER_JOGOS).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 38);

  [60, 130, 140, 100, 90, 160, 160, 75, 75, 120].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.getRange('D2:D250').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('E2:E250').setNumberFormat('HH:mm');
  sheet.getRange('H2:I250').setNumberFormat('0');

  const statusRng = sheet.getRange('J2:J250');
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ENCERRADO').setBackground('#c8e6c9').setFontColor('#1b5e20')
      .setRanges([statusRng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('EM ANDAMENTO').setBackground('#fff9c4').setFontColor('#e65100')
      .setRanges([statusRng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('AGENDADO').setBackground('#e3f2fd').setFontColor('#1565c0')
      .setRanges([statusRng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ADIADO').setBackground('#eeeeee').setFontColor('#757575')
      .setRanges([statusRng]).build(),
  ]);

  return sheet;
}

function _criarAbaClassificacao(ss) {
  let sheet = ss.getSheetByName(BOLAO.ABAS.CLASSIFICACAO);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(BOLAO.ABAS.CLASSIFICACAO, 1);

  const headers = ['POS', 'JOGADOR', 'EMAIL', 'PONTOS', 'EXATOS\n(3 pts)', 'VENCEDOR\n(1 pt)', 'PALPITES\nFEITOS', '% ACERTO'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(BOLAO.CORES.HEADER_CLASS).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true).setFontSize(10);

  sheet.setRowHeight(1, 55);
  sheet.setFrozenRows(1);

  [50, 160, 220, 90, 85, 85, 85, 90].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.getRange('D2:D60').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.getRange('H2:H60').setNumberFormat('0.0%').setHorizontalAlignment('center');
  sheet.getRange('A2:G60').setHorizontalAlignment('center');

  return sheet;
}

function _criarAbaConfig(ss) {
  let sheet = ss.getSheetByName(BOLAO.ABAS.CONFIG);
  if (sheet) return sheet;
  sheet = ss.insertSheet(BOLAO.ABAS.CONFIG);
  sheet.hideSheet();
  sheet.getRange('A1:B1').setValues([['NOME', 'EMAIL']]).setFontWeight('bold');
  return sheet;
}

// ============================================================
// JOGOS PLACEHOLDER (modo sem API)
// Copa do Mundo 2026 — 104 jogos: 72 grupos + 32 eliminatórias
// Preencha os times reais na aba JOGOS após o setup
// ============================================================

function _popularJogosPlaceholder(sheet) {
  if (!sheet) return;
  const rows = [];
  let id = 1;

  // Fase de Grupos: 12 grupos (A–L), 4 times cada, 6 jogos por grupo = 72 jogos
  'ABCDEFGHIJKL'.split('').forEach(g => {
    // Combinações: (T1×T2, T3×T4, T1×T3, T2×T4, T1×T4, T2×T3)
    [[1,2],[3,4],[1,3],[2,4],[1,4],[2,3]].forEach(([a, b]) => {
      rows.push([id++, 'GRUPOS', `Grupo ${g}`, '', '', `${g}${a}`, `${g}${b}`, '', '', 'AGENDADO']);
    });
  });

  // Rodada de 32: 16 jogos
  for (let i = 1; i <= 16; i++) {
    rows.push([id++, 'RODADA_32', `Rodada 32 — Jogo ${i}`, '', '', `R32/${i}A`, `R32/${i}B`, '', '', 'AGENDADO']);
  }
  // Oitavas de Final: 8 jogos
  for (let i = 1; i <= 8; i++) {
    rows.push([id++, 'OITAVAS', `Oitavas — Jogo ${i}`, '', '', `OIT/${i}A`, `OIT/${i}B`, '', '', 'AGENDADO']);
  }
  // Quartas de Final: 4 jogos
  for (let i = 1; i <= 4; i++) {
    rows.push([id++, 'QUARTAS', `Quartas — Jogo ${i}`, '', '', `QRT/${i}A`, `QRT/${i}B`, '', '', 'AGENDADO']);
  }
  // Semifinais: 2 jogos
  rows.push([id++, 'SEMI', 'Semifinal 1', '', '', 'Semi-1/A', 'Semi-1/B', '', '', 'AGENDADO']);
  rows.push([id++, 'SEMI', 'Semifinal 2', '', '', 'Semi-2/A', 'Semi-2/B', '', '', 'AGENDADO']);
  // 3º Lugar e Final
  rows.push([id++, 'TERCEIRO_LUGAR', '3º Lugar', '', '', '3Lug/A', '3Lug/B', '', '', 'AGENDADO']);
  rows.push([id++, 'FINAL', 'Final', '', '', 'Final/A', 'Final/B', '', '', 'AGENDADO']);

  sheet.getRange(2, 1, rows.length, 10).setValues(rows);
  _colorirFasesJogos(sheet, rows.length);
}

function _colorirFasesJogos(sheet, totalRows) {
  if (totalRows <= 0) return;
  const faseData = sheet.getRange(2, BOLAO.CJ.FASE, totalRows, 1).getValues();

  // Agrupa linhas consecutivas da mesma fase para batch de setBackground
  let startRow = 2, currentFase = faseData[0]?.[0] || '', count = 0;
  faseData.forEach((row, i) => {
    if (row[0] === currentFase) {
      count++;
    } else {
      if (count > 0) sheet.getRange(startRow, 1, count, 10).setBackground(BOLAO.CORES.FASE[currentFase] || '#f9f9f9');
      startRow = 2 + i;
      currentFase = row[0];
      count = 1;
    }
  });
  if (count > 0) sheet.getRange(startRow, 1, count, 10).setBackground(BOLAO.CORES.FASE[currentFase] || '#f9f9f9');
}

// ============================================================
// GESTÃO DE JOGADORES
// ============================================================

function adicionarJogadorBolao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!ss.getSheetByName(BOLAO.ABAS.JOGOS)) {
    ui.alert('❌ Execute o Setup primeiro!', 'Use "Bolão > ⚙️ Setup Inicial do Bolão".', ui.ButtonSet.OK);
    return;
  }

  const nomeResp = ui.prompt('➕ Novo Participante', 'Nome do jogador:', ui.ButtonSet.OK_CANCEL);
  if (nomeResp.getSelectedButton() !== ui.Button.OK) return;
  const nome = nomeResp.getResponseText().trim();
  if (!nome) return;

  const emailResp = ui.prompt('➕ Novo Participante', `E-mail Google de ${nome}:`, ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const email = emailResp.getResponseText().trim().toLowerCase();

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    ui.alert('❌ E-mail inválido', 'Informe um e-mail Google válido.', ui.ButtonSet.OK);
    return;
  }

  try {
    _criarAbaJogador(ss, nome, email);
    _salvarJogadorConfig(ss, nome, email);
    ui.alert(
      '✅ Participante adicionado!',
      `Aba criada para ${nome}.\nApenas ${email} poderá registrar palpites.\n\n` +
      '⚠️ O participante precisa ter acesso à planilha para editar sua aba.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    console.error(`[BOLAO] adicionarJogador: ${err.message}`);
    ui.alert('❌ Erro', err.message, ui.ButtonSet.OK);
  }
}

function _criarAbaJogador(ss, nome, email) {
  const nomeAba = `🎯 ${nome}`;
  if (ss.getSheetByName(nomeAba)) throw new Error(`Jogador "${nome}" já existe. Use outro nome.`);

  const sheet = ss.insertSheet(nomeAba);

  // Linha de título
  sheet.getRange(1, 1, 1, 11).merge()
    .setValue(`🏆 PALPITES DE ${nome.toUpperCase()} — COPA DO MUNDO 2026`)
    .setBackground(BOLAO.CORES.HEADER_JOGADOR).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  sheet.setRowHeight(1, 42);

  // Cabeçalhos das colunas
  const headers = ['ID', 'FASE', 'DATA', 'HORA\n(BRT)', 'PARTIDA', '⚽ PLACAR\nTIME A', '⚽ PLACAR\nTIME B', '🔒', 'RESULTADO\nOFICIAL', 'ACERTO', 'PTS'];
  sheet.getRange(2, 1, 1, 11)
    .setValues([headers])
    .setBackground('#37474f').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true).setFontSize(9);
  sheet.setRowHeight(2, 42);
  sheet.setFrozenRows(2);

  [55, 110, 90, 70, 215, 100, 100, 38, 110, 120, 65].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  _popularAbaJogador(sheet, ss);

  // Proteção: somente o jogador (e o admin) podem editar
  const protection = sheet.protect().setDescription(`Aba de palpites — ${nome}`);
  protection.removeEditors(protection.getEditors());
  protection.addEditor(email);
  const admin = Session.getEffectiveUser().getEmail();
  if (admin && admin !== email) protection.addEditor(admin);
}

function _popularAbaJogador(sheet, ss) {
  const jogosSheet = ss.getSheetByName(BOLAO.ABAS.JOGOS);
  if (!jogosSheet || jogosSheet.getLastRow() < 2) return;

  const jogosData = jogosSheet.getRange(2, 1, jogosSheet.getLastRow() - 1, 10).getValues();
  const rows = jogosData.map(j => [
    j[BOLAO.CJ.ID - 1],
    j[BOLAO.CJ.FASE - 1],
    j[BOLAO.CJ.DATA - 1],
    j[BOLAO.CJ.HORA - 1],
    `${j[BOLAO.CJ.TIME_A - 1]}  ×  ${j[BOLAO.CJ.TIME_B - 1]}`,
    '',  // PALPITE A (editável)
    '',  // PALPITE B (editável)
    '',  // LOCK (sistema)
    '',  // RESULTADO (sistema)
    '',  // ACERTO (sistema)
    '',  // PONTOS (sistema)
  ]);

  if (!rows.length) return;

  const startRow = 3;
  sheet.getRange(startRow, 1, rows.length, 11).setValues(rows);

  // Formatação de data/hora
  sheet.getRange(startRow, BOLAO.CP.DATA, rows.length, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(startRow, BOLAO.CP.HORA, rows.length, 1).setNumberFormat('HH:mm');

  // Colunas de palpite: fundo verde, texto grande, centralizado
  sheet.getRange(startRow, BOLAO.CP.PAL_A, rows.length, 2)
    .setBackground('#e8f5e9').setHorizontalAlignment('center')
    .setFontSize(13).setFontWeight('bold');

  // Colunas do sistema: fundo cinza discreto
  sheet.getRange(startRow, 1, rows.length, 5).setBackground('#f5f5f5').setFontColor('#546e7a');
  sheet.getRange(startRow, BOLAO.CP.LOCK, rows.length, 4).setBackground('#f5f5f5').setFontColor('#546e7a');

  // Coloração por fase
  _colorirFasesJogador(sheet, rows, startRow);
}

function _colorirFasesJogador(sheet, rows, startRow) {
  let currentFase = rows[0]?.[1] || '', count = 0, start = startRow;
  rows.forEach((row, i) => {
    if (row[1] === currentFase) {
      count++;
    } else {
      if (count > 0) {
        const cor = BOLAO.CORES.FASE[currentFase] || '#ffffff';
        sheet.getRange(start, 1, count, 5).setBackground(cor);
        sheet.getRange(start, BOLAO.CP.LOCK, count, 4).setBackground(cor);
      }
      start = startRow + i;
      currentFase = row[1];
      count = 1;
    }
  });
  if (count > 0) {
    const cor = BOLAO.CORES.FASE[currentFase] || '#ffffff';
    sheet.getRange(start, 1, count, 5).setBackground(cor);
    sheet.getRange(start, BOLAO.CP.LOCK, count, 4).setBackground(cor);
  }
}

function _salvarJogadorConfig(ss, nome, email) {
  const sheet = ss.getSheetByName(BOLAO.ABAS.CONFIG);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[nome, email]]);
}

// ============================================================
// CONFIGURAÇÃO DO TRIGGER HORÁRIO
// ============================================================

function _configurarTriggerHorario() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'buscarResultadosAutomatico') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('buscarResultadosAutomatico')
    .timeBased().everyHours(1).create();
  console.log('[BOLAO] Trigger horário configurado.');
}
