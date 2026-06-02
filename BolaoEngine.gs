/**
 * @fileoverview Bolão Copa 2026 — Motor de Pontuação, Classificação e onEdit
 * @version 1.0.0
 */

// ============================================================
// RECÁLCULO DE PONTUAÇÃO
// ============================================================

function recalcularPontuacao() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    _recalcularTudo(ss);
    ui.alert('✅ Pontuação recalculada!', 'Classificação atualizada com sucesso.', ui.ButtonSet.OK);
  } catch (err) {
    console.error(`[BOLAO] recalcularPontuacao: ${err.message}\n${err.stack}`);
    ui.alert('❌ Erro', err.message, ui.ButtonSet.OK);
  }
}

function _recalcularTudo(ss) {
  const jogosSheet = ss.getSheetByName(BOLAO.ABAS.JOGOS);
  if (!jogosSheet) throw new Error('Aba JOGOS não encontrada. Execute o Setup primeiro.');

  const jogosMap = _construirJogosMap(jogosSheet);
  const jogadores = _listarJogadores(ss);
  if (!jogadores.length) return;

  const resumo = jogadores.map(({ nome, email, nomeAba }) => {
    const sheet = ss.getSheetByName(nomeAba);
    if (!sheet) return null;
    return _recalcularJogador(sheet, jogosMap, nome, email);
  }).filter(Boolean);

  _atualizarClassificacao(ss, resumo);
  SpreadsheetApp.flush();
}

function _recalcularJogador(sheet, jogosMap, nome, email) {
  if (sheet.getLastRow() < 3) return { nome, email, totalPontos: 0, exatos: 0, vencedores: 0, totalPalpites: 0 };

  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 11).getValues();
  const updates = [];
  let totalPontos = 0, exatos = 0, vencedores = 0, totalPalpites = 0;

  data.forEach(row => {
    const idJogo  = row[BOLAO.CP.ID_JOGO - 1];
    const palpA   = row[BOLAO.CP.PAL_A - 1];
    const palpB   = row[BOLAO.CP.PAL_B - 1];
    const jogo    = jogosMap[idJogo];

    // Jogo não iniciado ainda
    if (!jogo || jogo.status === BOLAO.STATUS.AGENDADO) {
      updates.push(['', '', '', '']);
      return;
    }

    const lock = '🔒';

    // Jogo em andamento: bloqueia mas sem resultado final
    if (jogo.status === BOLAO.STATUS.EM_ANDAMENTO) {
      updates.push([lock, '⏳ Em andamento', '', '']);
      return;
    }

    // Jogo encerrado
    const resultado = `${jogo.golsA} × ${jogo.golsB}`;
    const temPalpite = palpA !== '' && palpB !== '';

    if (!temPalpite) {
      updates.push([lock, resultado, '—', 0]);
      return;
    }

    totalPalpites++;
    const calc = _calcularPontos(Number(palpA), Number(palpB), jogo.golsA, jogo.golsB);
    totalPontos += calc.pontos;
    if (calc.acerto === '🎯 EXATO')    exatos++;
    if (calc.acerto === '✅ VENCEDOR') vencedores++;

    updates.push([lock, resultado, calc.acerto, calc.pontos]);
  });

  sheet.getRange(3, BOLAO.CP.LOCK, updates.length, 4).setValues(updates);
  _colorirAcertos(sheet, updates, 3);

  return { nome, email, totalPontos, exatos, vencedores, totalPalpites };
}

function _calcularPontos(palpA, palpB, golsA, golsB) {
  if (palpA === golsA && palpB === golsB) {
    return { pontos: BOLAO.PONTUACAO.EXATO, acerto: '🎯 EXATO' };
  }
  if (Math.sign(palpA - palpB) === Math.sign(golsA - golsB)) {
    return { pontos: BOLAO.PONTUACAO.VENCEDOR, acerto: '✅ VENCEDOR' };
  }
  return { pontos: 0, acerto: '❌ ERROU' };
}

function _colorirAcertos(sheet, updates, startRow) {
  const corMap = {
    '🎯 EXATO':    { bg: '#a5d6a7', fg: '#1b5e20' },
    '✅ VENCEDOR': { bg: '#c8e6c9', fg: '#2e7d32' },
    '❌ ERROU':    { bg: '#ffcdd2', fg: '#c62828' },
    '—':           { bg: '#eeeeee', fg: '#9e9e9e' },
  };

  updates.forEach((row, i) => {
    const cor = corMap[row[2]];
    if (!cor) return;
    sheet.getRange(startRow + i, BOLAO.CP.ACERTO, 1, 2)
      .setBackground(cor.bg).setFontColor(cor.fg).setHorizontalAlignment('center');
  });
}

// ============================================================
// DADOS DE JOGOS E JOGADORES
// ============================================================

function _construirJogosMap(jogosSheet) {
  const data = jogosSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[BOLAO.CJ.ID - 1];
    if (!id && id !== 0) continue;

    // Monta dataHora para comparação de prazo
    let dataHora = null;
    const dataVal = row[BOLAO.CJ.DATA - 1];
    const horaVal = row[BOLAO.CJ.HORA - 1];
    if (dataVal instanceof Date && horaVal instanceof Date) {
      dataHora = new Date(dataVal);
      dataHora.setHours(horaVal.getHours(), horaVal.getMinutes(), 0, 0);
    }

    map[id] = {
      fase:     row[BOLAO.CJ.FASE - 1],
      timeA:    row[BOLAO.CJ.TIME_A - 1],
      timeB:    row[BOLAO.CJ.TIME_B - 1],
      golsA:    Number(row[BOLAO.CJ.GOLS_A - 1]),
      golsB:    Number(row[BOLAO.CJ.GOLS_B - 1]),
      status:   row[BOLAO.CJ.STATUS - 1],
      dataHora,
    };
  }

  return map;
}

function _listarJogadores(ss) {
  const sheet = ss.getSheetByName(BOLAO.ABAS.CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .filter(r => r[0])
    .map(r => ({ nome: r[0], email: r[1], nomeAba: `🎯 ${r[0]}` }));
}

// ============================================================
// CLASSIFICAÇÃO
// ============================================================

function _atualizarClassificacao(ss, resumo) {
  const sheet = ss.getSheetByName(BOLAO.ABAS.CLASSIFICACAO);
  if (!sheet) return;

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 8)
      .clearContent().setBackground(null).setFontColor('#000000').setFontWeight('normal');
  }

  const ranking = resumo
    .sort((a, b) => b.totalPontos - a.totalPontos || b.exatos - a.exatos || b.vencedores - a.vencedores)
    .map((j, i) => {
      const acertos = j.exatos + j.vencedores;
      const perc = j.totalPalpites > 0 ? acertos / j.totalPalpites : 0;
      return [i + 1, j.nome, j.email, j.totalPontos, j.exatos, j.vencedores, j.totalPalpites, perc];
    });

  if (!ranking.length) return;

  sheet.getRange(2, 1, ranking.length, 8).setValues(ranking);
  sheet.getRange(2, 8, ranking.length, 1).setNumberFormat('0.0%');

  // Destaque top 3
  [
    { bg: '#ffd700', fg: '#5d4037', icon: '🥇' },
    { bg: '#e0e0e0', fg: '#37474f', icon: '🥈' },
    { bg: '#ffcc80', fg: '#4e342e', icon: '🥉' },
  ].forEach(({ bg, fg }, i) => {
    if (i >= ranking.length) return;
    sheet.getRange(2 + i, 1, 1, 8).setBackground(bg).setFontColor(fg).setFontWeight('bold');
  });
}

// ============================================================
// onEdit — Bloquear palpites após início do jogo
// ============================================================

function onEditBolao(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Só atua em abas de jogadores
  if (!sheetName.startsWith('🎯 ')) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (row < 3) return; // ignora cabeçalhos

  // Colunas do sistema: reverte silenciosamente (apenas palpites cols 6 e 7 são editáveis)
  if (col !== BOLAO.CP.PAL_A && col !== BOLAO.CP.PAL_B) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    return;
  }

  // Busca o ID do jogo na linha editada
  const idJogo = sheet.getRange(row, BOLAO.CP.ID_JOGO).getValue();
  if (!idJogo && idJogo !== 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jogosSheet = ss.getSheetByName(BOLAO.ABAS.JOGOS);
  if (!jogosSheet) return;

  const jogosData = jogosSheet.getDataRange().getValues();
  const jogoRow = jogosData.find(r => r[BOLAO.CJ.ID - 1] === idJogo);
  if (!jogoRow) return;

  const status = jogoRow[BOLAO.CJ.STATUS - 1];

  // Bloqueia se jogo já encerrou ou está em andamento
  if (status === BOLAO.STATUS.ENCERRADO || status === BOLAO.STATUS.EM_ANDAMENTO) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    ss.toast(
      `⛔ Palpite não aceito — jogo ${status === BOLAO.STATUS.ENCERRADO ? 'encerrado' : 'em andamento'}.`,
      '🏆 Bolão Copa 2026', 5
    );
    return;
  }

  // Bloqueia se data/hora de início já passou
  const dataVal = jogoRow[BOLAO.CJ.DATA - 1];
  const horaVal = jogoRow[BOLAO.CJ.HORA - 1];

  if (dataVal instanceof Date && horaVal instanceof Date) {
    const jogoDateTime = new Date(dataVal);
    jogoDateTime.setHours(horaVal.getHours(), horaVal.getMinutes(), 0, 0);

    if (new Date() >= jogoDateTime) {
      e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
      ss.toast('⛔ Prazo encerrado! Este jogo já começou.', '🏆 Bolão Copa 2026', 5);
    }
  }
}

// ============================================================
// GUIA DE PONTUAÇÃO
// ============================================================

function mostrarGuiaPontuacao() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; font-size: 13px; color: #333; }
        h2 { color: #1565c0; margin-bottom: 4px; }
        h3 { color: #555; font-size: 12px; margin-top: 16px; margin-bottom: 6px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 12px; }
        td, th { border: 1px solid #ddd; padding: 9px 12px; text-align: center; }
        th { background: #1565c0; color: #fff; font-weight: bold; }
        .exato   { background: #a5d6a7; font-weight: bold; color: #1b5e20; }
        .venced  { background: #c8e6c9; color: #2e7d32; }
        .errou   { background: #ffcdd2; color: #c62828; }
        .note    { font-size: 11px; color: #777; border-left: 3px solid #1565c0; padding-left: 10px; }
        ul { margin: 4px 0; padding-left: 18px; }
      </style>
    </head>
    <body>
      <h2>🏆 Sistema de Pontuação</h2>
      <p style="font-size:12px;color:#666;">Bolão Copa do Mundo 2026 · Pontuação flat em todas as fases</p>

      <table>
        <tr><th>Tipo de Acerto</th><th>Pontos</th><th>Exemplo</th></tr>
        <tr class="exato">
          <td>🎯 Placar Exato</td>
          <td><strong>3 pts</strong></td>
          <td>Palpite: 2×1 · Resultado: 2×1</td>
        </tr>
        <tr class="venced">
          <td>✅ Vencedor / Empate</td>
          <td><strong>1 pt</strong></td>
          <td>Palpite: 3×1 · Resultado: 2×0</td>
        </tr>
        <tr class="errou">
          <td>❌ Errou</td>
          <td><strong>0 pts</strong></td>
          <td>Palpite: 0×1 · Resultado: 2×0</td>
        </tr>
      </table>

      <h3>📌 Regras</h3>
      <ul>
        <li>Palpites <strong>bloqueados automaticamente</strong> quando o jogo começa</li>
        <li>Resultados buscados via API a cada hora automaticamente</li>
        <li>Em caso de empate de pontos: mais acertos exatos (3pts) decide</li>
      </ul>

      <h3>🔒 Proteção dos palpites</h3>
      <p class="note">Cada jogador só pode editar sua própria aba.<br>
      As colunas verdes são os únicos campos editáveis.</p>
    </body>
    </html>
  `).setWidth(520).setHeight(340);

  SpreadsheetApp.getUi().showModalDialog(html, '📊 Guia de Pontuação — Bolão Copa 2026');
}
