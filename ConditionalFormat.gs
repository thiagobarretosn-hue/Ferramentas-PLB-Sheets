/**
 * @fileoverview Script para aplicar formatação condicional de comparação
 *               entre colunas pareadas (Building vs Project)
 * @version 1.0.0
 *
 * USO: Execute a função setupComparisonFormatting() no editor do Apps Script
 *      ou pelo menu customizado.
 *
 * INTERVALO: D134:L169
 * PARES:
 *   E vs F  → Códigos (destaca faltantes)
 *   G vs H  → Material (destaca maior/menor)
 *   I vs J  → Labor (destaca maior/menor)
 *   K vs L  → Equipment (destaca maior/menor)
 */

// ============================================================
// CONFIGURAÇÕES
// ============================================================
const CF_CONFIG = {
  START_ROW: 134,
  END_ROW: 169,
  COLORS: {
    GREEN:  '#D9EAD3',  // Valores IGUAIS
    RED:    '#F4CCCC',  // Valor MAIOR
    YELLOW: '#FFF2CC',  // Valor MENOR
    ORANGE: '#FCE5CD'   // Código faltante / diferente
  }
};

// ============================================================
// FUNÇÃO PRINCIPAL
// ============================================================

/**
 * Aplica formatação condicional para comparação entre colunas pareadas.
 * Pares: E-F (códigos), G-H (material), I-J (labor), K-L (equipment)
 */
function setupComparisonFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { START_ROW, END_ROW, COLORS } = CF_CONFIG;
  const numRows = END_ROW - START_ROW + 1;

  // Regras existentes (preservar)
  const existingRules = sheet.getConditionalFormatRules();
  const newRules = [];

  // ─── Par E-F: Comparação de CÓDIGOS (faltantes) ───
  _addCodeComparisonRules(sheet, newRules, START_ROW, numRows, COLORS.ORANGE);

  // ─── Pares de VALORES: G-H, I-J, K-L ───
  const valuePairs = [
    { leftCol: 7,  rightCol: 8,  leftLetter: 'G', rightLetter: 'H' },  // Material
    { leftCol: 9,  rightCol: 10, leftLetter: 'I', rightLetter: 'J' },  // Labor
    { leftCol: 11, rightCol: 12, leftLetter: 'K', rightLetter: 'L' }   // Equipment
  ];

  valuePairs.forEach(pair => {
    _addValueComparisonRules(sheet, newRules, START_ROW, numRows, pair, COLORS);
  });

  // Aplicar: novas regras primeiro (prioridade), depois as existentes
  sheet.setConditionalFormatRules([...newRules, ...existingRules]);

  const totalRules = newRules.length;
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${totalRules} regras criadas para o intervalo ${START_ROW}:${END_ROW}`,
    'Formatação Condicional Aplicada',
    5
  );
}

/**
 * Remove TODAS as regras de formatação condicional criadas por este script
 * no intervalo D134:L169. Útil para limpar e reaplicar.
 */
function clearComparisonFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { START_ROW, END_ROW } = CF_CONFIG;

  const rules = sheet.getConditionalFormatRules();
  const filtered = rules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => {
      const row = r.getRow();
      const lastRow = r.getLastRow();
      const col = r.getColumn();
      return row >= START_ROW && lastRow <= END_ROW && col >= 4 && col <= 12;
    });
  });

  sheet.setConditionalFormatRules(filtered);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Regras removidas do intervalo ${START_ROW}:${END_ROW}`,
    'Formatação Condicional Removida',
    3
  );
}

// ============================================================
// FUNÇÕES AUXILIARES
// ============================================================

/**
 * Adiciona regras para comparação de códigos (E vs F)
 * - Laranja quando código existe num lado mas falta no outro
 * @private
 */
function _addCodeComparisonRules(sheet, rules, startRow, numRows, yellowColor) {
  const rangeE = sheet.getRange(startRow, 5, numRows, 1);  // E
  const rangeF = sheet.getRange(startRow, 6, numRows, 1);  // F

  // E tem valor mas F está vazio → falta no projeto
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(E' + startRow + '<>"", F' + startRow + '="")')
      .setBackground(yellowColor)
      .setRanges([rangeE])
      .build()
  );

  // F tem valor mas E está vazio → falta no building
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(F' + startRow + '<>"", E' + startRow + '="")')
      .setBackground(yellowColor)
      .setRanges([rangeF])
      .build()
  );

  // E e F iguais (ambos preenchidos e batem) → verde
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(E' + startRow + '<>"", F' + startRow + '<>"", E' + startRow + '=F' + startRow + ')')
      .setBackground(CF_CONFIG.COLORS.GREEN)
      .setRanges([rangeE, rangeF])
      .build()
  );

  // E e F diferentes (ambos preenchidos mas não batem) → laranja
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(E' + startRow + '<>"", F' + startRow + '<>"", E' + startRow + '<>F' + startRow + ')')
      .setBackground(CF_CONFIG.COLORS.ORANGE)
      .setRanges([rangeE, rangeF])
      .build()
  );
}

/**
 * Adiciona regras de comparação de valores (maior/menor) para um par de colunas
 * @private
 */
function _addValueComparisonRules(sheet, rules, startRow, numRows, pair, colors) {
  const rangeLeft  = sheet.getRange(startRow, pair.leftCol, numRows, 1);
  const rangeRight = sheet.getRange(startRow, pair.rightCol, numRows, 1);
  const L = pair.leftLetter;
  const R = pair.rightLetter;
  const r = startRow;

  // Ambos IGUAIS → verde nos dois lados
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + L + r + '=' + R + r)
      .setBackground(colors.GREEN)
      .setRanges([rangeLeft, rangeRight])
      .build()
  );

  // Coluna esquerda (Building) MAIOR → vermelho
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + L + r + '>' + R + r)
      .setBackground(colors.RED)
      .setRanges([rangeLeft])
      .build()
  );

  // Coluna esquerda (Building) MENOR → amarelo
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + L + r + '<' + R + r)
      .setBackground(colors.YELLOW)
      .setRanges([rangeLeft])
      .build()
  );

  // Coluna direita (Project) MAIOR → vermelho
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + R + r + '>' + L + r)
      .setBackground(colors.RED)
      .setRanges([rangeRight])
      .build()
  );

  // Coluna direita (Project) MENOR → amarelo
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + R + r + '<' + L + r)
      .setBackground(colors.YELLOW)
      .setRanges([rangeRight])
      .build()
  );
}
