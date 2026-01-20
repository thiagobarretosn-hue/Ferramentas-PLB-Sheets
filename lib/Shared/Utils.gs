/**
 * @fileoverview Utilitários Compartilhados - Ferramentas PLB Sheets
 * @version 1.0.0
 *
 * Este arquivo centraliza funções utilitárias usadas por múltiplos módulos.
 * IMPORTANTE: Não modifique as assinaturas das funções para manter compatibilidade.
 *
 * Módulos que usam este arquivo:
 * - BOM.gs (via Utils object)
 * - Template.gs (funções diretas)
 * - SuperBusca.gs (funções diretas)
 */

// ============================================================================
// CONVERSÃO DE COLUNAS
// ============================================================================

/**
 * Converte número de coluna para letra(s)
 * Exemplos: 1 → "A", 26 → "Z", 27 → "AA", 703 → "AAA"
 *
 * @param {number} columnNumber - Número da coluna (1-indexed)
 * @returns {string} Letra(s) da coluna
 */
function SharedUtils_numberToColumnLetter(columnNumber) {
  if (!columnNumber || columnNumber < 1) return '';

  let result = '';
  let num = columnNumber;

  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }

  return result;
}

/**
 * Converte letra(s) de coluna para número
 * Exemplos: "A" → 1, "Z" → 26, "AA" → 27, "AAA" → 703
 *
 * @param {string} columnLetter - Letra(s) da coluna
 * @returns {number} Número da coluna (1-indexed)
 */
function SharedUtils_columnLetterToIndex(columnLetter) {
  if (!columnLetter) return 0;

  const letters = String(columnLetter).toUpperCase().replace(/[^A-Z]/g, '');
  if (!letters) return 0;

  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }

  return index;
}

/**
 * Extrai índice de coluna de configuração no formato "A - Nome" ou "A"
 * Usado pelo sistema BOM
 *
 * @param {string} colConfig - Configuração da coluna (ex: "A - Nome da Coluna" ou "A")
 * @returns {number} Índice da coluna (1-indexed) ou -1 se inválido
 */
function SharedUtils_getColumnIndexFromConfig(colConfig) {
  if (!colConfig || typeof colConfig !== 'string') return -1;

  // Extrai apenas a primeira letra (ou letras até o espaço/hífen)
  const match = colConfig.match(/^([A-Za-z]+)/);
  if (!match) return -1;

  return SharedUtils_columnLetterToIndex(match[1]);
}

/**
 * Extrai nome do cabeçalho de configuração no formato "A - Nome"
 * Usado pelo sistema BOM
 *
 * @param {string} colConfig - Configuração da coluna (ex: "A - Nome da Coluna")
 * @returns {string} Nome do cabeçalho ou string vazia
 */
function SharedUtils_getColumnHeaderFromConfig(colConfig) {
  if (!colConfig || typeof colConfig !== 'string') return '';

  const parts = colConfig.split(' - ');
  return parts.length > 1 ? parts.slice(1).join(' - ').trim() : '';
}

// ============================================================================
// MANIPULAÇÃO DE STRINGS
// ============================================================================

/**
 * Sanitiza nome de aba removendo caracteres inválidos do Google Sheets
 * Caracteres inválidos: / \ ? * [ ] :
 *
 * @param {string} name - Nome original
 * @returns {string} Nome sanitizado (max 100 caracteres)
 */
function SharedUtils_sanitizeSheetName(name) {
  if (!name) return '';

  return String(name)
    .replace(/[\/\\\?\*\[\]:]/g, '_')
    .replace(/_+/g, '_')           // Remove underscores duplicados
    .replace(/^_|_$/g, '')         // Remove underscores no início/fim
    .substring(0, 100);            // Limite do Google Sheets
}

/**
 * Formata versão com zeros à esquerda
 * Exemplos: "1" → "01", "12" → "12", "" → ""
 *
 * @param {string|number} input - Versão de entrada
 * @returns {string} Versão formatada
 */
function SharedUtils_formatVersion(input) {
  if (input === null || input === undefined || input === '') return '';

  const str = String(input).trim();
  if (!str) return '';

  const num = parseInt(str, 10);
  if (isNaN(num) || num < 1) return str;

  return num < 10 ? `0${num}` : `${num}`;
}

/**
 * Cria ID seguro a partir de texto (para uso em HTML/CSS)
 * Remove caracteres especiais, substitui por underscore
 *
 * @param {string} text - Texto original
 * @returns {string} ID seguro
 */
function SharedUtils_createSafeId(text) {
  if (!text) return '';

  return String(text)
    .replace(/[^A-Za-z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

// ============================================================================
// VALIDAÇÃO
// ============================================================================

/**
 * Valida e converte para inteiro dentro de um range
 *
 * @param {*} value - Valor a validar
 * @param {number} [min=1] - Valor mínimo
 * @param {number} [max=10000] - Valor máximo
 * @returns {number} Valor validado
 * @throws {Error} Se valor estiver fora do range
 */
function SharedUtils_validateInteger(value, min = 1, max = 10000) {
  if (value === null || value === undefined || value === '') return min;

  const number = parseInt(value, 10);

  if (isNaN(number)) {
    throw new Error(`Valor inválido: "${value}". Esperado um número.`);
  }

  if (number < min || number > max) {
    throw new Error(`Valor ${number} fora do range permitido (${min}-${max}).`);
  }

  return number;
}

/**
 * Verifica se valor é vazio (null, undefined, string vazia ou só espaços)
 *
 * @param {*} value - Valor a verificar
 * @returns {boolean} True se vazio
 */
function SharedUtils_isEmpty(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  return false;
}

// ============================================================================
// EXTRAÇÃO DE DADOS
// ============================================================================

/**
 * Extrai diâmetro de descrição de tubo
 * Usado pelo sistema de fixadores do BOM
 *
 * @param {string} desc - Descrição do item (ex: "PIPE 1-1/2 IN...")
 * @returns {string|null} Diâmetro encontrado ou null
 */
function SharedUtils_extractPipeDiameter(desc) {
  if (!desc) return null;

  const patterns = [
    /PIPE\s+(\d+-\d+\/\d+)\s+IN/i,   // 1-1/2
    /PIPE\s+(\d+\/\d+)\s+IN/i,       // 1/2
    /PIPE\s+(\d+)\s+IN/i             // 1, 2, etc
  ];

  for (const pattern of patterns) {
    const match = String(desc).match(pattern);
    if (match) {
      const diam = match[1].replace(/\s+/g, '');

      // Mapeamento para valores padrão
      const validDiameters = [
        '1/2', '3/4', '1', '1-1/4', '1-1/2', '2',
        '2-1/2', '3', '4', '6', '8', '10', '12'
      ];

      return validDiameters.includes(diam) ? diam : null;
    }
  }

  return null;
}

// ============================================================================
// NAMESPACE PARA COMPATIBILIDADE
// ============================================================================

/**
 * Namespace SharedUtils para acesso organizado
 * Pode ser usado como: SharedUtils.columnToLetter(1)
 */
const SharedUtils = {
  // Conversão de colunas
  numberToColumnLetter: SharedUtils_numberToColumnLetter,
  columnLetterToIndex: SharedUtils_columnLetterToIndex,
  getColumnIndexFromConfig: SharedUtils_getColumnIndexFromConfig,
  getColumnHeaderFromConfig: SharedUtils_getColumnHeaderFromConfig,

  // Strings
  sanitizeSheetName: SharedUtils_sanitizeSheetName,
  formatVersion: SharedUtils_formatVersion,
  createSafeId: SharedUtils_createSafeId,

  // Validação
  validateInteger: SharedUtils_validateInteger,
  isEmpty: SharedUtils_isEmpty,

  // Extração
  extractPipeDiameter: SharedUtils_extractPipeDiameter,

  // Respostas padronizadas
  success: SharedUtils_successResponse,
  error: SharedUtils_errorResponse,
  wrapAsync: SharedUtils_wrapAsync
};

// ============================================================================
// RESPOSTAS PADRONIZADAS
// Helpers para criar respostas consistentes em todas as funcoes publicas
// ============================================================================

/**
 * Cria resposta de sucesso padronizada
 *
 * @param {*} [data] - Dados a retornar
 * @param {string} [message] - Mensagem opcional
 * @returns {Object} { success: true, data?, message? }
 *
 * @example
 * return SharedUtils_successResponse({ items: 10 });
 * // { success: true, data: { items: 10 } }
 */
function SharedUtils_successResponse(data, message) {
  const response = { success: true };

  if (data !== undefined) {
    response.data = data;
  }

  if (message) {
    response.message = message;
  }

  return response;
}

/**
 * Cria resposta de erro padronizada
 *
 * @param {string|Error} error - Mensagem de erro ou objeto Error
 * @param {string} [code] - Codigo de erro opcional
 * @returns {Object} { success: false, error, code? }
 *
 * @example
 * return SharedUtils_errorResponse('Aba nao encontrada');
 * // { success: false, error: 'Aba nao encontrada' }
 *
 * return SharedUtils_errorResponse(new Error('Falha'), 'NOT_FOUND');
 * // { success: false, error: 'Falha', code: 'NOT_FOUND' }
 */
function SharedUtils_errorResponse(error, code) {
  const response = {
    success: false,
    error: error instanceof Error ? error.message : String(error)
  };

  if (code) {
    response.code = code;
  }

  return response;
}

/**
 * Wrapper para funcoes async que podem falhar
 * Captura erros e retorna resposta padronizada
 *
 * @param {Function} fn - Funcao a executar
 * @param {string} [context] - Contexto para log de erro
 * @returns {Object} Resposta padronizada
 *
 * @example
 * function minhaFuncao(params) {
 *   return SharedUtils_wrapAsync(() => {
 *     // codigo que pode falhar
 *     return dados;
 *   }, 'minhaFuncao');
 * }
 */
function SharedUtils_wrapAsync(fn, context) {
  try {
    const result = fn();
    return SharedUtils_successResponse(result);

  } catch (error) {
    // Log do erro com contexto
    const errorMsg = error instanceof Error ? error.message : String(error);
    console.error(`[ERRO${context ? ` em ${context}` : ''}] ${errorMsg}`);

    if (error instanceof Error && error.stack) {
      console.error(`Stack: ${error.stack}`);
    }

    return SharedUtils_errorResponse(error);
  }
}

// ============================================================================
// SANITIZACAO HTML
// Previne XSS quando dados da planilha sao exibidos em sidebars
// ============================================================================

/**
 * Escapa caracteres HTML para prevenir XSS
 *
 * @param {string} str - String a sanitizar
 * @returns {string} String com caracteres HTML escapados
 *
 * @example
 * SharedUtils_escapeHtml('<script>alert("xss")</script>');
 * // '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;'
 */
function SharedUtils_escapeHtml(str) {
  if (str === null || str === undefined) return '';

  const htmlEntities = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
    '/': '&#x2F;',
    '`': '&#x60;',
    '=': '&#x3D;'
  };

  return String(str).replace(/[&<>"'`=\/]/g, char => htmlEntities[char]);
}

/**
 * Sanitiza array de dados para exibicao em HTML
 *
 * @param {Array} data - Array de dados (pode ser 2D)
 * @returns {Array} Array com valores sanitizados
 */
function SharedUtils_sanitizeArrayForHtml(data) {
  if (!Array.isArray(data)) return [];

  return data.map(item => {
    if (Array.isArray(item)) {
      return item.map(val => SharedUtils_escapeHtml(val));
    }
    return SharedUtils_escapeHtml(item);
  });
}

/**
 * Sanitiza objeto para exibicao em HTML
 *
 * @param {Object} obj - Objeto a sanitizar
 * @returns {Object} Objeto com valores string sanitizados
 */
function SharedUtils_sanitizeObjectForHtml(obj) {
  if (!obj || typeof obj !== 'object') return obj;

  const sanitized = {};

  for (const key in obj) {
    const value = obj[key];

    if (typeof value === 'string') {
      sanitized[key] = SharedUtils_escapeHtml(value);
    } else if (Array.isArray(value)) {
      sanitized[key] = SharedUtils_sanitizeArrayForHtml(value);
    } else if (typeof value === 'object' && value !== null) {
      sanitized[key] = SharedUtils_sanitizeObjectForHtml(value);
    } else {
      sanitized[key] = value;
    }
  }

  return sanitized;
}

// Adicionar ao namespace SharedUtils
SharedUtils.escapeHtml = SharedUtils_escapeHtml;
SharedUtils.sanitizeArray = SharedUtils_sanitizeArrayForHtml;
SharedUtils.sanitizeObject = SharedUtils_sanitizeObjectForHtml;

// ============================================================================
// VALIDACAO DE TIPOS
// Funcoes robustas para conversao e validacao de tipos
// ============================================================================

/**
 * Converte valor para numero de forma segura
 * Trata strings, numeros, null, undefined e valores invalidos
 *
 * @param {*} value - Valor a converter
 * @param {number} [defaultValue=0] - Valor padrao se conversao falhar
 * @returns {number} Numero convertido ou valor padrao
 *
 * @example
 * SharedUtils_toNumber('123.45') // 123.45
 * SharedUtils_toNumber('abc', -1) // -1
 * SharedUtils_toNumber(null) // 0
 * SharedUtils_toNumber('  42  ') // 42
 */
function SharedUtils_toNumber(value, defaultValue = 0) {
  // Null, undefined, empty string
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }

  // Ja e numero
  if (typeof value === 'number') {
    return isNaN(value) || !isFinite(value) ? defaultValue : value;
  }

  // Boolean
  if (typeof value === 'boolean') {
    return value ? 1 : 0;
  }

  // String - limpar e converter
  if (typeof value === 'string') {
    const cleaned = value.trim();
    if (cleaned === '') return defaultValue;

    // Tenta converter
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) || !isFinite(parsed) ? defaultValue : parsed;
  }

  // Outros tipos
  return defaultValue;
}

/**
 * Converte valor para inteiro de forma segura
 *
 * @param {*} value - Valor a converter
 * @param {number} [defaultValue=0] - Valor padrao se conversao falhar
 * @returns {number} Inteiro convertido ou valor padrao
 *
 * @example
 * SharedUtils_toInteger('123.7') // 123
 * SharedUtils_toInteger('abc') // 0
 */
function SharedUtils_toInteger(value, defaultValue = 0) {
  const num = SharedUtils_toNumber(value, defaultValue);
  return Math.floor(num);
}

/**
 * Converte valor para inteiro positivo (>= 0)
 *
 * @param {*} value - Valor a converter
 * @param {number} [defaultValue=0] - Valor padrao se conversao falhar ou negativo
 * @returns {number} Inteiro positivo ou valor padrao
 */
function SharedUtils_toPositiveInteger(value, defaultValue = 0) {
  const num = SharedUtils_toInteger(value, defaultValue);
  return num >= 0 ? num : defaultValue;
}

/**
 * Verifica se valor e um numero valido (finito, nao NaN)
 *
 * @param {*} value - Valor a verificar
 * @returns {boolean} True se for numero valido
 *
 * @example
 * SharedUtils_isValidNumber(123) // true
 * SharedUtils_isValidNumber('123') // true
 * SharedUtils_isValidNumber('abc') // false
 * SharedUtils_isValidNumber(NaN) // false
 * SharedUtils_isValidNumber(Infinity) // false
 */
function SharedUtils_isValidNumber(value) {
  if (value === null || value === undefined || value === '') {
    return false;
  }

  const num = typeof value === 'number' ? value : parseFloat(value);
  return !isNaN(num) && isFinite(num);
}

/**
 * Converte valor para string de forma segura
 *
 * @param {*} value - Valor a converter
 * @param {string} [defaultValue=''] - Valor padrao se null/undefined
 * @returns {string} String convertida
 */
function SharedUtils_toString(value, defaultValue = '') {
  if (value === null || value === undefined) {
    return defaultValue;
  }
  return String(value);
}

/**
 * Converte valor para boolean de forma segura
 * Reconhece: true, false, 1, 0, 'true', 'false', 'sim', 'nao', 'yes', 'no'
 *
 * @param {*} value - Valor a converter
 * @param {boolean} [defaultValue=false] - Valor padrao se nao reconhecido
 * @returns {boolean} Boolean convertido
 */
function SharedUtils_toBoolean(value, defaultValue = false) {
  if (value === null || value === undefined) {
    return defaultValue;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  if (typeof value === 'string') {
    const lower = value.toLowerCase().trim();
    if (['true', '1', 'sim', 'yes', 's', 'y'].includes(lower)) {
      return true;
    }
    if (['false', '0', 'nao', 'no', 'n'].includes(lower)) {
      return false;
    }
  }

  return defaultValue;
}

/**
 * Garante que valor esta dentro de um range
 *
 * @param {number} value - Valor a verificar
 * @param {number} min - Valor minimo
 * @param {number} max - Valor maximo
 * @returns {number} Valor dentro do range (clampado)
 *
 * @example
 * SharedUtils_clamp(5, 1, 10) // 5
 * SharedUtils_clamp(-5, 1, 10) // 1
 * SharedUtils_clamp(15, 1, 10) // 10
 */
function SharedUtils_clamp(value, min, max) {
  const num = SharedUtils_toNumber(value, min);
  return Math.min(Math.max(num, min), max);
}

/**
 * Valida se valor e um array
 *
 * @param {*} value - Valor a verificar
 * @returns {boolean} True se for array
 */
function SharedUtils_isArray(value) {
  return Array.isArray(value);
}

/**
 * Converte valor para array de forma segura
 * Se ja for array, retorna como esta
 * Se for valor unico, retorna array com esse valor
 * Se for null/undefined, retorna array vazio
 *
 * @param {*} value - Valor a converter
 * @returns {Array} Array resultante
 */
function SharedUtils_toArray(value) {
  if (value === null || value === undefined) {
    return [];
  }
  if (Array.isArray(value)) {
    return value;
  }
  return [value];
}

/**
 * Valida se valor e um objeto (nao null, nao array)
 *
 * @param {*} value - Valor a verificar
 * @returns {boolean} True se for objeto valido
 */
function SharedUtils_isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

// Adicionar ao namespace SharedUtils
SharedUtils.toNumber = SharedUtils_toNumber;
SharedUtils.toInteger = SharedUtils_toInteger;
SharedUtils.toPositiveInteger = SharedUtils_toPositiveInteger;
SharedUtils.isValidNumber = SharedUtils_isValidNumber;
SharedUtils.toString = SharedUtils_toString;
SharedUtils.toBoolean = SharedUtils_toBoolean;
SharedUtils.clamp = SharedUtils_clamp;
SharedUtils.isArray = SharedUtils_isArray;
SharedUtils.toArray = SharedUtils_toArray;
SharedUtils.isObject = SharedUtils_isObject;
