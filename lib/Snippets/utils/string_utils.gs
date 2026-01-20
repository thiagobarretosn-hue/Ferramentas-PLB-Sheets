/**
 * @fileoverview Utilitários para manipulação de strings
 * @version 1.0.0
 */

/**
 * Utilitários de string
 * @namespace
 */
const StringUtils = {

  /**
   * Remove espaços extras e trim
   * @param {string} str - String a limpar
   * @returns {string} String limpa
   */
  clean(str) {
    if (!str) return '';
    return str.toString().trim().replace(/\s+/g, ' ');
  },

  /**
   * Capitaliza primeira letra
   * @param {string} str - String
   * @returns {string} String capitalizada
   */
  capitalize(str) {
    if (!str) return '';
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  },

  /**
   * Capitaliza cada palavra
   * @param {string} str - String
   * @returns {string} String com palavras capitalizadas
   */
  titleCase(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
  },

  /**
   * Converte para camelCase
   * @param {string} str - String
   * @returns {string} String em camelCase
   */
  toCamelCase(str) {
    if (!str) return '';
    return str
      .toLowerCase()
      .replace(/[^a-zA-Z0-9]+(.)/g, (_, char) => char.toUpperCase());
  },

  /**
   * Converte para snake_case
   * @param {string} str - String
   * @returns {string} String em snake_case
   */
  toSnakeCase(str) {
    if (!str) return '';
    return str
      .replace(/([A-Z])/g, '_$1')
      .toLowerCase()
      .replace(/^_/, '')
      .replace(/[^a-z0-9]+/g, '_');
  },

  /**
   * Converte para kebab-case
   * @param {string} str - String
   * @returns {string} String em kebab-case
   */
  toKebabCase(str) {
    if (!str) return '';
    return str
      .replace(/([A-Z])/g, '-$1')
      .toLowerCase()
      .replace(/^-/, '')
      .replace(/[^a-z0-9]+/g, '-');
  },

  /**
   * Trunca string com ellipsis
   * @param {string} str - String
   * @param {number} maxLength - Tamanho máximo
   * @param {string} [suffix='...'] - Sufixo
   * @returns {string} String truncada
   */
  truncate(str, maxLength, suffix = '...') {
    if (!str || str.length <= maxLength) return str || '';
    return str.slice(0, maxLength - suffix.length) + suffix;
  },

  /**
   * Pad left com zeros ou caractere
   * @param {string|number} value - Valor
   * @param {number} length - Tamanho final
   * @param {string} [char='0'] - Caractere de pad
   * @returns {string} String com pad
   */
  padLeft(value, length, char = '0') {
    return String(value).padStart(length, char);
  },

  /**
   * Pad right
   * @param {string|number} value - Valor
   * @param {number} length - Tamanho final
   * @param {string} [char=' '] - Caractere de pad
   * @returns {string} String com pad
   */
  padRight(value, length, char = ' ') {
    return String(value).padEnd(length, char);
  },

  /**
   * Verifica se string contém substring (case insensitive)
   * @param {string} str - String a buscar
   * @param {string} search - Termo de busca
   * @returns {boolean} Se contém
   */
  contains(str, search) {
    if (!str || !search) return false;
    return str.toLowerCase().includes(search.toLowerCase());
  },

  /**
   * Escapa caracteres HTML
   * @param {string} str - String
   * @returns {string} String escapada
   */
  escapeHtml(str) {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },

  /**
   * Remove acentos
   * @param {string} str - String
   * @returns {string} String sem acentos
   */
  removeAccents(str) {
    if (!str) return '';
    return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  },

  /**
   * Gera slug (URL-friendly)
   * @param {string} str - String
   * @returns {string} Slug
   */
  slugify(str) {
    if (!str) return '';
    return this.removeAccents(str)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '');
  },

  /**
   * Formata número com separadores de milhar
   * @param {number} num - Número
   * @param {string} [decimalSep=','] - Separador decimal
   * @param {string} [thousandSep='.'] - Separador de milhar
   * @returns {string} Número formatado
   */
  formatNumber(num, decimalSep = ',', thousandSep = '.') {
    const parts = num.toString().split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, thousandSep);
    return parts.join(decimalSep);
  },

  /**
   * Extrai números de uma string
   * @param {string} str - String
   * @returns {number[]} Array de números encontrados
   */
  extractNumbers(str) {
    if (!str) return [];
    const matches = str.match(/-?\d+\.?\d*/g);
    return matches ? matches.map(Number) : [];
  },

  /**
   * Verifica se string é vazia ou só espaços
   * @param {string} str - String
   * @returns {boolean} Se é vazia
   */
  isEmpty(str) {
    return !str || str.toString().trim() === '';
  },

  /**
   * Conta ocorrências de substring
   * @param {string} str - String
   * @param {string} search - Substring
   * @returns {number} Contagem
   */
  countOccurrences(str, search) {
    if (!str || !search) return 0;
    return (str.match(new RegExp(search, 'gi')) || []).length;
  }
};

// Exportar para uso global
if (typeof module !== 'undefined') {
  module.exports = StringUtils;
}
