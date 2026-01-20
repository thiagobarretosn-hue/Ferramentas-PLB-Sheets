/**
 * @fileoverview Sistema de Logging Estruturado
 * @version 1.0.0
 *
 * Fornece logging consistente com niveis, timestamp e contexto.
 * Pode ser ativado/desativado via AppConfig.DEBUG_MODE
 *
 * NIVEIS: DEBUG < INFO < WARN < ERROR
 *
 * USO:
 *   Log.debug('Mensagem de debug', 'MeuModulo');
 *   Log.info('Operacao iniciada');
 *   Log.warn('Aviso importante');
 *   Log.error('Erro critico', error);
 */

// ============================================================================
// CONFIGURACAO DE LOGGING
// ============================================================================

const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  OFF: 99
};

// ============================================================================
// SISTEMA DE LOGGING
// ============================================================================

/**
 * Sistema de logging estruturado
 * @namespace
 */
const Log = {

  /**
   * Obtem nivel de log configurado
   * @private
   */
  _getLevel() {
    const levelStr = AppConfig ? AppConfig.get('LOG_LEVEL', 'ERROR') : 'ERROR';
    return LOG_LEVELS[levelStr.toUpperCase()] || LOG_LEVELS.ERROR;
  },

  /**
   * Verifica se modo debug esta ativo
   * @private
   */
  _isDebugMode() {
    return AppConfig ? AppConfig.get('DEBUG_MODE', false) : false;
  },

  /**
   * Formata timestamp
   * @private
   */
  _timestamp() {
    return new Date().toISOString().replace('T', ' ').substring(0, 19);
  },

  /**
   * Formata e envia log
   * @private
   */
  _log(level, levelName, message, context, extra) {
    // Verifica se deve logar
    const currentLevel = this._getLevel();
    if (level < currentLevel) return;

    // Monta mensagem
    const prefix = context ? `[${context}]` : '';
    const timestamp = this._timestamp();
    const fullMessage = `[${levelName}] ${timestamp} ${prefix} ${message}`;

    // Envia para console apropriado
    switch (levelName) {
      case 'ERROR':
        console.error(fullMessage);
        if (extra instanceof Error) {
          console.error(`Stack: ${extra.stack || 'N/A'}`);
        }
        break;
      case 'WARN':
        console.warn(fullMessage);
        break;
      case 'INFO':
        console.info(fullMessage);
        break;
      default:
        console.log(fullMessage);
    }

    // Log extra data se fornecido
    if (extra && !(extra instanceof Error)) {
      console.log('  Data:', JSON.stringify(extra, null, 2));
    }
  },

  /**
   * Log nivel DEBUG - detalhes de desenvolvimento
   * So aparece quando DEBUG_MODE = true E LOG_LEVEL = DEBUG
   *
   * @param {string} message - Mensagem
   * @param {string} [context] - Contexto/Modulo
   * @param {*} [data] - Dados extras
   */
  debug(message, context, data) {
    if (this._isDebugMode()) {
      this._log(LOG_LEVELS.DEBUG, 'DEBUG', message, context, data);
    }
  },

  /**
   * Log nivel INFO - informacoes gerais
   *
   * @param {string} message - Mensagem
   * @param {string} [context] - Contexto/Modulo
   * @param {*} [data] - Dados extras
   */
  info(message, context, data) {
    this._log(LOG_LEVELS.INFO, 'INFO', message, context, data);
  },

  /**
   * Log nivel WARN - avisos
   *
   * @param {string} message - Mensagem
   * @param {string} [context] - Contexto/Modulo
   * @param {*} [data] - Dados extras
   */
  warn(message, context, data) {
    this._log(LOG_LEVELS.WARN, 'WARN', message, context, data);
  },

  /**
   * Log nivel ERROR - erros
   *
   * @param {string} message - Mensagem
   * @param {string|Error} [contextOrError] - Contexto ou objeto Error
   * @param {Error} [error] - Objeto Error se contexto foi fornecido
   */
  error(message, contextOrError, error) {
    if (contextOrError instanceof Error) {
      this._log(LOG_LEVELS.ERROR, 'ERROR', message, null, contextOrError);
    } else {
      this._log(LOG_LEVELS.ERROR, 'ERROR', message, contextOrError, error);
    }
  },

  /**
   * Log de inicio de funcao (para profiling)
   * So aparece em DEBUG
   *
   * @param {string} funcName - Nome da funcao
   * @param {Object} [params] - Parametros recebidos
   * @returns {number} Timestamp de inicio (para calcular duracao)
   */
  start(funcName, params) {
    if (this._isDebugMode()) {
      this.debug(`>>> Iniciando`, funcName, params);
    }
    return Date.now();
  },

  /**
   * Log de fim de funcao (para profiling)
   * So aparece em DEBUG
   *
   * @param {string} funcName - Nome da funcao
   * @param {number} startTime - Timestamp retornado por start()
   * @param {*} [result] - Resultado da funcao
   */
  end(funcName, startTime, result) {
    if (this._isDebugMode()) {
      const duration = Date.now() - startTime;
      this.debug(`<<< Concluido em ${duration}ms`, funcName, result);
    }
  },

  /**
   * Cria logger com contexto fixo
   * Util para usar em um modulo especifico
   *
   * @param {string} context - Nome do modulo/contexto
   * @returns {Object} Logger com contexto
   *
   * @example
   * const log = Log.withContext('BOM');
   * log.info('Processando...'); // [INFO] ... [BOM] Processando...
   */
  withContext(context) {
    const self = this;
    return {
      debug: (msg, data) => self.debug(msg, context, data),
      info: (msg, data) => self.info(msg, context, data),
      warn: (msg, data) => self.warn(msg, context, data),
      error: (msg, err) => self.error(msg, context, err),
      start: (func, params) => self.start(`${context}.${func}`, params),
      end: (func, time, result) => self.end(`${context}.${func}`, time, result)
    };
  }
};

// ============================================================================
// FUNCOES DE CONVENIENCIA
// ============================================================================

/**
 * Log rapido de debug (atalho)
 */
function logDebug(message, context) {
  Log.debug(message, context);
}

/**
 * Log rapido de erro (atalho)
 */
function logError(message, error) {
  Log.error(message, error);
}
