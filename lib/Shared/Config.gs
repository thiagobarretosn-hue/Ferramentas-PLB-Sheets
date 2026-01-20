/**
 * @fileoverview Sistema de Configuracao Centralizado
 * @version 1.0.0
 *
 * Gerencia configuracoes persistentes usando PropertiesService.
 * Permite externalizar valores que antes eram hardcoded.
 *
 * USO:
 *   // Obter valor
 *   const id = AppConfig.get('CENTRAL_SPREADSHEET_ID');
 *
 *   // Definir valor
 *   AppConfig.set('CENTRAL_SPREADSHEET_ID', 'novo-id');
 *
 *   // Obter com fallback
 *   const sheet = AppConfig.get('SOURCE_SHEET', 'Dados');
 */

// ============================================================================
// VALORES PADRAO (FALLBACK)
// Se nao houver valor configurado, usa estes
// ============================================================================

const APP_CONFIG_DEFAULTS = {
  // Template.gs - ID da planilha central de templates
  CENTRAL_SPREADSHEET_ID: '1KD_NXdtEjORJJFYwiFc9uzrZeO57TPPSXWh7bokZMdU',
  CENTRAL_SHEET_NAME: 'DATA BASE',

  // SuperBusca.gs - Configuracoes padrao
  SUPER_BUSCA_SOURCE_SHEET: '5.COST.LIST',
  SUPER_BUSCA_DEFAULT_COL: 6,

  // Geral
  DEFAULT_CACHE_TTL: 300,      // 5 minutos
  DEBUG_MODE: false,
  LOG_LEVEL: 'ERROR'           // DEBUG, INFO, WARN, ERROR
};

// ============================================================================
// SERVICO DE CONFIGURACAO
// ============================================================================

/**
 * Servico de configuracao da aplicacao
 * Usa ScriptProperties para persistencia entre execucoes
 */
const AppConfig = {

  /**
   * Prefixo para chaves no PropertiesService
   * @private
   */
  _PREFIX: 'PLB_CONFIG_',

  /**
   * Cache em memoria para evitar leituras repetidas
   * @private
   */
  _cache: null,

  /**
   * Obtem valor de configuracao
   *
   * @param {string} key - Chave da configuracao
   * @param {*} [defaultValue] - Valor padrao se nao encontrar
   * @returns {*} Valor da configuracao
   *
   * @example
   * const id = AppConfig.get('CENTRAL_SPREADSHEET_ID');
   * const debug = AppConfig.get('DEBUG_MODE', false);
   */
  get(key, defaultValue) {
    // Primeiro verifica cache em memoria
    if (this._cache === null) {
      this._loadCache();
    }

    if (this._cache.hasOwnProperty(key)) {
      return this._cache[key];
    }

    // Fallback para valor padrao do sistema
    if (APP_CONFIG_DEFAULTS.hasOwnProperty(key)) {
      return APP_CONFIG_DEFAULTS[key];
    }

    // Fallback para valor fornecido
    return defaultValue;
  },

  /**
   * Define valor de configuracao
   *
   * @param {string} key - Chave da configuracao
   * @param {*} value - Valor a definir
   * @returns {boolean} True se sucesso
   *
   * @example
   * AppConfig.set('CENTRAL_SPREADSHEET_ID', 'novo-id');
   */
  set(key, value) {
    try {
      const props = PropertiesService.getScriptProperties();
      const stringValue = JSON.stringify(value);
      props.setProperty(this._PREFIX + key, stringValue);

      // Atualiza cache
      if (this._cache === null) {
        this._cache = {};
      }
      this._cache[key] = value;

      return true;
    } catch (error) {
      console.error(`[AppConfig] Erro ao salvar ${key}: ${error.message}`);
      return false;
    }
  },

  /**
   * Define multiplos valores de uma vez
   *
   * @param {Object} configs - Objeto com pares chave-valor
   * @returns {boolean} True se sucesso
   *
   * @example
   * AppConfig.setAll({
   *   CENTRAL_SPREADSHEET_ID: 'id',
   *   DEBUG_MODE: true
   * });
   */
  setAll(configs) {
    try {
      const props = PropertiesService.getScriptProperties();
      const toSave = {};

      for (const key in configs) {
        toSave[this._PREFIX + key] = JSON.stringify(configs[key]);
      }

      props.setProperties(toSave);

      // Atualiza cache
      if (this._cache === null) {
        this._cache = {};
      }
      Object.assign(this._cache, configs);

      return true;
    } catch (error) {
      console.error(`[AppConfig] Erro ao salvar configs: ${error.message}`);
      return false;
    }
  },

  /**
   * Remove uma configuracao (volta ao valor padrao)
   *
   * @param {string} key - Chave a remover
   */
  remove(key) {
    try {
      const props = PropertiesService.getScriptProperties();
      props.deleteProperty(this._PREFIX + key);

      if (this._cache) {
        delete this._cache[key];
      }
    } catch (error) {
      console.error(`[AppConfig] Erro ao remover ${key}: ${error.message}`);
    }
  },

  /**
   * Retorna todas as configuracoes (incluindo defaults)
   *
   * @returns {Object} Todas as configuracoes
   */
  getAll() {
    if (this._cache === null) {
      this._loadCache();
    }

    return { ...APP_CONFIG_DEFAULTS, ...this._cache };
  },

  /**
   * Reseta todas as configuracoes para valores padrao
   */
  resetToDefaults() {
    try {
      const props = PropertiesService.getScriptProperties();
      const allProps = props.getProperties();

      // Remove apenas propriedades com nosso prefixo
      const keysToDelete = Object.keys(allProps)
        .filter(key => key.startsWith(this._PREFIX));

      keysToDelete.forEach(key => props.deleteProperty(key));

      this._cache = {};
      return true;
    } catch (error) {
      console.error(`[AppConfig] Erro ao resetar: ${error.message}`);
      return false;
    }
  },

  /**
   * Carrega configuracoes do PropertiesService para cache
   * @private
   */
  _loadCache() {
    this._cache = {};

    try {
      const props = PropertiesService.getScriptProperties();
      const allProps = props.getProperties();

      for (const fullKey in allProps) {
        if (fullKey.startsWith(this._PREFIX)) {
          const key = fullKey.substring(this._PREFIX.length);
          try {
            this._cache[key] = JSON.parse(allProps[fullKey]);
          } catch (e) {
            // Se nao for JSON valido, usa como string
            this._cache[key] = allProps[fullKey];
          }
        }
      }
    } catch (error) {
      console.error(`[AppConfig] Erro ao carregar cache: ${error.message}`);
    }
  },

  /**
   * Invalida cache (forca recarga na proxima leitura)
   */
  invalidateCache() {
    this._cache = null;
  }
};

// ============================================================================
// FUNCOES DE CONVENIENCIA
// ============================================================================

/**
 * Obtem ID da planilha central de templates
 * @returns {string} ID da planilha
 */
function getConfigCentralSpreadsheetId() {
  return AppConfig.get('CENTRAL_SPREADSHEET_ID');
}

/**
 * Obtem nome da aba central de templates
 * @returns {string} Nome da aba
 */
function getConfigCentralSheetName() {
  return AppConfig.get('CENTRAL_SHEET_NAME');
}

/**
 * Verifica se modo debug esta ativo
 * @returns {boolean} True se debug ativo
 */
function isDebugMode() {
  return AppConfig.get('DEBUG_MODE', false);
}

/**
 * Abre dialog para configurar parametros do sistema
 * Pode ser chamado do menu
 */
function showConfigDialog() {
  const config = AppConfig.getAll();

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 16px; }
        .form-group { margin-bottom: 12px; }
        label { display: block; margin-bottom: 4px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; box-sizing: border-box; }
        .btn { padding: 10px 20px; margin-right: 8px; cursor: pointer; }
        .btn-primary { background: #1a73e8; color: white; border: none; }
        .btn-secondary { background: #f1f1f1; border: 1px solid #ccc; }
        .section { margin-top: 20px; padding-top: 16px; border-top: 1px solid #ddd; }
        h3 { margin-bottom: 12px; color: #333; }
      </style>
    </head>
    <body>
      <h2>Configuracoes do Sistema</h2>

      <div class="section">
        <h3>Templates</h3>
        <div class="form-group">
          <label>ID da Planilha Central:</label>
          <input type="text" id="CENTRAL_SPREADSHEET_ID" value="${config.CENTRAL_SPREADSHEET_ID || ''}">
        </div>
        <div class="form-group">
          <label>Nome da Aba Central:</label>
          <input type="text" id="CENTRAL_SHEET_NAME" value="${config.CENTRAL_SHEET_NAME || ''}">
        </div>
      </div>

      <div class="section">
        <h3>Super Busca</h3>
        <div class="form-group">
          <label>Aba Padrao:</label>
          <input type="text" id="SUPER_BUSCA_SOURCE_SHEET" value="${config.SUPER_BUSCA_SOURCE_SHEET || ''}">
        </div>
        <div class="form-group">
          <label>Coluna Padrao (numero):</label>
          <input type="number" id="SUPER_BUSCA_DEFAULT_COL" value="${config.SUPER_BUSCA_DEFAULT_COL || 6}">
        </div>
      </div>

      <div class="section">
        <h3>Debug</h3>
        <div class="form-group">
          <label>
            <input type="checkbox" id="DEBUG_MODE" ${config.DEBUG_MODE ? 'checked' : ''}>
            Modo Debug Ativo
          </label>
        </div>
      </div>

      <div style="margin-top: 24px;">
        <button class="btn btn-primary" onclick="saveConfig()">Salvar</button>
        <button class="btn btn-secondary" onclick="google.script.host.close()">Cancelar</button>
        <button class="btn btn-secondary" onclick="resetConfig()">Resetar Padrao</button>
      </div>

      <script>
        function saveConfig() {
          const config = {
            CENTRAL_SPREADSHEET_ID: document.getElementById('CENTRAL_SPREADSHEET_ID').value,
            CENTRAL_SHEET_NAME: document.getElementById('CENTRAL_SHEET_NAME').value,
            SUPER_BUSCA_SOURCE_SHEET: document.getElementById('SUPER_BUSCA_SOURCE_SHEET').value,
            SUPER_BUSCA_DEFAULT_COL: parseInt(document.getElementById('SUPER_BUSCA_DEFAULT_COL').value) || 6,
            DEBUG_MODE: document.getElementById('DEBUG_MODE').checked
          };

          google.script.run
            .withSuccessHandler(() => {
              alert('Configuracoes salvas!');
              google.script.host.close();
            })
            .withFailureHandler(err => alert('Erro: ' + err.message))
            .saveAppConfig(config);
        }

        function resetConfig() {
          if (confirm('Resetar todas as configuracoes para valores padrao?')) {
            google.script.run
              .withSuccessHandler(() => {
                alert('Configuracoes resetadas!');
                google.script.host.close();
              })
              .withFailureHandler(err => alert('Erro: ' + err.message))
              .resetAppConfig();
          }
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Configuracoes');
}

/**
 * Salva configuracoes do dialog
 * @param {Object} config - Configuracoes a salvar
 */
function saveAppConfig(config) {
  return AppConfig.setAll(config);
}

/**
 * Reseta configuracoes para padrao
 */
function resetAppConfig() {
  return AppConfig.resetToDefaults();
}
