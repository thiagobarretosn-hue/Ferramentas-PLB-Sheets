# DIRETRIZES DE DESENVOLVIMENTO - Google Apps Script
> **Versão:** 1.0 | **Atualizado:** 2026-01-20

Este documento define os padrões e práticas obrigatórias para desenvolvimento de Google Apps Script no projeto Ferramentas PLB Sheets.

---

## 1. WORKFLOW OBRIGATÓRIO

### Fase 1: Análise
1. Ler e entender o código existente relacionado
2. Identificar padrões já utilizados no projeto
3. Verificar snippets disponíveis em `lib/Snippets/`
4. Consultar `CLAUDE.md` para erros conhecidos

### Fase 2: Planejamento
1. Definir estrutura da solução
2. Listar serviços GAS necessários
3. Planejar tratamento de erros
4. Considerar impacto em performance

### Fase 3: Implementação
1. Usar template base de `CLAUDE.md`
2. Seguir convenções de nomenclatura
3. Implementar logging adequado
4. Reutilizar snippets existentes

### Fase 4: Validação
1. Testar com dados reais
2. Verificar edge cases
3. Validar performance
4. Documentar no código

---

## 2. ESTRUTURA DE ARQUIVOS

### Arquivos .gs (Google Apps Script)
```
src/
├── Menu.gs           # Menus e triggers
├── BOM.gs            # Sistema de relatórios
├── Template.gs       # Sistema de templates
├── SuperBusca.gs     # Sistema de busca
└── [NovoArquivo].gs  # Novos módulos
```

### Arquivos HTML (Interface)
```
html/
├── BomSidebar.html
├── ConfigSidebar.html
├── template-sidebar.html
└── [NovaSidebar].html
```

### Bibliotecas
```
lib/
└── Snippets/
    ├── cache/
    ├── sheets/
    ├── drive/
    ├── ui/
    └── utils/
```

---

## 3. CONVENÇÕES DE NOMENCLATURA

### Funções
```javascript
// Públicas: camelCase com verbo
function getUserData() {}
function processReport() {}
function validateInput() {}

// Privadas: prefixo underscore
function _parseRow() {}
function _formatCell() {}

// Handlers de UI: prefixo 'on' ou 'handle'
function onButtonClick() {}
function handleFormSubmit() {}

// Callbacks do Google: padrão Google
function onOpen(e) {}
function onEdit(e) {}
function doGet(e) {}
function doPost(e) {}
```

### Variáveis
```javascript
// Variáveis locais: camelCase
let rowCount = 0;
const sheetData = [];

// Constantes: UPPER_SNAKE_CASE
const MAX_ROWS = 10000;
const CACHE_TTL = 300;

// Objetos de configuração: CONFIG ou Settings
const CONFIG = {};
const SETTINGS = {};
```

### Arquivos
```
// Scripts: PascalCase ou descriptivo
Menu.gs
BOM.gs
Template.gs
SuperBusca.gs

// HTML: kebab-case
bom-sidebar.html
config-dialog.html

// Snippets: snake_case com prefixo de categoria
cache_manager.gs
sheet_utils.gs
```

---

## 4. DOCUMENTAÇÃO DE CÓDIGO

### Cabeçalho de Arquivo
```javascript
/**
 * @fileoverview Sistema de geração de relatórios dinâmicos
 * @version 2.13
 * @author PLB Development Team
 *
 * Este módulo gerencia:
 * - Processamento de dados BOM
 * - Geração de relatórios agrupados
 * - Exportação para PDF
 *
 * Changelog:
 * - v2.13: Correção exportação PDF
 * - v2.12: Nomeação por kojoSuffix
 */
```

### Documentação de Função
```javascript
/**
 * Processa dados e gera relatório agrupado
 *
 * @param {string} sourceSheet - Nome da aba de origem
 * @param {Object} grouping - Configuração de agrupamento
 * @param {string} grouping.level1 - Primeiro nível
 * @param {string} [grouping.level2] - Segundo nível (opcional)
 * @returns {Object} Resultado do processamento
 * @returns {boolean} result.success - Se operação teve sucesso
 * @returns {string[]} result.sheets - Abas criadas
 * @throws {Error} Se aba de origem não existir
 *
 * @example
 * const result = processReport('RawData', { level1: 'Category' });
 * if (result.success) {
 *   console.log(`Criadas ${result.sheets.length} abas`);
 * }
 */
function processReport(sourceSheet, grouping) {
  // ...
}
```

---

## 5. TRATAMENTO DE ERROS

### Padrão Try-Catch
```javascript
function operacaoComErro() {
  try {
    // Operação que pode falhar
    const data = fetchData();
    return { success: true, data: data };

  } catch (error) {
    // Log detalhado
    console.error(`[ERRO] ${error.message}`);
    console.error(`Stack: ${error.stack}`);

    // Feedback ao usuário
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Erro: ${error.message}`,
      'Falha na operação',
      5
    );

    // Retorno estruturado
    return { success: false, error: error.message };
  }
}
```

### Validação de Entrada
```javascript
function processData(sheetName, options = {}) {
  // Validar parâmetros obrigatórios
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Nome da aba é obrigatório');
  }

  // Validar existência de recursos
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Aba "${sheetName}" não encontrada`);
  }

  // Continuar com lógica...
}
```

### Controle de Concorrência
```javascript
function operacaoCritica() {
  const lock = LockService.getScriptLock();

  try {
    // Aguardar até 30 segundos pelo lock
    if (!lock.tryLock(30000)) {
      throw new Error('Não foi possível obter acesso exclusivo');
    }

    // Operação crítica aqui

  } finally {
    // SEMPRE liberar o lock
    lock.releaseLock();
  }
}
```

---

## 6. PERFORMANCE

### SEMPRE Usar Batch Operations
```javascript
// ❌ NUNCA fazer isso
for (let i = 0; i < data.length; i++) {
  sheet.getRange(i + 1, 1).setValue(data[i]);
}

// ✅ SEMPRE fazer isso
const values = data.map(item => [item]);
sheet.getRange(1, 1, values.length, 1).setValues(values);
```

### Minimizar Chamadas à API
```javascript
// ❌ Múltiplas chamadas
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName('Data');
const lastRow = sheet.getLastRow();
const lastCol = sheet.getLastColumn();
const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

// ✅ Encadeamento eficiente
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
const data = sheet.getDataRange().getValues();
```

### Usar Cache para Dados Frequentes
```javascript
const CacheManager = {
  _cache: CacheService.getScriptCache(),
  TTL: 300, // 5 minutos

  get(key) {
    const data = this._cache.get(key);
    return data ? JSON.parse(data) : null;
  },

  set(key, value) {
    this._cache.put(key, JSON.stringify(value), this.TTL);
  },

  invalidate(key) {
    this._cache.remove(key);
  }
};
```

---

## 7. INTERFACE DO USUÁRIO

### Sidebars
```javascript
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('MySidebar')
    .setTitle('Título da Sidebar')
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}
```

### Dialogs
```javascript
function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MyDialog')
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Título do Dialog');
}
```

### Feedback ao Usuário
```javascript
// Toast - notificação não-bloqueante
SpreadsheetApp.getActiveSpreadsheet().toast(
  'Mensagem',           // corpo
  'Título',             // título (opcional)
  5                     // duração em segundos
);

// Alert - bloqueante
SpreadsheetApp.getUi().alert('Mensagem');

// Confirmação
const response = SpreadsheetApp.getUi().alert(
  'Título',
  'Deseja continuar?',
  SpreadsheetApp.getUi().ButtonSet.YES_NO
);

if (response === SpreadsheetApp.getUi().Button.YES) {
  // Usuário confirmou
}
```

---

## 8. HTML/CSS PARA SIDEBARS

### Estrutura Base HTML
```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <link rel="stylesheet"
    href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
  <style>
    .section { margin-bottom: 16px; }
    .label { font-weight: 500; margin-bottom: 4px; }
    .error { color: #d93025; }
    .success { color: #188038; }
  </style>
</head>
<body>
  <div class="section">
    <!-- Conteúdo -->
  </div>

  <script>
    // Comunicação com backend
    function callBackend() {
      google.script.run
        .withSuccessHandler(onSuccess)
        .withFailureHandler(onError)
        .backendFunction();
    }

    function onSuccess(result) {
      console.log('Sucesso:', result);
    }

    function onError(error) {
      console.error('Erro:', error);
    }
  </script>
</body>
</html>
```

### Padrões CSS
```css
/* Usar classes do Google add-ons CSS */
.sidebar { padding: 16px; }
.branding-below { border-bottom: 1px solid #e0e0e0; }

/* Botões */
.action { /* botão primário - azul */ }
.create { /* botão verde */ }

/* Inputs */
input[type="text"], select {
  width: 100%;
  margin-bottom: 8px;
}
```

---

## 9. SEGURANÇA

### Nunca Hardcode Credenciais
```javascript
// ❌ NUNCA
const API_KEY = 'sk-1234567890';

// ✅ Usar PropertiesService
const API_KEY = PropertiesService.getScriptProperties()
  .getProperty('API_KEY');
```

### Validar Inputs do Usuário
```javascript
function processUserInput(input) {
  // Sanitizar strings
  const sanitized = input.toString().trim();

  // Validar formato
  if (!/^[a-zA-Z0-9_-]+$/.test(sanitized)) {
    throw new Error('Input contém caracteres inválidos');
  }

  return sanitized;
}
```

### Limitar Escopo de Permissões
```javascript
// No appsscript.json, usar apenas escopos necessários
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/drive.file"
  ]
}
```

---

## 10. TESTES

### Funções de Teste
```javascript
/**
 * Testa o sistema de cache
 * Executar manualmente no editor
 */
function testCacheSystem() {
  console.log('=== TESTE: Cache System ===');

  try {
    // Setup
    const testKey = 'test_key_' + Date.now();
    const testData = { name: 'Test', value: 123 };

    // Test set
    CacheManager.set(testKey, testData);
    console.log('✓ Cache.set() OK');

    // Test get
    const retrieved = CacheManager.get(testKey);
    if (JSON.stringify(retrieved) !== JSON.stringify(testData)) {
      throw new Error('Dados recuperados não conferem');
    }
    console.log('✓ Cache.get() OK');

    // Cleanup
    CacheManager.invalidate(testKey);
    console.log('✓ Cache.invalidate() OK');

    console.log('=== TODOS OS TESTES PASSARAM ===');

  } catch (error) {
    console.error('✗ TESTE FALHOU:', error.message);
  }
}
```

---

## 11. CHECKLIST PRÉ-COMMIT

- [ ] Código segue convenções de nomenclatura
- [ ] Funções estão documentadas com JSDoc
- [ ] Tratamento de erros implementado
- [ ] Batch operations usadas (não loops individuais)
- [ ] Cache utilizado onde apropriado
- [ ] Sem credenciais hardcoded
- [ ] Testado com dados reais
- [ ] Console.log de debug removido

---

**Referência rápida:** Consulte `CLAUDE.md` para erros conhecidos e soluções.
