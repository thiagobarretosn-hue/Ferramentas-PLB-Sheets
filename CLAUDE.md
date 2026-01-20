# CLAUDE.md - Google Apps Script Development Hub
> **Versão:** 1.0 | **Autor:** Claude Code Agent | **Atualizado:** 2026-01-20

Este arquivo é o centro de conhecimento para desenvolvimento de Google Apps Script com assistência de IA.

---

## SEÇÃO 0: MODO DE OPERAÇÃO - PODERES

### REGRA ZERO - Respeitar a API do Google Apps Script
Nunca invente soluções que a API não suporta. Sempre verifique a documentação oficial.

### CHECKLIST PRÉ-DESENVOLVIMENTO (OBRIGATÓRIO)

**FASE 1 - ANÁLISE**
- [ ] Ler código existente relacionado
- [ ] Identificar padrões já utilizados no projeto
- [ ] Verificar snippets disponíveis em `lib/Snippets/`
- [ ] Consultar documentação oficial se necessário

**FASE 2 - PLANEJAMENTO**
- [ ] Definir estrutura da solução
- [ ] Identificar serviços GAS necessários (SpreadsheetApp, DriveApp, etc.)
- [ ] Planejar tratamento de erros
- [ ] Considerar performance (cache, batch operations)

**FASE 3 - IMPLEMENTAÇÃO**
- [ ] Usar template base (SEÇÃO 3)
- [ ] Seguir convenções de nomenclatura
- [ ] Implementar logging adequado
- [ ] Adicionar validações de entrada

**FASE 4 - VALIDAÇÃO**
- [ ] Testar com dados reais
- [ ] Verificar tratamento de erros
- [ ] Validar performance com volumes maiores
- [ ] Documentar função no código

### LIBERDADE CRIATIVA COM ANÁLISE CRÍTICA
- Tenho autonomia para propor soluções melhores
- Devo questionar requisitos ambíguos
- Posso discordar com justificativa técnica
- Devo sugerir alternativas quando apropriado

### ORDEM DE CONSULTA
1. **Snippets locais** (`lib/Snippets/`)
2. **Código existente** no projeto
3. **Documentação oficial** Google Apps Script
4. **Web search** como último recurso

---

## SEÇÃO 1: AMBIENTE DE DESENVOLVIMENTO

### Runtime
- **V8 Runtime** (obrigatório desde Jan/2026)
- Rhino runtime descontinuado
- Suporte a ES6+ (arrow functions, const/let, classes, etc.)

### Ferramentas Recomendadas
- **clasp** - CLI para desenvolvimento local
- **gas-fakes** - Emulação local do ambiente GAS
- **TypeScript** - Opcional, com bundler (Rollup)

### Estrutura do Projeto
```
Ferramentas-PLB-Sheets/
├── CLAUDE.md              # Este arquivo
├── README.md              # Documentação pública
├── .clasp.json            # Configuração clasp
├── appsscript.json        # Manifesto do projeto
├── docs/
│   ├── context/           # Contexto para IA
│   │   ├── DIRETRIZES.md
│   │   ├── PROJECT_CONTEXT.md
│   │   └── API_REFERENCE.md
│   └── guides/            # Guias e tutoriais
├── lib/
│   └── Snippets/          # Funções reutilizáveis
│       ├── cache/
│       ├── sheets/
│       ├── drive/
│       ├── ui/
│       └── utils/
├── src/
│   ├── Menu.gs
│   ├── BOM.gs
│   ├── Template.gs
│   └── SuperBusca.gs
├── html/                  # Arquivos HTML para sidebars
└── .claude/               # Configuração do agente
```

---

## SEÇÃO 2: ERROS CRÍTICOS E SOLUÇÕES

### 2.1 Cache Service - Limite de 100KB
```javascript
// ❌ ERRO: Dados maiores que 100KB
CacheService.getScriptCache().put('key', hugeData);

// ✅ SOLUÇÃO: Dividir em chunks
function putLargeCache(key, data, ttl = 21600) {
  const cache = CacheService.getScriptCache();
  const chunks = chunkString(JSON.stringify(data), 90000);
  chunks.forEach((chunk, i) => cache.put(`${key}_${i}`, chunk, ttl));
  cache.put(`${key}_count`, chunks.length.toString(), ttl);
}
```

### 2.2 Timeout de 6 minutos (scripts normais)
```javascript
// ❌ ERRO: Loop muito longo
for (let i = 0; i < 100000; i++) {
  sheet.getRange(i, 1).setValue(data[i]); // Muito lento!
}

// ✅ SOLUÇÃO: Batch operations
const values = data.map(d => [d]);
sheet.getRange(1, 1, values.length, 1).setValues(values);
```

### 2.3 Quota de Leitura/Escrita
```javascript
// ❌ ERRO: Muitas chamadas à API
for (const row of rows) {
  const value = sheet.getRange(row, 1).getValue();
}

// ✅ SOLUÇÃO: Ler tudo de uma vez
const allValues = sheet.getDataRange().getValues();
```

### 2.4 HTML Service - Sandboxing
```javascript
// ❌ ERRO: Scripts inline bloqueados
'<button onclick="doSomething()">Click</button>'

// ✅ SOLUÇÃO: Event listeners
'<button id="btn">Click</button>'
'<script>document.getElementById("btn").addEventListener("click", doSomething);</script>'
```

### 2.5 SpreadsheetApp.flush() - Uso Correto
```javascript
// Usar flush() apenas quando necessário forçar escrita
// antes de operações que dependem dos dados escritos
sheet.getRange('A1').setValue('test');
SpreadsheetApp.flush(); // Força escrita imediata
const value = sheet.getRange('A1').getValue(); // Agora lê o valor correto
```

### 2.6 Triggers - Limitações
```javascript
// ❌ ERRO: Trigger simples não tem acesso a UI
function onEdit(e) {
  SpreadsheetApp.getUi().alert('Hello'); // ERRO!
}

// ✅ SOLUÇÃO: Usar installable trigger ou toast
function onEdit(e) {
  SpreadsheetApp.getActiveSpreadsheet().toast('Editado!');
}
```

### 2.7 Logger vs Console
```javascript
// Logger.log() - Perde mensagens em crash
// console.log() - V8 runtime, melhor para debug
// Stackdriver Logging - Para produção

// ✅ PADRÃO RECOMENDADO
function log(message, level = 'INFO') {
  console.log(`[${level}] ${new Date().toISOString()} - ${message}`);
}
```

---

## SEÇÃO 3: TEMPLATE BASE

```javascript
/**
 * @fileoverview [Descrição do arquivo]
 * @version 1.0.0
 * @author [Nome]
 *
 * Changelog:
 * - v1.0.0: Versão inicial
 */

// ============================================================
// CONFIGURAÇÕES
// ============================================================
const CONFIG = {
  CACHE_TTL: 300,        // 5 minutos
  MAX_ROWS: 10000,
  DEBUG: false
};

// ============================================================
// FUNÇÕES PRINCIPAIS
// ============================================================

/**
 * Função principal
 * @param {Object} params - Parâmetros de entrada
 * @returns {Object} Resultado da operação
 */
function mainFunction(params) {
  try {
    // Validação de entrada
    if (!params) {
      throw new Error('Parâmetros obrigatórios não fornecidos');
    }

    // Lógica principal
    const result = processData(params);

    // Retorno de sucesso
    return {
      success: true,
      data: result
    };

  } catch (error) {
    console.error(`[ERRO] ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================
// FUNÇÕES AUXILIARES
// ============================================================

/**
 * Processa dados
 * @private
 */
function processData(params) {
  // Implementação
}

// ============================================================
// UTILITÁRIOS
// ============================================================

/**
 * Log com timestamp
 */
function log(message, level = 'INFO') {
  if (CONFIG.DEBUG || level === 'ERROR') {
    console.log(`[${level}] ${new Date().toISOString()} - ${message}`);
  }
}
```

---

## SEÇÃO 4: SERVIÇOS GOOGLE APPS SCRIPT

### Serviços Principais
| Serviço | Uso | Documentação |
|---------|-----|--------------|
| `SpreadsheetApp` | Manipulação de planilhas | [Link](https://developers.google.com/apps-script/reference/spreadsheet) |
| `DriveApp` | Arquivos e pastas | [Link](https://developers.google.com/apps-script/reference/drive) |
| `DocumentApp` | Google Docs | [Link](https://developers.google.com/apps-script/reference/document) |
| `FormApp` | Google Forms | [Link](https://developers.google.com/apps-script/reference/forms) |
| `GmailApp` | Email | [Link](https://developers.google.com/apps-script/reference/gmail) |
| `CalendarApp` | Calendário | [Link](https://developers.google.com/apps-script/reference/calendar) |

### Serviços Utilitários
| Serviço | Uso |
|---------|-----|
| `CacheService` | Cache temporário (max 6h, 100KB) |
| `PropertiesService` | Armazenamento persistente |
| `LockService` | Controle de concorrência |
| `UrlFetchApp` | Requisições HTTP |
| `HtmlService` | Interfaces HTML |
| `Utilities` | Funções utilitárias (base64, hash, etc.) |

---

## SEÇÃO 5: BIBLIOTECA DE SNIPPETS

### Categorias Disponíveis

#### `lib/Snippets/cache/`
- `cache_manager.gs` - Gerenciador de cache com chunks
- `cache_invalidation.gs` - Invalidação seletiva

#### `lib/Snippets/sheets/`
- `sheet_utils.gs` - Utilitários de planilha
- `data_validation.gs` - Validação de dados
- `formatting.gs` - Formatação de células

#### `lib/Snippets/drive/`
- `folder_manager.gs` - Gerenciamento de pastas
- `pdf_export.gs` - Exportação para PDF

#### `lib/Snippets/ui/`
- `sidebar_base.gs` - Base para sidebars
- `dialog_utils.gs` - Utilitários para diálogos
- `toast_notifications.gs` - Notificações toast

#### `lib/Snippets/utils/`
- `string_utils.gs` - Manipulação de strings
- `date_utils.gs` - Manipulação de datas
- `array_utils.gs` - Operações em arrays

---

## SEÇÃO 6: PADRÕES DE CÓDIGO

### Nomenclatura
```javascript
// Funções: camelCase (verbos)
function getUserData() {}
function processItems() {}
function validateInput() {}

// Variáveis: camelCase (substantivos)
const userData = {};
let itemCount = 0;

// Constantes: UPPER_SNAKE_CASE
const MAX_ROWS = 1000;
const API_KEY = 'xxx';

// Classes: PascalCase
class DataProcessor {}
class CacheManager {}

// Funções privadas: prefixo _
function _helperFunction() {}
```

### Estrutura de Funções
```javascript
/**
 * Descrição breve da função
 *
 * @param {string} param1 - Descrição do parâmetro
 * @param {number} [param2=10] - Parâmetro opcional com default
 * @returns {Object} Descrição do retorno
 * @throws {Error} Quando ocorre erro X
 *
 * @example
 * const result = minhaFuncao('teste', 20);
 */
function minhaFuncao(param1, param2 = 10) {
  // Implementação
}
```

### Tratamento de Erros
```javascript
// Padrão try-catch-finally
function operacaoCritica() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    // Operação crítica

  } catch (error) {
    console.error(`[ERRO] ${error.message}\n${error.stack}`);
    throw error; // Re-throw se necessário

  } finally {
    lock.releaseLock();
  }
}
```

---

## SEÇÃO 7: FERRAMENTAS EXISTENTES

### Menu.gs
| Função | Descrição |
|--------|-----------|
| `onOpen()` | Inicializa menus na abertura |
| `onEdit(e)` | Processa edições (trigger) |

### BOM.gs
| Função | Descrição |
|--------|-----------|
| `processBomCore()` | Engine de processamento de relatórios |
| `forceCreateConfig()` | Cria aba Config |
| `updateGroupingPanel()` | Atualiza painel de agrupamento |
| `exportPDFsWithFeedback()` | Exporta PDFs para Drive |
| `processarFixadoresSelecionados()` | Adiciona fixadores |

### Template.gs
| Função | Descrição |
|--------|-----------|
| `loadTemplatesWithCache()` | Carrega templates com cache |
| `insertLocalTemplates()` | Insere template |
| `createTemplateFromSelection()` | Cria template da seleção |
| `applyGroupColors()` | Aplica cores por grupo |
| `showSheetManager()` | Gerenciador de abas |

### SuperBusca.gs
| Função | Descrição |
|--------|-----------|
| `abrirSuperBuscaSidebar()` | Abre sidebar de busca |
| `getDadosParaBusca()` | Obtém dados para busca |
| `inserirItensSelecionados()` | Insere itens selecionados |

---

## SEÇÃO 8: PERFORMANCE

### Batch Operations (SEMPRE usar)
```javascript
// ❌ Lento
for (let i = 0; i < 1000; i++) {
  sheet.getRange(i+1, 1).setValue(data[i]);
}

// ✅ Rápido
const values = data.map(d => [d]);
sheet.getRange(1, 1, data.length, 1).setValues(values);
```

### Cache Strategy
```javascript
// Cache de dados frequentes
const CACHE_TTL = 300; // 5 minutos

function getDataWithCache(key, fetchFn) {
  const cache = CacheService.getScriptCache();
  let data = cache.get(key);

  if (data) {
    return JSON.parse(data);
  }

  data = fetchFn();
  cache.put(key, JSON.stringify(data), CACHE_TTL);
  return data;
}
```

### Minimize API Calls
```javascript
// ❌ Múltiplas chamadas
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName('Data');
const range = sheet.getDataRange();
const values = range.getValues();

// ✅ Encadeamento
const values = SpreadsheetApp.getActiveSpreadsheet()
  .getSheetByName('Data')
  .getDataRange()
  .getValues();
```

---

## SEÇÃO 9: RECURSOS EXTERNOS

### Documentação Oficial
- [Apps Script Reference](https://developers.google.com/apps-script/reference)
- [Apps Script Guides](https://developers.google.com/apps-script/guides)
- [Release Notes](https://developers.google.com/apps-script/release-notes)

### Ferramentas
- [clasp - CLI](https://github.com/google/clasp)
- [gas-fakes - Local testing](https://github.com/nicetomeetyou1208/gas-fakes)
- [Apps Script Samples](https://github.com/googleworkspace/apps-script-samples)

### Comunidade
- [Stack Overflow - google-apps-script](https://stackoverflow.com/questions/tagged/google-apps-script)
- [Google Apps Script Community](https://groups.google.com/g/google-apps-script-community)

---

## SEÇÃO 10: QUICK REFERENCE

### SpreadsheetApp
```javascript
// Obter planilha ativa
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getActiveSheet();

// Obter/definir valores
const value = sheet.getRange('A1').getValue();
const values = sheet.getRange('A1:B10').getValues();
sheet.getRange('A1').setValue('Hello');
sheet.getRange('A1:B2').setValues([['A','B'],['C','D']]);

// Última linha/coluna com dados
const lastRow = sheet.getLastRow();
const lastCol = sheet.getLastColumn();

// Criar nova aba
const newSheet = ss.insertSheet('NovaAba');

// Deletar aba
ss.deleteSheet(sheet);
```

### DriveApp
```javascript
// Obter pasta por ID
const folder = DriveApp.getFolderById('folder_id');

// Criar arquivo
const file = folder.createFile('nome.txt', 'conteúdo');

// Criar PDF de planilha
const blob = sheet.getParent().getAs('application/pdf');
folder.createFile(blob.setName('relatorio.pdf'));
```

### HtmlService
```javascript
// Criar sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Minha Sidebar')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Criar dialog
function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Título');
}
```

### CacheService
```javascript
const cache = CacheService.getScriptCache();

// Salvar (max 6h = 21600s)
cache.put('key', 'value', 21600);

// Obter
const value = cache.get('key');

// Remover
cache.remove('key');

// Múltiplos
cache.putAll({'key1': 'val1', 'key2': 'val2'}, 21600);
const values = cache.getAll(['key1', 'key2']);
```

---

**Última atualização:** 2026-01-20 | **Versão:** 1.0
