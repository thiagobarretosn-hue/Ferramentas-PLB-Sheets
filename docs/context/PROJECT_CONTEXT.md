# PROJECT CONTEXT - Ferramentas PLB Sheets
> **Versão:** 1.0 | **Atualizado:** 2026-01-20

Este documento fornece contexto completo do projeto para assistência de IA.

---

## 1. VISÃO GERAL

**Ferramentas PLB Sheets** é um sistema unificado de automação para Google Sheets que combina três subsistemas principais:

1. **Relatórios Dinâmicos (BOM)** - Geração de relatórios hierárquicos
2. **PLB Templates** - Biblioteca centralizada de templates
3. **Super Busca** - Busca avançada de materiais

**Propósito:** Simplificar operações complexas de planilhas para a equipe PLB.

---

## 2. ARQUITETURA

```
┌─────────────────────────────────────────────────────────────┐
│                      Google Sheets                          │
├─────────────────────────────────────────────────────────────┤
│  Menu.gs (Orquestração)                                     │
│    ├── onOpen() → Cria menus                                │
│    └── onEdit() → Processa edições                          │
├─────────────────────────────────────────────────────────────┤
│  BOM.gs              │  Template.gs       │  SuperBusca.gs  │
│  ├── Config Sheet    │  ├── Cache 5min    │  ├── Search     │
│  ├── Grouping        │  ├── Colors        │  └── Insert     │
│  ├── PDF Export      │  └── Sheet Mgmt    │                 │
│  └── Fasteners       │                    │                 │
├─────────────────────────────────────────────────────────────┤
│  HTML Sidebars (8 arquivos)                                 │
├─────────────────────────────────────────────────────────────┤
│  Services: SpreadsheetApp, DriveApp, CacheService,          │
│            PropertiesService, HtmlService, LockService      │
└─────────────────────────────────────────────────────────────┘
```

---

## 3. MÓDULOS DETALHADOS

### 3.1 Menu.gs (102 linhas)

**Responsabilidade:** Orquestração de menus e triggers

**Menus Criados:**
| Menu | Funções |
|------|---------|
| Relatórios Dinâmicos | BOM functions |
| PLB Templates | Template functions |
| Gerenciar Abas | Sheet management |
| Super Busca | Search functions |

**Triggers:**
- `onOpen(e)` - Inicializa menus
- `onEdit(e)` - Processa edições em Config e auto-coloração

---

### 3.2 BOM.gs (1.343 linhas)

**Responsabilidade:** Geração de relatórios dinâmicos com agrupamento hierárquico

**Componentes Principais:**

```javascript
// Configurações na aba Config
const CONFIG_KEYS = {
  SOURCE_SHEET,    // Aba de origem
  GROUP_L1,        // Agrupamento nível 1
  GROUP_L2,        // Agrupamento nível 2
  GROUP_L3,        // Agrupamento nível 3
  COL_1..COL_5,    // Colunas do relatório
  PROJECT, BOM,    // Identificação
  DRIVE_FOLDER_ID, // Pasta para PDFs
  // ...
};
```

**Fluxo de Processamento:**
1. Usuário configura na aba Config
2. `updateGroupingPanel()` mostra valores disponíveis
3. Usuário seleciona combinações
4. `processBomCore()` gera relatórios
5. `exportPDFsWithFeedback()` exporta PDFs

**Sistema de Fixadores:**
```javascript
// Biblioteca de componentes
const FASTENER_LIBRARY = {
  RISER: [clamps, nuts, washers, anchors, rods],
  LOOP: [hangs, nuts, washers, anchors, rods]
};

// Intervalos
RISER: a cada 10 unidades
LOOP: a cada 3 unidades
```

---

### 3.3 Template.gs (1.173 linhas)

**Responsabilidade:** Gerenciamento de biblioteca de templates

**Banco de Dados Central:**
```javascript
const CENTRAL_SPREADSHEET_ID = '1KD_NXdtEjORJJFYwiFc9uzrZeO57TPPSXWh7bokZMdU';

// Estrutura:
// DATA BASE → Templates
// Task → Cores por tarefa
// SubTrade → Cores por subtrade
```

**Sistema de Cache:**
```javascript
class TemplateCache {
  static EXPIRY = 5 * 60 * 1000; // 5 minutos
  _cache = null;
  _timestamp = 0;

  get() { /* ... */ }
  set(data) { /* ... */ }
  isValid() { /* ... */ }
}
```

**Hierarquia de Templates:**
```
TASK → SUB-TASK → SUB-TRADE → LOCAL → DESCRIPTION, QUANTITY
```

**Sistema de Cores:**
```javascript
const TASK_COLORS = {
  'UNITS': '#E8F5E9',
  'COMMON AREAS': '#E3F2FD',
  'SHELL': '#FFF3E0',
  // ...
};

const SUBTRADE_COLORS = {
  'SEWER': '#FFEBEE',
  'WS': '#E3F2FD',
  // ...
};
```

---

### 3.4 SuperBusca.gs (106 linhas)

**Responsabilidade:** Busca rápida e inserção de materiais

**Fluxo:**
1. `abrirSuperBuscaSidebar()` - Abre interface
2. `getSheetList()` - Lista abas disponíveis
3. `getColumnHeaders()` - Obtém cabeçalhos
4. `getDadosParaBusca()` - Busca dados únicos
5. `inserirItensSelecionados()` - Insere na célula ativa

---

## 4. PADRÕES DE CÓDIGO

### ConfigService (BOM.gs)
```javascript
const ConfigService = {
  getAll() {
    // Retorna todas as configurações como objeto
  },
  get(key, defaultValue) {
    // Retorna valor específico
  }
};
```

### CacheManager (BOM.gs)
```javascript
const CacheManager = {
  _cache: CacheService.getScriptCache(),
  TTL: 180, // 3 minutos

  get(key) { /* ... */ },
  put(key, value, ttl) { /* ... */ },
  invalidateAll() { /* ... */ }
};
```

### Utilitários Comuns
```javascript
function getColumnIndex(letter)     // A → 1, B → 2
function getColumnHeader(config)    // 'A|Nome' → 'Nome'
function numberToColumnLetter(n)    // 1 → A, 27 → AA
function sanitizeSheetName(name)    // Remove caracteres inválidos
```

---

## 5. INTERFACES HTML

| Arquivo | Propósito |
|---------|-----------|
| `BomSidebar.html` | Painel principal BOM |
| `ConfigSidebar.html` | Painel de controle BOM |
| `FixadoresSidebar.html` | Seletor de fixadores |
| `SuperBuscaSidebar.html` | Interface de busca |
| `template-sidebar.html` | Navegador de templates |
| `SheetManager.html` | Gerenciador de abas |
| `color-config-sidebar.html` | Configuração de cores |
| `duplicate-dialog.html` | Resolução de conflitos |

### Comunicação Frontend-Backend
```javascript
// No HTML
google.script.run
  .withSuccessHandler(callback)
  .withFailureHandler(errorHandler)
  .backendFunction(params);

// No .gs
function backendFunction(params) {
  return { success: true, data: result };
}
```

---

## 6. DEPENDÊNCIAS EXTERNAS

### Planilha Central de Templates
- **ID:** `1KD_NXdtEjORJJFYwiFc9uzrZeO57TPPSXWh7bokZMdU`
- **Abas:** DATA BASE, Task, SubTrade

### Google Drive
- Usado para exportação de PDFs
- Pasta configurável via `DRIVE_FOLDER_ID`

---

## 7. PONTOS DE EXTENSÃO

### Adicionar Novo Menu
```javascript
// Em Menu.gs, dentro de onOpen()
ui.createMenu('Novo Menu')
  .addItem('Nova Função', 'novaFuncao')
  .addToUi();
```

### Adicionar Nova Sidebar
1. Criar arquivo HTML em `html/`
2. Criar função de abertura em `.gs`
3. Adicionar ao menu

### Adicionar Novo Snippet
1. Criar em `lib/Snippets/[categoria]/`
2. Documentar em `CLAUDE.md` seção 5

---

## 8. LIMITAÇÕES CONHECIDAS

| Limitação | Impacto | Mitigação |
|-----------|---------|-----------|
| Cache max 100KB | Dados grandes | Chunking |
| Timeout 6min | Processamento longo | Batch + progress |
| Trigger simples sem UI | Sem alerts em onEdit | Usar toast |
| 30 requests/min para Drive | Export em massa | Queue + delay |

---

## 9. AMBIENTE DE DESENVOLVIMENTO

### Local (Recomendado)
```bash
# Instalar clasp
npm install -g @google/clasp

# Login
clasp login

# Clonar projeto existente
clasp clone <script-id>

# Push alterações
clasp push

# Pull do servidor
clasp pull
```

### Editor Web
- [script.google.com](https://script.google.com)
- Útil para debug rápido

---

## 10. RECURSOS PARA CONSULTA

### Ordem de Prioridade
1. `CLAUDE.md` - Referência rápida
2. `docs/context/DIRETRIZES.md` - Padrões de código
3. Código existente no projeto
4. [Documentação oficial](https://developers.google.com/apps-script/reference)

### Links Úteis
- [SpreadsheetApp](https://developers.google.com/apps-script/reference/spreadsheet)
- [DriveApp](https://developers.google.com/apps-script/reference/drive)
- [HtmlService](https://developers.google.com/apps-script/reference/html)
- [CacheService](https://developers.google.com/apps-script/reference/cache)
