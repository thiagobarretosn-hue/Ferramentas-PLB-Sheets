# 🛠️ Ferramentas PLB Sheets

Sistema unificado de ferramentas para Google Sheets, combinando funcionalidades de gerenciamento de BOM (Bill of Materials), Templates de tarefas e ferramentas avançadas de busca.

## 📋 Descrição

Este projeto combina três sistemas principais:

1. **Relatórios Dinâmicos (BOM)** - Sistema de geração automática de relatórios BOM com fixadores
2. **PLB Templates** - Gerenciamento de templates de tarefas para projetos
3. **Super Busca** - Ferramenta avançada de busca em planilhas

## 📂 Estrutura do Projeto

```
Ferramentas-PLB-Sheets/
├── Menu.gs                        # Menu principal e triggers onOpen/onEdit
├── BOM.gs                         # Sistema de Relatórios Dinâmicos (BOM)
├── Template.gs                    # Sistema de Templates PLB
│
├── BomSidebar.html               # Painel gerador de BOM
├── ConfigSidebar.html            # Painel de controle BOM
├── FixadoresSidebar.html         # Seletor de fixadores
├── SuperBuscaSidebar.html        # Painel Super Busca
│
├── template-sidebar.html         # Sidebar de templates
├── SheetManager.html             # Gerenciador de abas
├── color-config-sidebar.html    # Configuração de cores
├── duplicate-dialog.html        # Dialog de conflito de templates
│
└── README.md                     # Este arquivo
```

## 🎯 Funcionalidades

### 1. Sistema BOM (Relatórios Dinâmicos)

#### Características:
- ✅ Geração automática de relatórios agrupados
- ✅ Exportação de PDFs
- ✅ Sistema de fixadores inteligente
- ✅ Painéis de agrupamento interativos
- ✅ Cache otimizado para performance

#### Arquivos:
- **BOM.gs** - Código principal do sistema BOM
- **BomSidebar.html** - Interface do gerador de BOM
- **ConfigSidebar.html** - Painel de configuração
- **FixadoresSidebar.html** - Seletor de fixadores

#### Funções principais:
```javascript
// Criação e gerenciamento
forceCreateConfig()           // Cria aba Config
ensureConfigExists()          // Garante existência da aba Config
updateGroupingPanel()         // Atualiza painéis de agrupamento

// Processamento
runProcessing()               // Processa relatórios (via Config)
runProcessingFromHtml()       // Processa relatórios (via HTML)
processBomCore()              // Núcleo do processamento

// Exportação
exportPDFsWithFeedback()      // Exporta PDFs com feedback
runPdfExportFromHtml()        // Exporta PDFs via HTML

// Fixadores
abrirSeletorFixadores()       // Abre seletor de fixadores
getPipesElegiveis()           // Lista pipes elegíveis
processarFixadoresSelecionados() // Adiciona fixadores
removerFixadoresSelecionados()   // Remove fixadores

// UI
openBomSidebar()              // Abre painel BOM
openConfigSidebar()           // Abre painel de controle
```

### 2. Sistema de Templates PLB

#### Características:
- ✅ Biblioteca centralizada de templates
- ✅ Inserção rápida de tarefas
- ✅ Configuração de cores automáticas
- ✅ Gerenciamento avançado de abas
- ✅ Cache de 5 minutos

#### Arquivos:
- **Template.gs** - Código principal do sistema Templates
- **template-sidebar.html** - Sidebar de navegação de templates
- **SheetManager.html** - Gerenciador de abas
- **color-config-sidebar.html** - Configuração de cores
- **duplicate-dialog.html** - Resolução de conflitos

#### Funções principais:
```javascript
// Templates
loadTemplatesWithCache()      // Carrega templates com cache
insertLocalTemplates()        // Insere template único
insertMultipleLocals()        // Insere múltiplos templates
createTemplateFromSelection() // Cria template da seleção
refreshTemplates()            // Atualiza cache

// Cores
applyGroupColors()            // Aplica cores por grupo
saveColorConfiguration()      // Salva configuração de cores
getAllColorConfigs()          // Obtém todas as configs

// Abas
showSheetManager()            // Abre gerenciador de abas
renameSelected()              // Renomeia abas selecionadas
duplicateSelected()           // Duplica abas selecionadas
findAndReplaceSelected()      // Localizar/substituir em nomes

// UI
openTemplateSidebar()         // Abre sidebar de templates
openColorConfig()             // Abre config de cores
openSystemConfig()            // Abre config do sistema
testSystemTemplate()          // Testa o sistema
```

### 3. Menu Principal

#### Arquivo:
- **Menu.gs** - Centraliza todos os menus e triggers

#### Estrutura:
```javascript
function onOpen() {
  // Cria 4 menus principais:
  // 1. 🔧 Relatórios Dinâmicos
  // 2. 🏗️ PLB Templates
  // 3. 📑 Gerenciar Abas
  // 4. 🔍 Super Busca
}

function onEdit(e) {
  // Gerencia triggers de edição
  // - BOM Config
  // - Cores automáticas
}
```

## 🚀 Como Usar

### Instalação

1. Abra seu Google Sheets
2. Vá em **Extensões** > **Apps Script**
3. Cole os arquivos na seguinte ordem:
   - Menu.gs
   - BOM.gs
   - Template.gs
4. Adicione os arquivos HTML como arquivos separados
5. Salve e recarregue a planilha

### Configuração Inicial

#### BOM:
1. Menu **🔧 Relatórios Dinâmicos** > **Recriar Config**
2. Preencha as configurações na aba "Config"
3. Use **Painel de Controle** para gerenciar

#### Templates:
1. Menu **🏗️ PLB Templates** > **Configurar Sistema**
2. Defina a linha padrão de inserção
3. Configure o ID da planilha central

## 🔧 Configuração

### Constantes Importantes (BOM.gs)

```javascript
const CONFIG = {
  SHEETS: {
    CONFIG: 'Config'  // Nome da aba de configuração
  },
  CACHE_TTL: 180,     // Tempo de cache (segundos)
  DELIMITER: '|||'    // Delimitador de combinações
};
```

### Constantes Importantes (Template.gs)

```javascript
const CENTRAL_SPREADSHEET_ID = "ID_DA_PLANILHA_CENTRAL";
const CENTRAL_SHEET_NAME = "DATA BASE";
```

## 📊 Estrutura de Dados

### Aba Config (BOM)

| Seção | Configurações |
|-------|--------------|
| Agrupamento | Aba Origem, Níveis 1-3 |
| Dados BOMS | Colunas 1-5 |
| Cabeçalho | Project, BOM, Prefixo KOJO, Engenheiro, Versão |
| Opções | Classificação |
| Salvamento | Pasta Drive, Prefixo PDF |
| Fixadores | Colunas de mapeamento |

### Base de Templates

| Coluna | Descrição |
|--------|-----------|
| O (15) | TASK |
| P (16) | SUB-TASK |
| Q (17) | SUB-TRADE |
| R (18) | LOCAL |
| S (19) | DESC |
| T (20) | QTY |

## 🎨 Menus

### Menu: 🔧 Relatórios Dinâmicos
- ⚙️ Painel de Controle (Sidebar)
- 📊 Gerador de BOM (Painel)
- 🔧 Fixadores → Fonte
- 📄 Exportar PDFs (da Aba Config)
- 🗑️ Limpar Relatórios
- 🔄 Limpar Cache
- 🧪 Diagnóstico
- 🔧 Recriar Config

### Menu: 🏗️ PLB Templates
- 📋 Abrir Sidebar
- 🔄 Atualizar Templates
- ➕ Criar Template da Seleção
- ⚙️ Configurar Sistema
- 📂 Abrir Base de Dados
- 🧪 Testar Sistema
- Substituir SHELL em FIRESTOP

### Menu: 📑 Gerenciar Abas
- Gerenciador de Abas
- 🎨 Configurar Cores
- ✨ Aplicar Cores

### Menu: 🔍 Super Busca
- 🚀 Abrir Painel

## 🐛 Debug e Testes

### BOM
```javascript
testSystem()           // Diagnóstico completo
forceRefreshCache()    // Limpa cache
```

### Templates
```javascript
testSystemTemplate()   // Testa sistema de templates
refreshTemplates()     // Atualiza cache de templates
```

## 📝 Notas de Versão

### V2.13 (BOM)
- ✅ Correção na exportação de PDFs
- ✅ Função `getReportSheetNames()` readicionada
- ✅ Carregamento automático na aba "Exportar"

### Versão Inicial (Templates)
- ✅ Sistema de cache de templates
- ✅ Configuração de cores automáticas
- ✅ Gerenciador de abas completo

## 🤝 Contribuindo

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/NovaFuncionalidade`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/NovaFuncionalidade`)
5. Abra um Pull Request

## 📄 Licença

Este projeto é de uso interno da PLB.

## 👥 Autores

- **Thiago Barreto** - [thiagobarretosn-hue](https://github.com/thiagobarretosn-hue)

## 🔗 Links Úteis

- [Repositório Original BOM](https://github.com/thiagobarretosn-hue/bom)
- [Repositório Original TEMPLATEPLB](https://github.com/thiagobarretosn-hue/TEMPLATEPLB)
- [Documentação Google Apps Script](https://developers.google.com/apps-script)

---

**Desenvolvido com ❤️ para PLB**
