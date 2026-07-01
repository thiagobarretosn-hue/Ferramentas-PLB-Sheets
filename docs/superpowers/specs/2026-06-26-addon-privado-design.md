# Design — Ferramentas PLB como Google Workspace Add-on Privado

**Data:** 2026-06-26  
**Status:** Aprovado pelo usuário  
**Projeto:** Ferramentas-PLB-Sheets  

---

## Problema

O projeto Ferramentas-PLB-Sheets é atualmente um script container-bound — precisa ser copiado manualmente para cada planilha. Com 15+ planilhas em uso e outros engenheiros criando novas sem intervenção do desenvolvedor, qualquer atualização de ferramenta exige atualizar N cópias manualmente. Isso é insustentável.

---

## Solução

Converter o projeto para um **Google Workspace Add-on privado**. O add-on é instalado uma vez por usuário (ou pelo admin do Workspace org-wide) e injeta o menu `Sheet Tools` em qualquer Sheets aberto. Atualizações são feitas com um único redeploy.

---

## Arquitetura

### Fluxo de distribuição

```
[Desenvolvedor — clasp push]
        │
        ▼
[script.google.com — projeto standalone]
        │  deploy como Add-on
        ▼
[Link de instalação enviado ao time]
        │
        ▼
[Usuário instala uma vez]
        │
        ▼
[Menu "Sheet Tools" aparece em qualquer Sheets automaticamente]
```

### Tipo de projeto

- **Antes:** container-bound (ligado a uma planilha específica)
- **Depois:** standalone script, desacoplado de qualquer planilha

O add-on acessa a planilha ativa via `SpreadsheetApp.getActiveSpreadsheet()` — mesmo comportamento de antes, mas disponível em qualquer Sheets.

---

## Mudanças necessárias

### 1. Setup GCP (uma vez, fora do código)

- Criar projeto em console.cloud.google.com
- Vincular ao script em Apps Script > Project Settings > Google Cloud Platform
- Habilitar a API `Google Workspace Add-ons`

### 2. `appsscript.json` — adicionar seção `addOns`

```json
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/drive"
  ],
  "addOns": {
    "common": {
      "name": "PLB Sheet Tools",
      "logoUrl": "https://www.gstatic.com/images/icons/material/system/1x/table_chart_black_48dp.png",
      "homepageTrigger": {
        "runFunction": "onHomepage",
        "enabled": true
      }
    },
    "sheets": {
      "homepageTrigger": {
        "runFunction": "onHomepage",
        "enabled": true
      },
      "onFileScopeGrantedTrigger": {
        "runFunction": "onFileScopeGranted"
      }
    }
  }
}
```

### 3. `Menu.gs` — adaptar `onOpen`

O add-on usa `onOpen(e)` com guard de `authMode` para criar o menu corretamente.

```javascript
// Antes
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Sheet Tools')
    ...
    .addToUi();
}

// Depois
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🛠️ Sheet Tools');

  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    // modo limitado — só adiciona itens que não precisam de auth
    menu.addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar');
  } else {
    menu
      .addItem('📊 Gerador de BOM', 'openBomSidebar')
      .addItem('📋 Gerador de Request', 'openRequestSidebar')
      .addSeparator()
      .addItem('📑 Gerenciador de Abas', 'showSheetManager')
      .addItem('🎨 Cores das Abas', 'openColorConfig')
      .addSeparator()
      .addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar')
      .addSeparator()
      .addItem('📊 Summary All', 'openSummaryAllSidebar');
  }
  menu.addToUi();
}

// Função obrigatória para homepage card (pode retornar card vazio)
function onHomepage(e) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('PLB Sheet Tools'))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText('Acesse as ferramentas pelo menu 🛠️ Sheet Tools na planilha.'))
    )
    .build();
}

function onFileScopeGranted(e) {
  onOpen(e);
}
```

### 4. Sem mudanças

Todo o restante do código `.gs` e todas as sidebars HTML permanecem idênticos:
- `BOM.gs`, `Request.gs`, `Template.gs`, `SuperBusca.gs`
- `ColorConfig.gs`, `SheetManager.gs`, `SummaryAll.gs`, `ConditionalFormat.gs`
- Todos os arquivos `.html`
- `ExportProReceiver.gs` (continua como está — é um `doPost` independente)

---

## Estrutura de pastas

O projeto add-on é criado em uma **pasta nova separada**, preservando o projeto atual intacto durante a implementação:

```text
C:\DEV\Sheets\
├── Ferramentas-PLB-Sheets\        ← projeto atual (container-bound, preservado)
└── Ferramentas-PLB-Sheets-AddOn\  ← novo projeto add-on (esta implementação)
```

A pasta atual não é tocada até o add-on estar validado e aprovado para substituição.

---

## Conta Google

O add-on será implantado na **conta pessoal do desenvolvedor** (não na conta Ambar US).

- Sem acesso de admin ao Google Workspace da empresa por enquanto
- Distribuição: link de instalação compartilhado manualmente com cada membro do time
- Cada usuário instala em sua conta Google pessoal ou de trabalho via o link

---

## Deploy

1. Apps Script > Deploy > New deployment > **Add-on**
2. Tipo: Add-on (não Web App)
3. Gerar link de instalação (formato interno, não marketplace público)
4. Compartilhar link com o time (instalação em 30 segundos por usuário)

---

## Distribuição futura de atualizações

1. `clasp push` para subir as alterações
2. Apps Script > Deploy > Manage deployments > criar nova versão
3. Usuários recebem automaticamente na próxima sessão

---

## O que NÃO é coberto nesta spec

- Migração do `ExportProReceiver.gs` (funciona independente, não precisa mudar)
- Triggers instaláveis para `onEdit` (se necessário, é adicionado separadamente por escopo futuro)
- Publicação no Google Workspace Marketplace público (fora do escopo — uso interno Ambar US)

---

## Critérios de sucesso

- [ ] Add-on instalado em 1 planilha de teste aparece automaticamente em outra planilha nova
- [ ] Todas as ferramentas (BOM, Request, Template, SuperBusca, SheetManager, ColorConfig, SummaryAll) funcionam identicamente ao comportamento atual
- [ ] Um redeploy propaga a atualização sem ação dos usuários
