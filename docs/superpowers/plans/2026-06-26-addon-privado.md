# PLB Sheet Tools — Add-on Privado: Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Converter as Ferramentas PLB de script container-bound para um Google Workspace Add-on privado instalável em qualquer Sheets.

**Architecture:** O projeto atual é copiado para uma nova pasta `Ferramentas-PLB-Sheets-AddOn`. Apenas `appsscript.json` e `Menu.gs` são modificados. Todo o resto — BOM, Request, SuperBusca, sidebars HTML — é copiado sem alteração. O add-on é implantado na conta pessoal do desenvolvedor e distribuído via link de instalação.

**Tech Stack:** Google Apps Script V8, CardService (para homepage card), HtmlService (sidebars existentes), clasp (deploy opcional).

## Global Constraints

- Pasta original `C:\DEV\Sheets\Ferramentas-PLB-Sheets\` **não deve ser tocada** — é preservada como backup.
- Nova pasta: `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\`
- Conta Google: pessoal (não Ambar US Workspace)
- `timeZone` permanece `"America/Fortaleza"` (copiado do original)
- Nenhuma ferramenta existente (BOM, Request, Template, SuperBusca, etc.) pode mudar de comportamento

---

## Mapa de Arquivos

| Arquivo | Ação | Mudança |
|---|---|---|
| Todos os `.gs` exceto `Menu.gs` | Copiar | Nenhuma |
| Todos os `.html` | Copiar | Nenhuma |
| `lib/` (inteiro) | Copiar | Nenhuma |
| `templates/` (inteiro) | Copiar | Nenhuma |
| `appsscript.json` | Copiar + Modificar | Adicionar `oauthScopes` e `addOns` |
| `Menu.gs` | Copiar + Modificar | Guard `authMode`, `onHomepage`, `onFileScopeGranted` |

---

## Task 1 — Criar pasta do add-on e copiar projeto

**Files:**
- Create: `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\` (todos os arquivos do projeto original)

**Interfaces:**
- Produces: pasta nova pronta para edição, projeto original intacto

- [ ] **Step 1: Copiar o projeto inteiro para a nova pasta**

PowerShell:
```powershell
Copy-Item "C:\DEV\Sheets\Ferramentas-PLB-Sheets" `
          "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" `
          -Recurse
```

- [ ] **Step 2: Verificar que a cópia está completa**

```powershell
Get-ChildItem "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" -Recurse -File |
  Select-Object -ExpandProperty Name | Sort-Object
```

Esperado: todos os `.gs` e `.html` listados (BOM.gs, Menu.gs, SuperBusca.gs, etc.).

- [ ] **Step 3: Verificar que o original não foi tocado**

```powershell
Get-ChildItem "C:\DEV\Sheets\Ferramentas-PLB-Sheets" -Recurse -File | Measure-Object
Get-ChildItem "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" -Recurse -File | Measure-Object
```

Esperado: contagens iguais.

- [ ] **Step 4: Commit inicial da nova pasta**

```powershell
cd "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn"
git init
git add .
git commit -m "chore: copia inicial do projeto para versao add-on"
```

> Nota: se o projeto já usa git no diretório pai, apenas `git add` e `git commit` na raiz do repo existente.

---

## Task 2 — Atualizar `appsscript.json`

**Files:**
- Modify: `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\appsscript.json`

**Interfaces:**
- Consumes: arquivo copiado da Task 1 (conteúdo atual: `timeZone`, `dependencies`, `exceptionLogging`, `runtimeVersion`)
- Produces: manifesto com `oauthScopes` e `addOns` declarados — habilita o add-on no Apps Script

- [ ] **Step 1: Substituir o conteúdo de `appsscript.json`**

Conteúdo completo do arquivo (substitui o original inteiro):

```json
{
  "timeZone": "America/Fortaleza",
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

- [ ] **Step 2: Verificar JSON válido**

```powershell
Get-Content "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\appsscript.json" |
  ConvertFrom-Json | ConvertTo-Json -Depth 10
```

Esperado: JSON impresso sem erro de parse.

- [ ] **Step 3: Commit**

```powershell
git add appsscript.json
git commit -m "feat: adiciona configuracao de add-on e oauth scopes ao manifesto"
```

---

## Task 3 — Atualizar `Menu.gs`

**Files:**
- Modify: `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\Menu.gs`

**Interfaces:**
- Consumes: funções existentes `openBomSidebar`, `openRequestSidebar`, `showSheetManager`, `openColorConfig`, `abrirSuperBuscaSidebar`, `openSummaryAllSidebar` (definidas nos outros `.gs` — não mudam)
- Produces: `onOpen(e)` com guard de authMode, `onHomepage(e)` (retorna Card obrigatório), `onFileScopeGranted(e)` (reconstrói menu após auth)

- [ ] **Step 1: Substituir o conteúdo de `Menu.gs`**

```javascript
/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets (Add-on)
 * @version 4.0.0 - Convertido para Google Workspace Add-on privado
 */

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🛠️ Sheet Tools');

  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    // Auth limitada: adiciona apenas item para disparar autorização completa
    menu.addItem('🔓 Ativar ferramentas PLB', 'showAuthPrompt');
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

function onFileScopeGranted(e) {
  onOpen(e);
}

function onHomepage(e) {
  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader().setTitle('PLB Sheet Tools')
    )
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph()
            .setText('Ferramentas PLB ativas. Acesse pelo menu 🛠️ Sheet Tools na barra superior.')
        )
    )
    .build();
}

function showAuthPrompt() {
  SpreadsheetApp.getUi().alert(
    'PLB Sheet Tools',
    'Clique em OK para autorizar as ferramentas. O menu completo aparecerá em seguida.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
}
```

- [ ] **Step 2: Commit**

```powershell
git add Menu.gs
git commit -m "feat: adapta Menu.gs para add-on (authMode guard, onHomepage, onFileScopeGranted)"
```

---

## Task 4 — Setup GCP + Deploy + Instalar + Testar

> Esta task é executada manualmente no browser. Não há código a escrever — apenas passos de configuração no Google Cloud e no Apps Script.

**Files:**
- Nenhum arquivo novo

**Interfaces:**
- Consumes: projeto `Ferramentas-PLB-Sheets-AddOn` com as 3 tasks anteriores completas
- Produces: link de instalação do add-on funcional

### 4A — Subir código para o Apps Script

- [ ] **Step 1: Criar novo projeto standalone no Apps Script**

1. Abrir [script.google.com](https://script.google.com) na conta pessoal
2. Clicar em **New project**
3. Renomear para `PLB Sheet Tools - AddOn`
4. Apagar o código padrão `function myFunction() {}`

- [ ] **Step 2: Copiar os arquivos para o Apps Script**

Para cada arquivo `.gs` da pasta `Ferramentas-PLB-Sheets-AddOn`:
1. No Apps Script, clicar em **+** > **Script** para criar arquivo com o mesmo nome
2. Colar o conteúdo do arquivo local

Arquivos a copiar (ordem não importa):
- `Menu.gs` ← versão modificada (Task 3)
- `BOM.gs`
- `Request.gs`
- `Template.gs`
- `SuperBusca.gs`
- `ColorConfig.gs`
- `SheetManager.gs`
- `SummaryAll.gs`
- `ConditionalFormat.gs`
- `ExportProReceiver.gs`
- `lib/Shared/Config.gs`
- `lib/Shared/Logger.gs`
- `lib/Shared/Utils.gs`
- `lib/Snippets/cache/cache_manager.gs`
- `lib/Snippets/drive/pdf_export.gs`
- `lib/Snippets/sheets/sheet_utils.gs`
- `lib/Snippets/ui/sidebar_utils.gs`
- `lib/Snippets/utils/string_utils.gs`

Para cada arquivo `.html`:
1. No Apps Script, clicar em **+** > **HTML** para criar arquivo com o mesmo nome (sem extensão)
2. Colar o conteúdo

Arquivos HTML a copiar:
- `BomSidebar`
- `RequestSidebar`
- `SuperBuscaSidebar`
- `ConfigSidebar`
- `FixadoresSidebar`
- `SheetManager`
- `SummaryAllSidebar`
- `color-config-sidebar`
- `duplicate-dialog`
- `template-sidebar`

- [ ] **Step 3: Substituir o `appsscript.json`**

1. No Apps Script, clicar em **Project Settings** (ícone de engrenagem)
2. Marcar **Show "appsscript.json" manifest file in editor**
3. Abrir `appsscript.json` no editor
4. Substituir pelo conteúdo da Task 2

### 4B — Vincular projeto GCP

- [ ] **Step 4: Criar projeto GCP**

1. Abrir [console.cloud.google.com](https://console.cloud.google.com) na conta pessoal
2. Clicar em **Select a project** > **New project**
3. Nome: `PLB Sheet Tools`
4. Clicar em **Create**
5. Anotar o **Project number** (formato: `123456789012`)

- [ ] **Step 5: Habilitar a API Google Workspace Add-ons**

1. No GCP, ir em **APIs & Services** > **Enable APIs and Services**
2. Buscar `Google Workspace Add-ons API`
3. Clicar em **Enable**

- [ ] **Step 6: Vincular GCP ao Apps Script**

1. No Apps Script, ir em **Project Settings** > **Google Cloud Platform (GCP) Project**
2. Clicar em **Change project**
3. Colar o **Project number** do Step 4
4. Clicar em **Set project**

### 4C — Deploy como Add-on

- [ ] **Step 7: Criar o deployment**

1. No Apps Script, clicar em **Deploy** > **New deployment**
2. Clicar no ícone de engrenagem ao lado de **Select type** > escolher **Add-on**
3. Description: `v1.0 — PLB Sheet Tools Add-on privado`
4. Clicar em **Deploy**
5. **Copiar e salvar o Deployment ID** (será usado para gerar o link de instalação)

- [ ] **Step 8: Gerar link de instalação**

O link de instalação tem o formato:
```
https://script.google.com/macros/d/{DEPLOYMENT_ID}/edit
```

Para distribuir ao time, usar o link gerado pela própria interface do Apps Script em **Deploy > Manage deployments** > ícone de compartilhamento.

### 4D — Testar

- [ ] **Step 9: Instalar o add-on na conta pessoal**

1. Abrir uma planilha do Google Sheets **nova e vazia** (não a planilha de desenvolvimento)
2. Ir em **Extensions** > **Add-ons** > **Get add-ons**
3. Colar o link do deployment na barra de busca, ou usar o link direto de instalação
4. Autorizar as permissões solicitadas

Alternativa: em Apps Script, clicar em **Deploy** > **Test deployments** > **Install**.

- [ ] **Step 10: Verificar menu em nova planilha**

1. Abrir qualquer planilha Sheets diferente na mesma conta
2. Verificar que o menu `🛠️ Sheet Tools` aparece na barra superior
3. Clicar em `📊 Gerador de BOM` — deve abrir a sidebar normalmente

- [ ] **Step 11: Verificar menu em planilha de outro usuário (smoke test)**

1. Compartilhar o link de instalação com outro membro do time
2. Pedir para instalar e abrir uma planilha qualquer
3. Confirmar que o menu `🛠️ Sheet Tools` aparece

- [ ] **Step 12: Commit final**

```powershell
git add .
git commit -m "docs: plano executado — add-on v1.0 deployado e testado"
```

---

## Troubleshooting

### Menu não aparece após instalação

Causa comum: add-on instalado mas `onOpen` ainda não rodou.  
Solução: fechar e reabrir a planilha, ou ir em **Extensions** > **PLB Sheet Tools** > **Start**.

### Erro "Authorization required" ao clicar em ferramenta

Causa: `authMode === NONE` na primeira abertura.  
Solução: o menu mostrará `🔓 Ativar ferramentas PLB` — clicar nele dispara o fluxo de autorização. Após autorizar, fechar e reabrir a planilha para carregar o menu completo.

### `CardService is not defined`

Causa: script ainda configurado como container-bound (não standalone).  
Solução: criar novo projeto standalone em script.google.com (não via Extensions > Apps Script dentro do Sheets).

### Sidebar abre mas não carrega dados

Causa: o add-on não tem permissão `spreadsheets` para a planilha atual.  
Solução: verificar `oauthScopes` no `appsscript.json` e redeployar.

### Arquivos `lib/` ficam "soltos" no Apps Script UI

No editor online do Apps Script não há pastas — todos os `.gs` ficam no nível raiz com o nome do arquivo. Os nomes base (`Config`, `Logger`, `Utils`, `cache_manager`, etc.) são todos únicos neste projeto, portanto não há conflito. Se usar clasp em vez de cópia manual, a estrutura de pastas é preservada automaticamente.

### `ExportProReceiver.gs` — comportamento no add-on

O `doPost` copiado **não ficará ativo** no deployment de add-on (o tipo "Add-on" não expõe endpoint HTTP). O receiver HTTP atual do ExportPro é o `GAS_doPost.js` no projeto pyRevit (deployment separado como Web App) — não depende deste projeto. Copiar o `ExportProReceiver.gs` é seguro: apenas fica inativo.
