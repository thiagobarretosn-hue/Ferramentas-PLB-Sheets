# Task 3 Brief — Atualizar `Menu.gs`

## Contexto

Task 3 de 4. O manifesto `appsscript.json` já declara `onHomepage` e `onFileScopeGranted`.
Agora `Menu.gs` precisa implementar essas funções e adaptar `onOpen` para o modo add-on.

## Arquivo a modificar

`C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\Menu.gs`

Conteúdo atual:
```javascript
/**
 * @fileoverview Menu Principal - Ferramentas PLB Sheets
 * @version 3.1.0 - Cores unificadas em um item; Gerenciador de Abas atualizado
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Sheet Tools')
    .addItem('📊 Gerador de BOM', 'openBomSidebar')
    .addItem('📋 Gerador de Request', 'openRequestSidebar')
    .addSeparator()
    .addItem('📑 Gerenciador de Abas', 'showSheetManager')
    .addItem('🎨 Cores das Abas', 'openColorConfig')
    .addSeparator()
    .addItem('🔍 Super Busca', 'abrirSuperBuscaSidebar')
    .addSeparator()
    .addItem('📊 Summary All', 'openSummaryAllSidebar')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
}
```

## Novo conteúdo (substitui o arquivo inteiro)

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

## Steps

1. Substituir o conteúdo de `Menu.gs` pelo novo conteúdo acima (substituição completa)
2. Commitar:
   ```powershell
   git -C "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" add Menu.gs
   git -C "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" commit -m "feat: adapta Menu.gs para add-on (authMode guard, onHomepage, onFileScopeGranted)"
   ```

## Constraints

- Somente `Menu.gs` pode ser modificado
- Nenhum outro arquivo deve ser tocado
- As funções `openBomSidebar`, `openRequestSidebar`, `showSheetManager`, `openColorConfig`, `abrirSuperBuscaSidebar`, `openSummaryAllSidebar` já existem nos outros `.gs` — não precisam ser criadas aqui
- `onEdit(e)` deve ser mantido no final do arquivo

## Report

Escreva em:
`C:\DEV\Sheets\Ferramentas-PLB-Sheets\.superpowers\sdd\task-3-report.md`

Inclua:
- Status: DONE / BLOCKED / NEEDS_CONTEXT
- Hash do commit
- Confirmação que apenas Menu.gs foi modificado
