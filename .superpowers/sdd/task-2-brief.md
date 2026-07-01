# Task 2 Brief — Atualizar `appsscript.json`

## Contexto

Task 2 de 4. O projeto foi copiado para `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\` no Task 1.
Agora precisa ter seu manifesto atualizado para habilitar o modo Add-on do Google Workspace.

## Arquivo a modificar

`C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\appsscript.json`

Conteúdo atual:
```json
{
  "timeZone": "America/Fortaleza",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

## Novo conteúdo (substitui o arquivo inteiro)

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

## Steps

1. Escrever o novo conteúdo no arquivo (substituição completa)
2. Validar JSON:
   ```powershell
   Get-Content "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\appsscript.json" | ConvertFrom-Json | ConvertTo-Json -Depth 10
   ```
   Esperado: JSON impresso sem erro.
3. Commitar:
   ```powershell
   git -C "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" add appsscript.json
   git -C "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" commit -m "feat: adiciona configuracao de add-on e oauth scopes ao manifesto"
   ```

## Constraint

- `timeZone` deve permanecer `"America/Fortaleza"` (não alterar)
- Nenhum outro arquivo deve ser modificado

## Report

Escreva em:
`C:\DEV\Sheets\Ferramentas-PLB-Sheets\.superpowers\sdd\task-2-report.md`

Inclua:
- Status: DONE / BLOCKED / NEEDS_CONTEXT
- Hash do commit
- Confirmação que JSON é válido
