# Task 3 Report — Adaptação de Menu.gs

## Status
**DONE**

## Commit Hash
`3461a6b`

## Changes Made
- **File modified:** `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\Menu.gs` (única arquivo tocado)
- **Version bumped:** 3.1.0 → 4.0.0
- **Lines added:** 50 | **Lines removed:** 15

## Implementation Details

### Functions Added
1. **onFileScopeGranted(e)** — Chama `onOpen()` quando permissões de arquivo são concedidas
2. **onHomepage(e)** — Retorna Card com welcome message para o homepage do add-on
3. **showAuthPrompt()** — Alert para disparar fluxo de autorização completa

### onOpen() Enhanced
- Agora recebe parâmetro `e` (evento do add-on)
- **AuthMode guard:** Se `e.authMode === ScriptApp.AuthMode.NONE`, mostra apenas "Ativar ferramentas PLB"
- Menu completo aparece após autorização

### onEdit() Preserved
- Mantido no final do arquivo (sem alterações)

## Verification
✅ Menu.gs é o **único arquivo modificado**  
✅ Funções de callback `openBomSidebar`, `openRequestSidebar`, etc. **já existem** em outros `.gs` (sem mudança necessária)  
✅ Commit message segue convenção de commit  

## Notes
- LF→CRLF warning (expected no ambiente Windows)
- Pronto para test em Google Sheets como add-on privado
