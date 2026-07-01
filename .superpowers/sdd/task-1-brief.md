# Task 1 Brief — Criar pasta do add-on e copiar projeto

## Contexto

Este é o Task 1 de 4 de um plano para converter as Ferramentas PLB de script
container-bound para Google Workspace Add-on privado.

A pasta de origem (`Ferramentas-PLB-Sheets`) deve ser **preservada intacta**.
O add-on será desenvolvido na pasta nova (`Ferramentas-PLB-Sheets-AddOn`).

## Constraint Global

- Pasta original `C:\DEV\Sheets\Ferramentas-PLB-Sheets\` **NÃO deve ser modificada**.
- Nova pasta: `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\`
- Novo repo git independente na pasta nova (não submodule, não subpasta do repo atual).
- A nova pasta começa com exatamente os mesmos arquivos da origem.

## Steps

### Step 1: Copiar projeto inteiro para nova pasta

```powershell
Copy-Item "C:\DEV\Sheets\Ferramentas-PLB-Sheets" `
          "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" `
          -Recurse
```

**Atenção:** se a pasta destino já existir, o comando acima cria uma subpasta.
Verificar antes:

```powershell
Test-Path "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn"
```

Se retornar `True`, a pasta já existe — não recriar, apenas verificar se está correta.

### Step 2: Verificar cópia completa

```powershell
$src = (Get-ChildItem "C:\DEV\Sheets\Ferramentas-PLB-Sheets" -Recurse -File | Measure-Object).Count
$dst = (Get-ChildItem "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn" -Recurse -File | Measure-Object).Count
"Origem: $src | Destino: $dst"
```

Esperado: contagens iguais.

### Step 3: Inicializar git na nova pasta

```powershell
Set-Location "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn"
git init
git add .
git commit -m "chore: copia inicial do projeto para versao add-on"
```

**Excluir a pasta `.git` copiada da origem** (se houver — o Copy-Item copia tudo):

```powershell
# Antes do git init, remover .git copiado se existir
if (Test-Path "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\.git") {
    Remove-Item "C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\.git" -Recurse -Force
}
```

### Step 4: Verificar que original está intacto

```powershell
git -C "C:\DEV\Sheets\Ferramentas-PLB-Sheets" status
```

Esperado: `nothing to commit, working tree clean` (sem mudanças no original).

## Deliverable

- Pasta `C:\DEV\Sheets\Ferramentas-PLB-Sheets-AddOn\` criada com todos os arquivos
- Git inicializado com commit inicial
- Original `C:\DEV\Sheets\Ferramentas-PLB-Sheets\` sem alterações

## Report

Escreva o relatório em:
`C:\DEV\Sheets\Ferramentas-PLB-Sheets\.superpowers\sdd\task-1-report.md`

Inclua:
- Status: DONE / BLOCKED / NEEDS_CONTEXT
- Commits criados (hash curto)
- Contagem de arquivos copiados
- Resultado do `git status` no original
- Qualquer preocupação (DONE_WITH_CONCERNS se houver)
