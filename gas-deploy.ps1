# Google Apps Script Deploy Tool
# Rode via gas-deploy.bat

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = $PSScriptRoot

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

function Show-Header {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Google Apps Script - Deploy Tool" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Section([string]$text) {
    Write-Host ""
    Write-Host "--- $text ---" -ForegroundColor DarkCyan
}

function Wait-Return {
    Write-Host ""
    Write-Host "Pressione ENTER para voltar ao menu..." -ForegroundColor DarkGray
    $null = Read-Host
}

# ─────────────────────────────────────────────────────────────────────────────
# Encoding: SEMPRE UTF-8 sem BOM
# Windows PowerShell 5.1: `Set-Content -Encoding UTF8` e [Encoding]::UTF8 gravam
# BOM (EF BB BF). O clasp (Node.js) rejeita JSON com BOM:
#   "Unexpected token '(BOM)' ... is not valid JSON"
# Por isso TODA gravacao deste script passa por Write-Utf8NoBom.
# ─────────────────────────────────────────────────────────────────────────────

$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Write-Utf8NoBom([string]$path, [string]$content) {
    [System.IO.File]::WriteAllText($path, $content, $script:Utf8NoBom)
}

function Repair-BomFile([string]$path) {
    # Remove BOM in-place se presente. Retorna $true se reparou.
    if (-not (Test-Path $path)) { return $false }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        [System.IO.File]::WriteAllBytes($path, $bytes[3..($bytes.Length - 1)])
        return $true
    }
    return $false
}

function Repair-AllBoms {
    # Conserta arquivos ja contaminados por versoes antigas do script.
    # Evita ter que deletar .clasprc.json e reautenticar.
    $candidates = @("$env:USERPROFILE\.clasprc.json", (Join-Path $root ".clasp.json"))
    Get-ChildItem "$env:USERPROFILE\.clasprc-*.json" -ErrorAction SilentlyContinue |
        ForEach-Object { $candidates += $_.FullName }

    foreach ($f in $candidates) {
        if (Repair-BomFile $f) {
            Write-Host "  BOM removido: $f" -ForegroundColor DarkYellow
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Contas
# ─────────────────────────────────────────────────────────────────────────────

function Get-EmailFromClasprc([string]$filePath) {
    $cacheFile = $filePath -replace '\.json$', '-email.txt'
    if (Test-Path $cacheFile) { return (Get-Content $cacheFile -Raw).Trim() }
    return "autenticado"
}

function Get-SavedAccounts {
    $accounts = [ordered]@{}

    $main = "$env:USERPROFILE\.clasprc.json"
    if (Test-Path $main) {
        $accounts["principal"] = @{
            file  = $main
            label = (Get-EmailFromClasprc $main)
        }
    }

    Get-ChildItem "$env:USERPROFILE\.clasprc-*.json" -ErrorAction SilentlyContinue | ForEach-Object {
        $key = $_.BaseName -replace '^\.clasprc-', ''
        $accounts[$key] = @{
            file  = $_.FullName
            label = (Get-EmailFromClasprc $_.FullName)
        }
    }

    return $accounts
}

function Select-Account {
    $accounts = Get-SavedAccounts

    if ($accounts.Count -eq 0) {
        Write-Host "  Nenhuma conta autenticada. Use a opcao [2] Autenticar." -ForegroundColor Yellow
        return $null
    }

    if ($accounts.Count -eq 1) {
        $acc = $accounts[($accounts.Keys | Select-Object -First 1)]
        Write-Host "  Conta: $($acc.label)" -ForegroundColor DarkGray
        return $acc.file
    }

    Show-Section "Selecione a conta Google"
    $keys = @($accounts.Keys)
    for ($i = 0; $i -lt $keys.Count; $i++) {
        Write-Host "  [$($i+1)] $($accounts[$keys[$i]].label)" -ForegroundColor White
    }
    Write-Host ""
    $sel = Read-Host "  Numero"
    $idx = [int]$sel - 1
    if ($idx -lt 0 -or $idx -ge $keys.Count) {
        Write-Host "  Opcao invalida." -ForegroundColor Red
        return $null
    }
    return $accounts[$keys[$idx]].file
}

function Get-AccessToken([string]$authFile) {
    $data  = Get-Content $authFile -Raw | ConvertFrom-Json
    $creds = if ($data.PSObject.Properties['tokens']) { $data.tokens.default } else { $data.token }

    # Reusar access_token em cache se ainda valido (com 5 min de margem)
    $nowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    if ($creds.access_token -and $creds.expiry_date -and ($creds.expiry_date - $nowMs) -gt 300000) {
        return $creds.access_token
    }

    # Renovar via refresh_token (permanente, nao expira por tempo)
    $resp = Invoke-RestMethod "https://oauth2.googleapis.com/token" -Method Post -Body @{
        client_id     = $creds.client_id
        client_secret = $creds.client_secret
        refresh_token = $creds.refresh_token
        grant_type    = "refresh_token"
    }

    # Persistir novo access_token + validade no arquivo (evita refresh desnecessario)
    $expiryMs = $nowMs + ([long]$resp.expires_in * 1000)
    if ($creds.PSObject.Properties['access_token']) {
        $creds.access_token = $resp.access_token
    } else {
        $creds | Add-Member -NotePropertyName 'access_token' -NotePropertyValue $resp.access_token -Force
    }
    if ($creds.PSObject.Properties['expiry_date']) {
        $creds.expiry_date = $expiryMs
    } else {
        $creds | Add-Member -NotePropertyName 'expiry_date' -NotePropertyValue $expiryMs -Force
    }
    Write-Utf8NoBom $authFile ($data | ConvertTo-Json -Depth 10)

    # Cachear email para exibicao no menu
    try {
        $userinfo = Invoke-RestMethod "https://www.googleapis.com/oauth2/v3/userinfo" `
            -Headers @{ Authorization = "Bearer $($resp.access_token)" } -TimeoutSec 5
        $cacheFile = $authFile -replace '\.json$', '-email.txt'
        Write-Utf8NoBom $cacheFile $userinfo.email
    } catch {}

    return $resp.access_token
}

<# Atualiza o cache de email de uma conta (mostrado no menu). Retorna o email ou $null. #>
function Update-EmailCache([string]$authFile) {
    try {
        $token = Get-AccessToken $authFile
        $userinfo = Invoke-RestMethod "https://www.googleapis.com/oauth2/v3/userinfo" `
            -Headers @{ Authorization = "Bearer $token" } -TimeoutSec 5
        Write-Utf8NoBom ($authFile -replace '\.json$', '-email.txt') $userinfo.email
        return $userinfo.email
    } catch {
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Projetos salvos (gas-projects.json na raiz do repo)
# Permite dar push no mesmo codigo para varias planilhas (1, varias ou todas).
# ─────────────────────────────────────────────────────────────────────────────

$script:ProjectsFile = Join-Path $root "gas-projects.json"

function Get-SavedProjects {
    if (-not (Test-Path $script:ProjectsFile)) { return @() }
    try {
        $data = Get-Content $script:ProjectsFile -Raw | ConvertFrom-Json
        if ($data.projects) { return @($data.projects) }
    } catch {
        Write-Host "  Aviso: gas-projects.json invalido - ignorado." -ForegroundColor Yellow
    }
    return @()
}

function Save-Projects($projects) {
    $json = @{ projects = @($projects) } | ConvertTo-Json -Depth 5
    Write-Utf8NoBom $script:ProjectsFile $json
}

function Get-DefaultScriptId {
    $claspPath = Join-Path $root ".clasp.json"
    if (Test-Path $claspPath) {
        try { return (Get-Content $claspPath -Raw | ConvertFrom-Json).scriptId } catch {}
    }
    return ""
}

<# Nome amigavel de um scriptId, se estiver salvo #>
function Get-ProjectLabel([string]$scriptId) {
    $match = Get-SavedProjects | Where-Object { $_.scriptId -eq $scriptId } | Select-Object -First 1
    if ($match) { return $match.name } else { return $null }
}

<#
 Seleciona os destinos do push. Retorna array de @{ name; scriptId } ou $null.
 Opcoes: numero(s) da lista (ex: "1" ou "1,3"), T = todos, M = ID manual,
 ENTER = padrao do .clasp.json.
#>
function Select-PushTargets {
    $projects = Get-SavedProjects
    $default  = Get-DefaultScriptId

    Show-Section "Destino do push"
    if ($default) {
        $defLabel = Get-ProjectLabel $default
        if (-not $defLabel) { $defLabel = "(.clasp.json)" } else { $defLabel = "$defLabel (.clasp.json)" }
        Write-Host "  ENTER = padrao: $defLabel  $default" -ForegroundColor DarkGray
    }
    for ($i = 0; $i -lt $projects.Count; $i++) {
        Write-Host ("  [{0}] {1,-22} {2}" -f ($i + 1), $projects[$i].name, $projects[$i].scriptId) -ForegroundColor White
    }
    if ($projects.Count -gt 1) {
        Write-Host "  [T] TODOS os $($projects.Count) projetos salvos" -ForegroundColor Cyan
        Write-Host "      (numeros separados por virgula tambem valem, ex: 1,3)" -ForegroundColor DarkGray
    }
    Write-Host "  [M] Digitar outro Script ID" -ForegroundColor Gray
    if ($projects.Count -eq 0) {
        Write-Host "  Dica: salve projetos na opcao [5] do menu para push em lote." -ForegroundColor DarkGray
    }
    Write-Host ""

    $sel = (Read-Host "  Destino").Trim()

    # ENTER → padrao do .clasp.json
    if (-not $sel) {
        if (-not $default) {
            Write-Host "  Nenhum padrao definido no .clasp.json." -ForegroundColor Red
            return $null
        }
        $label = Get-ProjectLabel $default
        if (-not $label) { $label = "padrao" }
        return @(@{ name = $label; scriptId = $default })
    }

    # T → todos
    if ($sel -match '^[tT]$') {
        if ($projects.Count -eq 0) {
            Write-Host "  Nenhum projeto salvo." -ForegroundColor Red
            return $null
        }
        return @($projects | ForEach-Object { @{ name = $_.name; scriptId = $_.scriptId } })
    }

    # M → ID manual
    if ($sel -match '^[mM]$') {
        $typed = (Read-Host "  Script ID").Trim()
        if (-not $typed) { Write-Host "  ID vazio." -ForegroundColor Red; return $null }
        $label = Get-ProjectLabel $typed
        if (-not $label) { $label = "manual" }
        return @(@{ name = $label; scriptId = $typed })
    }

    # Numero(s): "2" ou "1,3"
    $targets = @()
    foreach ($part in ($sel -split ',')) {
        $n = 0
        if (-not [int]::TryParse($part.Trim(), [ref]$n) -or $n -lt 1 -or $n -gt $projects.Count) {
            Write-Host "  Opcao invalida: '$($part.Trim())'" -ForegroundColor Red
            return $null
        }
        $p = $projects[$n - 1]
        $targets += @{ name = $p.name; scriptId = $p.scriptId }
    }
    return $targets
}

# ─────────────────────────────────────────────────────────────────────────────
# Coletar arquivos
# ─────────────────────────────────────────────────────────────────────────────

function Get-FilesToPush([string]$folder) {
    # Manter alinhado com .claspignore: lib/Snippets e templates sao referencia
    # dev e NAO devem ir ao GAS (higiene 07/2026)
    $ignoreDirs  = @('.git', '.claude', '.superpowers', 'docs', 'node_modules', 'templates')
    $ignorePaths = @('lib\Snippets')
    $ignoreFiles = @('CLAUDE.md', 'README.md')
    $result      = [System.Collections.Generic.List[PSCustomObject]]::new()

    $manifest = Join-Path $folder "appsscript.json"
    if (Test-Path $manifest) {
        $result.Add([PSCustomObject]@{
            name   = "appsscript"
            type   = "JSON"
            source = (Get-Content $manifest -Raw -Encoding UTF8)
        })
    }

    Get-ChildItem -Path $folder -Recurse -File | Where-Object {
        $rel    = $_.FullName.Substring($folder.Length + 1)
        $topDir = $rel.Split([IO.Path]::DirectorySeparatorChar)[0]
        $ext    = $_.Extension.ToLower()

        if ($topDir -in $ignoreDirs)        { return $false }
        foreach ($p in $ignorePaths) {
            if ($rel.StartsWith($p, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
        }
        if ($_.Name -in $ignoreFiles)       { return $false }
        if ($_.Name -eq "appsscript.json")  { return $false }
        if ($ext -notin @('.gs', '.html'))  { return $false }
        return $true
    } | Sort-Object FullName | ForEach-Object {
        $rel = $_.FullName.Substring($folder.Length + 1) -replace '\\', '/'
        $src = Get-Content $_.FullName -Raw -Encoding UTF8

        if ($_.Extension.ToLower() -eq '.gs') {
            $result.Add([PSCustomObject]@{
                name   = ($rel -replace '\.gs$', '')
                type   = "SERVER_JS"
                source = $src
            })
        } else {
            $result.Add([PSCustomObject]@{
                name   = $rel
                type   = "HTML"
                source = $src
            })
        }
    }

    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# Acoes
# ─────────────────────────────────────────────────────────────────────────────

<#
 Envia os arquivos para UM projeto. Retorna @{ ok; message }.
 Nao lanca excecao -o chamador decide como reportar (loop de varios projetos).
#>
function Push-ToProject([string]$token, [string]$scriptId, $files) {
    try {
        # Verificar projeto
        $proj = Invoke-RestMethod `
            -Uri         "https://script.googleapis.com/v1/projects/$scriptId" `
            -Method      Get `
            -Headers     @{ Authorization = "Bearer $token" } `
            -TimeoutSec  30
        Write-Host "    Projeto: '$($proj.title)'" -ForegroundColor DarkGray

        # Serializar payload
        Add-Type -AssemblyName System.Web.Extensions
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $ser.MaxJsonLength = [int]::MaxValue
        $filesArray = @($files | ForEach-Object {
            @{ name=[string]$_.name; type=[string]$_.type; source=[string]$_.source }
        })
        $jsonBody = $ser.Serialize(@{ files = $filesArray })

        $tempJson = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.json'
        Write-Utf8NoBom $tempJson $jsonBody

        $url      = "https://script.googleapis.com/v1/projects/$scriptId/content"
        $respJson = & curl.exe -s -X PUT `
            -H "Authorization: Bearer $token" `
            -H "Content-Type: application/json" `
            --data-binary "@$tempJson" `
            --max-time 60 `
            $url

        Remove-Item $tempJson -Force -ErrorAction SilentlyContinue

        if (-not $respJson) { throw "curl nao retornou resposta. Verifique a conexao." }
        $result = $respJson | ConvertFrom-Json
        if ($result.error) { throw "API error $($result.error.code): $($result.error.message)" }

        return @{ ok = $true; message = "$($result.files.Count) arquivos em '$($proj.title)'" }

    } catch {
        $errMsg = "$_"
        $hint = ""
        if ($errMsg -match 'invalid_rapt|invalid_grant|reauth') {
            $hint = "Sessao expirada pelo Google Workspace - use [2] Autenticar."
        } elseif ($errMsg -match '400') {
            $hint = "Credenciais revogadas (refresh token invalido) - use [2] Autenticar."
        } elseif ($errMsg -match '401') {
            $hint = "Token expirado - use [2] Autenticar."
        } elseif ($errMsg -match '403') {
            $hint = "Sem permissao - verifique o Script ID ou a conta."
        } elseif ($errMsg -match '404') {
            $hint = "Projeto nao encontrado - Script ID invalido?"
        }
        $msg = $errMsg
        if ($hint) { $msg = "$errMsg`n    >> $hint" }
        return @{ ok = $false; message = $msg }
    }
}

function Action-Push {
    Show-Header
    Write-Host "  [1] PUSH - Enviar arquivos para Apps Script" -ForegroundColor White
    Write-Host ""

    $authFile = Select-Account
    if (-not $authFile) { Wait-Return; return }

    $targets = Select-PushTargets
    if (-not $targets -or $targets.Count -eq 0) { Wait-Return; return }

    Show-Section "Arquivos encontrados"
    $files = Get-FilesToPush $root
    $files | ForEach-Object {
        Write-Host ("  {0,-12} {1}" -f $_.type, $_.name) -ForegroundColor Gray
    }
    Write-Host "  Total: $($files.Count) arquivos -> $($targets.Count) projeto(s)" -ForegroundColor White

    # Confirmacao apenas para push em lote
    if ($targets.Count -gt 1) {
        Write-Host ""
        $names = ($targets | ForEach-Object { $_.name }) -join ', '
        Write-Host "  Enviar para: $names" -ForegroundColor Yellow
        $go = Read-Host "  Confirmar? (s/n)"
        if ($go.ToLower() -ne 's') { Wait-Return; return }
    }

    # Token uma vez so - reutilizado em todos os pushes
    Show-Section "Autenticacao"
    Write-Host "  Obtendo token..." -ForegroundColor DarkGray
    try {
        $token = Get-AccessToken $authFile
        Write-Host "  OK" -ForegroundColor Green
    } catch {
        Write-Host "  ERRO: $_" -ForegroundColor Red
        Write-Host "  >> Use opcao [2] Autenticar." -ForegroundColor Yellow
        Wait-Return; return
    }

    # Push por projeto (falha em um nao aborta os demais)
    $summary = @()
    foreach ($t in $targets) {
        Show-Section "Push: $($t.name)  ($($t.scriptId))"
        $r = Push-ToProject $token $t.scriptId $files
        if ($r.ok) {
            Write-Host "    OK - $($r.message)" -ForegroundColor Green
        } else {
            Write-Host "    ERRO: $($r.message)" -ForegroundColor Red
        }
        $summary += @{ name = $t.name; scriptId = $t.scriptId; ok = $r.ok }
    }

    # Resumo
    $okCount   = @($summary | Where-Object { $_.ok }).Count
    $failCount = $summary.Count - $okCount
    Write-Host ""
    if ($failCount -eq 0) {
        Write-Host "  ================================" -ForegroundColor Green
        Write-Host "  PUSH OK - $okCount de $($summary.Count) projeto(s)" -ForegroundColor Green
        Write-Host "  ================================" -ForegroundColor Green
    } else {
        Write-Host "  ================================" -ForegroundColor Red
        Write-Host "  PUSH: $okCount OK, $failCount FALHOU" -ForegroundColor Red
        $summary | Where-Object { -not $_.ok } | ForEach-Object {
            Write-Host "    FALHOU: $($_.name)  $($_.scriptId)" -ForegroundColor Red
        }
        Write-Host "  ================================" -ForegroundColor Red
    }

    # Oferecer salvar projeto usado com sucesso e ainda nao registrado
    $unsaved = @($summary | Where-Object { $_.ok -and -not (Get-ProjectLabel $_.scriptId) })
    foreach ($u in $unsaved) {
        Write-Host ""
        Write-Host "  O projeto $($u.scriptId) nao esta salvo." -ForegroundColor Gray
        $nm = (Read-Host "  Nome para salvar (ENTER para pular)").Trim()
        if ($nm) {
            $projects = @(Get-SavedProjects) + @([PSCustomObject]@{ name = $nm; scriptId = $u.scriptId })
            Save-Projects $projects
            Write-Host "  Salvo em gas-projects.json" -ForegroundColor Green
        }
    }

    Wait-Return
}

function Action-Reauth {
    Show-Header
    Write-Host "  [2] AUTENTICAR - Login Google" -ForegroundColor White
    Write-Host ""
    Write-Host "  Vai abrir o browser para login." -ForegroundColor Gray
    Write-Host "  Credenciais salvas em: $env:USERPROFILE\.clasprc.json" -ForegroundColor Gray
    Write-Host ""

    $confirm = Read-Host "  Continuar? (s/n)"
    if ($confirm.ToLower() -ne 's') { return }

    # BOM em .clasprc.json quebra o clasp ANTES do login
    # ("Unexpected token '(BOM)' ... is not valid JSON") - reparar antes.
    Repair-AllBoms

    clasp login
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "  clasp login falhou (exit $LASTEXITCODE)." -ForegroundColor Red
        Wait-Return; return
    }

    # clasp pode gravar com BOM dependendo da versao - garantir limpeza
    Repair-AllBoms

    Write-Host ""
    $mainAuth = "$env:USERPROFILE\.clasprc.json"
    $email = Update-EmailCache $mainAuth
    if ($email) {
        Write-Host "  Autenticado como: $email" -ForegroundColor Green
    } else {
        Write-Host "  Autenticado!" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  Quer salvar esta conta com um nome especifico?" -ForegroundColor Gray
    Write-Host "  (util para ter mais de uma conta, ex: work, pessoal)" -ForegroundColor Gray
    $apelido = (Read-Host "  Nome ou ENTER para pular").Trim()

    if ($apelido) {
        $dest = "$env:USERPROFILE\.clasprc-$apelido.json"
        Copy-Item $mainAuth $dest
        $emailCache = $mainAuth -replace '\.json$', '-email.txt'
        if (Test-Path $emailCache) {
            Copy-Item $emailCache ($dest -replace '\.json$', '-email.txt')
        }
        Write-Host "  Salvo em: $dest" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ATENCAO: agora faca login na sua conta principal novamente." -ForegroundColor Yellow
        $novamente = Read-Host "  Fazer login agora? (s/n)"
        if ($novamente.ToLower() -eq 's') {
            clasp login
            Repair-AllBoms
            $emailMain = Update-EmailCache $mainAuth
            Write-Host "  Conta principal restaurada$(if ($emailMain) { ": $emailMain" })." -ForegroundColor Green
        }
    }

    Wait-Return
}

function Action-ListAccounts {
    Show-Header
    Write-Host "  [3] CONTAS SALVAS" -ForegroundColor White
    Write-Host ""

    $accounts = Get-SavedAccounts
    if ($accounts.Count -eq 0) {
        Write-Host "  Nenhuma conta encontrada." -ForegroundColor Yellow
    } else {
        foreach ($key in $accounts.Keys) {
            Write-Host "  [$key]" -ForegroundColor White -NoNewline
            Write-Host "  $($accounts[$key].label)" -ForegroundColor Cyan
            Write-Host "        $($accounts[$key].file)" -ForegroundColor DarkGray
        }
    }

    Wait-Return
}

function Action-SetDefault {
    Show-Header
    Write-Host "  [4] DEFINIR PROJETO PADRAO" -ForegroundColor White
    Write-Host ""

    $claspPath = Join-Path $root ".clasp.json"
    if (Test-Path $claspPath) {
        $current = (Get-Content $claspPath -Raw | ConvertFrom-Json).scriptId
        Write-Host "  Atual: $current" -ForegroundColor DarkGray
    }

    Write-Host "  Cole o novo Script ID:" -ForegroundColor Gray
    $newId = (Read-Host "  ID").Trim()
    if (-not $newId) {
        Write-Host "  Cancelado." -ForegroundColor Yellow
        Wait-Return; return
    }

    # Sem BOM: o clasp (Node) tambem le este arquivo
    Write-Utf8NoBom $claspPath (@{ scriptId = $newId; rootDir = "." } | ConvertTo-Json)
    Write-Host "  .clasp.json atualizado: $newId" -ForegroundColor Green

    Wait-Return
}

function Action-Projects {
    while ($true) {
        Show-Header
        Write-Host "  [5] PROJETOS SALVOS (gas-projects.json)" -ForegroundColor White
        Write-Host ""

        $projects = Get-SavedProjects
        if ($projects.Count -eq 0) {
            Write-Host "  Nenhum projeto salvo." -ForegroundColor Yellow
        } else {
            for ($i = 0; $i -lt $projects.Count; $i++) {
                Write-Host ("  [{0}] {1,-22} {2}" -f ($i + 1), $projects[$i].name, $projects[$i].scriptId) -ForegroundColor White
            }
        }

        Write-Host ""
        Write-Host "  [A] Adicionar projeto" -ForegroundColor Gray
        Write-Host "  [R] Remover projeto"   -ForegroundColor Gray
        Write-Host "  [N] Renomear projeto"  -ForegroundColor Gray
        Write-Host "  [0] Voltar"            -ForegroundColor DarkGray
        Write-Host ""

        $op = (Read-Host "  Opcao").Trim().ToLower()

        switch ($op) {
            'a' {
                $nm = (Read-Host "  Nome do projeto").Trim()
                if (-not $nm) { Write-Host "  Nome vazio." -ForegroundColor Red; Start-Sleep 1; continue }
                if ($projects | Where-Object { $_.name -eq $nm }) {
                    Write-Host "  Ja existe projeto com esse nome." -ForegroundColor Red; Start-Sleep 2; continue
                }
                $id = (Read-Host "  Script ID").Trim()
                if (-not $id) { Write-Host "  ID vazio." -ForegroundColor Red; Start-Sleep 1; continue }
                if ($projects | Where-Object { $_.scriptId -eq $id }) {
                    Write-Host "  Esse Script ID ja esta salvo." -ForegroundColor Red; Start-Sleep 2; continue
                }
                Save-Projects (@($projects) + @([PSCustomObject]@{ name = $nm; scriptId = $id }))
                Write-Host "  Adicionado: $nm" -ForegroundColor Green
                Start-Sleep 1
            }
            'r' {
                if ($projects.Count -eq 0) { continue }
                $n = (Read-Host "  Numero do projeto a remover").Trim()
                $idx = 0
                if (-not [int]::TryParse($n, [ref]$idx) -or $idx -lt 1 -or $idx -gt $projects.Count) {
                    Write-Host "  Numero invalido." -ForegroundColor Red; Start-Sleep 1; continue
                }
                $victim = $projects[$idx - 1]
                $ok = Read-Host "  Remover '$($victim.name)'? (s/n)"
                if ($ok.ToLower() -eq 's') {
                    Save-Projects (@($projects | Where-Object { $_.scriptId -ne $victim.scriptId }))
                    Write-Host "  Removido." -ForegroundColor Green
                    Start-Sleep 1
                }
            }
            'n' {
                if ($projects.Count -eq 0) { continue }
                $n = (Read-Host "  Numero do projeto a renomear").Trim()
                $idx = 0
                if (-not [int]::TryParse($n, [ref]$idx) -or $idx -lt 1 -or $idx -gt $projects.Count) {
                    Write-Host "  Numero invalido." -ForegroundColor Red; Start-Sleep 1; continue
                }
                $novo = (Read-Host "  Novo nome").Trim()
                if (-not $novo) { Write-Host "  Nome vazio." -ForegroundColor Red; Start-Sleep 1; continue }
                $projects[$idx - 1].name = $novo
                Save-Projects $projects
                Write-Host "  Renomeado." -ForegroundColor Green
                Start-Sleep 1
            }
            '0' { return }
            default {
                Write-Host "  Opcao invalida." -ForegroundColor Red
                Start-Sleep 1
            }
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Menu principal
# ─────────────────────────────────────────────────────────────────────────────

# Auto-reparo na inicializacao: conserta .clasprc*.json/.clasp.json contaminados
# com BOM por versoes antigas deste script (sem precisar deletar/reautenticar).
Repair-AllBoms

while ($true) {
    Show-Header

    $curId = Get-DefaultScriptId
    if ($curId) {
        $curLabel = Get-ProjectLabel $curId
        if ($curLabel) {
            Write-Host "  Projeto padrao: $curLabel  ($curId)" -ForegroundColor DarkGray
        } else {
            Write-Host "  Projeto padrao: $curId" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "  Projeto padrao: nao definido" -ForegroundColor Yellow
    }

    $savedCount = (Get-SavedProjects).Count
    if ($savedCount -gt 0) {
        Write-Host "  Salvos:  $savedCount projeto(s) para push em lote" -ForegroundColor DarkGray
    }

    $mainAuth = "$env:USERPROFILE\.clasprc.json"
    if (Test-Path $mainAuth) {
        $email = Get-EmailFromClasprc $mainAuth
        Write-Host "  Conta:   $email" -ForegroundColor DarkGray
    } else {
        Write-Host "  Conta:   nao autenticado" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "  [1] Push - enviar para 1, varios ou TODOS os projetos" -ForegroundColor White
    Write-Host "  [2] Autenticar / trocar conta Google"                  -ForegroundColor White
    Write-Host "  [3] Ver contas salvas"                                 -ForegroundColor White
    Write-Host "  [4] Definir projeto padrao"                            -ForegroundColor White
    Write-Host "  [5] Projetos salvos (adicionar/remover/renomear)"      -ForegroundColor White
    Write-Host "  [0] Sair"                                              -ForegroundColor DarkGray
    Write-Host ""

    $choice = (Read-Host "  Opcao").Trim()

    switch ($choice) {
        '1' { Action-Push }
        '2' { Action-Reauth }
        '3' { Action-ListAccounts }
        '4' { Action-SetDefault }
        '5' { Action-Projects }
        '0' { Clear-Host; exit }
        default {
            Write-Host "  Opcao invalida." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}
