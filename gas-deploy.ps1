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
    $data | ConvertTo-Json -Depth 10 | Set-Content $authFile -Encoding UTF8

    # Cachear email para exibicao no menu
    try {
        $userinfo = Invoke-RestMethod "https://www.googleapis.com/oauth2/v3/userinfo" `
            -Headers @{ Authorization = "Bearer $($resp.access_token)" } -TimeoutSec 5
        $cacheFile = $authFile -replace '\.json$', '-email.txt'
        $userinfo.email | Set-Content $cacheFile -Encoding UTF8
    } catch {}

    return $resp.access_token
}

# ─────────────────────────────────────────────────────────────────────────────
# Script ID
# ─────────────────────────────────────────────────────────────────────────────

function Select-ScriptId {
    $claspPath = Join-Path $root ".clasp.json"
    $default   = ""
    if (Test-Path $claspPath) {
        $default = (Get-Content $claspPath -Raw | ConvertFrom-Json).scriptId
    }

    Show-Section "Script ID"
    if ($default) {
        Write-Host "  Padrao (.clasp.json): $default" -ForegroundColor DarkGray
        Write-Host "  ENTER = usar padrao, ou cole outro ID:" -ForegroundColor Gray
    } else {
        Write-Host "  Cole o Script ID do projeto Apps Script:" -ForegroundColor Gray
    }

    $typed = (Read-Host "  ID").Trim()
    if ($typed) { return $typed } else { return $default }
}

# ─────────────────────────────────────────────────────────────────────────────
# Coletar arquivos
# ─────────────────────────────────────────────────────────────────────────────

function Get-FilesToPush([string]$folder) {
    $ignoreDirs  = @('.git', '.claude', 'docs', 'node_modules')
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

function Action-Push {
    Show-Header
    Write-Host "  [1] PUSH - Enviar arquivos para Apps Script" -ForegroundColor White
    Write-Host ""

    $authFile = Select-Account
    if (-not $authFile) { Wait-Return; return }

    $scriptId = Select-ScriptId
    if (-not $scriptId) {
        Write-Host "  Script ID nao informado." -ForegroundColor Red
        Wait-Return; return
    }

    Show-Section "Arquivos encontrados"
    $files = Get-FilesToPush $root
    $files | ForEach-Object {
        Write-Host ("  {0,-12} {1}" -f $_.type, $_.name) -ForegroundColor Gray
    }
    Write-Host "  Total: $($files.Count) arquivos" -ForegroundColor White

    Show-Section "Enviando para $scriptId"
    $success = $false
    try {
        # Passo 1: token
        Write-Host "  [1/3] Obtendo token..." -ForegroundColor DarkGray
        $token = Get-AccessToken $authFile
        Write-Host "        OK" -ForegroundColor Green

        # Passo 2: verificar projeto
        Write-Host "  [2/3] Verificando projeto..." -ForegroundColor DarkGray
        $proj = Invoke-RestMethod `
            -Uri         "https://script.googleapis.com/v1/projects/$scriptId" `
            -Method      Get `
            -Headers     @{ Authorization = "Bearer $token" } `
            -TimeoutSec  30
        Write-Host "        OK - '$($proj.title)'" -ForegroundColor Green

        # Passo 3: push via curl.exe
        Write-Host "  [3/3] Enviando $($files.Count) arquivos..." -ForegroundColor DarkGray

        Add-Type -AssemblyName System.Web.Extensions
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $ser.MaxJsonLength = [int]::MaxValue
        $filesArray = @($files | ForEach-Object {
            @{ name=[string]$_.name; type=[string]$_.type; source=[string]$_.source }
        })
        $jsonBody = $ser.Serialize(@{ files = $filesArray })

        $tempJson = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.json'
        [System.IO.File]::WriteAllText($tempJson, $jsonBody, [System.Text.Encoding]::UTF8)

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

        Write-Host "        OK - $($result.files.Count) arquivos no projeto" -ForegroundColor Green

        $success = $true
    } catch {
        Write-Host ""
        $errMsg = "$_"
        Write-Host "  ERRO: $errMsg" -ForegroundColor Red
        if ($errMsg -match 'invalid_rapt|invalid_grant|reauth') {
            Write-Host ""
            Write-Host "  >> Sessao expirada pelo Google Workspace." -ForegroundColor Yellow
            Write-Host "  >> Use opcao [2] Autenticar para fazer login novamente." -ForegroundColor Yellow
        } elseif ($errMsg -match '400') {
            Write-Host "  >> Credenciais revogadas (refresh token invalido)." -ForegroundColor Yellow
            Write-Host "  >> Use opcao [2] Autenticar para fazer login novamente." -ForegroundColor Yellow
        } elseif ($errMsg -match '401') {
            Write-Host "  >> Token expirado. Use opcao [2] Autenticar." -ForegroundColor Yellow
        } elseif ($errMsg -match '403') {
            Write-Host "  >> Sem permissao. Verifique o Script ID ou a conta." -ForegroundColor Yellow
        } elseif ($errMsg -match '404') {
            Write-Host "  >> Projeto nao encontrado. Script ID invalido?" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    if ($success) {
        Write-Host "  ================================" -ForegroundColor Green
        Write-Host "  PUSH OK - $($result.files.Count) arquivos enviados" -ForegroundColor Green
        Write-Host "  ================================" -ForegroundColor Green
    } else {
        Write-Host "  ================================" -ForegroundColor Red
        Write-Host "  PUSH FALHOU" -ForegroundColor Red
        Write-Host "  ================================" -ForegroundColor Red
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

    clasp login

    Write-Host ""
    Write-Host "  Autenticado!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Quer salvar esta conta com um nome especifico?" -ForegroundColor Gray
    Write-Host "  (util para ter mais de uma conta, ex: work, pessoal)" -ForegroundColor Gray
    $apelido = (Read-Host "  Nome ou ENTER para pular").Trim()

    if ($apelido) {
        $dest = "$env:USERPROFILE\.clasprc-$apelido.json"
        Copy-Item "$env:USERPROFILE\.clasprc.json" $dest
        Write-Host "  Salvo em: $dest" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ATENCAO: agora faca login na sua conta principal novamente." -ForegroundColor Yellow
        $novamente = Read-Host "  Fazer login agora? (s/n)"
        if ($novamente.ToLower() -eq 's') {
            clasp login
            Write-Host "  Conta principal restaurada." -ForegroundColor Green
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

    @{ scriptId = $newId; rootDir = "." } | ConvertTo-Json | Set-Content $claspPath -Encoding UTF8
    Write-Host "  .clasp.json atualizado: $newId" -ForegroundColor Green

    Wait-Return
}

# ─────────────────────────────────────────────────────────────────────────────
# Menu principal
# ─────────────────────────────────────────────────────────────────────────────

while ($true) {
    Show-Header

    $claspPath = Join-Path $root ".clasp.json"
    if (Test-Path $claspPath) {
        $curId = (Get-Content $claspPath -Raw | ConvertFrom-Json).scriptId
        Write-Host "  Projeto: $curId" -ForegroundColor DarkGray
    } else {
        Write-Host "  Projeto: nao definido" -ForegroundColor Yellow
    }

    $mainAuth = "$env:USERPROFILE\.clasprc.json"
    if (Test-Path $mainAuth) {
        $email = Get-EmailFromClasprc $mainAuth
        Write-Host "  Conta:   $email" -ForegroundColor DarkGray
    } else {
        Write-Host "  Conta:   nao autenticado" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "  [1] Push - enviar arquivos para Apps Script" -ForegroundColor White
    Write-Host "  [2] Autenticar / trocar conta Google"        -ForegroundColor White
    Write-Host "  [3] Ver contas salvas"                       -ForegroundColor White
    Write-Host "  [4] Definir projeto padrao"                  -ForegroundColor White
    Write-Host "  [0] Sair"                                    -ForegroundColor DarkGray
    Write-Host ""

    $choice = (Read-Host "  Opcao").Trim()

    switch ($choice) {
        '1' { Action-Push }
        '2' { Action-Reauth }
        '3' { Action-ListAccounts }
        '4' { Action-SetDefault }
        '0' { Clear-Host; exit }
        default {
            Write-Host "  Opcao invalida." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}
