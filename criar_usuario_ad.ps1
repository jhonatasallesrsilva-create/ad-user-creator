param(
    [string]$PrimeiroNome,
    [string]$Sobrenome,
    [string]$Username,
    [string]$Template
)

# ============================================
# CONFIGURACOES - Edite o arquivo config.ps1
# ============================================
$ConfigFile = "$PSScriptRoot\..\config.ps1"

if (-not (Test-Path $ConfigFile)) {
    Write-Host "  [ERRO] Arquivo de configuracao nao encontrado: $ConfigFile" -ForegroundColor Red
    Write-Host "  Copie o arquivo 'config.example.ps1' para 'config.ps1' e preencha os dados." -ForegroundColor Yellow
    Read-Host "  Pressione ENTER para sair"
    exit 1
}

. $ConfigFile  # Importa as variaveis de configuracao
# ============================================

# --- Funcao de Log ---
function Write-Log {
    param([string]$Mensagem, [string]$Tipo = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linha = "[$timestamp] [$Tipo] $Mensagem"

    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $logFile = Join-Path $LogDir "criacao_usuarios_$(Get-Date -Format 'yyyy-MM').log"
    Add-Content -Path $logFile -Value $linha -Encoding UTF8

    switch ($Tipo) {
        "OK"    { Write-Host "  [OK] $Mensagem" -ForegroundColor Green }
        "ERRO"  { Write-Host "  [ERRO] $Mensagem" -ForegroundColor Red }
        "AVISO" { Write-Host "  [AVISO] $Mensagem" -ForegroundColor Yellow }
        "INFO"  { Write-Host "  [...] $Mensagem" -ForegroundColor Cyan }
    }
}

# Capitaliza primeira letra de cada palavra
$textInfo     = (Get-Culture).TextInfo
$PrimeiroNome = $textInfo.ToTitleCase($PrimeiroNome.ToLower().Trim())
$Sobrenome    = $textInfo.ToTitleCase($Sobrenome.ToLower().Trim())

$Senha  = ConvertTo-SecureString $SenhaInicial -AsPlainText -Force
$Email  = "$Username@$Dominio"

Write-Host ""
Write-Host "  ========================================" -ForegroundColor Cyan
Write-Host "   Criando: $PrimeiroNome $Sobrenome" -ForegroundColor Cyan
Write-Host "  ========================================" -ForegroundColor Cyan
Write-Host ""

Write-Log "Inicio - Usuario: $Username ($PrimeiroNome $Sobrenome)"

# --- ETAPA 1: Verificar pre-requisitos ---
Write-Host "  --- ETAPA 1: Verificando pre-requisitos ---" -ForegroundColor White

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "Modulo ActiveDirectory carregado." "OK"
} catch {
    Write-Log "Modulo ActiveDirectory nao encontrado!" "ERRO"
    Read-Host "  Pressione ENTER para sair"
    exit 1
}

# Verifica se usuario ja existe
try {
    $existente = Get-ADUser -Identity $Username -ErrorAction Stop
    Write-Log "Usuario '$Username' ja existe no AD! Abortando." "ERRO"
    Read-Host "  Pressione ENTER para sair"
    exit 1
} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    Write-Log "Username '$Username' disponivel." "OK"
} catch {
    Write-Log "Aviso ao verificar duplicata: $_" "AVISO"
}

# --- ETAPA 2: Buscar template ---
Write-Host ""
Write-Host "  --- ETAPA 2: Buscando usuario template ---" -ForegroundColor White

try {
    $templateUser = Get-ADUser -Identity $Template -Properties MemberOf, Department, Title, Company, Office, Description, Manager
} catch {
    Write-Log "Usuario template '$Template' nao encontrado!" "ERRO"
    Read-Host "  Pressione ENTER para sair"
    exit 1
}

$OU          = $templateUser.DistinguishedName -replace '^CN=[^,]+,', ''
$Setor       = $templateUser.Department
$NomeExibido = "$PrimeiroNome $Sobrenome - $Setor"

Write-Log "Template encontrado - Setor: $Setor" "OK"
Write-Log "OU destino: $OU" "INFO"

# --- ETAPA 3: Criar usuario no AD ---
Write-Host ""
Write-Host "  --- ETAPA 3: Criando usuario no Active Directory ---" -ForegroundColor White

try {
    $params = @{
        GivenName             = $PrimeiroNome
        Surname               = $Sobrenome
        Name                  = $NomeExibido
        DisplayName           = $NomeExibido
        SamAccountName        = $Username
        UserPrincipalName     = $Email
        EmailAddress          = $Email
        OfficePhone           = $Telefone
        HomePage              = $PaginaWeb
        AccountPassword       = $Senha
        Enabled               = $true
        ChangePasswordAtLogon = $false
        Path                  = $OU
    }

    if ($templateUser.Department)  { $params["Department"]  = $templateUser.Department }
    if ($templateUser.Title)       { $params["Title"]       = $templateUser.Title }
    if ($templateUser.Company)     { $params["Company"]     = $templateUser.Company }
    if ($templateUser.Office)      { $params["Office"]      = $templateUser.Office }
    if ($templateUser.Description) { $params["Description"] = $templateUser.Description }

    New-ADUser @params

    if ($templateUser.Manager) {
        Set-ADUser -Identity $Username -Manager $templateUser.Manager
    }

    Write-Log "Usuario '$Username' criado no AD!" "OK"

} catch {
    Write-Log "Falha ao criar usuario: $_" "ERRO"
    Read-Host "  Pressione ENTER para sair"
    exit 1
}

# --- ETAPA 4: Copiar grupos ---
Write-Host ""
Write-Host "  --- ETAPA 4: Copiando grupos do template ---" -ForegroundColor White

$gruposOk = 0
foreach ($grupo in $templateUser.MemberOf) {
    try {
        Add-ADGroupMember -Identity $grupo -Members $Username -ErrorAction Stop
        $nome = ($grupo -split ',')[0] -replace 'CN=', ''
        Write-Log "Grupo: $nome" "OK"
        $gruposOk++
    } catch {
        $nome = ($grupo -split ',')[0] -replace 'CN=', ''
        Write-Log "Falha no grupo '$nome': $_" "AVISO"
    }
}
Write-Log "Total: $gruposOk grupo(s) copiado(s)" "INFO"

# --- ETAPA 5: Sincronizar com Azure AD ---
Write-Host ""
Write-Host "  --- ETAPA 5: Sincronizando com Microsoft 365 ---" -ForegroundColor White

$syncOk = $false
try {
    Import-Module ADSync -ErrorAction Stop
    Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
    Write-Log "Sincronizacao Delta iniciada." "OK"
    $syncOk = $true
} catch {
    Write-Log "ADSync nao disponivel: $_" "AVISO"
}

$espera = if ($syncOk) { 90 } else { 120 }
Write-Log "Aguardando $espera segundos..." "INFO"
Start-Sleep -Seconds $espera

# --- ETAPA 6: Atribuir licenca M365 ---
Write-Host ""
Write-Host "  --- ETAPA 6: Atribuindo licenca Microsoft 365 ---" -ForegroundColor White

try {
    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }
    $token = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body).access_token
    $headers = @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }
    Write-Log "Token Graph API obtido." "OK"

    $usuario365 = $null
    for ($i = 1; $i -le 10; $i++) {
        Write-Host "  [...] Procurando usuario no M365... tentativa $i/10" -ForegroundColor Yellow
        try {
            $usuario365 = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$Email" -Headers $headers -ErrorAction Stop
            break
        } catch {
            if ($i -lt 10) { Start-Sleep -Seconds 30 }
        }
    }

    if (-not $usuario365) {
        Write-Log "Usuario nao encontrado no M365 apos 5 min. Atribua manualmente." "AVISO"
    } else {
        Write-Log "Usuario encontrado no M365!" "OK"

        $bodyLoc = @{ usageLocation = $UsageLocation } | ConvertTo-Json
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$Email" -Method Patch -Headers $headers -Body $bodyLoc | Out-Null
        Write-Log "UsageLocation definido: $UsageLocation" "OK"

        $licencas = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers).value

        $standard = $licencas | Where-Object { $_.skuPartNumber -eq "SPB" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
        $appsEnt  = $licencas | Where-Object { $_.skuPartNumber -eq "OFFICESUBSCRIPTION" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
        $e1       = $licencas | Where-Object { $_.skuPartNumber -eq "STANDARDPACK" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }

        $skuIds      = @()
        $nomeLicenca = ""

        if ($standard) {
            $skuIds      = @($standard.skuId)
            $nomeLicenca = "Microsoft 365 Business Standard"
        } elseif ($appsEnt -and $e1) {
            $skuIds      = @($appsEnt.skuId, $e1.skuId)
            $nomeLicenca = "M365 Apps Enterprise + E1"
        } else {
            Write-Log "Nenhuma licenca disponivel! Atribua em admin.microsoft.com" "AVISO"
        }

        if ($skuIds.Count -gt 0) {
            foreach ($id in $skuIds) {
                $bodyLic = @{ addLicenses = @(@{ skuId = $id }); removeLicenses = @() } | ConvertTo-Json -Depth 5
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$Email/assignLicense" -Method Post -Headers $headers -Body $bodyLic | Out-Null
            }
            Write-Log "Licenca atribuida: $nomeLicenca" "OK"
        }
    }
} catch {
    Write-Log "Erro na licenca M365: $_" "AVISO"
    Write-Log "Atribua manualmente em admin.microsoft.com" "AVISO"
}

# --- RESUMO FINAL ---
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "   USUARIO CRIADO COM SUCESSO!" -ForegroundColor Green
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "   Nome     : $NomeExibido" -ForegroundColor Green
Write-Host "   Login    : $Username" -ForegroundColor Green
Write-Host "   Email    : $Email" -ForegroundColor Green
Write-Host "   Setor    : $Setor" -ForegroundColor Green
Write-Host "   Grupos   : $gruposOk copiado(s)" -ForegroundColor Green
Write-Host "   Senha    : $SenhaInicial" -ForegroundColor Green
if ($nomeLicenca) {
    Write-Host "   Licenca  : $nomeLicenca" -ForegroundColor Green
}
Write-Host "  ============================================" -ForegroundColor Green
Write-Host ""

Write-Log "Processo finalizado para '$Username'."
