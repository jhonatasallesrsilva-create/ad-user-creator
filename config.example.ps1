# ============================================================
# CONFIGURACOES - AD User Creator
# ============================================================
# INSTRUCOES:
#   1. Copie este arquivo e renomeie para: config.ps1
#   2. Preencha com os dados da SUA organizacao
#   3. NUNCA faca commit do arquivo config.ps1 (ja esta no .gitignore)
# ============================================================

# --- Microsoft Graph API (Azure App Registration) ---
# Acesse: https://portal.azure.com > Azure Active Directory > App registrations
$TenantId     = "SEU-TENANT-ID-AQUI"           # Ex: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
$ClientId     = "SEU-CLIENT-ID-AQUI"            # Application (client) ID
$ClientSecret = "SEU-CLIENT-SECRET-AQUI"        # Secrets > New client secret

# --- Configuracoes da Organizacao ---
$Dominio       = "suaempresa.com.br"            # Dominio do Active Directory / M365
$SenhaInicial  = "SenhaForte@2026"             # Senha padrao para novos usuarios
$Telefone      = "(XX) XXXX-XXXX"              # Telefone da empresa (campo AD)
$PaginaWeb     = "www.suaempresa.com.br"        # Site da empresa (campo AD)
$UsageLocation = "BR"                           # Codigo do pais para licenca M365

# --- Diretorios de Log ---
$LogDir = "$PSScriptRoot\logs"
