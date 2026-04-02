# ============================================================
# CRIACAO DE USUARIO - AD + Microsoft 365 | Interface Grafica
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# CARREGAR CONFIGURACOES
# ============================================================
$ConfigFile = "$PSScriptRoot\config.ps1"

if (-not (Test-Path $ConfigFile)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Arquivo de configuracao nao encontrado!`n`nCopie 'config.example.ps1' para 'config.ps1' e preencha os dados.",
        "Configuracao ausente", "OK", "Error"
    )
    exit
}

. $ConfigFile
# ============================================================

# ============================================================
# GERENCIAR CREDENCIAIS GRAPH API (salvas encriptadas)
# ============================================================
$CredFile = "$PSScriptRoot\creds.xml"

# Chave de criptografia - gerada automaticamente por maquina (DPAPI)
# Nao usa chave hardcoded; usa ConvertFrom-SecureString sem -Key
# para que apenas o usuario atual nesta maquina possa descriptografar.

function Get-StoredCredentials {
    if (Test-Path $CredFile) {
        try {
            $config = Import-Clixml -Path $CredFile
            if (-not $config -or -not $config.ClientSecret) { throw "Config vazio" }
            return @{
                TenantId     = $config.TenantId
                ClientId     = $config.ClientId
                ClientSecret = $config.ClientSecret | ConvertTo-SecureString
            }
        } catch {
            Remove-Item -Path $CredFile -Force -ErrorAction SilentlyContinue
            return $null
        }
    }
    return $null
}

function Save-Credentials {
    param([string]$TenantId, [string]$ClientId, [System.Security.SecureString]$ClientSecret)
    $config = @{
        TenantId     = $TenantId
        ClientId     = $ClientId
        ClientSecret = $ClientSecret | ConvertFrom-SecureString  # DPAPI: apenas este usuario/maquina
    }
    $config | Export-Clixml -Path $CredFile -Force
}

function Show-CredentialForm {
    $credForm = New-Object System.Windows.Forms.Form
    $credForm.Text = "Configuracao de Credenciais - Microsoft Graph API"
    $credForm.Size = New-Object System.Drawing.Size(520, 340)
    $credForm.StartPosition = "CenterScreen"
    $credForm.FormBorderStyle = "FixedDialog"
    $credForm.MaximizeBox = $false
    $credForm.MinimizeBox = $false
    $credForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 46)
    $credForm.ForeColor = [System.Drawing.Color]::White
    $credForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Configure as credenciais do Azure AD (App Registration).`nOs dados serao salvos de forma segura (DPAPI) para este usuario."
    $lblInfo.Size = New-Object System.Drawing.Size(470, 45)
    $lblInfo.Location = New-Object System.Drawing.Point(20, 15)
    $lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 200)
    $credForm.Controls.Add($lblInfo)

    $corInput = [System.Drawing.Color]::FromArgb(55, 58, 78)

    $lblT = New-Object System.Windows.Forms.Label
    $lblT.Text = "Tenant ID:"
    $lblT.Location = New-Object System.Drawing.Point(20, 70)
    $lblT.AutoSize = $true
    $credForm.Controls.Add($lblT)
    $txtT = New-Object System.Windows.Forms.TextBox
    $txtT.Size = New-Object System.Drawing.Size(460, 28)
    $txtT.Location = New-Object System.Drawing.Point(20, 93)
    $txtT.BackColor = $corInput
    $txtT.ForeColor = [System.Drawing.Color]::White
    $txtT.BorderStyle = "FixedSingle"
    $credForm.Controls.Add($txtT)

    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = "Client ID (Application ID):"
    $lblC.Location = New-Object System.Drawing.Point(20, 128)
    $lblC.AutoSize = $true
    $credForm.Controls.Add($lblC)
    $txtC = New-Object System.Windows.Forms.TextBox
    $txtC.Size = New-Object System.Drawing.Size(460, 28)
    $txtC.Location = New-Object System.Drawing.Point(20, 151)
    $txtC.BackColor = $corInput
    $txtC.ForeColor = [System.Drawing.Color]::White
    $txtC.BorderStyle = "FixedSingle"
    $credForm.Controls.Add($txtC)

    $lblS = New-Object System.Windows.Forms.Label
    $lblS.Text = "Client Secret:"
    $lblS.Location = New-Object System.Drawing.Point(20, 186)
    $lblS.AutoSize = $true
    $credForm.Controls.Add($lblS)
    $txtS = New-Object System.Windows.Forms.TextBox
    $txtS.Size = New-Object System.Drawing.Size(460, 28)
    $txtS.Location = New-Object System.Drawing.Point(20, 209)
    $txtS.BackColor = $corInput
    $txtS.ForeColor = [System.Drawing.Color]::White
    $txtS.BorderStyle = "FixedSingle"
    $txtS.UseSystemPasswordChar = $true
    $credForm.Controls.Add($txtS)

    $btnSalvar = New-Object System.Windows.Forms.Button
    $btnSalvar.Text = "SALVAR E CONTINUAR"
    $btnSalvar.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnSalvar.Size = New-Object System.Drawing.Size(460, 40)
    $btnSalvar.Location = New-Object System.Drawing.Point(20, 250)
    $btnSalvar.FlatStyle = "Flat"
    $btnSalvar.FlatAppearance.BorderSize = 0
    $btnSalvar.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnSalvar.ForeColor = [System.Drawing.Color]::White
    $btnSalvar.Cursor = [System.Windows.Forms.Cursors]::Hand

    $btnSalvar.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtT.Text) -or
            [string]::IsNullOrWhiteSpace($txtC.Text) -or
            [string]::IsNullOrWhiteSpace($txtS.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Preencha todos os campos.", "Erro", "OK", "Warning")
            return
        }
        $secSecret = ConvertTo-SecureString $txtS.Text -AsPlainText -Force
        Save-Credentials -TenantId $txtT.Text.Trim() -ClientId $txtC.Text.Trim() -ClientSecret $secSecret
        $credForm.Tag = @{
            TenantId     = $txtT.Text.Trim()
            ClientId     = $txtC.Text.Trim()
            ClientSecret = $secSecret
        }
        $credForm.DialogResult = "OK"
        $credForm.Close()
    })
    $credForm.Controls.Add($btnSalvar)

    $existing = Get-StoredCredentials
    if ($existing) {
        $txtT.Text = $existing.TenantId
        $txtC.Text = $existing.ClientId
    }

    $result = $credForm.ShowDialog()
    if ($result -eq "OK") { return $credForm.Tag }
    return $null
}

# ============================================================
# CARREGAR OU SOLICITAR CREDENCIAIS
# ============================================================
$creds = Get-StoredCredentials
if (-not $creds) {
    $creds = Show-CredentialForm
    if (-not $creds) {
        [System.Windows.Forms.MessageBox]::Show("Credenciais nao configuradas. O aplicativo sera encerrado.", "Erro", "OK", "Error")
        exit
    }
}

$TenantId     = $creds.TenantId
$ClientId     = $creds.ClientId
$SecureSecret = $creds.ClientSecret

# ============================================================
# FUNCOES AUXILIARES
# ============================================================
function Write-Log {
    param([string]$Mensagem, [string]$Tipo = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linha = "[$timestamp] [$Tipo] $Mensagem"
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $logFile = Join-Path $LogDir "criacao_usuarios_$(Get-Date -Format 'yyyy-MM').log"
    Add-Content -Path $logFile -Value $linha -Encoding UTF8
}

function Add-LogMessage {
    param([string]$Mensagem, [string]$Tipo = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $icon = switch ($Tipo) {
        "OK"    { [char]0x2714 }
        "ERRO"  { [char]0x2718 }
        "AVISO" { [char]0x26A0 }
        "INFO"  { [char]0x27A4 }
        "ETAPA" { [char]0x25B6 }
        default { " " }
    }
    $cor = switch ($Tipo) {
        "OK"    { [System.Drawing.Color]::FromArgb(46, 204, 113) }
        "ERRO"  { [System.Drawing.Color]::FromArgb(231, 76, 60) }
        "AVISO" { [System.Drawing.Color]::FromArgb(241, 196, 15) }
        "INFO"  { [System.Drawing.Color]::FromArgb(52, 152, 219) }
        "ETAPA" { [System.Drawing.Color]::FromArgb(255, 255, 255) }
        default { [System.Drawing.Color]::FromArgb(180, 180, 200) }
    }
    $texto = "[$timestamp] $icon $Mensagem`r`n"
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor = $cor
    $txtLog.AppendText($texto)
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    Write-Log $Mensagem $Tipo
}

function Save-HistoricoCsv {
    param(
        [string]$Nome, [string]$Username, [string]$Email,
        [string]$Setor, [string]$Template, [string]$Licenca, [int]$Grupos
    )
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $HistoricoCsv = Join-Path $LogDir "historico_usuarios.csv"
    $registro = [PSCustomObject]@{
        DataHora  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Nome      = $Nome
        Username  = $Username
        Email     = $Email
        Setor     = $Setor
        Template  = $Template
        Licenca   = $Licenca
        Grupos    = $Grupos
        CriadoPor = $env:USERNAME
    }
    $registro | Export-Csv -Path $HistoricoCsv -Append -NoTypeInformation -Encoding UTF8
}

function Update-Username {
    $nome = $txtNome.Text.Trim()
    $sobrenome = $txtSobrenome.Text.Trim()
    if ($nome -ne "" -and $sobrenome -ne "") {
        $user = ("$nome.$sobrenome").ToLower() -replace ' ','' -replace '[^a-z0-9.]',''
        $txtUsername.Text = $user
        $txtEmail.Text = "$user@$Dominio"
    } else {
        $txtUsername.Text = ""
        $txtEmail.Text = ""
    }
}

function DoEvents-Sleep {
    param([int]$Seconds)
    for ($i = 0; $i -lt ($Seconds * 4); $i++) {
        Start-Sleep -Milliseconds 250
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ============================================================
# INTERFACE
# ============================================================
$corFundo       = [System.Drawing.Color]::FromArgb(30, 30, 46)
$corPainel      = [System.Drawing.Color]::FromArgb(40, 42, 58)
$corPainelClaro = [System.Drawing.Color]::FromArgb(50, 52, 70)
$corAzul        = [System.Drawing.Color]::FromArgb(0, 120, 212)
$corAzulHover   = [System.Drawing.Color]::FromArgb(30, 144, 235)
$corVerde       = [System.Drawing.Color]::FromArgb(46, 204, 113)
$corVermelho    = [System.Drawing.Color]::FromArgb(231, 76, 60)
$corTexto       = [System.Drawing.Color]::White
$corTextoCinza  = [System.Drawing.Color]::FromArgb(180, 180, 200)
$corInput       = [System.Drawing.Color]::FromArgb(55, 58, 78)
$corBorda       = [System.Drawing.Color]::FromArgb(80, 84, 110)

$fonteTitulo    = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$fonteSubtitulo = New-Object System.Drawing.Font("Segoe UI", 9)
$fonteLabel     = New-Object System.Drawing.Font("Segoe UI", 10)
$fonteLabelBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$fonteInput     = New-Object System.Drawing.Font("Segoe UI", 11)
$fonteBotao     = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$fonteLog       = New-Object System.Drawing.Font("Consolas", 9)
$fonteInfo      = New-Object System.Drawing.Font("Segoe UI", 10)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Criacao de Usuario - AD + Microsoft 365"
$form.Size = New-Object System.Drawing.Size(750, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = $corFundo
$form.ForeColor = $corTexto
$form.Font = $fonteLabel

# Cabecalho
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Size = New-Object System.Drawing.Size(750, 70)
$pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
$pnlHeader.BackColor = $corAzul

$lblTitulo = New-Object System.Windows.Forms.Label
$lblTitulo.Text = "  Criacao de Usuario"
$lblTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblTitulo.ForeColor = [System.Drawing.Color]::White
$lblTitulo.Size = New-Object System.Drawing.Size(400, 32)
$lblTitulo.Location = New-Object System.Drawing.Point(10, 8)
$pnlHeader.Controls.Add($lblTitulo)

$lblSubtitulo = New-Object System.Windows.Forms.Label
$lblSubtitulo.Text = "  Active Directory + Microsoft 365"
$lblSubtitulo.Font = $fonteSubtitulo
$lblSubtitulo.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
$lblSubtitulo.Size = New-Object System.Drawing.Size(500, 20)
$lblSubtitulo.Location = New-Object System.Drawing.Point(10, 40)
$pnlHeader.Controls.Add($lblSubtitulo)

# Logo opcional
$logoPath = Join-Path $PSScriptRoot "logo.png"
if (Test-Path $logoPath) {
    $logoImage = [System.Drawing.Image]::FromFile($logoPath)
    $picLogo = New-Object System.Windows.Forms.PictureBox
    $picLogo.SizeMode = "Zoom"
    $picLogo.Size = New-Object System.Drawing.Size(140, 42)
    $picLogo.Location = New-Object System.Drawing.Point(510, 14)
    $picLogo.BackColor = [System.Drawing.Color]::Transparent
    $picLogo.Image = $logoImage
    $pnlHeader.Controls.Add($picLogo)
    try {
        $bmp = New-Object System.Drawing.Bitmap($logoImage, 32, 32)
        $form.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    } catch {}
}

# Botao de configuracoes
$btnConfig = New-Object System.Windows.Forms.Button
$btnConfig.Text = [char]0x2699
$btnConfig.Font = New-Object System.Drawing.Font("Segoe UI", 16)
$btnConfig.Size = New-Object System.Drawing.Size(45, 45)
$btnConfig.Location = New-Object System.Drawing.Point(680, 13)
$btnConfig.FlatStyle = "Flat"
$btnConfig.FlatAppearance.BorderSize = 0
$btnConfig.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 190)
$btnConfig.ForeColor = [System.Drawing.Color]::White
$btnConfig.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnConfig.Add_Click({
    $result = Show-CredentialForm
    if ($result) {
        $script:TenantId     = $result.TenantId
        $script:ClientId     = $result.ClientId
        $script:SecureSecret = $result.ClientSecret
        [System.Windows.Forms.MessageBox]::Show("Credenciais atualizadas com sucesso!", "Configuracao", "OK", "Information")
    }
})
$pnlHeader.Controls.Add($btnConfig)
$form.Controls.Add($pnlHeader)

# Formulario
$pnlForm = New-Object System.Windows.Forms.Panel
$pnlForm.Size = New-Object System.Drawing.Size(710, 250)
$pnlForm.Location = New-Object System.Drawing.Point(15, 80)
$pnlForm.BackColor = $corPainel
$form.Controls.Add($pnlForm)

function New-FormLabel($text, $x, $y, $bold = $false) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Font = if ($bold) { $fonteLabelBold } else { $fonteLabel }
    $lbl.ForeColor = $corTextoCinza
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    return $lbl
}
function New-FormInput($x, $y, $w) {
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Font = $fonteInput
    $txt.Size = New-Object System.Drawing.Size($w, 32)
    $txt.Location = New-Object System.Drawing.Point($x, $y)
    $txt.BackColor = $corInput
    $txt.ForeColor = $corTexto
    $txt.BorderStyle = "FixedSingle"
    return $txt
}

$pnlForm.Controls.Add((New-FormLabel "Primeiro Nome *" 20 15))
$txtNome = New-FormInput 20 35 310
$txtNome.Add_TextChanged({ Update-Username })
$pnlForm.Controls.Add($txtNome)

$pnlForm.Controls.Add((New-FormLabel "Sobrenome *" 350 15))
$txtSobrenome = New-FormInput 350 35 325
$txtSobrenome.Add_TextChanged({ Update-Username })
$pnlForm.Controls.Add($txtSobrenome)

$pnlForm.Controls.Add((New-FormLabel "Usuario Template *" 20 72))
$txtTemplate = New-FormInput 20 92 310
$pnlForm.Controls.Add($txtTemplate)

$pnlForm.Controls.Add((New-FormLabel "Senha Inicial" 350 72))
$txtSenha = New-FormInput 350 92 325
$txtSenha.Text = $SenhaInicial
$txtSenha.ReadOnly = $true
$txtSenha.ForeColor = $corTextoCinza
$pnlForm.Controls.Add($txtSenha)

$pnlForm.Controls.Add((New-FormLabel "Licenca Microsoft 365" 20 125))
$cmbLicenca = New-Object System.Windows.Forms.ComboBox
$cmbLicenca.Font = $fonteInput
$cmbLicenca.Size = New-Object System.Drawing.Size(310, 28)
$cmbLicenca.Location = New-Object System.Drawing.Point(20, 145)
$cmbLicenca.BackColor = $corInput
$cmbLicenca.ForeColor = $corTexto
$cmbLicenca.FlatStyle = "Flat"
$cmbLicenca.DropDownStyle = "DropDownList"
$cmbLicenca.Items.AddRange(@(
    "Automatico (recomendado)",
    "Business Standard (SPB)",
    "Apps Enterprise + E1",
    "Nao atribuir licenca"
))
$cmbLicenca.SelectedIndex = 0
$pnlForm.Controls.Add($cmbLicenca)

$lblPreview = New-FormLabel "DADOS GERADOS" 350 125 $true
$lblPreview.ForeColor = $corAzul
$pnlForm.Controls.Add($lblPreview)

$pnlForm.Controls.Add((New-FormLabel "Login:" 350 150))
$txtUsername = New-FormInput 405 147 270
$txtUsername.ReadOnly = $true
$txtUsername.ForeColor = $corVerde
$pnlForm.Controls.Add($txtUsername)

$pnlForm.Controls.Add((New-FormLabel "Email:" 350 185))
$txtEmail = New-FormInput 405 182 270
$txtEmail.ReadOnly = $true
$txtEmail.ForeColor = $corVerde
$pnlForm.Controls.Add($txtEmail)

# Botoes
$btnCriar = New-Object System.Windows.Forms.Button
$btnCriar.Text = "CRIAR USUARIO"
$btnCriar.Font = $fonteBotao
$btnCriar.Size = New-Object System.Drawing.Size(345, 38)
$btnCriar.Location = New-Object System.Drawing.Point(15, 338)
$btnCriar.FlatStyle = "Flat"
$btnCriar.FlatAppearance.BorderSize = 0
$btnCriar.BackColor = $corAzul
$btnCriar.ForeColor = [System.Drawing.Color]::White
$btnCriar.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCriar.Add_MouseEnter({ $btnCriar.BackColor = $corAzulHover })
$btnCriar.Add_MouseLeave({ $btnCriar.BackColor = $corAzul })
$form.Controls.Add($btnCriar)

$btnLimpar = New-Object System.Windows.Forms.Button
$btnLimpar.Text = "LIMPAR CAMPOS"
$btnLimpar.Font = $fonteBotao
$btnLimpar.Size = New-Object System.Drawing.Size(345, 38)
$btnLimpar.Location = New-Object System.Drawing.Point(380, 338)
$btnLimpar.FlatStyle = "Flat"
$btnLimpar.FlatAppearance.BorderSize = 0
$btnLimpar.BackColor = $corPainelClaro
$btnLimpar.ForeColor = $corTexto
$btnLimpar.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLimpar.Add_MouseEnter({ $btnLimpar.BackColor = $corBorda })
$btnLimpar.Add_MouseLeave({ $btnLimpar.BackColor = $corPainelClaro })
$form.Controls.Add($btnLimpar)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(710, 6)
$progressBar.Location = New-Object System.Drawing.Point(15, 382)
$progressBar.Style = "Continuous"
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$form.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Pronto para criar usuario."
$lblStatus.Font = $fonteInfo
$lblStatus.ForeColor = $corTextoCinza
$lblStatus.Size = New-Object System.Drawing.Size(710, 20)
$lblStatus.Location = New-Object System.Drawing.Point(17, 392)
$form.Controls.Add($lblStatus)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Size = New-Object System.Drawing.Size(710, 195)
$txtLog.Location = New-Object System.Drawing.Point(15, 415)
$txtLog.Font = $fonteLog
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 30)
$txtLog.ForeColor = $corTextoCinza
$txtLog.ReadOnly = $true
$txtLog.BorderStyle = "None"
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

# ============================================================
# EVENTOS
# ============================================================
$btnLimpar.Add_Click({
    $txtNome.Text = ""
    $txtSobrenome.Text = ""
    $txtTemplate.Text = ""
    $txtUsername.Text = ""
    $txtEmail.Text = ""
    $cmbLicenca.SelectedIndex = 0
    $txtLog.Clear()
    $progressBar.Value = 0
    $lblStatus.Text = "Pronto para criar usuario."
    $lblStatus.ForeColor = $corTextoCinza
    $txtNome.Focus()
})

$btnCriar.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtNome.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Preencha o PRIMEIRO NOME.", "Campo obrigatorio", "OK", "Warning")
        $txtNome.Focus(); return
    }
    if ([string]::IsNullOrWhiteSpace($txtSobrenome.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Preencha o SOBRENOME.", "Campo obrigatorio", "OK", "Warning")
        $txtSobrenome.Focus(); return
    }
    if ([string]::IsNullOrWhiteSpace($txtTemplate.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Preencha o USUARIO TEMPLATE.", "Campo obrigatorio", "OK", "Warning")
        $txtTemplate.Focus(); return
    }
    if ([string]::IsNullOrWhiteSpace($txtUsername.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Username nao gerado. Verifique nome e sobrenome.", "Erro", "OK", "Error")
        return
    }

    $textInfo     = (Get-Culture).TextInfo
    $PrimeiroNome = $textInfo.ToTitleCase($txtNome.Text.ToLower().Trim())
    $Sobrenome    = $textInfo.ToTitleCase($txtSobrenome.Text.ToLower().Trim())
    $Username     = $txtUsername.Text.Trim()
    $Template     = $txtTemplate.Text.Trim()
    $Email        = $txtEmail.Text.Trim()
    $licencaEscolhida = $cmbLicenca.SelectedItem.ToString()

    $confirma = [System.Windows.Forms.MessageBox]::Show(
        "Confirma a criacao do usuario?`n`nNome: $PrimeiroNome $Sobrenome`nLogin: $Username`nEmail: $Email`nTemplate: $Template`nLicenca: $licencaEscolhida`nSenha: $SenhaInicial",
        "Confirmacao", "YesNo", "Question"
    )
    if ($confirma -ne "Yes") { return }

    $btnCriar.Enabled = $false; $btnLimpar.Enabled = $false
    $txtNome.Enabled = $false; $txtSobrenome.Enabled = $false
    $txtTemplate.Enabled = $false; $cmbLicenca.Enabled = $false
    $btnCriar.Text = "PROCESSANDO..."
    $txtLog.Clear()
    $progressBar.Value = 0

    $Senha = ConvertTo-SecureString $SenhaInicial -AsPlainText -Force

    function Restore-Controls {
        $btnCriar.Enabled = $true; $btnLimpar.Enabled = $true
        $txtNome.Enabled = $true; $txtSobrenome.Enabled = $true
        $txtTemplate.Enabled = $true; $cmbLicenca.Enabled = $true
        $btnCriar.Text = "CRIAR USUARIO"
    }

    # ETAPA 1
    $lblStatus.Text = "Etapa 1/6: Verificando pre-requisitos..."
    $lblStatus.ForeColor = $corAzul
    Add-LogMessage "ETAPA 1: Verificando pre-requisitos" "ETAPA"
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Add-LogMessage "Modulo ActiveDirectory carregado." "OK"
    } catch {
        Add-LogMessage "Modulo ActiveDirectory nao encontrado!" "ERRO"
        $lblStatus.Text = "ERRO: Modulo ActiveDirectory nao disponivel."
        $lblStatus.ForeColor = $corVermelho
        Restore-Controls; return
    }
    try {
        $existente = Get-ADUser -Identity $Username -ErrorAction Stop
        Add-LogMessage "Usuario '$Username' ja existe no AD!" "ERRO"
        $lblStatus.Text = "ERRO: Usuario ja existe."
        $lblStatus.ForeColor = $corVermelho
        Restore-Controls; return
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Add-LogMessage "Username '$Username' disponivel." "OK"
    } catch {
        Add-LogMessage "Aviso ao verificar: $_" "AVISO"
    }
    $progressBar.Value = 10

    # ETAPA 2
    $lblStatus.Text = "Etapa 2/6: Buscando usuario template..."
    Add-LogMessage "ETAPA 2: Buscando template" "ETAPA"
    try {
        $templateUser = Get-ADUser -Identity $Template -Properties MemberOf, Department, Title, Company, Office, Description, Manager
    } catch {
        Add-LogMessage "Template '$Template' nao encontrado!" "ERRO"
        $lblStatus.Text = "ERRO: Template nao encontrado."
        $lblStatus.ForeColor = $corVermelho
        Restore-Controls; return
    }
    $OU          = $templateUser.DistinguishedName -replace '^CN=[^,]+,', ''
    $Setor       = $templateUser.Department
    $NomeExibido = "$PrimeiroNome $Sobrenome - $Setor"
    Add-LogMessage "Template: $Template | Setor: $Setor" "OK"
    Add-LogMessage "OU: $OU" "INFO"
    $progressBar.Value = 20

    # ETAPA 3
    $lblStatus.Text = "Etapa 3/6: Criando usuario no Active Directory..."
    Add-LogMessage "ETAPA 3: Criando usuario no AD" "ETAPA"
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
            Add-LogMessage "Gestor copiado do template." "OK"
        }
        Add-LogMessage "Usuario '$Username' criado no AD!" "OK"
    } catch {
        Add-LogMessage "Falha ao criar usuario: $_" "ERRO"
        $lblStatus.Text = "ERRO ao criar usuario no AD."
        $lblStatus.ForeColor = $corVermelho
        Restore-Controls; return
    }
    $progressBar.Value = 40

    # ETAPA 4
    $lblStatus.Text = "Etapa 4/6: Copiando grupos do template..."
    Add-LogMessage "ETAPA 4: Copiando grupos" "ETAPA"
    $gruposOk = 0
    foreach ($grupo in $templateUser.MemberOf) {
        try {
            Add-ADGroupMember -Identity $grupo -Members $Username -ErrorAction Stop
            $nomeGrupo = ($grupo -split ',')[0] -replace 'CN=', ''
            Add-LogMessage "Grupo: $nomeGrupo" "OK"
            $gruposOk++
        } catch {
            $nomeGrupo = ($grupo -split ',')[0] -replace 'CN=', ''
            Add-LogMessage "Falha: $nomeGrupo" "AVISO"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    Add-LogMessage "Total: $gruposOk grupo(s) copiado(s)" "INFO"
    $progressBar.Value = 55

    # ETAPA 5
    $lblStatus.Text = "Etapa 5/6: Sincronizando com Microsoft 365..."
    Add-LogMessage "ETAPA 5: Sincronizacao Azure AD" "ETAPA"
    $syncOk = $false
    try {
        Import-Module ADSync -ErrorAction Stop
        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
        Add-LogMessage "Sincronizacao Delta iniciada." "OK"
        $syncOk = $true
    } catch {
        Add-LogMessage "ADSync nao disponivel nesta maquina." "AVISO"
    }
    $espera = if ($syncOk) { 90 } else { 120 }
    Add-LogMessage "Aguardando $espera segundos para propagacao..." "INFO"
    for ($s = $espera; $s -gt 0; $s -= 5) {
        $lblStatus.Text = "Etapa 5/6: Aguardando sincronizacao... $s seg restantes"
        $progressBar.Value = [math]::Min(70, 55 + ((($espera - $s) / $espera) * 15))
        DoEvents-Sleep -Seconds 5
    }
    $progressBar.Value = 70

    # ETAPA 6
    $nomeLicenca = ""
    if ($licencaEscolhida -eq "Nao atribuir licenca") {
        Add-LogMessage "ETAPA 6: Licenca ignorada (opcao do usuario)" "ETAPA"
        Add-LogMessage "Licenca nao atribuida por escolha do usuario." "AVISO"
        $progressBar.Value = 100
    } else {
        $lblStatus.Text = "Etapa 6/6: Atribuindo licenca Microsoft 365..."
        Add-LogMessage "ETAPA 6: Licenca Microsoft 365" "ETAPA"
        try {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureSecret)
            $plainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

            $body = @{
                grant_type    = "client_credentials"
                scope         = "https://graph.microsoft.com/.default"
                client_id     = $ClientId
                client_secret = $plainSecret
            }
            $token = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body).access_token
            $headers = @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }
            Add-LogMessage "Token Graph API obtido." "OK"
            $progressBar.Value = 75

            $usuario365 = $null
            for ($i = 1; $i -le 10; $i++) {
                $lblStatus.Text = "Etapa 6/6: Procurando usuario no M365... tentativa $i/10"
                Add-LogMessage "Procurando no M365... tentativa $i/10" "INFO"
                try {
                    $usuario365 = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$Email" -Headers $headers -ErrorAction Stop
                    break
                } catch {
                    if ($i -lt 10) { DoEvents-Sleep -Seconds 30 }
                }
                $progressBar.Value = [math]::Min(90, 75 + ($i * 1.5))
            }

            if (-not $usuario365) {
                Add-LogMessage "Usuario nao encontrado no M365 apos 5 min." "AVISO"
                Add-LogMessage "Atribua manualmente em admin.microsoft.com" "AVISO"
            } else {
                Add-LogMessage "Usuario encontrado no M365!" "OK"
                $bodyLoc = @{ usageLocation = $UsageLocation } | ConvertTo-Json
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$Email" -Method Patch -Headers $headers -Body $bodyLoc | Out-Null
                Add-LogMessage "UsageLocation: $UsageLocation" "OK"
                $progressBar.Value = 92

                $licencas = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers).value
                $skuIds = @()

                if ($licencaEscolhida -eq "Business Standard (SPB)") {
                    $lic = $licencas | Where-Object { $_.skuPartNumber -eq "SPB" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    if ($lic) { $skuIds = @($lic.skuId); $nomeLicenca = "Business Standard" }
                    else { Add-LogMessage "Licenca SPB nao disponivel!" "ERRO" }
                } elseif ($licencaEscolhida -eq "Apps Enterprise + E1") {
                    $apps = $licencas | Where-Object { $_.skuPartNumber -eq "OFFICESUBSCRIPTION" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    $e1   = $licencas | Where-Object { $_.skuPartNumber -eq "STANDARDPACK" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    if ($apps -and $e1) { $skuIds = @($apps.skuId, $e1.skuId); $nomeLicenca = "Apps Enterprise + E1" }
                    else { Add-LogMessage "Licencas Apps+E1 nao disponiveis!" "ERRO" }
                } else {
                    $standard = $licencas | Where-Object { $_.skuPartNumber -eq "SPB" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    $appsEnt  = $licencas | Where-Object { $_.skuPartNumber -eq "OFFICESUBSCRIPTION" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    $e1       = $licencas | Where-Object { $_.skuPartNumber -eq "STANDARDPACK" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    if ($standard) {
                        $skuIds = @($standard.skuId); $nomeLicenca = "Business Standard"
                    } elseif ($appsEnt -and $e1) {
                        $skuIds = @($appsEnt.skuId, $e1.skuId); $nomeLicenca = "Apps Enterprise + E1"
                    } else {
                        Add-LogMessage "Nenhuma licenca disponivel!" "AVISO"
                    }
                }

                if ($skuIds.Count -gt 0) {
                    foreach ($id in $skuIds) {
                        $bodyLic = @{ addLicenses = @(@{ skuId = $id }); removeLicenses = @() } | ConvertTo-Json -Depth 5
                        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$Email/assignLicense" -Method Post -Headers $headers -Body $bodyLic | Out-Null
                    }
                    Add-LogMessage "Licenca atribuida: $nomeLicenca" "OK"
                }
            }
        } catch {
            Add-LogMessage "Erro na licenca M365: $_" "AVISO"
            Add-LogMessage "Atribua manualmente em admin.microsoft.com" "AVISO"
        }
        $progressBar.Value = 100
    }

    # RESUMO
    Add-LogMessage "========================================" "ETAPA"
    Add-LogMessage "  USUARIO CRIADO COM SUCESSO!" "OK"
    Add-LogMessage "  Nome     : $NomeExibido" "OK"
    Add-LogMessage "  Login    : $Username" "OK"
    Add-LogMessage "  Email    : $Email" "OK"
    Add-LogMessage "  Setor    : $Setor" "OK"
    Add-LogMessage "  Grupos   : $gruposOk" "OK"
    Add-LogMessage "  Senha    : $SenhaInicial" "OK"
    if ($nomeLicenca) { Add-LogMessage "  Licenca  : $nomeLicenca" "OK" }
    Add-LogMessage "========================================" "ETAPA"

    Save-HistoricoCsv -Nome $NomeExibido -Username $Username -Email $Email `
        -Setor $Setor -Template $Template -Licenca $(if ($nomeLicenca) { $nomeLicenca } else { "N/A" }) `
        -Grupos $gruposOk

    $lblStatus.Text = "CONCLUIDO! Usuario '$Username' criado com sucesso."
    $lblStatus.ForeColor = $corVerde

    [System.Windows.Forms.MessageBox]::Show(
        "Usuario criado com sucesso!`n`nNome: $NomeExibido`nLogin: $Username`nEmail: $Email`nSenha: $SenhaInicial$(if($nomeLicenca){"`nLicenca: $nomeLicenca"})",
        "Sucesso", "OK", "Information"
    )

    Restore-Controls
})

# ============================================================
$txtLog.SelectionColor = $corTextoCinza
$txtLog.AppendText("Pronto. Preencha os campos e clique em CRIAR USUARIO.`r`n")
$txtNome.Focus()
[void]$form.ShowDialog()
$form.Dispose()
