Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

# Instala ImportExcel caso não esteja presente
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Módulo 'ImportExcel' não encontrado. Tentando instalar..."
    try {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
        Import-Module ImportExcel -ErrorAction Stop
        Write-Host "Módulo 'ImportExcel' instalado com sucesso."
    } catch {
        Write-Error "Erro ao instalar o módulo ImportExcel. Verifique sua conexão ou permissões."
        exit
    }
} else {
    Import-Module ImportExcel
}

$logPath = "$env:TEMP\\log_ad_grupos.txt"
if (-not (Test-Path "$env:TEMP")) { New-Item -ItemType Directory -Path "$env:TEMP" | Out-Null }

function Write-Log {
    param([string]$mensagem)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $mensagem" | Out-File -FilePath $logPath -Append
}

function Abrir-DialogoArquivo($titulo, $filtro) {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $titulo
    $dialog.Filter = $filtro
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq "OK") {
        return $dialog.FileName
    }
    return $null
}

function Show-InputBox {
    param(
        [string]$Message = "Digite um valor:",
        [string]$Title = "Input",
        [string]$Default = ""
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400,150)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,10)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Default
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(360,20)
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(200,80)
    $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.Location = New-Object System.Drawing.Point(290,80)
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Prometheus_Beta4.0.2"
$form.Size = New-Object System.Drawing.Size(600, 700)
$form.StartPosition = "CenterScreen"

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Size = New-Object System.Drawing.Size(780, 640)
$tabs.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($tabs)

# =================== ABA 1: Adicionar a Grupos ===================

$tab1 = New-Object System.Windows.Forms.TabPage
$tab1.Text = "Adicionar a Grupos"

$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Text = "Usuário (Login ou Email):"
$userLabel.Location = New-Object System.Drawing.Point(10,20)
$tab1.Controls.Add($userLabel)

$userBox = New-Object System.Windows.Forms.TextBox
$userBox.Location = New-Object System.Drawing.Point(200,18)
$userBox.Size = New-Object System.Drawing.Size(420,20)
$tab1.Controls.Add($userBox)

$groupFilterLabel = New-Object System.Windows.Forms.Label
$groupFilterLabel.Text = "Filtrar grupos:"
$groupFilterLabel.Location = New-Object System.Drawing.Point(10,60)
$tab1.Controls.Add($groupFilterLabel)

$groupFilterBox = New-Object System.Windows.Forms.TextBox
$groupFilterBox.Location = New-Object System.Drawing.Point(200,58)
$groupFilterBox.Size = New-Object System.Drawing.Size(420,20)
$tab1.Controls.Add($groupFilterBox)

$groupList = New-Object System.Windows.Forms.CheckedListBox
$groupList.Location = New-Object System.Drawing.Point(10,90)
$groupList.Size = New-Object System.Drawing.Size(640,170)
$groupList.CheckOnClick = $true

$allGroups = Get-ADGroup -Filter * | Select-Object -ExpandProperty Name | Sort-Object
$allGroups | ForEach-Object { [void]$groupList.Items.Add($_) }
$tab1.Controls.Add($groupList)

$groupFilterBox.Add_TextChanged({
    $groupList.Items.Clear()
    $allGroups | Where-Object { $_ -like "*$($groupFilterBox.Text)*" } | ForEach-Object {
        [void]$groupList.Items.Add($_)
    }
})

$presetPath = "$env:TEMP\\group_presets.json"
function Ler-PresetsGrupos {
    if (Test-Path $presetPath) {
        try {
            $content = Get-Content $presetPath -Raw
            if ($content) {
                $hash = ConvertFrom-Json $content
                $table = @{}
                foreach ($key in $hash.PSObject.Properties.Name) {
                    $table[$key] = @($hash.$key)
                }
                return $table
            }
        } catch { }
    }
    return @{}
}
function Salvar-PresetsGrupos($table) {
    $table | ConvertTo-Json | Out-File -Encoding UTF8 $presetPath
}
$presetes = Ler-PresetsGrupos

$presetComboBox = New-Object System.Windows.Forms.ComboBox
$presetComboBox.Location = New-Object System.Drawing.Point(10, 270)
$presetComboBox.Size = New-Object System.Drawing.Size(220, 20)
$presetComboBox.DropDownStyle = 'DropDownList'
$presetComboBox.Items.Add("<Selecionar preset>") | Out-Null
foreach ($preset in $presetes.Keys) {
    $presetComboBox.Items.Add($preset) | Out-Null
}
$presetComboBox.SelectedIndex = 0
$tab1.Controls.Add($presetComboBox)

$btnSalvarPreset = New-Object System.Windows.Forms.Button
$btnSalvarPreset.Text = "Salvar Preset"
$btnSalvarPreset.Location = New-Object System.Drawing.Point(240, 270)
$btnSalvarPreset.Size = New-Object System.Drawing.Size(110, 25)
$btnSalvarPreset.Add_Click({
    $nomePreset = Show-InputBox -Message "Nome do novo preset:" -Title "Salvar Preset" -Default "Preset1"
    if (-not [string]::IsNullOrWhiteSpace($nomePreset)) {
        $gruposSelecionados = @($groupList.CheckedItems)
        if ($gruposSelecionados.Count -gt 0) {
            $presetes[$nomePreset] = $gruposSelecionados
            Salvar-PresetsGrupos $presetes
            if (-not $presetComboBox.Items.Contains($nomePreset)) {
                $presetComboBox.Items.Add($nomePreset) | Out-Null
            }
            [System.Windows.Forms.MessageBox]::Show("Preset '$nomePreset' salvo com sucesso.", "Sucesso")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Selecione pelo menos um grupo para salvar no preset.", "Aviso")
        }
    }
})
$tab1.Controls.Add($btnSalvarPreset)

$btnDeletarPreset = New-Object System.Windows.Forms.Button
$btnDeletarPreset.Text = "Deletar Preset"
$btnDeletarPreset.Location = New-Object System.Drawing.Point(360, 270)
$btnDeletarPreset.Size = New-Object System.Drawing.Size(110, 25)
$btnDeletarPreset.Add_Click({
    if ($presetComboBox.SelectedIndex -gt 0) {
        $nomePreset = $presetComboBox.SelectedItem
        $presetes.Remove($nomePreset)
        Salvar-PresetsGrupos $presetes
        $presetComboBox.Items.Remove($nomePreset)
        $presetComboBox.SelectedIndex = 0
        [System.Windows.Forms.MessageBox]::Show("Preset removido.", "Removido")
    }
})
$tab1.Controls.Add($btnDeletarPreset)

$presetComboBox.Add_SelectedIndexChanged({
    if ($presetComboBox.SelectedIndex -gt 0) {
        $nome = $presetComboBox.SelectedItem
        $gruposDoPreset = $presetes[$nome]
        for ($i = 0; $i -lt $groupList.Items.Count; $i++) {
            if ($gruposDoPreset -contains $groupList.Items[$i]) {
                $groupList.SetItemChecked($i, $true)
            } else {
                $groupList.SetItemChecked($i, $false)
            }
        }
    } else {
        for ($i = 0; $i -lt $groupList.Items.Count; $i++) {
            $groupList.SetItemChecked($i, $false)
        }
    }
})

$groupOutputBox = New-Object System.Windows.Forms.TextBox
$groupOutputBox.Location = New-Object System.Drawing.Point(10,320)
$groupOutputBox.Size = New-Object System.Drawing.Size(640,130)
$groupOutputBox.Multiline = $true
$groupOutputBox.ScrollBars = "Vertical"
$groupOutputBox.ReadOnly = $true
$tab1.Controls.Add($groupOutputBox)

$btnAddToGroup = New-Object System.Windows.Forms.Button
$btnAddToGroup.Text = "Adicionar Usuário aos Grupos Selecionados"
$btnAddToGroup.Location = New-Object System.Drawing.Point(10,295)
$btnAddToGroup.Size = New-Object System.Drawing.Size(640,25)
$btnAddToGroup.Add_Click({
    $username = $userBox.Text.Trim()
    $selectedGroups = @($groupList.CheckedItems)

    if (-not $username -or $selectedGroups.Count -eq 0) {
        $groupOutputBox.Text = "Preencha o nome do usuário e selecione pelo menos um grupo."
        return
    }

    $user = $null
    foreach ($domain in (Get-ADForest).Domains) {
        try {
            $user = Get-ADUser -Server $domain -Filter { SamAccountName -eq $username -or Mail -eq $username } -Properties DistinguishedName -ErrorAction Stop
            if ($user) { break }
        } catch {}
    }

    if (-not $user) {
        $msg = "Usuário '$username' não encontrado."
        $groupOutputBox.Text = $msg
        Write-Log $msg
        return
    }

    $results = @()
    foreach ($groupName in $selectedGroups) {
        try {
            $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
            Add-ADGroupMember -Identity $group.DistinguishedName -Members $user.DistinguishedName -ErrorAction Stop
            $msg = "✅ Adicionado ao grupo: $groupName"
            $results += $msg
            Write-Log $msg
        } catch {
            $erro = "❌ Falha ao adicionar '$username' ao grupo '$groupName': $($_.Exception.Message)"
            $results += $erro
            Write-Log $erro
        }
    }
    $groupOutputBox.Lines = $results
})
$tab1.Controls.Add($btnAddToGroup)

$tabs.TabPages.Add($tab1)

# =================== ABA 2: Criar Usuários em Lote (com fallback de OU) ===================

$tab2 = New-Object System.Windows.Forms.TabPage
$tab2.Text = "Criar Usuários em Lote"

$labelExcelUsuarios = New-Object System.Windows.Forms.Label
$labelExcelUsuarios.Text = "Planilha de Usuários (.xlsx):"
$labelExcelUsuarios.Location = New-Object System.Drawing.Point(10,20)
$tab2.Controls.Add($labelExcelUsuarios)

$textExcelUsuarios = New-Object System.Windows.Forms.TextBox
$textExcelUsuarios.Location = New-Object System.Drawing.Point(200,18)
$textExcelUsuarios.Size = New-Object System.Drawing.Size(320,20)
$tab2.Controls.Add($textExcelUsuarios)

$btnSelecionarUsuarios = New-Object System.Windows.Forms.Button
$btnSelecionarUsuarios.Text = "Selecionar"
$btnSelecionarUsuarios.Location = New-Object System.Drawing.Point(530, 16)
$btnSelecionarUsuarios.Add_Click({
    $arquivo = Abrir-DialogoArquivo "Selecionar Planilha de Usuários" "Excel (*.xlsx)|*.xlsx"
    if ($arquivo) { $textExcelUsuarios.Text = $arquivo }
})
$tab2.Controls.Add($btnSelecionarUsuarios)

$labelExcelCC = New-Object System.Windows.Forms.Label
$labelExcelCC.Text = "Planilha de Centros de Custo (.xlsx):"
$labelExcelCC.Location = New-Object System.Drawing.Point(10,60)
$tab2.Controls.Add($labelExcelCC)

$textExcelCC = New-Object System.Windows.Forms.TextBox
$textExcelCC.Location = New-Object System.Drawing.Point(200,58)
$textExcelCC.Size = New-Object System.Drawing.Size(320,20)
$tab2.Controls.Add($textExcelCC)

$btnSelecionarCC = New-Object System.Windows.Forms.Button
$btnSelecionarCC.Text = "Selecionar"
$btnSelecionarCC.Location = New-Object System.Drawing.Point(530, 56)
$btnSelecionarCC.Add_Click({
    $arquivo = Abrir-DialogoArquivo "Selecionar Planilha de CC" "Excel (*.xlsx)|*.xlsx"
    if ($arquivo) { $textExcelCC.Text = $arquivo }
})
$tab2.Controls.Add($btnSelecionarCC)

$btnExecutarCriacao = New-Object System.Windows.Forms.Button
$btnExecutarCriacao.Text = "Executar Criação de Usuários"
$btnExecutarCriacao.Size = New-Object System.Drawing.Size(640,30)
$btnExecutarCriacao.Location = New-Object System.Drawing.Point(10,100)
$tab2.Controls.Add($btnExecutarCriacao)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10,140)
$outputBox.Size = New-Object System.Drawing.Size(760,410)
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$tab2.Controls.Add($outputBox)

function Show-InputBox {
    param(
        [string]$Message = "Digite um valor:",
        [string]$Title = "Input",
        [string]$Default = ""
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(450,150)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,10)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Default
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(410,20)
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(250,80)
    $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.Location = New-Object System.Drawing.Point(340,80)
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

$btnExecutarCriacao.Add_Click({
    $usuariosPath = $textExcelUsuarios.Text.Trim()
    $ccPath = $textExcelCC.Text.Trim()

    $outputBox.Clear()
    $outputBox.AppendText("Iniciando criação em lote...`r`n")
    $outputBox.AppendText("Arquivo Usuários: $usuariosPath`r`n")
    $outputBox.AppendText("Arquivo CC: $ccPath`r`n")

    if (-not (Test-Path $usuariosPath) -or -not (Test-Path $ccPath)) {
        $outputBox.AppendText("Erro: Caminho(s) inválido(s). Verifique os arquivos.`r`n")
        return
    }

    try {
        $outputBox.AppendText("Importando arquivos Excel...`r`n")
        $usuarios = Import-Excel -Path $usuariosPath -ErrorAction Stop
        $ccs = Import-Excel -Path $ccPath -ErrorAction Stop
        $outputBox.AppendText("Arquivos importados com sucesso!`r`n")
    } catch {
        $outputBox.AppendText("Erro ao importar arquivos: $($_.Exception.Message)`r`n")
        return
    }

    # Monta mapa CC x OU (base centros de custo)
    $mapCCOU = @{}
    foreach ($row in $ccs) {
        $ccNumber = $row.'Department Number'
        $ouValue = $row.OU
        if ($ccNumber -and $ouValue) {
            $ccKey = $ccNumber.Trim().ToUpper()
            $mapCCOU[$ccKey] = $ouValue.Trim()
        }
    }
    $outputBox.AppendText("Mapa CC_OU pronto!`r`n")

    $linha = 0
    foreach ($usuario in $usuarios) {
        $linha++
        # Só processa linhas preenchidas (nome e sobrenome)
        $nome = $usuario.NOMBRES
        $sobrenome = $usuario.APELLIDOS
        if (-not $nome -or -not $sobrenome) {
            $outputBox.AppendText("Linha ignorada: nome ou sobrenome ausente.`r`n")
            continue
        }

        $nomeDisplay = "$($usuario.NOMBRES) $($usuario.APELLIDOS)"
        $outputBox.AppendText("Processando usuário $linha : $nomeDisplay`r`n")

        $ou = $usuario.OU
        $cc = $usuario.CC
        $sam = $usuario.SAM

        # Busca OU por CC se OU estiver vazio
        if (-not $ou -or $ou -eq "") {
            $ccBusca = $null
            if ($cc) { $ccBusca = $cc.Trim().ToUpper() }
            if ($ccBusca -and $mapCCOU.ContainsKey($ccBusca)) {
                $ou = $mapCCOU[$ccBusca]
                $outputBox.AppendText("OU preenchido por CC: $ou`r`n")
            }
        }

        # Caso ainda não haja OU, pede manualmente
        while (-not $ou -or $ou -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Faltando OU para o usuário $sam. Informe manualmente.","Preenchimento obrigatório")
            $ouManual = Show-InputBox -Message "Informe a OU no AD para o usuário $sam ($nomeDisplay):" -Title "Preencher OU"
            if ($ouManual) {
                $ou = $ouManual
                $outputBox.AppendText("Campo OU preenchido manualmente.`r`n")
            } else {
                $outputBox.AppendText("Usuário '$nomeDisplay' ignorado (sem OU).`r`n")
                $ou = $null
                break
            }
        }
        if (-not $ou) { continue }

        # Monta os parâmetros
        $params = @{
            Name               = $nomeDisplay
            GivenName          = $usuario.NOMBRES
            Surname            = $usuario.APELLIDOS
            DisplayName        = $usuario.DISPLAY
            SamAccountName     = $usuario.SAM
            UserPrincipalName  = $usuario.CORREO
            Title              = $usuario.CARGO
            EmployeeID         = $usuario.RUT
            Company            = $usuario.EMPRESA
            Office             = $usuario.PLANTA
            Department         = $usuario.AREA
            StreetAddress      = $usuario.DIRECCION
            City               = $usuario.CIUDAD
            State              = $usuario.ESTADO
            Country            = $usuario.PAIS
            EmailAddress       = $usuario.CORREO
            AccountPassword    = (ConvertTo-SecureString (Nova-SenhaAleatoria 12) -AsPlainText -Force)
            Enabled            = $true
            ChangePasswordAtLogon = $true
            Path               = $ou
        }
        $params["OtherAttributes"] = @{
            employeeType     = $usuario.TIPO
            departmentNumber = $usuario.CC
        }

        # Se o campo dominio estiver presente na planilha, use. Senão, padrão local
        $domAD = $usuario.DOMAD
        if ($domAD) { $params["Server"] = $domAD }

        # Tenta criar e faz tratamento se a OU não existe no AD
        try {
            New-ADUser @params
            $outputBox.AppendText("Usuário '$nomeDisplay' criado com sucesso!`r`n")
        } catch {
            $erroMsg = $_.Exception.Message
            $outputBox.AppendText("Erro ao criar usuário '$nomeDisplay': $erroMsg`r`n")
            # Se erro de OU não existe ou permissão, oferece OU manual
            if ($erroMsg -like "*unwilling*" -or $erroMsg -like "*does not exist*" -or $erroMsg -like "*The specified directory service attribute or value does not exist*") {
                [System.Windows.Forms.MessageBox]::Show("A OU '$ou' não existe ou não está acessível no AD.`nInforme manualmente a OU para o usuário $nomeDisplay.","OU não encontrada")
                $ouManual = Show-InputBox -Message "Informe uma OU válida no AD para o usuário $sam ($nomeDisplay):" -Title "OU não encontrada"
                if ($ouManual) {
                    $params.Path = $ouManual
                    try {
                        New-ADUser @params
                        $outputBox.AppendText("Usuário '$nomeDisplay' criado com sucesso na OU manual!`r`n")
                    } catch {
                        $outputBox.AppendText("Falha ao criar usuário '$nomeDisplay' na OU informada: $($_.Exception.Message)`r`n")
                    }
                } else {
                    $outputBox.AppendText("Usuário '$nomeDisplay' não criado (OU não fornecida).`r`n")
                }
            }
        }
    }
    $outputBox.AppendText("Processo finalizado!`r`n")
})

$tabs.TabPages.Add($tab2)





# ========== ABA 3: Criar Usuário Unitário (PROMETHEUS, responsiva e botões sempre visíveis) ==========

function Nova-SenhaAleatoria([int]$tamanho = 12) {
    $maiusc = "ABCDEFGHJKLMNPQRSTUVWXYZ".ToCharArray()
    $minusc = "abcdefghjkmnpqrstuvwxyz".ToCharArray()
    $nums   = "23456789".ToCharArray()
    $espec  = "@#\$%!&*".ToCharArray()

    # Garante pelo menos um de cada:
    $senha = @()
    $senha += $maiusc | Get-Random
    $senha += $minusc | Get-Random
    $senha += $nums   | Get-Random
    $senha += $espec  | Get-Random

    $todos = $maiusc + $minusc + $nums + $espec
    for ($i = $senha.Count; $i -lt $tamanho; $i++) {
        $senha += $todos | Get-Random
    }
    # Embaralha:
    -join ($senha | Get-Random -Count $senha.Count)
}

$tab3 = New-Object System.Windows.Forms.TabPage
$tab3.Text = "Criar Usuário Unitário"

# Recomenda-se aumentar o tamanho do formulário principal e dos tabs:
$form.Size = New-Object System.Drawing.Size(1050, 850)
$tabs.Size = New-Object System.Drawing.Size(1020, 790)

# Painel esquerdo (inputs)
$panelFields = New-Object System.Windows.Forms.Panel
$panelFields.Location = New-Object System.Drawing.Point(10,10)
$panelFields.Size = New-Object System.Drawing.Size(600, 700)
$panelFields.Anchor = "Top,Left"
$panelFields.AutoScroll = $true
$tab3.Controls.Add($panelFields)

# Painel direito (output/ação)
$panelOutput = New-Object System.Windows.Forms.Panel
$panelOutput.Location = New-Object System.Drawing.Point(620, 10)
$panelOutput.Size = New-Object System.Drawing.Size(380, 760)
$panelOutput.Anchor = "Top,Right,Bottom"
$tab3.Controls.Add($panelOutput)

# Domínio e OU
$labDominio = New-Object System.Windows.Forms.Label
$labDominio.Text = "Domínio:"
$labDominio.Location = New-Object System.Drawing.Point(10,13)
$labDominio.ForeColor = 'Red'
$panelFields.Controls.Add($labDominio)

$comboDominios = New-Object System.Windows.Forms.ComboBox
$comboDominios.Location = New-Object System.Drawing.Point(130, 10)
$comboDominios.Size = New-Object System.Drawing.Size(320, 22)
$comboDominios.DropDownStyle = 'DropDownList'
(Get-ADForest).Domains | ForEach-Object { $comboDominios.Items.Add($_) | Out-Null }
$comboDominios.SelectedIndex = 0
$comboDominios.BackColor = 'Yellow'
$panelFields.Controls.Add($comboDominios)

$labOU = New-Object System.Windows.Forms.Label
$labOU.Text = "OU Destino:"
$labOU.Location = New-Object System.Drawing.Point(10,41)
$labOU.ForeColor = 'Red'
$panelFields.Controls.Add($labOU)

$comboOUs = New-Object System.Windows.Forms.ComboBox
$comboOUs.Location = New-Object System.Drawing.Point(130, 38)
$comboOUs.Size = New-Object System.Drawing.Size(320,22)
$comboOUs.DropDownStyle = 'DropDownList'
$comboOUs.BackColor = 'Yellow'
$panelFields.Controls.Add($comboOUs)

function AtualizaOUs($dom) {
    $comboOUs.Items.Clear()
    Get-ADOrganizationalUnit -Server $dom -Filter * | Sort-Object Name | ForEach-Object {
        $comboOUs.Items.Add($_.DistinguishedName) | Out-Null
    }
    if ($comboOUs.Items.Count -gt 0) { $comboOUs.SelectedIndex = 0 }
}
$comboDominios.Add_SelectedIndexChanged({ AtualizaOUs $comboDominios.SelectedItem })
AtualizaOUs $comboDominios.SelectedItem

# Campos do formulário
$fields = @(
    @{label="RUT"; var="txtRut"},
    @{label="Nombres"; var="txtNome"},
    @{label="Apellidos"; var="txtSobrenome"},
    @{label="Tipo Usuário"; var="txtTipo"},
    @{label="Centro de Custo"; var="txtCC"},
    @{label="Empresa"; var="txtEmpresa"},
    @{label="Gerência/Planta"; var="txtGerencia"},
    @{label="Sub-Gerência/Área"; var="txtSubger"},
    @{label="Direção"; var="txtEndereco"},
    @{label="Cidade"; var="txtCidade"},
    @{label="Estado/Província"; var="txtEstado"},
    @{label="País/Região"; var="txtPais"},
    @{label="Supervisor (login)"; var="txtSupervisor"},
    @{label="Cargo ou Função"; var="txtCargo"},
    @{label="Login"; var="txtLogin"},
    @{label="E-mail"; var="txtEmail"}
)
$startY = 70
$stepY = 28
$labelVars = @{}
foreach ($i in 0..($fields.Count-1)) {
    $f = $fields[$i]
    $y = $startY + ($i * $stepY)
    $lab = New-Object System.Windows.Forms.Label
    $lab.Text = $f.label
    $lab.Location = New-Object System.Drawing.Point(10, ($y + 3))
    $lab.AutoSize = $true
    if ($f.label -in @("Tipo Usuário","Login","E-mail","País/Região","DisplayName")) {
        $lab.ForeColor = 'Red'
    }
    $labelVars[$f.var] = $lab
    $panelFields.Controls.Add($lab)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point(130, $y)
    $box.Size = New-Object System.Drawing.Size(320, 22)
    Set-Variable -Name $f.var -Value $box -Scope Global

    if ($f.label -in @("Tipo Usuário","Login","E-mail","País/Região","DisplayName")) {
        $box.BackColor = 'Yellow'
    }
    $panelFields.Controls.Add($box)
}

# Campo DisplayName (obrigatório)
$labDisplayName = New-Object System.Windows.Forms.Label
$labDisplayName.Text = "DisplayName:"
$labDisplayName.Location = New-Object System.Drawing.Point(10, ($startY + $fields.Count * $stepY + 3))
$labDisplayName.ForeColor = 'Red'
$labDisplayName.AutoSize = $true
$panelFields.Controls.Add($labDisplayName)

$txtDisplayName = New-Object System.Windows.Forms.TextBox
$txtDisplayName.Location = New-Object System.Drawing.Point(130, ($startY + $fields.Count * $stepY))
$txtDisplayName.Size = New-Object System.Drawing.Size(320, 22)
$txtDisplayName.BackColor = 'Yellow'
$panelFields.Controls.Add($txtDisplayName)

# Campo Descrição (opcional)
$labDescricao = New-Object System.Windows.Forms.Label
$labDescricao.Text = "Descrição:"
$labDescricao.Location = New-Object System.Drawing.Point(10, ($startY + $fields.Count * $stepY + 32))
$labDescricao.AutoSize = $true
$panelFields.Controls.Add($labDescricao)

$txtDescricao = New-Object System.Windows.Forms.TextBox
$txtDescricao.Location = New-Object System.Drawing.Point(130, ($startY + $fields.Count * $stepY + 29))
$txtDescricao.Size = New-Object System.Drawing.Size(320, 22)
$panelFields.Controls.Add($txtDescricao)

# Checkbox google-no-sync
$chkGoogle = New-Object System.Windows.Forms.CheckBox
$chkGoogle.Text = "google-no-sync?"
$chkGoogle.Location = New-Object System.Drawing.Point(130, ($startY + $fields.Count * $stepY + 60))
$panelFields.Controls.Add($chkGoogle)

# ========== Botões logo após o painel de campos, fora do painel ==========
$totalCampos = $fields.Count + 2 # displayname + descricao
$btnY = $panelFields.Location.Y + $panelFields.Size.Height + 10  # Posição logo após o painel

$btnAutoPreencher = New-Object System.Windows.Forms.Button
$btnAutoPreencher.Text = "Auto preencher"
$btnAutoPreencher.Size = New-Object System.Drawing.Size(140,28)
$btnAutoPreencher.Location = New-Object System.Drawing.Point(10, ($panelFields.Location.Y + $panelFields.Size.Height + 10))
$tab3.Controls.Add($btnAutoPreencher)

$btnCriarUni = New-Object System.Windows.Forms.Button
$btnCriarUni.Text = "Criar Usuário"
$btnCriarUni.Size = New-Object System.Drawing.Size(140,28)
$btnCriarUni.Location      = New-Object System.Drawing.Point(180, ($panelFields.Location.Y + $panelFields.Size.Height + 10))
$tab3.Controls.Add($btnCriarUni)

# Output copiável (Painel direito)
$outResultado = New-Object System.Windows.Forms.TextBox
$outResultado.Location = New-Object System.Drawing.Point(0,0)
$outResultado.Size = New-Object System.Drawing.Size(350,70)
$outResultado.Multiline = $true
$outResultado.ScrollBars = "Vertical"
$outResultado.ReadOnly = $false
$outResultado.Anchor = "Top,Left,Right"
$panelOutput.Controls.Add($outResultado)

$outUni = New-Object System.Windows.Forms.TextBox
$outUni.Location = New-Object System.Drawing.Point(0,80)
$outUni.Size = New-Object System.Drawing.Size(350,600)
$outUni.Multiline = $true
$outUni.ScrollBars = "Vertical"
$outUni.ReadOnly = $true
$outUni.Anchor = "Top,Bottom,Left,Right"
$panelOutput.Controls.Add($outUni)


# --------- Botão Auto Preencher (DOCX e CSV) ---------
$btnAutoPreencher.Add_Click({
    $file = Abrir-DialogoArquivo "Selecione o formulário" "Arquivos (*.docx;*.csv)|*.docx;*.csv"
    if (-not $file) { return }
    try {
        $ext = [IO.Path]::GetExtension($file).ToLower()
        if ($ext -eq ".docx") {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $tmp = [System.IO.Path]::GetTempFileName()
            Copy-Item $file $tmp -Force
            $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
            $entry = $zip.Entries | Where-Object { $_.FullName -eq "word/document.xml" }
            if (-not $entry) { throw "Arquivo docx inválido." }
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $xmlText = $reader.ReadToEnd()
            $reader.Close()
            $zip.Dispose()
            Remove-Item $tmp -Force
            $xmlText = $xmlText -replace "</w:p>", "`n"
            $xmlText = $xmlText -replace "<[^>]+>", ""
            if ($xmlText -match "Rut\s*:?[\s\r\n]+(.+)") { $txtRut.Text = $matches[1].Trim() }
            if ($xmlText -match "Nombres\s*:?[\s\r\n]+(.+)") { $txtNome.Text = $matches[1].Trim() }
            if ($xmlText -match "Apellidos\s*:?[\s\r\n]+(.+)") { $txtSobrenome.Text = $matches[1].Trim() }
            if ($xmlText -match "Tipo Usuario\s*:?[\s\r\n]+(.+)") { $txtTipo.Text = $matches[1].Trim() }
            if ($xmlText -match "Centro de Costo\s*:?[\s\r\n]+(.+)") { $txtCC.Text = $matches[1].Trim() }
            if ($xmlText -match "Empresa\s*:?[\s\r\n]+(.+)") { $txtEmpresa.Text = $matches[1].Trim() }
            if ($xmlText -match "Gerencia o Planta\s*:?[\s\r\n]+(.+)") { $txtGerencia.Text = $matches[1].Trim() }
            if ($xmlText -match "Sub-Gerencia o Área\s*:?[\s\r\n]+(.+)") { $txtSubger.Text = $matches[1].Trim() }
            if ($xmlText -match "Dirección\s*:?[\s\r\n]+(.+)") { $txtEndereco.Text = $matches[1].Trim() }
            if ($xmlText -match "Ciudad\s*:?[\s\r\n]+(.+)") { $txtCidade.Text = $matches[1].Trim() }
            if ($xmlText -match "Estado/Provincia\s*:?[\s\r\n]+(.+)") { $txtEstado.Text = $matches[1].Trim() }
            if ($xmlText -match "País/Región\s*:?[\s\r\n]+(.+)") { $txtPais.Text = "CL" }
            if ($xmlText -match "Supervisor\s*:?[\s\r\n]+(.+)") { $txtSupervisor.Text = $matches[1].Trim() }
            if ($xmlText -match "Cargo o Función\s*:?[\s\r\n]+(.+)") { $txtCargo.Text = $matches[1].Trim() }
            if ($xmlText -match "Login\s*:?[\s\r\n]+(.+)") { $txtLogin.Text = $matches[1].Trim() }
            if ($xmlText -match "Correo\s*:?[\s\r\n]+(.+)") { $txtEmail.Text = "" }
            $txtDisplayName.Text = "$($txtNome.Text) $($txtSobrenome.Text) (CMPC)"
            $txtDescricao.Text = ""
            [System.Windows.Forms.MessageBox]::Show("Auto preenchimento via Word concluído.","Sucesso")
        }
        elseif ($ext -eq ".csv") {
            $linhas = Get-Content $file
            $sep = ";"

            $colRUT = $linhas[8] -split $sep
            $txtRut.Text = "$($colRUT[7])-$($colRUT[13])"

            $colApPat = $linhas[9] -split $sep
            $colApMat = $linhas[10] -split $sep
            $txtSobrenome.Text = "$($colApPat[7]) $($colApMat[7])".Trim()

            $colNome = $linhas[11] -split $sep
            $txtNome.Text = $colNome[7].Trim()

            $txtTipo.Text = "" # Defina lógica adicional se houver no CSV

            $colCC = $linhas[17] -split $sep
            $txtCC.Text = $colCC[20].Trim()

            $txtEmpresa.Text = "CMPC"

            $colGerencia = $linhas[19] -split $sep
            $txtGerencia.Text = $colGerencia[7].Trim()

            $colArea = $linhas[20] -split $sep
            $txtSubger.Text = $colArea[7].Trim()

            $colEndereco = $linhas[18] -split $sep
            $txtEndereco.Text = $colEndereco[7].Trim()

            $colCidade = $linhas[17] -split $sep
            $txtCidade.Text = $colCidade[7].Trim()
            $txtPais.Text = "CL"

            $txtEstado.Text = ""

            $colSupervisor = $linhas[16] -split $sep
            $txtSupervisor.Text = $colSupervisor[20].Trim()
            $txtCargo.Text = $colSupervisor[7].Trim()

            $txtLogin.Text = ""
            $txtEmail.Text = ""

            $txtDisplayName.Text = "$($txtNome.Text) $($txtSobrenome.Text) (CMPC)"
            $txtDescricao.Text = ""
            [System.Windows.Forms.MessageBox]::Show("Auto preenchimento via CSV concluído.","Sucesso")
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Formato de arquivo não suportado. Use apenas DOCX ou CSV.","Atenção")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro no auto preenchimento: $($_.Exception.Message)", "Erro")
    }
})

# --------- Botão Criar Usuário ---------
$btnCriarUni.Add_Click({
    $dominio = $comboDominios.SelectedItem
    $ou = $comboOUs.SelectedItem
    $webpage = if ($chkGoogle.Checked) { "google-no-sync" } else { "" }
    $faltando = @()
    if (-not $txtRut.Text) { $faltando += "RUT" }
    if (-not $txtNome.Text) { $faltando += "Nome" }
    if (-not $txtSobrenome.Text) { $faltando += "Sobrenome" }
    if (-not $txtTipo.Text) { $faltando += "Tipo Usuário" }
    if (-not $txtCC.Text) { $faltando += "Centro de Custo" }
    if (-not $txtEmpresa.Text) { $faltando += "Empresa" }
    if (-not $txtGerencia.Text) { $faltando += "Gerência/Planta" }
    if (-not $txtSubger.Text) { $faltando += "Sub-Gerência/Área" }
    if (-not $txtEndereco.Text) { $faltando += "Direção" }
    if (-not $txtCidade.Text) { $faltando += "Cidade" }
    if (-not $txtPais.Text) { $faltando += "País/Região" }
    if (-not $txtSupervisor.Text) { $faltando += "Supervisor" }
    if (-not $txtCargo.Text) { $faltando += "Cargo ou Função" }
    if (-not $txtLogin.Text) { $faltando += "Login" }
    if (-not $txtEmail.Text) { $faltando += "E-mail" }
    if (-not $txtDisplayName.Text) { $faltando += "DisplayName" }
    if ($faltando.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Preencha os campos obrigatórios: $($faltando -join ', ')", "Faltando campos")
        $outUni.AppendText("Preencha os campos obrigatórios: $($faltando -join ', ')\r\n")
        return
    }

    # Converta nome de país em código ISO (exemplo, personalize se quiser mais)
    $codigosPais = @{
        "Chile" = "CL"
        "Brasil" = "BR"
        "Colombia" = "CO"
        "Peru" = "PE"
        "Argentina" = "AR"
    }
    $pais = $txtPais.Text.Trim()
    if ($codigosPais.ContainsKey($pais)) { $pais = $codigosPais[$pais] }

    try {
        $senhaAleatoria = Nova-SenhaAleatoria 12
        $senha = ConvertTo-SecureString $senhaAleatoria -AsPlainText -Force

        $params = @{
            Server = $dominio
            Name = "$($txtNome.Text) $($txtSobrenome.Text)"
            GivenName = $txtNome.Text
            Surname = $txtSobrenome.Text
            SamAccountName = $txtLogin.Text
            UserPrincipalName = $txtEmail.Text
            EmailAddress = $txtEmail.Text
            EmployeeID = $txtRut.Text
            Company = $txtEmpresa.Text
            Office = $txtGerencia.Text
            Department = $txtSubger.Text
            StreetAddress = $txtEndereco.Text
            City = $txtCidade.Text
            State = $txtEstado.Text
            Manager = $txtSupervisor.Text
            Title = $txtCargo.Text
            HomePage = $webpage
            AccountPassword = $senha
            Enabled = $true
            Path = $ou
            ChangePasswordAtLogon = $true
            DisplayName = $txtDisplayName.Text
        }
        if ($pais) { $params["Country"] = $pais }
        if ($txtDescricao.Text) { $params["Description"] = $txtDescricao.Text }

        $params["OtherAttributes"] = @{
            employeeType = $txtTipo.Text.Trim()
            departmentNumber = $txtCC.Text.Trim()
        }

        New-ADUser @params

        $msg = "Usuário criado com sucesso!`r`n`r`nLogin: $($txtLogin.Text)`r`nEmail: $($txtEmail.Text)`r`nSenha: $senhaAleatoria"
        $outResultado.Text = "Login: $($txtLogin.Text)`r`nEmail: $($txtEmail.Text)`r`nSenha: $senhaAleatoria"
        $outResultado.Focus()
        $outResultado.SelectAll()
        [System.Windows.Forms.Clipboard]::SetText($outResultado.Text)  # só depois do Text!

        # Pop-up só para informar sucesso
        [System.Windows.Forms.MessageBox]::Show("Usuário criado! Os dados estão ao lado (copiados) e prontos para colar.", "Sucesso")

        $outUni.AppendText("$msg`r`n")
        Write-Log "Usuário criado: $($txtLogin.Text)"

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro ao criar usuário: $($_.Exception.Message)", "Erro")
        $outUni.AppendText("Erro ao criar usuário: $($_.Exception.Message)`r`n")
        Write-Log "Erro ao criar usuário: $($_.Exception.Message)"
    }
})

$tabs.TabPages.Add($tab3)


[void]$form.ShowDialog()