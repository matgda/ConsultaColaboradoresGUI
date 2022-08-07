[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();

#Código para ocultar o prompt de comando do powershell após a incialização do formulário
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#Declaração de elementos da GUI 
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='CONSULTA DE COLABORADORES'
$main_form.Width = 570
$main_form.Height = 450
$main_form.BackColor = "white"
$main_form.MaximizeBox = $false
$main_form.FormBorderStyle = 'Fixed3D'
$Font = New-Object System.Drawing.Font("Verdana",8)
$main_form.Font = $Font
$text1 = New-Object System.Windows.Forms.label
$text1.Location = New-Object System.Drawing.Size(7,10)
$text1.Size = New-Object System.Drawing.Size(270,15)
$text1.ForeColor = "black"
$text1.Text = "DIGITE O NOME DO COLABORADOR"
$text2 = New-Object System.Windows.Forms.label
$text2.Location = New-Object System.Drawing.Size(7,395)
$text2.Size = New-Object System.Drawing.Size(270,15)
$text2.ForeColor = "black"
$text2.Text = " "
$TextBoxPatch = New-Object System.Windows.Forms.TextBox
$TextBoxPatch.Location = '10,30'
$TextBoxPatch.Size = '250,50'

#Variavel global do usuário de rede para pesquisa
$global:textField 

#Função para chamar a pesquisa geral ao prescionar Enter no campo Nome
$TextBoxPatch.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        pesquisaGeral
    }
})

#Declaração de botões na tela inical
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(270,25)
$Button.Size = New-Object System.Drawing.Size(150,25)
$Button.Text = "PESQUISAR"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(270,55)
$Button2.Size = New-Object System.Drawing.Size(150,25)
$Button2.Text = "BUSCAR POR USUÁRIO"
#Delcaração do Lista de elementos 
$listBox = New-Object System.Windows.Forms.ListView
$listBox.Location = New-Object System.Drawing.Point(10,90)
$listBox.Size = New-Object System.Drawing.Size(540,20)
$listBox.Height = 300
$listBox.Columns.Add("Login")
$listBox.Columns.Add("Nome")
$listBox.Columns.Add("Status")
$listBox.View = [System.Windows.Forms.View]::Details
$listBox.MultiSelect = $false
$listBox.FullRowSelect = $true

#Função que após prescionar o botão pesquisar chamar a função pesquisa Geral
$Button.Add_Click({ 
    pesquisaGeral
})

#Função do botão buscar por usuário que incializa a tela de pesquisa por usuário de rede, e limpa se tiver essa tela estiver preenchida 
$Button2.Add_Click({
$listBox.SelectedItems.Clear()
buscarUsuario

})
#Função que pesquisa usuário por meio do nome
function pesquisaGeral{
#limpando o list view caso anteriomente foi preenchido
$listBox.Items.Clear();
#Pegando o valor do nome do campo de input
$nome =  $TextBoxPatch.Text
#Adicionando os caracteres especiais entre o nome para poder fazer a buscar sem precisar do nome exato 
$var = -join("*", $nome, "*");
#incializado a variavel contador
$contador = 0

#Verificação se de pesquisa não está vazia
if($TextBoxPatch.TextLength -eq 0){
     [System.Windows.MessageBox]::Show('Preencha o campo nome do colaborador!','Consultar Colaboradores')
    }else{
    #Consulta ao Active Diretory para buscar , usuário de rede, nome e status 
$dados = ForEach-Object {Get-Aduser -Filter {name -like $var } -properties name,samaccountname,enabled   | select name,samaccountname,enabled   }

#Laço de repetição para percorrer o resultado da busca
foreach ($dado in $dados) {

#Conventendo o valor da variavel enabled para string pois list view não suporta valores boolean
if($dado.enabled -eq $true){
   $status= "Ativo"
}else {
    $status= "Inativo"
}

#Adcionando cada valor em sua respecitiva coluna
$listBox.Items.Add($dado.samaccountname)
$listBox.Items[$contador].SubItems.Add($dado.name);
$listBox.Items[$contador].SubItems.Add($status);
#Varivael para percorrer 2 coluna em diante 
$contador +=1
} 

#Contador de resultados encontrados 
$text2.Text = 'Resultados Encontrados: '+$contador
}}

#Função para identificar qual linha foi precionanda com 2 clikes 
Function selecionarLinha(){
    if($listBox.SelectedItems.Count -eq 1){
    $usuarioRedeLinhaSelecionada = $listBox.SelectedItems[0].Text
   buscarUsuario
    }
}


#Função que cria a segunda tela buscar usuário de rede 
Function buscarUsuario{
#Inicialização do formulário
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();
#Declaração de elementos da GUI 
$form = New-Object Windows.Forms.Form
$form.text = "BUSCAR USUÁRIO DE REDE"
$form.BackColor = "white"
$form.Width = 350
$form.Height = 650
$form.MaximizeBox = $false
$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point 20,10
$label.Size = New-Object Drawing.Point 170,15
$label.text = "DIGITE O USUÁRIO DE REDE"
$global:textField = New-Object Windows.Forms.TextBox
$global:textField.Location = New-Object Drawing.Point 20,30
$global:textField.Size = New-Object Drawing.Point 200,15
$global:textField.Text = $listBox.SelectedItems[0].Text
$Font = New-Object System.Drawing.Font("Verdana",8)
$form.Font = $Font

$textField.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        buscarDados
    }
})

#Delcaração dos campos do formulário
$labelMatricula = New-Object Windows.Forms.Label
$labelMatricula.Location = New-Object Drawing.Point 20,60
$labelMatricula.Size = New-Object Drawing.Point 130,20
$labelMatricula.text = "MATRÍCULA"
$textFieldMatricula = New-Object Windows.Forms.TextBox
$textFieldMatricula.Location = New-Object Drawing.Point 20,80
$textFieldMatricula.Size = New-Object Drawing.Point 200,20
$textFieldMatricula.ReadOnly = "true"
$labelNome = New-Object Windows.Forms.Label
$labelNome.Location = New-Object Drawing.Point 20,110
$labelNome.Size = New-Object Drawing.Point 130,20
$labelNome.text = "NOME COMPLETO"
$textFieldNome = New-Object Windows.Forms.TextBox 
$textFieldNome.Location = New-Object Drawing.Point 20,130
$textFieldNome.Size = New-Object Drawing.Point 200,20
$textFieldNome.ReadOnly = "true"
$labelEmail = New-Object Windows.Forms.Label
$labelEmail.Location = New-Object Drawing.Point 20,160
$labelEmail.Size = New-Object Drawing.Point 130,20
$labelEmail.text = "E-MAIL"
$textFieldEmail = New-Object Windows.Forms.TextBox
$textFieldEmail.Location = New-Object Drawing.Point 20,180
$textFieldEmail.Size = New-Object Drawing.Point 200,20
$textFieldEmail.ReadOnly = "true"
$labelCargo = New-Object Windows.Forms.Label
$labelCargo.Location = New-Object Drawing.Point 20,210
$labelCargo.Size = New-Object Drawing.Point 130,20
$labelCargo.text = "CARGO"
$textFieldCargo = New-Object Windows.Forms.TextBox
$textFieldCargo.Location = New-Object Drawing.Point 20,230
$textFieldCargo.Size = New-Object Drawing.Point 200,20
$textFieldCargo.ReadOnly = "true"
$labelDepartamento = New-Object Windows.Forms.Label
$labelDepartamento.Location = New-Object Drawing.Point 20,260
$labelDepartamento.Size = New-Object Drawing.Point 130,20
$labelDepartamento.text = "DEPARTAMENTO"
$textFieldDepartamento = New-Object Windows.Forms.TextBox
$textFieldDepartamento.Location = New-Object Drawing.Point 20,280
$textFieldDepartamento.Size = New-Object Drawing.Point 200,20
$textFieldDepartamento.ReadOnly = "true"
$labelGestor = New-Object Windows.Forms.Label
$labelGestor.Location = New-Object Drawing.Point 20,310
$labelGestor.Size = New-Object Drawing.Point 130,20
$labelGestor.text = "GESTOR"
$textFieldGestor = New-Object Windows.Forms.TextBox
$textFieldGestor.Location = New-Object Drawing.Point 20,330
$textFieldGestor.Size = New-Object Drawing.Point 200,20
$textFieldGestor.ReadOnly = "true"
$labelCPF = New-Object Windows.Forms.Label
$labelCPF.Location = New-Object Drawing.Point 20,360
$labelCPF.Size = New-Object Drawing.Point 130,20
$labelCPF.text = "CPF"
$textFieldCPF = New-Object Windows.Forms.TextBox
$textFieldCPF.Location = New-Object Drawing.Point 20,380
$textFieldCPF.Size = New-Object Drawing.Point 200,20
$textFieldCPF.ReadOnly = "true"
$labelCC = New-Object Windows.Forms.Label
$labelCC.Location = New-Object Drawing.Point 20,410
$labelCC.Size = New-Object Drawing.Point 130,20
$labelCC.text = "CENTRO DE CUSTO"
$textFieldCC = New-Object Windows.Forms.TextBox
$textFieldCC.Location = New-Object Drawing.Point 20,430
$textFieldCC.Size = New-Object Drawing.Point 200,20
$textFieldCC.ReadOnly = "true"
$labelRamal = New-Object Windows.Forms.Label
$labelRamal.Location = New-Object Drawing.Point 20,460
$labelRamal.Size = New-Object Drawing.Point 130,20
$labelRamal.text = "TELEFONE/RAMAL"
$textFieldRamal = New-Object Windows.Forms.TextBox
$textFieldRamal.Location = New-Object Drawing.Point 20,480
$textFieldRamal.Size = New-Object Drawing.Point 200,20
$textFieldRamal.ReadOnly = "true"
$labelEmpresa = New-Object Windows.Forms.Label
$labelEmpresa.Location = New-Object Drawing.Point 20,510
$labelEmpresa.Size = New-Object Drawing.Point 130,20
$labelEmpresa.text = "EMPRESA"
$textFieldEmpresa = New-Object Windows.Forms.TextBox
$textFieldEmpresa.Location = New-Object Drawing.Point 20,530
$textFieldEmpresa.Size = New-Object Drawing.Point 200,20
$textFieldEmpresa.ReadOnly = "true"
$button = New-Object Windows.Forms.Button
$button.text = "Buscar"
$button.Location = New-Object Drawing.Point 20,570
$buttonCopiar = New-Object Windows.Forms.Button
$buttonCopiar.text = "Copiar"
$buttonCopiar.Location = New-Object Drawing.Point 140,570

#Verificação se o campo usuário de rede foi prenchido na tela anterior por meio linha selecicionada, caso sim é pesquisado as informações do colaborador e é preenchido o formulário
    if($textField.Text){
    $user = Get-Aduser -Filter {samaccountname -eq $textField.Text } -properties GivenName,SurName, mail , Title, Manager, Department,Description,Departmentnumber,EmployeeID,Company,telephoneNumber | Select-Object  GivenName,SurName, mail , Title, Manager, Department,Description,@{N='Departmentnumber';E={$_.Departmentnumber[0]}},EmployeeID,Company,telephoneNumber
    $nomeCompleto = -join($user.GivenName," ", $user.SurName);
    $textFieldNome.Text = $nomeCompleto
    $textFieldEmail.Text = $user.mail
    $textFieldCargo.Text = $user.Title
    $textFieldDepartamento.Text = $user.Department
    $textFieldGestor.Text = $user.Manager
    $textFieldCPF.Text = $user.Description
    $textFieldCC.Text = $user.Departmentnumber
    $textFieldMatricula.Text = $user.EmployeeID
    $textFieldEmpresa.Text = $user.Company
    $textFieldRamal.Text=$user.telephoneNumber

    }
    #Evendo do botão buscar na tela buscar por usuário de rede
    $button.add_click({
   buscarDados
})
    #Função que pesquisa as informações do colaborador e é preenchido o formulário
    function buscarDados{

    #Veirifcação se o campo usuário foi preenchido antes de realizar a busca
        if($textField.TextLength -eq 0){
         [System.Windows.MessageBox]::Show('Preencha o campo usuário de rede!','Consultar Colaboradores')
        }else{
        $user = Get-Aduser -Filter {samaccountname -eq $textField.Text } -properties GivenName,SurName, mail , Title, Manager, Department,Description,Departmentnumber,EmployeeID,Company,telephoneNumber | Select-Object  GivenName,SurName, mail , Title, Manager, Department,Description,@{N='Departmentnumber';E={$_.Departmentnumber[0]}},EmployeeID,Company,telephoneNumber
        $nomeCompleto = -join($user.GivenName," ", $user.SurName);
        $textFieldNome.Text = $nomeCompleto
        $textFieldEmail.Text = $user.mail
        $textFieldCargo.Text = $user.Title
        $textFieldDepartamento.Text = $user.Department
        $textFieldGestor.Text = $user.Manager
        $textFieldCPF.Text = $user.Description
        $textFieldCC.Text = $user.Departmentnumber
        $textFieldMatricula.Text = $user.EmployeeID
        $textFieldEmpresa.Text = $user.Company
        $textFieldRamal.Text=$user.telephoneNumber
            }
        }
# Adicionado elementos no form.
$form.controls.add($button)
$form.controls.add($buttonCopiar)
$form.controls.add($label)
$form.controls.add($textField)
$form.controls.add($labelMatricula)
$form.controls.add($textfieldMatricula)
$form.controls.add($labelNome)
$form.controls.add($textfieldNome)
$form.controls.add($labelEmail)
$form.controls.add($textfieldEmail)
$form.controls.add($labelCargo)
$form.controls.add($textfieldCargo)
$form.controls.add($labelDepartamento)
$form.controls.add($textfieldDepartamento)
$form.controls.add($labelGestor)
$form.controls.add($textfieldGestor)
$form.controls.add($labelCPF)
$form.controls.add($textfieldCPF)
$form.controls.add($labelCC)
$form.controls.add($textfieldCC)
$form.controls.add($labelRamal)
$form.controls.add($textfieldRamal)
$form.controls.add($labelEmpresa)
$form.controls.add($textfieldEmpresa)

#Função para copiar elementos para área de transferência
function copiarDados{
    $textFieldCopy = "Usuário de rede: "+$textField.Text
    $nomeCompletoCopy = "Nome completo: "+$textFieldNome.Text
    $textFieldEmailCopy = "E-mail: "+ $textFieldEmail.Text
    $textFieldMatriculaCopy ="Matrícula: "+$textFieldMatricula.Text 
    $textFieldCargoCopy = "Cargo: "+$textFieldCargo.Text
    $textFieldDepartamentoCopy = "Departamento: "+$textFieldDepartamento.Text
    $textFieldCPFCopy = "CPF: "+$textFieldCPF.Text
    $textFieldCCCopy = "Centro de custo: "+$textFieldCC.Text
    $textFieldRamalCopy = "Ramal/Telefone: "+$textFieldRamal.Text
    $textFieldEmpresaCopy = "Empresa: "+$textFieldEmpresa.Text

    Set-Clipboard $nomeCompletocopy,$textFieldMatriculaCopy,$textFieldCopy,$textFieldEmailCopy,$textFieldCargoCopy,$textFieldDepartamentoCopy,$textFieldCPFCopy,$textFieldCCCopy,$textFieldRamalCopy,$textFieldEmpresaCopy 
    [System.Windows.MessageBox]::Show('Dados foram copiados para área de transferência.','Consultar Colaboradores')
    }

$buttonCopiar.add_click({
      copiarDados
       })

#Mostrando o formulário
$form.ShowDialog()
}

#Instanciando formuário da tela inicial
$listBox.Add_ItemActivate({selecionarLinha})
$main_form.Controls.Add($text1)
$main_form.Controls.Add($text2)
$main_form.Controls.Add($TextBoxPatch)
$main_form.Controls.Add($Button)
$main_form.Controls.Add($Button2)
$main_form.Controls.Add($listBox)
$main_form.ShowDialog()


# SIG # Begin signature block
# MIIboQYJKoZIhvcNAQcCoIIbkjCCG44CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUkpmsVhkNQKs3NgANc2/k3b2D
# BvegghYXMIIDBjCCAe6gAwIBAgIQHa7OK3g1d4RGHwMUs61HGjANBgkqhkiG9w0B
# AQsFADAbMRkwFwYDVQQDDBBBVEEgQXV0aGVudGljb2RlMB4XDTIyMDcxMzIxNTM0
# MFoXDTIzMDcxMzIyMTM0MFowGzEZMBcGA1UEAwwQQVRBIEF1dGhlbnRpY29kZTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMigdeB/oM/Wsi9uY1AVH05U
# IhZyOezzTzMcpRKDQc5ktRABfinfkxWlr+Sz9eH1qJnjsTnXcAFv+EVh763nj6Ab
# i5JAmTPuTL0bvltw9f1Jngq4NPVD3jaPhq2Bbf4amhDWlqRuXD7EE+iXJnhbcG2M
# zuP35GNhE/JkPSbJd1+equZaxiisV7CpJJ9VtRbX6Y+tnz7D862+3bCKJS1JO9gq
# 3O/3XGlbIKrUkRWi16YrMiFw/wpnBGGSynviQqyVEKf2MT3/Mdcs0MB1JBDHyFw5
# 3/Dc/6Umn+adXhNtW14HvY/me9t0h0HH453/kNYZv1utonmFW1OGynDu0nHYUjUC
# AwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0G
# A1UdDgQWBBSZRT6t2sX0ecfhRNrDD+w5ZrwVVjANBgkqhkiG9w0BAQsFAAOCAQEA
# Tj+E4HkTLVtYBfiLkTn23Rgom2ux9CHAz/SRTp86o91m3aGj4CwzZ0O+N91rbr3v
# 6Zo3m42PPuktMeBRDyazQ0nKasQCBQtPBRCY+Xv86FLpOgEb98RKvBkurzb3hWnt
# BAT3IaUNqAWrNpakwVEsbLkP9c7RYEneYU9hwLTlh7FURz67D8icgP0zfSkj5eXu
# S3UGqDfeJIbjEnKXcFIP0O7BTa0jirt4Y/unpem1C0cqofxhs+VzML2lHOX/jOMy
# QGGvK5M4VrYjt5YwnnjxJn50nVyzII3WQg4VybMQPhvqwPSeHXqlFn2ytYX8bxQA
# +ty1imlIGY2p9xhqppQ7aDCCBY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFow
# DQYJKoZIhvcNAQEMBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNl
# cnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIz
# NTk1OVowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3Rl
# ZCBSb290IEc0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2je
# u+RdSjwwIjBpM+zCpyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bG
# l20dq7J58soR0uRf1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBE
# EC7fgvMHhOZ0O21x4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/N
# rDRAX7F6Zu53yEioZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A
# 2raRmECQecN4x7axxLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8
# IUzUvK4bA3VdeGbZOjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfB
# aYh2mHY9WV1CdoeJl2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaa
# RBkrfsCUtNJhbesz2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZi
# fvaAsPvoZKYz0YkH4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXe
# eqxfjT/JvNNBERJb5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g
# /KEexcCPorF+CiaZ9eRpL5gdLfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB
# /wQFMAMBAf8wHQYDVR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1Ud
# IAQKMAgwBgYEVR0gADANBgkqhkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22
# Ftf3v1cHvZqsoYcs7IVeqRq7IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih
# 9/Jy3iS8UgPITtAq3votVs/59PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYD
# E3cnRNTnf+hZqPC/Lwum6fI0POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c
# 2PR3WlxUjG/voVA9/HYJaISfb8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88n
# q2x2zm8jLfR+cWojayL/ErhULSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5
# lDCCBq4wggSWoAMCAQICEAc2N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290
# IEc0MB4XDTIyMDMyMzAwMDAwMFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBU
# cnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJ
# KoZIhvcNAQEBBQADggIPADCCAgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh
# 1tKD0Z5Mom2gsMyD+Vr2EaFEFUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+Feo
# An39Q7SE2hHxc7Gz7iuAhIoiGN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1
# decfBmWNlCnT2exp39mQh0YAe9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxnd
# X7RUCyFobjchu0CsX7LeSn3O9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6
# Th+xtVhNef7Xj3OTrCw54qVI1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPj
# Q2OAe3VuJyWQmDo4EbP29p7mO1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlREr
# WHRAKKtzQ87fSqEcazjFKfPKqpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JM
# q++bPf4OuGQq+nUoJEHtQr8FnGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh
# 3pP+OcD5sjClTNfpmEpYPtMDiP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8j
# u2TjY+Cm4T72wnSyPx4JduyrXUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnS
# DmuZDNIztM2xAgMBAAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1Ud
# DgQWBBS6FtltTYUvcyl2mi91jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzf
# Lmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgw
# dwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAG
# A1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOC
# AgEAfVmOwJO2b5ipRCIBfmbW2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp
# /GnBzx0H6T5gyNgL5Vxb122H+oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40B
# IiXOlWk/R3f7cnQU1/+rT4osequFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2d
# fNBwCnzvqLx1T7pa96kQsl3p/yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibB
# t94q6/aesXmZgaNWhqsKRcnfxI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7
# T6NJuXdmkfFynOlLAlKnN36TU6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZA
# myEhQNC3EyTN3B14OuSereU0cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdB
# eHo46Zzh3SP9HSjTx/no8Zhf+yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnK
# cPA3v5gA3yAWTyf7YGcWoWa63VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/
# pNHzV9m8BPqC3jLfBInwAM1dwvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yY
# lvZVVCsfgPrA8g4r5db7qS9EFUrnEw4d2zc4GqEr9u3WfPwwggbGMIIErqADAgEC
# AhAKekqInsmZQpAGYzhNhpedMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1
# c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjIwMzI5
# MDAwMDAwWhcNMzMwMzE0MjM1OTU5WjBMMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xJDAiBgNVBAMTG0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDIy
# IC0gMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALkqliOmXLxf1knw
# FYIY9DPuzFxs4+AlLtIx5DxArvurxON4XX5cNur1JY1Do4HrOGP5PIhp3jzSMFEN
# MQe6Rm7po0tI6IlBfw2y1vmE8Zg+C78KhBJxbKFiJgHTzsNs/aw7ftwqHKm9MMYW
# 2Nq867Lxg9GfzQnFuUFqRUIjQVr4YNNlLD5+Xr2Wp/D8sfT0KM9CeR87x5MHaGjl
# RDRSXw9Q3tRZLER0wDJHGVvimC6P0Mo//8ZnzzyTlU6E6XYYmJkRFMUrDKAz200k
# heiClOEvA+5/hQLJhuHVGBS3BEXz4Di9or16cZjsFef9LuzSmwCKrB2NO4Bo/tBZ
# mCbO4O2ufyguwp7gC0vICNEyu4P6IzzZ/9KMu/dDI9/nw1oFYn5wLOUrsj1j6siu
# gSBrQ4nIfl+wGt0ZvZ90QQqvuY4J03ShL7BUdsGQT5TshmH/2xEvkgMwzjC3iw9d
# RLNDHSNQzZHXL537/M2xwafEDsTvQD4ZOgLUMalpoEn5deGb6GjkagyP6+SxIXuG
# Z1h+fx/oK+QUshbWgaHK2jCQa+5vdcCwNiayCDv/vb5/bBMY38ZtpHlJrYt/YYcF
# aPfUcONCleieu5tLsuK2QT3nr6caKMmtYbCgQRgZTu1Hm2GV7T4LYVrqPnqYklHN
# P8lE54CLKUJy93my3YTqJ+7+fXprAgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8EBAMC
# B4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNVHSAE
# GTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2FL3Mp
# dpovdYxqII+eyG8wHQYDVR0OBBYEFI1kt4kh/lZYRIRhp+pvHDaP3a8NMFoGA1Ud
# HwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRy
# dXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsGAQUF
# BwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# WAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJKoZI
# hvcNAQELBQADggIBAA0tI3Sm0fX46kuZPwHk9gzkrxad2bOMl4IpnENvAS2rOLVw
# Eb+EGYs/XeWGT76TOt4qOVo5TtiEWaW8G5iq6Gzv0UhpGThbz4k5HXBw2U7fIyJs
# 1d/2WcuhwupMdsqh3KErlribVakaa33R9QIJT4LWpXOIxJiA3+5JlbezzMWn7g7h
# 7x44ip/vEckxSli23zh8y/pc9+RTv24KfH7X3pjVKWWJD6KcwGX0ASJlx+pedKZb
# NZJQfPQXpodkTz5GiRZjIGvL8nvQNeNKcEiptucdYL0EIhUlcAZyqUQ7aUcR0+7p
# x6A+TxC5MDbk86ppCaiLfmSiZZQR+24y8fW7OK3NwJMR1TJ4Sks3KkzzXNy2hcC7
# cDBVeNaY/lRtf3GpSBp43UZ3Lht6wDOK+EoojBKoc88t+dMj8p4Z4A2UKKDr2xpR
# oJWCjihrpM6ddt6pc6pIallDrl/q+A8GQp3fBmiW/iqgdFtjZt5rLLh4qk1wbfAs
# 8QcVfjW05rUMopml1xVrNQ6F1uAszOAMJLh8UgsemXzvyMjFjFhpr6s94c/MfRWu
# FL+Kcd/Kl7HYR+ocheBFThIcFClYzG/Tf8u+wQ5KbyCcrtlzMlkI5y2SoRoR/jKY
# pl0rl+CL05zMbbUNrkdjOEcXW28T2moQbh9Jt0RbtAgKh1pZBHYRoad3AhMcMYIE
# 9DCCBPACAQEwLzAbMRkwFwYDVQQDDBBBVEEgQXV0aGVudGljb2RlAhAdrs4reDV3
# hEYfAxSzrUcaMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBR6duBbS3Uha6kmhWNzp5KVTx4kBDAN
# BgkqhkiG9w0BAQEFAASCAQCYBFHtBPdvauRnK8YdnWv/8aU7rCA7XdQN8IgefUk6
# JMy65iNGGV4NR4K7WWFTbo71AmiPwjTkz7l6wGKg3CEv9wfDPGR82DX/Z/Tuk5Y/
# 2RcUevnNgExWCoDAeVdJNitOkjw9LcyxJiD4KGsmmz1Ap8Ac9SWb980vS2KYkOpC
# gnl4TYjV8yCCok5ta6UFFNUDp/d4+xC0HU0dc3MhandEBvTqugJ+JPgO+zH9Laic
# s+ec45EpmMVMTWXq/k3iFfDTj9ECD6etyZvk4lSR/ChGRKl7R1RY1kIuosjVINVk
# jSZ9JuCTvgXqhK2AVnl956m6BI4osg6bDmuElXcRl8groYIDIDCCAxwGCSqGSIb3
# DQEJBjGCAw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQQIQCnpKiJ7JmUKQBmM4TYaXnTANBglghkgB
# ZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkF
# MQ8XDTIyMDgwNzAyMjA1MFowLwYJKoZIhvcNAQkEMSIEIPf1gocpwwn+0c2smlXy
# X5tF5vrvwb569heEOnMdJ24jMA0GCSqGSIb3DQEBAQUABIICAIyJJkuNjz8wn7+6
# NHzmECOGk6ij/nLCZEU1N0TkNByDPQjZKLeV40417nLm7IWp9dlJcpruvy56SM87
# uHjLif0pZFXWh10jU0p+k7d89p7EbUtwiRdE4zGjipjQtf5o7XJn4+7Xhp/SUKeX
# ILRvlCNhhxB/9R6szffdZNb6RTWPrv4YS6dmkhAycvOW6OZCvuSecZFjwczN/JiP
# sG1Bd/UDdwg0K86dkl1rf0pc7VPQvpIA84aazEfhYX4nrtt1vkmw+CGL8zYxdyrH
# Pitjk8Itlh340OYGahnbgmYTCSNLhp6ZduoUtbjBkKwqGqN5LCefUhvFDJSfvtXQ
# 0mddueaFvbqU8mT8rPztaenCQXKueBuJ9wF8C1sWaT31N/ymBy23zpuS3QpkL2FE
# jfg6w9uE+ts9NHFwkpUX5wHT/01QIyL0Ys4p6gX1XtS0mkTXHaedlhWEqy5iK8sF
# 1IX8Bd/n+MEQR8tR7gEl4yPj+MKTHD9RDLNd/JS8pkSYr7WB3KkntfvVTPQUyG69
# l+rjLsFCC3a52RDI3Dd3k1OmuKrfrveeRrhxJ36d8mpUaWrDZh9/hKWZCcldnJgK
# RIP7cI7UO6o1Pf3rAE1CoqdiXK3A+f/540hforNSWiVFzL1+8RnioYNAqozB4kPF
# pzESReD6YbgUTbyxYq4gQggygxfG
# SIG # End signature block
