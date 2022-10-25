VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formGerenciador 
   Caption         =   "UserForm1"
   ClientHeight    =   9432.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15288
   OleObjectBlob   =   "formGerenciador.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SESS�O I - ELEMENTOS GERAIS DO FORMUL�RIO

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'(1) Inicializa��o e finaliza��o do formul�rio
Private Sub UserForm_Initialize()
On Error GoTo fim:

    'Ocultar barra de t�tulo do formul�rio
    Call BarradeTitulo(formGerenciador)
    
    'Definindo a lista de usu�rios para login
    lstUsuarios.RowSource = "usuarios"
    lstServicos.RowSource = "servicos"
    lstTipoServico.RowSource = "tipo"
    lstCategoriaServico.RowSource = "categoria"
    lstEquipamento.RowSource = "equipamento"
    lstMetrica.RowSource = "medida"
    lstClientes.RowSource = "clientes"
    
    comboPermissao.RowSource = "permissao"
    comboTipoServico.RowSource = "tipo"
    comboCategoriaServico.RowSource = "categoria"
    comboEquipamentoServico.RowSource = "equipamento"
    comboMetricaServico.RowSource = "medida"
    comboCaracteristica.RowSource = "caracteristica"
    comboUFCliente.RowSource = "uf"
    
Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
End Sub

Private Sub UserForm_Terminate()

    Planilha5.Range("D2:D11") = 0

End Sub

'==================================================================================================================================
'(2) Bot�es de comando
Private Sub cmdClientes_click()
On Error GoTo fim:
        With Gerenciador
        .Pages(0).Visible = False
        .Pages(1).Visible = False
        .Pages(2).Visible = True
        .Pages(3).Visible = False
        .Pages(4).Visible = False
        .Value = 2
        End With
        
        indice = Application.WorksheetFunction.CountA(Planilha1.Range("A:A"))
        
        If indice = 1 Then
            codigo = 1
        Else
            codigo = CDbl(Right(Planilha1.Cells(indice, 1), 4)) + 1
        End If
        
        txtCodigoCliente = "CL" & Format(codigo, "0000")
        
Exit Sub
fim:
    Call MsgBox("ERRO.", vbCritical, "Aten��o")
End Sub
Private Sub cmdServicos_click()
On Error GoTo fim:
        With Gerenciador
        .Pages(0).Visible = False
        .Pages(1).Visible = False
        .Pages(2).Visible = False
        .Pages(3).Visible = True
        .Pages(4).Visible = False
        .Value = 3
        End With
Exit Sub
fim:
    Call MsgBox("ERRO.", vbCritical, "Aten��o")
End Sub
Private Sub cmdUsuarios_Click()
On Error GoTo fim

    'Analisando a permiss�o do usu�rio
    permissao = Application.WorksheetFunction.XLookup(1, Planilha5.Range("D:D"), Planilha5.Range("B:B"), 0, 0, 1)
    
    If permissao = 1 Or permissao = 2 Then
        'Definindo o layout das p�ginas do multipage
        With Gerenciador
        .Pages(0).Visible = False
        .Pages(1).Visible = False
        .Pages(2).Visible = False
        .Pages(3).Visible = False
        .Pages(4).Visible = True
        .Value = 4
        End With
        
        'Ocultando os bot�es de editar e excluir usu�rios para usu�rios de n�vel 2
        If permissao = 2 Then
            cmdEditarUsuario.Visible = False
            cmdExcluirUsuario.Visible = False
            
            comboPermissao.RowSource = "permissao2"
        End If
    Else
        Call MsgBox("Usu�rio n�o tem permiss�o para acessar essa �rea.", vbOKOnly, "Aten��o")
    End If

Exit Sub

fim:
    Call MsgBox("ERRO.", vbCritical, "Aten��o")
End Sub
Private Sub cmdEncerrar_Click()
    'Descarregar formul�rio
    Unload Me
End Sub
'==================================================================================================================================
'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
'==================================================================================================================================

'SESS�O II - P�GINA DE CADASTRO DE USU�RIOS

'(1) Bot�es de Comando
Private Sub cmdIncluirUsuario_Click()
On Error GoTo fim:
    
    'Determinando a posi��o da linha aonde ser� inserido o novo usu�rio
    indice = Application.WorksheetFunction.CountA(Planilha5.Range("A:A")) + 1
    
    If PROCURAR(txtUsuario, Planilha5.Range("A:A")) <> 0 Then
        Call MsgBox("Usu�rio j� cadastrado.", vbCritical, "Aten��o")
        
    Else

        decisao = MsgBox("Deseja incluir o usu�rio?", vbYesNo, "Aten��o")
    
        If decisao = vbYes Then
            'Inserindo os valores do formul�rio na base de dados
            With Planilha5
            .Cells(indice, 1) = txtUsuario
            .Cells(indice, 2) = CDbl(comboPermissao)
            .Cells(indice, 3) = CDbl(txtSenha)
            End With
        End If
    
    End If
    
Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdEditarUsuario_click()
On Error GoTo fim:
    
    cmdSalvarEditarUsuario.Visible = True
    
    'Determinando o nome do usu�rio selecionado na listbox
    indice_lista = lstUsuarios.ListIndex
    usuario = lstUsuarios.List(indice_lista, 0)
    
    'Determinando a posi��o do usu�rio na planilha de cadastro
    indice_usuario = PROCURAR(usuario, Planilha5.Range("A:A"))
    
    'Preenchendo as informa��es para edi��o
    txtUsuario = Planilha5.Cells(indice_usuario, 1)
    txtUsuario.Locked = True
    txtSenha = Planilha5.Cells(indice_usuario, 3)
    comboPermissao = Planilha5.Cells(indice_usuario, 2)

Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdSalvarEditarUsuario_click()
On Error GoTo fim:
    
    decisao = MsgBox("Deseja confirmar a edi��o?", vbYesNo, "Aten��o")
    
    If decisao = vbYes Then
    
        'Determinando a posi��o do usu�rio na planilha de cadastro
        indice_usuario = PROCURAR(txtUsuario, Planilha5.Range("A:A"))
        
        'Inserindo as informa��es editadas na planilha
        With Planilha5
            .Cells(indice_usuario, 3) = CDbl(txtSenha)
            .Cells(indice_usuario, 2) = CDbl(comboPermissao)
        End With
        
    End If

Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")

End Sub
Private Sub cmdExcluirUsuario_click()
On Error GoTo fim:
    
    decisao = MsgBox("Deseja confirmar a exclus�o?", vbYesNo, "Aten��o")
    
    If decisao = vbYes Then
    
        'Determinando o nome do usu�rio selecionado na listbox
        indice_lista = lstUsuarios.ListIndex
        usuario = lstUsuarios.List(indice_lista, 0)
        
        'Determinando a posi��o do usu�rio na planilha de cadastro
        indice_usuario = PROCURAR(usuario, Planilha5.Range("A:A"))
        
        'Excluindo a linha com o usu�rio selecionado e subindo uma linha acima
        Planilha5.Rows(indice_usuario).Delete Shift:=x1up
        
    End If
    
Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
   
End Sub

'==================================================================================================================================
'(2) Fun��o da intera��o com a lista
Private Sub lstUsuarios_Click()

    cmdSalvarEditarUsuario.Visible = False

End Sub


'==================================================================================================================================
'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
'==================================================================================================================================

'SESS�O III - P�GINA DE CADASTRO DE SERVI�OS

'(1) Bot�es de Comando
Private Sub cmdLimparServico_click()
            'limpar os campos do formul�rio
            txtCodigoServico = Empty
            comboTipoServico = Empty
            comboCategoriaServico = Empty
            comboEquipamentoServico = Empty
            comboMetricaServico = Empty
            txtPrecoServico = Empty
            txtResumoServico = Empty
            txtDescricaoServico = Empty

End Sub
Private Sub cmdSalvarServico_click()
On Error GoTo fim
    
    If txtCodigoServico <> Empty And _
        comboTipoServico <> Empty And _
        comboCategoriaServico <> Empty And _
        comboEquipamentoServico <> Empty And _
        comboMetricaServico <> Empty And _
        txtPrecoServico <> Empty Then
        
        decisao = MsgBox("Deseja salvar o registro?", vbYesNo, "Aten��o")
        
        If decisao = vbYes Then
        
            'determinando a linha para inserir as informa��es
            indice = Application.WorksheetFunction.CountA(Planilha2.Range("A:A")) + 1
            
            'concatena os valores inseridos no formul�rio
            concatena = Application.WorksheetFunction.Concat(comboTipoServico, comboCategoriaServico, _
                                                            comboEquipamentoServico, comboMetricaServico)
            
            analise = 0
                        
            With Planilha2
                For i = 1 To indice
                        'concatena as vari�veis da observa��o, presentas na coluna i da tabela
                        verifica = Application.WorksheetFunction.Concat(.Cells(i, 2), .Cells(i, 3), .Cells(i, 4), _
                                                                        .Cells(i, 5))
                                                                                       
                        'verifica se as informa��es digitadas j� constam na planilha
                        If verifica = concatena Then
                            analise = 1
                        End If
                Next
            End With
            
            If analise = 0 Then
                    'inser��o das informa��es na planilha
                    With Planilha2
                        .Cells(indice, 1) = txtCodigoServico
                        .Cells(indice, 2) = comboTipoServico
                        .Cells(indice, 3) = comboCategoriaServico
                        .Cells(indice, 4) = comboEquipamentoServico
                        .Cells(indice, 5) = comboMetricaServico
                        .Cells(indice, 6) = txtPrecoServico
                        .Cells(indice, 7) = txtResumoServico
                        .Cells(indice, 8) = txtDescricaoServico
                    End With
                    
                    'limpar os campos do formul�rio
                    txtCodigoServico = Empty
                    comboTipoServico = Empty
                    comboCategoriaServico = Empty
                    comboEquipamentoServico = Empty
                    comboMetricaServico = Empty
                    txtPrecoServico = Empty
                    txtResumoServico = Empty
                    txtDescricaoServico = Empty
                    
            Else
                
                Call MsgBox("Servi�o j� cadastrado!", vbOKOnly, "Aten��o")
                
            End If
            
        End If
    Else
        
        Call MsgBox("Preencha todas as informa��es", vbOKOnly, "Aten��o")
    
    End If
        
Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
End Sub
Private Sub cmdIncluirCaracteristica_click()
On Error GoTo fim
    
    'Definindo os intervalos de refer�ncia e a coluna correspondente
    intervalo = Application.WorksheetFunction.XLookup(comboCaracteristica, Planilha4.Range("J:J"), Planilha4.Range("K:K"), 0, 0, 1)
    coluna = Application.WorksheetFunction.XLookup(comboCaracteristica, Planilha4.Range("J:J"), Planilha4.Range("L:L"), 0, 0, 1)
    
    'Definindo o �ndice de inser��o dos dados
    indice = Application.WorksheetFunction.CountA(Planilha4.Range(coluna & ":" & coluna)) + 1
    
    'Inserindo os dados
    Planilha4.Cells(indice, coluna) = txtDescricaoCaracteristica

Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdEditarServico_click()
On Error GoTo fim
    
    cmdSalvarEditarServico.Visible = True
    cmdCancelarEditarServico.Visible = True
    
        'Determinando o nome do usu�rio selecionado na listbox
        indice_lista = lstServicos.ListIndex
        servico = lstServicos.List(indice_lista, 0)
        
        'Determinando a posi��o do usu�rio na planilha de cadastro
        indice_servico = PROCURAR(servico, Planilha2.Range("A:A"))
        
        'Preenchendo as informa��es para edi��o
        txtCodigoServico = Planilha2.Cells(indice_servico, 1)
        txtCodigoServico.Locked = True
        comboTipoServico = Planilha2.Cells(indice_servico, 2)
        comboCategoriaServico = Planilha2.Cells(indice_servico, 3)
        comboEquipamentoServico = Planilha2.Cells(indice_servico, 4)
        comboMetricaServico = Planilha2.Cells(indice_servico, 5)
        txtPrecoServico = Planilha2.Cells(indice_servico, 6)
        txtResumoServico = Planilha2.Cells(indice_servico, 7)
        txtDescricaoServico = Planilha2.Cells(indice_servico, 8)

Exit Sub
fim:
    
    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdSalvarEditarServico_click()
On Error GoTo fim
    
    indice = Application.WorksheetFunction.Match(txtCodigoServico, Planilha2.Range("A:A"), 0)
    
    decisao = MsgBox("Deseja confirmar a edi��o?", vbYesNo, "Aten��o")
    
    If decisao = vbYes Then
        'inser��o das informa��es na planilha
        With Planilha2
            .Cells(indice, 2) = comboTipoServico
            .Cells(indice, 3) = comboCategoriaServico
            .Cells(indice, 4) = comboEquipamentoServico
            .Cells(indice, 5) = comboMetricaServico
            .Cells(indice, 6) = txtPrecoServico
            .Cells(indice, 7) = txtResumoServico
            .Cells(indice, 8) = txtDescricaoServico
        End With
        
        'limpar os campos do formul�rio
        txtCodigoServico = Empty
        comboTipoServico = Empty
        comboCategoriaServico = Empty
        comboEquipamentoServico = Empty
        comboMetricaServico = Empty
        txtPrecoServico = Empty
        txtResumoServico = Empty
        txtDescricaoServico = Empty
        
        'ocultar os bot�es de edi��o
        cmdSalvarEditarServico.Visible = False
        cmdCancelarEditarServico.Visible = False
    End If
    
Exit Sub
fim:
    
    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    cmdSalvarEditarServico.Visible = False
    cmdCancelarEditarServico.Visible = False
    
End Sub
Private Sub cmdCancelarEditarServico_click()
On Error GoTo fim
    
    'limpar os campos do formul�rio
    txtCodigoServico = Empty
    comboTipoServico = Empty
    comboCategoriaServico = Empty
    comboEquipamentoServico = Empty
    comboMetricaServico = Empty
    txtPrecoServico = Empty
    txtResumoServico = Empty
    txtDescricaoServico = Empty
    
    'ocultar os bot�es de edi��o
    cmdSalvarEditarServico.Visible = False
    cmdCancelarEditarServico.Visible = False
    
Exit Sub
fim:
    
    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
    cmdSalvarEditarServico.Visible = False
    cmdCancelarEditarServico.Visible = False
    
End Sub
Private Sub cmdExcluirServico_click()
On Error GoTo fim

    'determinando a permiss�o de usu�rio
    permissao = Application.WorksheetFunction.XLookup(1, Planilha5.Range("D:D"), Planilha5.Range("B:B"), 0, 0, 1)
    
    If permissao = 1 Or permissao = 2 Then
        
        decisao = MsgBox("Confirmar exclus�o de servi�o?", vbYesNo, "Aten��o")
        
        If decisao = vbYes Then
            'Determinando o nome do usu�rio selecionado na listbox
            indice_lista = lstServicos.ListIndex
            servico = lstServicos.List(indice_lista, 0)
            
            'Determinando a posi��o do usu�rio na planilha de cadastro
            indice_servico = PROCURAR(servico, Planilha2.Range("A:A"))
            
            'Excluindo a linha com o usu�rio selecionado e subindo uma linha acima
            Planilha2.Rows(indice_servico).Delete Shift:=x1up
        
        End If
    Else
        
        Call MsgBox("Usu�rio sem permiss�o", vbOKOnly, "Aten��o")
        
    End If
Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")

End Sub
'====================================================================================================================================
'(2) Comportamento dos campos
Private Sub comboTipoServico_AfterUpdate()
On Error GoTo fim

    'determinando as vari�veis de codifica��o dos servi�os
    contador = Application.WorksheetFunction.CountA(Planilha2.Range("A:A"))
    
    If contador = 1 Then
        numero = 1
    Else
        numero = CDbl(Right(Planilha2.Cells(contador, 1), 3)) + 1
    End If
    
    codigo = Application.WorksheetFunction.XLookup(comboTipoServico, Planilha4.Range("B:B"), Planilha4.Range("C:C"), 0, 0, 1)
    
    txtCodigoServico = codigo & Format(numero, "000")
        
Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub txtPrecoServico_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 44 Or KeyAscii > 57 Or _
        KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 47 Then
        KeyAscii = 0
    End If
End Sub


'==================================================================================================================================
'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
'==================================================================================================================================

'SESS�O IV - P�GINA DE CADASTRO DE CLIENTES

'(1) Comportamento dos campos de cadastro
Private Sub txtTelefoneCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCEPCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCNPJCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtNumeroCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End Sub

'(2) Bot�es de comando
Private Sub cmdSalvarCliente_click()
On Error GoTo fim:
    
    'permitir o cadastro apenas de clientes com as informa��es legais e de contato
    If txtCNPJCliente = Empty Or txtRazaoCliente = Empty _
        Or txtNomeCliente = Empty Or txtContatoCliente = Empty _
        Or txtTelefoneCliente = Empty Or txtEmailCliente = Empty Then
        
        Call MsgBox("Cadastro Incompleto", vbOKOnly, "Aten��o")
        
    Else
    
        decisao = MsgBox("Confirmar cadastro de cliente?", vbYesNo, "Aten��o")
        
        If decisao = vbYes Then
        
            indice = Application.WorksheetFunction.CountA(Planilha1.Range("A:A")) + 1
            
            'adicionando as informa��es do formul�rio na planilha
            With Planilha1
            .Cells(indice, 1) = txtCodigoCliente
            .Cells(indice, 2) = txtCNPJCliente
            .Cells(indice, 3) = txtRazaoCliente
            .Cells(indice, 4) = txtNomeCliente
            .Cells(indice, 5) = txtContatoCliente
            .Cells(indice, 6) = txtTelefoneCliente
            .Cells(indice, 7) = txtEmailCliente
            .Cells(indice, 8) = txtLogradouroCliente
            .Cells(indice, 9) = txtNumeroCliente
            .Cells(indice, 10) = txtComplementoCliente
            .Cells(indice, 11) = txtBairroCliente
            .Cells(indice, 12) = txtCidadeCliente
            .Cells(indice, 13) = comboUFCliente
            .Cells(indice, 14) = txtCEPCliente
            End With
            
            'limpando as informa��es do formul�rio
            txtCodigoCliente = Empty
            txtCNPJCliente = Empty
            txtRazaoCliente = Empty
            txtNomeCliente = Empty
            txtContatoCliente = Empty
            txtTelefoneCliente = Empty
            txtEmailCliente = Empty
            txtLogradouroCliente = Empty
            txtNumeroCliente = Empty
            txtComplementoCliente = Empty
            txtBairroCliente = Empty
            txtCidadeCliente = Empty
            comboUFCliente = Empty
            txtCEPCliente = Empty
        
        'reinserindo o c�digo de cliente no campo do formu�rio
        indice_codigo = Application.WorksheetFunction.CountA(Planilha1.Range("A:A"))
        
        If indice_codigo = 1 Then
            codigo = 1
        Else
            codigo = CDbl(Right(Planilha1.Cells(indice_codigo, 1), 4)) + 1
        End If
        
        txtCodigoCliente = "CL" & Format(codigo, "0000")
    
        End If
        
    End If
    
Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdLimparCliente_click()
On Error GoTo fim
        
        'limpando os campos do formul�rio
        txtCodigoCliente = Empty
        txtCNPJCliente = Empty
        txtRazaoCliente = Empty
        txtNomeCliente = Empty
        txtContatoCliente = Empty
        txtTelefoneCliente = Empty
        txtEmailCliente = Empty
        txtLogradouroCliente = Empty
        txtNumeroCliente = Empty
        txtComplementoCliente = Empty
        txtBairroCliente = Empty
        txtCidadeCliente = Empty
        comboUFCliente = Empty
        txtCEPCliente = Empty

        'reinserindo o c�digo de cliente no campo do formu�rio
        indice_codigo = Application.WorksheetFunction.CountA(Planilha1.Range("A:A"))
        
        If indice_codigo = 1 Then
            codigo = 1
        Else
            codigo = CDbl(Right(Planilha1.Cells(indice_codigo, 1), 4)) + 1
        End If
        
        txtCodigoCliente = "CL" & Format(codigo, "0000")
        
Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdEditarCliente_click()
On Error GoTo fim:
    
    'exibindo os bot�es de salvar e cancelar edi��o
    cmdSalvarEditarCliente.Visible = True
    cmdCancelarEditarCliente.Visible = True
    
    'Determinando o nome do usu�rio selecionado na listbox
    indice_lista = lstClientes.ListIndex
    cliente = lstClientes.List(indice_lista, 0)
    
    'Determinando a posi��o do usu�rio na planilha de cadastro
    indice_cliente = PROCURAR(cliente, Planilha1.Range("A:A"))
    
    'preenchendo os campos do formul�rio com o cliente desejado
    txtCodigoCliente = Planilha1.Cells(indice_cliente, 1)
    txtCNPJCliente = Planilha1.Cells(indice_cliente, 2)
    txtRazaoCliente = Planilha1.Cells(indice_cliente, 3)
    txtNomeCliente = Planilha1.Cells(indice_cliente, 4)
    txtContatoCliente = Planilha1.Cells(indice_cliente, 5)
    txtTelefoneCliente = Planilha1.Cells(indice_cliente, 6)
    txtEmailCliente = Planilha1.Cells(indice_cliente, 7)
    txtLogradouroCliente = Planilha1.Cells(indice_cliente, 8)
    txtNumeroCliente = Planilha1.Cells(indice_cliente, 9)
    txtComplementoCliente = Planilha1.Cells(indice_cliente, 10)
    txtBairroCliente = Planilha1.Cells(indice_cliente, 11)
    txtCidadeCliente = Planilha1.Cells(indice_cliente, 12)
    comboUFCliente = Planilha1.Cells(indice_cliente, 13)
    txtCEPCliente = Planilha1.Cells(indice_cliente, 14)

Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
End Sub
Private Sub cmdExcluirCliente_click()
On Error GoTo fim

    'determinando a permiss�o de usu�rio
    permissao = Application.WorksheetFunction.XLookup(1, Planilha5.Range("D:D"), Planilha5.Range("B:B"), 0, 0, 1)
    
    If permissao = 1 Or permissao = 2 Then
        
        decisao = MsgBox("Confirmar exclus�o de cliente?", vbYesNo, "Aten��o")
        
        If decisao = vbYes Then
            'Determinando o nome do usu�rio selecionado na listbox
            indice_lista = lstClientes.ListIndex
            cliente = lstClientes.List(indice_lista, 0)
            
            'Determinando a posi��o do usu�rio na planilha de cadastro
            indice_cliente = PROCURAR(cliente, Planilha1.Range("A:A"))
            
            'Excluindo a linha com o usu�rio selecionado e subindo uma linha acima
            Planilha1.Rows(indice_cliente).Delete Shift:=x1up
        
        End If
    Else
        
        Call MsgBox("Usu�rio sem permiss�o", vbOKOnly, "Aten��o")
        
    End If
Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")

End Sub
Private Sub cmdSalvarEditarCliente_click()
On Error GoTo fim

    'permitir o cadastro apenas de clientes com as informa��es legais e de contato
    If txtCNPJCliente = Empty Or txtRazaoCliente = Empty _
        Or txtNomeCliente = Empty Or txtContatoCliente = Empty _
        Or txtTelefoneCliente = Empty Or txtEmailCliente = Empty Then
        
        Call MsgBox("Cadastro Incompleto", vbOKOnly, "Aten��o")
        
    Else
    
        decisao = MsgBox("Confirmar cadastro de cliente?", vbYesNo, "Aten��o")
        
        If decisao = vbYes Then
        
            indice = Application.WorksheetFunction.Match(txtCodigoCliente, Planilha1.Range("A:A"), 0)
            
            'adicionando as informa��es do formul�rio na planilha
            With Planilha1
            .Cells(indice, 2) = txtCNPJCliente
            .Cells(indice, 3) = txtRazaoCliente
            .Cells(indice, 4) = txtNomeCliente
            .Cells(indice, 5) = txtContatoCliente
            .Cells(indice, 6) = txtTelefoneCliente
            .Cells(indice, 7) = txtEmailCliente
            .Cells(indice, 8) = txtLogradouroCliente
            .Cells(indice, 9) = txtNumeroCliente
            .Cells(indice, 10) = txtComplementoCliente
            .Cells(indice, 11) = txtBairroCliente
            .Cells(indice, 12) = txtCidadeCliente
            .Cells(indice, 13) = comboUFCliente
            .Cells(indice, 14) = txtCEPCliente
            End With
            
            'limpando as informa��es do formul�rio
            txtCodigoCliente = Empty
            txtCNPJCliente = Empty
            txtRazaoCliente = Empty
            txtNomeCliente = Empty
            txtContatoCliente = Empty
            txtTelefoneCliente = Empty
            txtEmailCliente = Empty
            txtLogradouroCliente = Empty
            txtNumeroCliente = Empty
            txtComplementoCliente = Empty
            txtBairroCliente = Empty
            txtCidadeCliente = Empty
            comboUFCliente = Empty
            txtCEPCliente = Empty
        
            'reinserindo o c�digo de cliente no campo do formu�rio
            indice_codigo = Application.WorksheetFunction.CountA(Planilha1.Range("A:A"))
            
            If indice_codigo = 1 Then
                codigo = 1
            Else
                codigo = CDbl(Right(Planilha1.Cells(indice_codigo, 1), 4)) + 1
            End If
            
            txtCodigoCliente = "CL" & Format(codigo, "0000")
            End If
    
    End If

    'ocultando os bot�es de salvar e cancelar edi��o
    cmdSalvarEditarCliente.Visible = False
    cmdCancelarEditarCliente.Visible = False
Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
    'ocultando os bot�es de salvar e cancelar edi��o
    cmdSalvarEditarCliente.Visible = False
    cmdCancelarEditarCliente.Visible = False
    
End Sub
Private Sub cmdCancelarEditarCliente_click()
On Error GoTo fim

    'limpando as informa��es do formul�rio
    txtCodigoCliente = Empty
    txtCNPJCliente = Empty
    txtRazaoCliente = Empty
    txtNomeCliente = Empty
    txtContatoCliente = Empty
    txtTelefoneCliente = Empty
    txtEmailCliente = Empty
    txtLogradouroCliente = Empty
    txtNumeroCliente = Empty
    txtComplementoCliente = Empty
    txtBairroCliente = Empty
    txtCidadeCliente = Empty
    comboUFCliente = Empty
    txtCEPCliente = Empty
    
    cmdSalvarEditarCliente.Visible = False
    cmdCancelarEditarCliente.Visible = False

    'reinserindo o c�digo de cliente no campo do formu�rio
    indice_codigo = Application.WorksheetFunction.CountA(Planilha1.Range("A:A"))
    
    If indice_codigo = 1 Then
        codigo = 1
    Else
        codigo = CDbl(Right(Planilha1.Cells(indice_codigo, 1), 4)) + 1
    End If
    
    txtCodigoCliente = "CL" & Format(codigo, "0000")

Exit Sub
fim:

    Call MsgBox("ERRO.", vbCritical, "Aten��o")
    
    'ocultando os bot�es de salvar e cancelar edi��o
    cmdSalvarEditarCliente.Visible = False
    cmdCancelarEditarCliente.Visible = False
    
End Sub
