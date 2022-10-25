VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLogin 
   Caption         =   "UserForm1"
   ClientHeight    =   4092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "formLogin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkVisualizar_Change()
    'alterar a cor da fonte
    If txtSenha.ForeColor = &H80000012 Then
        txtSenha.ForeColor = &H8000000E
    Else
        txtSenha.ForeColor = &H80000012
    End If
    
End Sub
Private Sub cmdEncerrar_Click()
    'descarrega o formulário
    Unload Me
End Sub
Private Sub cmdLogin_click()
On Error GoTo fim:

    'verifica se a senha está correta e caso esteja abre o formulário gerenciador
    indice = PROCURAR(comboUsuario, Planilha5.Range("A:A"))
    
    If comboUsuario = Empty Then
        Call MsgBox("Usuário inválido.", vbOKOnly, "Atenção")
    ElseIf txtSenha <> CDbl(Planilha5.Cells(indice, 3)) Then
        Call MsgBox("Senha inválida.", vbOKOnly, "Atenção")
    Else
        'Resetando o usuário logado
        Planilha5.Range("D2:D11") = 0
        'Definindo o usuário logado
        Planilha5.Cells(indice, 4) = 1
        'Abrindo o formulário gerenciador
        formGerenciador.Show
        'Encerrando o formulário de login
        Unload Me
    End If

Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Atenção")
    
End Sub
Private Sub comboUsuario_AfterUpdate()
On Error GoTo fim:

    'Verifica se o usuário está cadastrado, caso contrário, limpa o nome inserido
    If PROCURAR(comboUsuario, Planilha5.Range("A:A")) = 0 Then
        Call MsgBox("Usuário não cadastrado.", vbOKOnly, "Atenção")
        comboUsuario = Empty
    End If

Exit Sub

fim:

    Call MsgBox("ERRO.", vbCritical, "Atenção")
    
End Sub

Private Sub txtSenha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'permite apenas valores numéricos no txtbox
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()

    'Ocultar a barra de título do formulário
    Call BarradeTitulo(formLogin)

    'Definindo a cor padrão da letra da senha
    txtSenha.ForeColor = &H8000000E
    
    'Definindo a lista de usuários para login
    comboUsuario.RowSource = "usuarios"
End Sub
