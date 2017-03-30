VERSION 5.00
Begin VB.Form frmAdminLoginOpcoes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2025
   ClientLeft      =   4155
   ClientTop       =   4155
   ClientWidth     =   4485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opt 
      Caption         =   "&Desativar Login"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Retirar o meu Login da Se��o Ativa"
      Top             =   780
      Width           =   1455
   End
   Begin VB.OptionButton opt 
      Caption         =   "&Ativar Login"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Tornar o meu Login como padr�o"
      Top             =   540
      Width           =   1185
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   750
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Digite a Senha do Usu�rio"
      Top             =   1605
      Width           =   1335
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Index           =   0
      Left            =   750
      TabIndex        =   4
      ToolTipText     =   "Digite o Nome do Usu�rio (Nome Resumido)"
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   315
      Index           =   1
      Left            =   3420
      TabIndex        =   7
      ToolTipText     =   "Cancela e Fecha esta Janela"
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Confirma as op��es selecionadas"
      Top             =   1620
      Width           =   975
   End
   Begin VB.ListBox lstUsu�riosLogados 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   2370
      TabIndex        =   8
      ToolTipText     =   "Lista os Usu�rios Logados no Sistema"
      Top             =   240
      Width           =   2025
   End
   Begin VB.OptionButton opt 
      Caption         =   "&Logoff"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Finalizar todas as Se��es do meu Login"
      Top             =   1020
      Width           =   825
   End
   Begin VB.OptionButton opt 
      Caption         =   "&Novo Login"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Iniciar uma nova Se��o de Login e �rea de Trabalho"
      Top             =   300
      Width           =   1185
   End
   Begin VB.Label lbl 
      Caption         =   "Senha"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Usu�rio"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Usu�rios Logados"
      Height          =   195
      Left            =   2370
      TabIndex        =   10
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label lbl 
      Caption         =   "O que voc� deseja fazer?"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   30
      Width           =   2025
   End
End
Attribute VB_Name = "frmAdminLoginOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim log As New DLLSystemManager.log
Dim strEstacao As String
Private Enum Op��esLogin
    OPL_NovoLogin = 0
    OPL_LogOff = 1
    OPL_AtivarLogin = 4
    OPL_DesativarLogin = 2
End Enum
Dim Op��oDeLoginSelecionada As Op��esLogin

Private Sub cmd_Click(Index As Integer)
    
    Dim I As Integer
    Dim frm As frmAdminDeskTopCliente
    Dim Achou As Boolean
    Dim mtzUsu�riosLogados() As String
    Dim res As Boolean
    Dim strSql As String
    Dim rstcomparacao As New ADODB.Recordset
    Dim datValidade As Date
    
    strSql = "SELECT PKCodigo_TBUsuario,FKCodigo_TBEmpresa,DFNome_TBUsuario,DFSenha_TBUsuario,DFNivel_TBUsuario,IXData_validade_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & txtLogin(0).Text & "'"
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstcomparacao, "Otica", Me)
    
    Select Case Index
        Case 0
            Select Case Op��oDeLoginSelecionada
                Case Op��esLogin.OPL_AtivarLogin
                    res = ReativarLogin(Trim(txtLogin(0).Text), Trim(txtLogin(1).Text), "Otica", Me)
                    
                    If res = True Then
                       'Verificando a validade da conta do usu�rio
                       If datValidade <= Format(Now, "YYYYMMDD") Then
                          MsgBox "Seu usu�rio n�o � v�lido, sua conta expirou em: " & Format(rstcomparacao!IXData_validade_TBUsuario, "DD/MM/YYYY") & " !Verifique com o administrador do Sistema.", vbInformation, "Logicx"
                          Set rstcomparacao = Nothing
                          Exit Sub
                       End If
                       'Informa��es Variaveis para o log
                       log.Usuario = txtLogin(0).Text
                       log.Evento = "Novo Login"
                       log.Descricao = "Reativando o login do usu�rio da �rea de trabalho"
                       log.Tipo = 1
                       log.Data = Date
                       log.Hora = Format(Now, "hh:mm:ss")
                       
                       'Gravando o log
                       log.Gravar_log "Otica", Me
                       'Gravando as inf do usu�rio no registro para conting�ncia
                       Movimentacoes.Grava_Contingencia_Acessibilidade strEstacao, txtLogin(0).Text, Str(rstcomparacao!PKCodigo_TBUsuario), Str(rstcomparacao!FKCodigo_TBEmpresa), "Otica"
                       Unload Me
                    End If
                Case OPL_DesativarLogin
                    DesativarLogin Trim(txtLogin(0).Text), Trim(txtLogin(1).Text), "Otica", Me
                    'Informa��es Variaveis para o log
                    log.Usuario = txtLogin(0).Text
                    log.Evento = "Novo Login"
                    log.Descricao = "Desativando o usu�rio da �rea de trabalho"
                    log.Tipo = 1
                    log.Data = Date
                    log.Hora = Format(Now, "hh:mm:ss")
                    'Gravando o log
                    log.Gravar_log "Otica", Me
                    txtLogin(0).Text = Empty
                    txtLogin(1).Text = Empty
                Case Op��esLogin.OPL_LogOff
                    ValidarUsu�rioSenha = True
                    If ValidarUsu�rioSenha = True Then
                        LogOff Trim(txtLogin(0).Text), Trim(txtLogin(1).Text), "Otica", Me
                        'Atualizar Grid
                        'ExibirLoginOp��es True
                        
                        'Informa��es Variaveis para o log
                        log.Usuario = txtLogin(0).Text
                        log.Evento = "Novo Login"
                        log.Descricao = "Logoff de usu�rio da �rea de trabalho"
                        log.Tipo = 1
                        log.Data = Date
                        log.Hora = Format(Now, "hh:mm:ss")
                         
                        'Gravando o log
                        log.Gravar_log "Otica", Me
                        Unload Me
                    Else
                        MsgBox "Usu�rio/Senha Inv�lido"
                        txtLogin(0).Text = ""
                        txtLogin(1).Text = ""
                        txtLogin(0).SetFocus
                    End If
                   
                Case Op��esLogin.OPL_NovoLogin
                    Dim strUsuario As String
                    Dim strSenha As String
                    Dim intCodigo As Integer
                    Dim strEmpresa As String
                    Dim intNivel_usuario As Integer
                    
                    strUsuario = UCase(rstcomparacao!DFNome_TBUsuario)
                    strSenha = UCase(rstcomparacao!DFSenha_TBUsuario)
                    
                    intCodigo = rstcomparacao!PKCodigo_TBUsuario
                    strEmpresa = rstcomparacao!FKCodigo_TBEmpresa
                    intNivel_usuario = rstcomparacao!DFNivel_TBUsuario
                    If strUsuario <> txtLogin(0).Text Then
                       MsgBox "Usu�rio n�o cadastrado!", vbCritical, "Logicx"
                       txtLogin(0).SetFocus
                    Else
                      If strSenha <> txtLogin(1).Text Then
                         MsgBox "Senha n�o confere!", vbCritical, "Logicx"
                         txtLogin(1).SetFocus
                      Else
                        ValidarUsu�rioSenha = True
                        NovoLogin Trim(txtLogin(0).Text), Trim(txtLogin(1).Text), Str(intCodigo), strEmpresa, intNivel_usuario
                        
                        'Informa��es Variaveis para o log
                        log.Usuario = txtLogin(0).Text
                        log.Evento = "Novo Login"
                        log.Descricao = "Novo Usu�rio na �rea de trabalho"
                        log.Tipo = 1
                        log.Data = Date
                        log.Hora = Format(Now, "hh:mm:ss")
                         
                        'Gravando o log
                        log.Gravar_log "Otica", Me
                        
                        'Gravando as inf do usu�rio no registro para conting�ncia
                        Movimentacoes.Grava_Contingencia_Acessibilidade strEstacao, txtLogin(0).Text, Str(rstcomparacao!PKCodigo_TBUsuario), Str(rstcomparacao!FKCodigo_TBEmpresa), "Otica"
                        
                        Unload Me
                      End If
                    End If

                    Set rstcomparacao = Nothing
            End Select
        Case 1
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    'Setando e passando a esta��o local para a mensagem do intercomunicador
    Set FCRegistro = New DLLSystemManager.Registro
    strEstacao = FCRegistro.WinRegLerSequ�ncia("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
    
    'Informa��es Constantes para o log
    log.Estacao = strEstacao
    log.Programa = "Admin do Sistema"

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub lstUsu�riosLogados_Click()
    txtLogin(0).Text = Replace(lstUsu�riosLogados.List(lstUsu�riosLogados.ListIndex), " (Atual)", "")
    txtLogin(1).SetFocus
End Sub

Private Sub opt_Click(Index As Integer)
    Op��oDeLoginSelecionada = Index
    txtLogin(0).SetFocus
End Sub

Private Sub txtLogin_LostFocus(Index As Integer)
    txtLogin(Index).Text = UCase(txtLogin(Index).Text)
End Sub
