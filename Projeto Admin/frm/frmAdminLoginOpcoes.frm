VERSION 5.00
Begin VB.Form frmAdminLoginOpcoes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2025
   ClientLeft      =   4155
   ClientTop       =   4155
   ClientWidth     =   4485
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
      TabIndex        =   4
      ToolTipText     =   "Retirar o meu Login da Se��o Ativa"
      Top             =   780
      Width           =   1455
   End
   Begin VB.OptionButton opt 
      Caption         =   "&Ativar Login"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
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
      TabIndex        =   1
      ToolTipText     =   "Digite a Senha do Usu�rio"
      Top             =   1605
      Width           =   1335
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Index           =   0
      Left            =   750
      TabIndex        =   0
      ToolTipText     =   "Digite o Nome do Usu�rio (Nome Resumido)"
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   315
      Index           =   1
      Left            =   2370
      TabIndex        =   6
      ToolTipText     =   "Cancela e Fecha esta Janela"
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Height          =   315
      Index           =   0
      Left            =   3420
      TabIndex        =   7
      ToolTipText     =   "Confirma as op��es selecionadas"
      Top             =   1620
      Width           =   975
   End
   Begin VB.ListBox lstUsu�riosLogados 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   2370
      TabIndex        =   9
      ToolTipText     =   "Lista os Usu�rios Logados no Sistema"
      Top             =   240
      Width           =   2025
   End
   Begin VB.OptionButton opt 
      Caption         =   "&Logoff"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Finalizar todas as Se��es do meu Login"
      Top             =   1020
      Width           =   825
   End
   Begin VB.OptionButton opt 
      Caption         =   "&Novo Login"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
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
      TabIndex        =   8
      Top             =   30
      Width           =   2025
   End
End
Attribute VB_Name = "frmAdminLoginOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
    Select Case Index
        Case 0
            Select Case Op��oDeLoginSelecionada
                Case Op��esLogin.OPL_AtivarLogin
                    res = ReativarLogin(Trim(txtLogin(0).Text), Trim(txtLogin(1).Text))
                    If res = True Then Unload Me
                    
                Case OPL_DesativarLogin
                    DesativarLogin Trim(txtLogin(0).Text), Trim(txtLogin(1).Text)
                    
                Case Op��esLogin.OPL_LogOff
                    ValidarUsu�rioSenha = True
                    If ValidarUsu�rioSenha = True Then
                        LogOff Trim(txtLogin(0).Text), Trim(txtLogin(1).Text)
                        'Atualizar Grid
                        'ExibirLoginOp��es True
                        Unload Me
                    Else
                        MsgBox "Usu�rio/Senha Inv�lido"
                        txtLogin(0).Text = ""
                        txtLogin(1).Text = ""
                        txtLogin(0).SetFocus
                    End If
                   
                Case Op��esLogin.OPL_NovoLogin
                    Dim strSql As String
                    Dim rstcomparacao As New ADODB.Recordset
                    Dim conexao_login As New DLLConexao_Sistema.Conexao

                    strSql = "SELECT DFNome_TBUsuario,DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & txtLogin(0).Text & "'"
                    conexao_login.Abrir_conexao ("PDV")

                    Call Movimentacoes.Select_geral(strSql, "BDSupervisor", rstcomparacao, "PDV", Me)

                    If rstcomparacao!DFNome_TBUsuario <> txtLogin(0).Text Then
                       MsgBox "Usu�rio n�o cadastrado!", vbCritical, "Logicx"
                       txtLogin(0).SetFocus
                    Else
                      If rstcomparacao!DFSenha_TBUsuario <> txtLogin(1).Text Then
                         MsgBox "Senha n�o confere!", vbCritical, "Logicx"
                         txtLogin(1).SetFocus
                      Else
                        ValidarUsu�rioSenha = True
                        NovoLogin Trim(txtLogin(0).Text), Trim(txtLogin(1).Text)
                        Unload Me
                      End If
                    End If

                    Set rstcomparacao = Nothing
                    conexao_login.Fechar_conexao
            End Select
        Case 1
            Unload Me
    End Select
    
End Sub

Private Sub lstUsu�riosLogados_Click()
    txtLogin(0).Text = Replace(lstUsu�riosLogados.List(lstUsu�riosLogados.ListIndex), " (Atual)", "")
    txtLogin(1).SetFocus
    
End Sub

Private Sub opt_Click(Index As Integer)
    Op��oDeLoginSelecionada = Index
    txtLogin(0).SetFocus
End Sub
