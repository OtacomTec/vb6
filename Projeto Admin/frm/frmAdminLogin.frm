VERSION 5.00
Begin VB.Form frmAdminLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   5130
   ClientTop       =   3930
   ClientWidth     =   5235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAdminLogin.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLogin 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2610
      Width           =   1875
   End
   Begin VB.TextBox txtLogin 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   2040
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Configurador do Sistema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2940
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   720
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao_login As New DLLConexao_Sistema.Conexao
Dim rstcomparacao As New ADODB.Recordset
Dim strSql As String
Dim strSenha As String

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
   End If
   If KeyAscii = 27 Then
        End
   End If
End Sub
Private Sub txtLogin_LostFocus(Index As Integer)
   If Trim(txtLogin(Index).Text) = "" Then Exit Sub
    
    Select Case Index
        Case 0 'Usuário
        
            strSql = "SELECT DFNome_TBUsuario,DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & txtLogin(0).Text & "'"
            conexao_login.Abrir_conexao ("PDV")
            Call Movimentacoes.Select_geral(strSql, "BDSupervisor", rstcomparacao, "PDV", Me)
            
            If rstcomparacao.EOF = True And rstcomparacao.BOF = True Then
               MsgBox "Usuário não cadastrado", vbCritical, "Logicx"
               txtLogin(0).SetFocus
            Else
               strSenha = rstcomparacao!DFSenha_TBUsuario
            End If
            
            Set rstcomparacao = Nothing
            conexao_login.Fechar_conexao
            
        Case 1 'Senha
            If Trim(txtLogin(0).Text) = "" Then
                txtLogin(0).SetFocus
                Exit Sub
            Else
                If strSenha = txtLogin(1).Text Then
                    ValidarUsuárioSenha = True
                    frmAdminMDI.Show
                    'Adicionando um Novo Componente AplicativoUsuário
                    NovoLogin Trim(Me.txtLogin(0).Text), Trim(Me.txtLogin(1).Text)
                    Unload Me
                    KeyAscii = 0
                Else
                    MsgBox "Senha Inválida", vbCritical, "Logicx"
                    txtLogin(0).Text = ""
                    txtLogin(1).Text = ""
                    txtLogin(0).SetFocus
                End If
                
            End If
    End Select
End Sub
