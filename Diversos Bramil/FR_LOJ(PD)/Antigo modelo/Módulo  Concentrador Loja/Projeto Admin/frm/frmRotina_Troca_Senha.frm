VERSION 5.00
Begin VB.Form frmRotina_Troca_Senha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca de Senha"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRotina_Troca_Senha.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1290
      TabIndex        =   3
      Top             =   2430
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   90
      TabIndex        =   2
      Top             =   2430
      Width           =   1095
   End
   Begin VB.TextBox txtConfirmacao 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   90
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1860
      Width           =   2265
   End
   Begin VB.TextBox txtNova_Senha 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   90
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1140
      Width           =   2265
   End
   Begin VB.TextBox txtSenha_Atual 
      Enabled         =   0   'False
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   90
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   420
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Confirmação"
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   1590
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nova Senha"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   870
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha Atual"
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "frmRotina_Troca_Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                     '
' Módulo.................: Admin                                                          '
' Objetivo...............: MDI Principal                                                  '
' Data de Criação........: 23/07/2004                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.: 22/07/2004                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim CNconexao As New DLLConexao_Sistema.Conexao
Public strUsuario As String
Public strEstacao As String
Public dblCodigo_Usuario As Double
Dim IntData_Expiracao As Integer
Dim log As New DLLSystemManager.log
Dim intCodigo_usuario As Double

Private Sub cmdOk_Click()

    'Conexão Log Morto
    CNconexao.Initial_Catalog = "BDRetaguarda"
    CNconexao.Abrir_conexao ("Otica")
    
    Dim dtpProxima_Data As Date
    
    dtpProxima_Data = DateAdd("d", IntData_Expiracao, Now)
    
    CNconexao.CNconexao.BeginTrans
    
    strSql = Empty
    strSql = "UPDATE TBUsuario SET DFSenha_TBUsuario = '" & txtNova_Senha.Text & "',DFData_ultima_troca_senha_TBUsuario = '" & Format(Now, "YYYYMMDD") & "',DFProxima_troca_senha_TBUsuario = '" & Format(dtpProxima_Data, "YYYYMMDD") & "' WHERE PKCodigo_TBUsuario = " & intCodigo_usuario & ""
             
    'Incluindo Registro no LOG MORTO
    CNconexao.CNconexao.Execute strSql
     
    'Informações Constantes para o log
    log.Usuario = strUsuario
    log.Programa = "Login do Sistema"
    log.Estacao = strEstacao
    
    'Informações Variaveis para o log
    log.Evento = "Login do Sistema"
    log.Descricao = "Rotina de troca de senha do usuário."
    log.Tipo = 2
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
        
    CNconexao.CNconexao.CommitTrans
    CNconexao.Fechar_conexao
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    MsgBox "Troca de senha efetuada com sucesso!Favor Reinicie a aplicação", vbInformation, "Logicx"
    
    Unload Me
    
    Exit Sub
    
Erro:

    'ROLLBACK'S
    CNconexao.CNconexao.RollbackTrans
    
    MsgBox "Ocorreu um erro : " & Err.Description & "", vbCritical, "Logicx"
    
    End
    
    Exit Sub
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
   End If
   If KeyAscii = 27 Then
        End
   End If
End Sub
Private Sub Form_Load()

    Dim rstUsuarios As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT PKCodigo_TBUsuario,DFSenha_TBUsuario,DFPrazo_expira_senha_TBUsuario FROM TBUsuario WHERE PKCodigo_TBUsuario = " & frmAdminLogin.intCodigo_usuario & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstUsuarios, "Otica", Me
    
    IntData_Expiracao = rstUsuarios!DFPrazo_expira_senha_TBUsuario
    intCodigo_usuario = rstUsuarios!PKCodigo_TBUsuario
    
    txtSenha_Atual.Text = rstUsuarios!DFSenha_TBUsuario
    
    Set rstUsuarios = Nothing
    
End Sub

Private Sub txtConfirmacao_LostFocus()
    If txtConfirmacao.Text <> Me.txtNova_Senha.Text And txtConfirmacao.Text <> "" Then
       MsgBox "A confirmação de senha que vc digitou difere da nova senha digitada!Favor Redigite.", vbInformation, "Logicx"
       txtConfirmacao.SetFocus
       Exit Sub
    End If
End Sub

Private Sub txtNova_Senha_LostFocus()
    If Me.txtSenha_Atual.Text = Me.txtNova_Senha.Text Then
       MsgBox "A senha que vc acaba de digitar é igual a atual!Favor Redigite.", vbInformation, "Logicx"
       txtNova_Senha.SetFocus
       Exit Sub
    End If
End Sub
