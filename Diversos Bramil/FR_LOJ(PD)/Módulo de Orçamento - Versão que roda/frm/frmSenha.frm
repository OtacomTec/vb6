VERSION 5.00
Begin VB.Form frmSenha 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Senha"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSenha_Liberacao 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   315
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   870
      Width           =   5145
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   5850
      Picture         =   "frmSenha.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   0
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   810
      Width           =   5355
   End
   Begin VB.Line Line5 
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   1700
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1700
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   0
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      X1              =   6000
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   645
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   930
      Width           =   5355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   2040
      X2              =   0
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Senha de Liberação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   690
      TabIndex        =   2
      Top             =   180
      Width           =   3480
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim booIntegracao_online As Boolean
Dim rstAplicacao As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
    
       strSql = "SELECT PKCodigo_TBOperadores_ecf,DFNivel_TBOperadores_ecf FROM TBOperadores_ecf "

       If Len(txtSenha_Liberacao.Text) > 10 Then
          strSql = strSql + "WHERE DFNumero_cartao_TBOperadores_ecf = '" & txtSenha_Liberacao.Text & "' "
       Else
          strSql = strSql + "WHERE DFSenha_TBOperadores_ecf = '" & txtSenha_Liberacao.Text & "' "
       End If
       
       strSql = strSql + "AND DFNivel_TBOperadores_ecf <> 1"
       
       If booIntegracao_online = True Then
          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
       Else
          Select_geral strSql, "BDPDV", rstAplicacao, "PDV", Me
       End If
       
       If rstAplicacao.RecordCount <> 0 Then
          frmTela_Venda.intNivel = rstAplicacao.Fields("DFNivel_TBOperadores_ecf")
       Else
          MsgBox "SENHA OU NÚMERO DE CARTÃO INVÁLIDO!", vbCritical, "Only Tech"
          
          frmTela_Venda.intNivel = frmTela_Venda.intNivel_Operador
       End If
       
       Set rstAplicacao = Nothing
       
       Unload Me
       
       frmTela_Venda.txtCodigo_Produto.SetFocus
       
    End If
    
End Sub

Private Sub Form_Load()

    booIntegracao_online = frmTela_Venda.booIntegracao_Retaguarda

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a saida com ESC
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
