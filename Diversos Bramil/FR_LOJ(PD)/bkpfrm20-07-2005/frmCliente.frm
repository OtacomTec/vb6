VERSION 5.00
Begin VB.Form frmCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cliente"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   Icon            =   "frmCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCliente 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   270
      MaxLength       =   14
      TabIndex        =   0
      Top             =   780
      Width           =   5955
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5970
      ScaleHeight     =   615
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   810
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   240
      ScaleHeight     =   75
      ScaleWidth      =   5955
      TabIndex        =   5
      Top             =   750
      Width           =   5955
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   780
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   240
      ScaleHeight     =   75
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   1350
      Width           =   5955
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1890
      Width           =   1635
   End
   Begin VB.Label lblCliente 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   645
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   1950
      Width           =   1665
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   660
      Width           =   6165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente Especial (Cartão):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   3675
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   915
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   6135
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String

Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCliente_LostFocus()
    Dim rstCliente As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT IXCodigo_TBCliente,PKId_TBCliente,DFNome_TBCliente FROM TBCliente WHERE DFNumero_contrato_TBCliente = " & Me.txtCliente.Text & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCliente, "Otica", Me
    
    If rstCliente.BOF = True And rstCliente.EOF = True Then
       MsgBox "Cliente não cadastrado!Verifique.", vbCritical, "Only Tech"
       Me.txtCliente.Text = Empty
       Me.txtCliente.SetFocus
       Set rstCliente = Nothing
       Exit Sub
    Else
       frmFechamento_Cupom.Cod_Cliente = rstCliente!IXCodigo_TBCliente
       Me.lblCliente.Caption = rstCliente!IXCodigo_TBCliente & "-" & rstCliente!DFNome_TBCliente
       Set rstCliente = Nothing
    End If

End Sub
