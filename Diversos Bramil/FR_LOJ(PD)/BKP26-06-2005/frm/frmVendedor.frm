VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVendedor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Vendedor"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   Icon            =   "frmVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5970
      ScaleHeight     =   615
      ScaleWidth      =   255
      TabIndex        =   0
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
      TabIndex        =   3
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
      TabIndex        =   1
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
      TabIndex        =   2
      Top             =   1350
      Width           =   5955
   End
   Begin MSDataListLib.DataCombo dtcVendedor 
      Height          =   570
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Finalizadora"
      Top             =   810
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1005
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
      ForeColor       =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   ">>"
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
      TabIndex        =   6
      Top             =   1890
      Width           =   1635
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
      Caption         =   "Vendedor:"
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
      TabIndex        =   5
      Top             =   210
      Width           =   1470
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
Attribute VB_Name = "frmVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String

Private Sub cmdOK_Click()
    frmFechamento_Cupom.lngCodigo_vendedor = dtcVendedor.BoundText
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.dtcVendedor.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
  strSql = Empty
  strSql = "SELECT IXCodigo_TBVendedor,DFNome_TBVendedor FROM TBVendedor"
  Call Movimentacoes.Movimenta_DataCombo("IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSql, "BDRetaguarda", "Otica", Me)
  
End Sub
