VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTela_Venda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSCommLib.MSComm mscPDV 
      Left            =   5880
      Top             =   7590
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtData_Operacao 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   9390
      MaxLength       =   40
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8610
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   11490
      ScaleHeight     =   4665
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   720
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5340
      Top             =   7650
   End
   Begin VB.TextBox txtNumero_loja 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
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
      Height          =   345
      Left            =   5100
      MaxLength       =   40
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txtVersao_software 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
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
      Height          =   345
      Left            =   10080
      MaxLength       =   40
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   30
      Width           =   1845
   End
   Begin VB.TextBox txtNumero_Nome_Operadora 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
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
      Height          =   345
      Left            =   6030
      MaxLength       =   40
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   30
      Width           =   4035
   End
   Begin VB.TextBox txtNumero_check_out 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
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
      Height          =   345
      Left            =   3540
      MaxLength       =   40
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   30
      Width           =   1545
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   480
      Picture         =   "frmTela_Venda.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   1935
      TabIndex        =   15
      Top             =   7530
      Width           =   1935
   End
   Begin VB.TextBox txtCodigo_Produto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   600
      MaxLength       =   14
      TabIndex        =   0
      Top             =   3600
      Width           =   3225
   End
   Begin VB.TextBox txtData_Hora 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
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
      Height          =   345
      Left            =   90
      MaxLength       =   40
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   30
      Width           =   3435
   End
   Begin VB.TextBox txtPreco_total_cupom 
      Alignment       =   1  'Right Justify
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
      Height          =   615
      Left            =   8430
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7440
      Width           =   3195
   End
   Begin VB.TextBox txtDescricao_Produto 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   630
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4770
      Width           =   6285
   End
   Begin VB.TextBox txtQuantidade_Produto 
      Alignment       =   1  'Right Justify
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
      Height          =   615
      Left            =   660
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6450
      Width           =   2625
   End
   Begin VB.TextBox txtPreco_Total 
      Alignment       =   1  'Right Justify
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
      Height          =   615
      Left            =   8310
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6450
      Width           =   3345
   End
   Begin VB.TextBox txtPreco_Unitario 
      Alignment       =   1  'Right Justify
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
      Height          =   615
      Left            =   4230
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6450
      Width           =   3165
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   30
      MaxLength       =   40
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8610
      Width           =   9345
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HflexGrid 
      DragMode        =   1  'Automatic
      Height          =   4575
      Left            =   7290
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   780
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   8454143
      BackColorFixed  =   8454143
      BackColorBkg    =   8454143
      BackColorUnpopulated=   8454143
      GridColorFixed  =   8454143
      GridColorUnpopulated=   8454143
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imgInd_pouco_papel 
      Height          =   255
      Left            =   11130
      Picture         =   "frmTela_Venda.frx":340E
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   23
      Top             =   8640
      Width           =   255
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   11670
      Shape           =   3  'Circle
      Top             =   8640
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   7320
      X2              =   11400
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Image imgLogo_Empresa 
      Height          =   2055
      Left            =   540
      Stretch         =   -1  'True
      Top             =   750
      Width           =   2955
   End
   Begin VB.Image imgProduto 
      Height          =   3765
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   750
      Width           =   2955
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   5085
      Left            =   7170
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   4725
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   5025
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   4665
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   8190
      Shape           =   4  'Rounded Rectangle
      Top             =   7290
      Width           =   3645
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   8190
      Shape           =   4  'Rounded Rectangle
      Top             =   6270
      Width           =   3675
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   4050
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   3465
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   570
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   2835
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   6345
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   8100
      Shape           =   4  'Rounded Rectangle
      Top             =   7410
      Width           =   3585
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   8100
      Shape           =   4  'Rounded Rectangle
      Top             =   6390
      Width           =   3585
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   6390
      Width           =   3435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quantidade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5880
      Width           =   1620
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   6390
      Width           =   2835
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   915
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   4740
      Width           =   6315
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   3630
      Width           =   3285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "C�digo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   3120
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3570
      Width           =   3285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Descri��o"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   4260
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3570
      TabIndex        =   10
      Top             =   6540
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre�o Unit�rio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4110
      TabIndex        =   9
      Top             =   5880
      Width           =   1980
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre�o Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8220
      TabIndex        =   8
      Top             =   5880
      Width           =   1785
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7230
      TabIndex        =   7
      Top             =   7860
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7770
      TabIndex        =   6
      Top             =   6540
      Width           =   180
   End
End
Attribute VB_Name = "frmTela_Venda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' M�dulo.................: Frente de Loja                                                 '
' Objetivo...............: Tela de Vendas                                                 '
' Data de Cria��o........: 04/01/2005                                                     '
' Equipe Respons�vel.....: Giordano Vilela,Marcos Bai�o,Alex Bai�o,Rafael Gomes, S�rgio   '
' �ltima Manuten��o......:                                                                '
' Data �ltima manuten��o.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim rstInf_Produtos As New ADODB.Recordset
Dim booImpressora_lacrada As Boolean
Dim booAcionado_Fechamento_cupom As Boolean
Dim rstEmpresa As New ADODB.Recordset
Public strCodigo_Operador As String
Public strOperador As String
Public strEmpresa_Operador As String
Public booInterrompe_venda As Boolean
Public strDigito_Peso_Variavel As String
Public strPDV As String
Public booComanda As Boolean
Public strNumero_Comanda As String
Public strVendedor_Comanda As String
Public booIntegracao_Retaguarda As Boolean
Public intImpressoes_suportadas As Integer
Public intIP_Concentrador As String
Dim booICMS_cadastrado As Boolean
Public booPreco_online As Boolean
Public booCupom_fiscal As Boolean
Public strImpresora As String
Public booConsulta As Boolean
Public booComissao_vendedor As Boolean
Public strTipo_quantidade As String
Public strCasas_Decimais As String
Public strTipo_desconto As String
Public dtpData_operacao As Date
Dim intAliquota_ICMS As Integer
Dim strAliquota_imp_fisc As String
Public booLeitor_serial As Boolean
Public strCom_leitor_serial As String
Public intTipo_imp_orcamento As Integer
Public booGaveta_integrada As Boolean
Public intFinalizadora_sangria As Integer
Public booPreco_peso_balanca_TBParametros_ecf As Boolean
Public strCaminho_impComum As String
Public intPerfil_ECF As Integer
Public booEncerrante As Boolean
Public booErro_processamento_impressora As Boolean
Public booFinaliza_direto As Boolean
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intRetorno As Integer
    
    'Verifica se foi precionado F2 e abre tela de consulta
    If KeyCode = 113 Then
       Me.txtStatus.Text = "Consultando o Item..."
       booConsulta = True
       frmConsulta_Produto.Show (1)
       Me.txtStatus.Text = "Aguardando Venda...."
       booConsulta = False
    End If
    
    'Verifica se foi preciona do F3 e abre novo cupom
    If KeyCode = 114 Then
        Call Abre_Cupom
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F4 e Finaliza a Compra
    If KeyCode = 115 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Finalizando a compra..."
       If Me.txtPreco_total_cupom.Text > "" Then
          Call Finaliza_Cupom
       End If
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
    End If
    '
    'Verifica se foi precionado F5 e Cancela o �ltimo item
    If KeyCode = 116 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Cancelando �ltimo item..."
        
       If Me.HflexGrid.Rows - 3 > 0 Then
            
           'Subtraindo totalizador..........
           If strCasas_Decimais = 2 Then
              txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) - CDbl(Me.HflexGrid.Text), "#,###0.00")
           End If
            
           If strCasas_Decimais = 3 Then
              txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) - CDbl(Me.HflexGrid.Text), "#,###0.000")
           End If
            
           Me.HflexGrid.RemoveItem (Me.HflexGrid.Rows - 1)
         
           If Me.HflexGrid.Rows - 2 > 0 Then Me.HflexGrid.RemoveItem (Me.HflexGrid.Rows - 1)
        End If
        
        If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
           Call Cancela_item_cupom
        End If
        
        txtCodigo_Produto.SetFocus
        Me.txtStatus.Text = "Aguardando Venda...."
    End If
    
    'Verifica se foi precionado F6 e Cancela o cupom
    If KeyCode = 117 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Cancelando o cupom..."
       'Verifica se confirma a opera��o
       intRetorno = MsgBox("Se confirmar a opera��o este cupom ser� cancelado.Confirma a opera��o?", vbYesNo, "Only Tech")
       If intRetorno = 7 Then
          Exit Sub
       End If
       
       Call Cancela_cupom
       
       'Limpando o cupom em processamento
       Call Limpa_Processa_Cupom
       
       Call Limpa_Tela
       
       Me.HflexGrid.Clear
       Me.HflexGrid.ClearStructure
       Me.txtPreco_total_cupom.Text = ""
       Me.HflexGrid.Rows = 2
       Me.txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
       
    End If
    
    'Verifica se foi precionado F7 e Finaliza o Operador
    If KeyCode = 118 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Finalizando o Operador..."
       'Posto de Gasolina
       If frmTela_Venda.intPerfil_ECF = 2 Then
           frmFechamento_Encerrantes.Show (1)
       End If
       frmFechamento_Operador.Show (1)
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
    End If
    
    'Verifica se foi precionado F8 e Aciona Sangria
    If KeyCode = 119 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Acionando a Sangria..."
       frmSangria.Show (1)
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
    End If
    
    'Verifica se foi precionado F9 e libera a digita��o da quantidade
    If KeyCode = 120 Then
       'Libera para digita��o da quantidade
       Me.txtStatus.Text = "Digite a quantidade..."
       txtQuantidade_Produto.TabStop = True
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
    End If
    
    'Verifica se foi precionado F10 e libera a digita��o do pre�o
    If KeyCode = 121 Then
       'Libera para digita��o da qunantidade
       Me.txtStatus.Text = "Digite o valor do pre�o..."
       Me.txtPreco_Unitario.TabStop = True
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
    End If
    
    'Verifica se foi precionado F11 e Finaliza o dia
    If KeyCode = 122 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Finalizando o Dia..."
       'Posto de Gasolina
       If frmTela_Venda.intPerfil_ECF = 2 Then
          frmFechamento_Encerrantes.Show (1)
       End If
       frmFechamento_dia.Show (1)
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
       End
    End If
    
    'Verifica se foi precionado F12 e abre a tela de comanda
    If KeyCode = 123 Then
       booAcionado_Fechamento_cupom = True
       Me.txtStatus.Text = "Comanda..."
       frmComanda.Show (1)
       txtCodigo_Produto.SetFocus
       Me.txtStatus.Text = "Aguardando Venda...."
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    'Alterna cupom com or�amento
    If KeyAscii = 19 Then
       'Amarelo
       If Me.txtData_Hora.BackColor = &H80FFFF Then
          Me.HflexGrid.Col = 1
          Me.HflexGrid.Row = 0
          If Me.HflexGrid.Text = "" Then
             Me.txtData_Hora.BackColor = &HC0C0C0
             Me.txtNumero_check_out.BackColor = &HC0C0C0
             Me.txtNumero_loja.BackColor = &HC0C0C0
             Me.txtNumero_Nome_Operadora.BackColor = &HC0C0C0
             Me.txtStatus.BackColor = &HC0C0C0
             Me.txtVersao_software.BackColor = &HC0C0C0
             Me.txtData_Operacao.BackColor = &HC0C0C0
             booCupom_fiscal = False
             Me.txtCodigo_Produto.SetFocus
          End If
       Else
            'Cinza
            If Me.txtData_Hora.BackColor = &HC0C0C0 Then
               Me.HflexGrid.Col = 1
               Me.HflexGrid.Row = 0
               If Me.HflexGrid.Text = "" Then
                  Me.txtData_Hora.BackColor = &H80FFFF
                  Me.txtNumero_check_out.BackColor = &H80FFFF
                  Me.txtNumero_loja.BackColor = &H80FFFF
                  Me.txtNumero_Nome_Operadora.BackColor = &H80FFFF
                  Me.txtStatus.BackColor = &H80FFFF
                  Me.txtVersao_software.BackColor = &H80FFFF
                  Me.txtData_Operacao.BackColor = &H80FFFF
                  booCupom_fiscal = True
                  Me.txtCodigo_Produto.SetFocus
               End If
            End If
       End If
    End If
    
    
    'Se perfil for de posto e acionado o CTRL + E abre a op��o de venda por encerrante
    If KeyAscii = 5 And Me.intPerfil_ECF = 2 Then
       frmVenda_Encerrante.Show (1)
    End If
    
    'Tira um X da impressora - CTRL + X
    If KeyAscii = 24 Then
       Call Comandos_impressoras_fiscais.Leitura_x(strImpresora)
    End If
    
    'Abre a gaveta - CTRL + G
    If KeyAscii = 7 Then
       Call Comandos_impressoras_fiscais.Abrir_gaveta(strImpresora)
       txtCodigo_Produto.SetFocus
    End If
    
    'Configura��es da impressora - CTRL + B
    If KeyAscii = 2 Then
       Call Comandos_impressoras_fiscais.Configuracoes_impressora_fiscal(strImpresora)
       txtCodigo_Produto.SetFocus
    End If
    
    'Configura��es da impressora - CTRL + L
    If KeyAscii = 12 Then
       booConsulta = True
       frmConsulta_Cliente.Show (1)
       Me.txtStatus.Text = "Aguardando Venda...."
       booConsulta = False
    End If
    
End Sub

Private Sub Form_Load()
    
    'Indica se para o programa se a impressora est� lacrada ou n�o
    booImpressora_lacrada = True
    
    If booLeitor_serial = True Then
       mscPDV.CommPort = strCom_leitor_serial
       mscPDV.PortOpen = True
    End If
    
    strImpresora = Comandos_impressoras_fiscais.Fabricante_Bematech
    
    'Indica se este cupom recebeu carga ou n�o de uma comanda
    booComanda = False
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    'Data de opera��o
    txtData_Operacao.Text = "Dt.Op.: " & dtpData_operacao
    
    booCupom_fiscal = True
    
    If frmTela_Venda.booCupom_fiscal = True And intImpressoes_suportadas <> 1 Then
       Comandos_impressoras_fiscais.Abre_impressora_fiscal (strImpresora)
    End If
    
    'Consultas ---------------------------------------------------------------------------------------
    
    strSql = Empty
    strSql = "SELECT * FROM TBEmpresa Where PKCodigo_TBEmpresa = " & strEmpresa_Operador & ""
    
    If booIntegracao_Retaguarda = True Then
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstEmpresa, "Otica", Me
    Else
       Movimentacoes.Select_geral strSql, "BDPDV", rstEmpresa, "PDV", Me
    End If

    'Tratando quando n�o houver imagem gravada
    On Error GoTo carrega_imagem
    imgLogo_Empresa.Picture = LoadPicture(rstEmpresa!DFPath_logomarca_TBEmpresa)
    
carrega_imagem:

    If Err.Number <> 76 And Err.Number > 0 Then
       MsgBox "Ocorreu um erro n�:" & Err.Number & "-" & Err.Description
    End If
    
    On Error Resume Next
    
    If frmTela_Venda.booCupom_fiscal = False And strImpresora = "Bematech" And frmTela_Venda.intImpressoes_suportadas <> 2 Then
       strTipo_quantidade = "I"
       strCasas_Decimais = 2
       strTipo_desconto = "%"
    End If
    
    If frmTela_Venda.intImpressoes_suportadas <> 0 Then
        If frmTela_Venda.intTipo_imp_orcamento = 0 Then
            '-------------------------------------------------------------------------------------------------------
            'Abrindo Impressora n�o fiscal
            '-------------------------------------------------------------------------------------------------------
            Dim intPorta As Integer
            Dim strComunica As String
            
            ' Fecha a porta que est� aberta
            intPorta = FechaPorta()
            If intPorta <= 0 Then
               MsgBox "Problemas ao Fechar a Porta de Comunica��o com a imp. n�o fiscal.Reinicie a aplica��o", vbCritical, "Only Tech"
            End If
            ' Abre a porta de comunicacao com imp. n�o fiscal
            intPorta = IniciaPorta("LPT1")
            If intPorta <= 0 Then
               MsgBox "Problemas ao Abrir a Porta de Comunica��o com a imp. n�o fiscal.Reinicie a aplica��o", vbCritical, "Only Tech"
            End If
        End If
    End If
    
    txtStatus.Text = "Aguardando Venda....."
    
    Set rstEmpresa = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
        
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
       Retorno = Bematech_FI_FechaPortaSerial()
       Call VerificaRetornoImpressora("", "", "BemaFI32")
    End If
    
    If booLeitor_serial = True Then
       Me.mscPDV.PortOpen = False
    End If
    
    End
    
End Sub

Private Sub Timer1_Timer()
    txtData_Hora.Text = "Data: " & Format(Now, "DD/MM/YYYY") & " - " & Format(Now, "hh:mm:ss")
    
    If booLeitor_serial = True Then
        If Me.txtCodigo_Produto.Text = "" Then
           Me.txtCodigo_Produto.Text = mscPDV.Input
           If Me.txtCodigo_Produto.Text <> "" Then
              Call txtCodigo_Produto_LostFocus
           End If
        End If
    End If
    
End Sub

Private Sub txtCodigo_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Produto_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Public Sub txtCodigo_Produto_LostFocus()

    Dim strPreco_Peso_Parametro As String
    Dim strDigito_Produto_Digitado As String
    Dim strCodigo_Produto_Etiqueta As String
    Dim strPreco_Peso As String
    Dim strDecimal As String
    Dim strQuantidade As String
    Dim strTotal As String
    Dim strPreco_Tabela As String
    Dim strID_Produto As String

    If Me.txtCodigo_Produto.Text <> "" Then
       If IsNumeric(Me.txtCodigo_Produto.Text) = False Then
          txtCodigo_Produto.Text = Empty
          txtCodigo_Produto.SetFocus
          Exit Sub
       End If
    End If

    booICMS_cadastrado = True
    
    If booPreco_peso_balanca_TBParametros_ecf = False Then
       strPreco_Peso_Parametro = 0
    Else
       strPreco_Peso_Parametro = 1
    End If
    
    If txtCodigo_Produto.Text <> "" Then
        Me.txtStatus.Text = "Processando o item..."
        'C�digo Interno
        If Len(Me.txtCodigo_Produto.Text) < 7 Then
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
        
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco,DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & strEmpresa_Operador & ")" & _
                       "AND TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & " " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""

              Movimentacoes.Select_geral strSql, "BDRetaguarda", rstInf_Produtos, "Otica", Me
              
              If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                MsgBox "C�digo Interno n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                txtCodigo_Produto.Text = Empty
                txtCodigo_Produto.SetFocus
                Me.txtPreco_Unitario.TabStop = False
                Me.txtQuantidade_Produto.TabStop = False
                Set rstInf_Produtos = Nothing
                Exit Sub
              End If
            
              If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
                MsgBox "Pre�o do Item n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                txtCodigo_Produto.Text = Empty
                txtCodigo_Produto.SetFocus
                Me.txtPreco_Unitario.TabStop = False
                Me.txtQuantidade_Produto.TabStop = False
                Set rstInf_Produtos = Nothing
                Exit Sub
               End If
              
              txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
              If frmTela_Venda.booEncerrante = False Then
                 txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
              End If
              txtQuantidade_Produto.Text = 1
            
            Else
              'Acessando o Pdv
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPreco_venda_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "WHERE TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & " " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""

              Movimentacoes.Select_geral strSql, "BDPDV", rstInf_Produtos, "PDV", Me
              
              If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                 MsgBox "C�digo Interno n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                 txtCodigo_Produto.Text = Empty
                 txtCodigo_Produto.SetFocus
                 Me.txtPreco_Unitario.TabStop = False
                 Me.txtQuantidade_Produto.TabStop = False
                 Set rstInf_Produtos = Nothing
                 Exit Sub
              End If
            
              If rstInf_Produtos!DFPreco_venda_TBProduto <= 0 Or IsNull(rstInf_Produtos!DFPreco_venda_TBProduto) Then
                 MsgBox "Pre�o do Item n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                 txtCodigo_Produto.Text = Empty
                 txtCodigo_Produto.SetFocus
                 Me.txtPreco_Unitario.TabStop = False
                 Me.txtQuantidade_Produto.TabStop = False
                 Set rstInf_Produtos = Nothing
                 Exit Sub
              End If
              
              txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
              If frmTela_Venda.booEncerrante = False Then
                 txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_venda_TBProduto
              End If
              txtQuantidade_Produto.Text = 1
            
            End If
            
            Call Verifica_ICMS
            
            If booICMS_cadastrado = False And frmTela_Venda.intImpressoes_suportadas <> 1 Then
                MsgBox "Aliquota de ICMS ou codigo ref. para impressora n�o cadastrada.Verifique!", vbCritical, "Only Tech"
                txtCodigo_Produto.Text = Empty
                Me.txtDescricao_Produto.Text = Empty
                Me.txtQuantidade_Produto.Text = Empty
                Me.txtPreco_Total.Text = Empty
                Me.txtPreco_Unitario.Text = Empty
                txtCodigo_Produto.SetFocus
                Me.txtPreco_Unitario.TabStop = False
                Me.txtQuantidade_Produto.TabStop = False
                Set rstInf_Produtos = Nothing
                Exit Sub
            End If
                        
            If Not IsNull(rstInf_Produtos!DFPath_imagem_TBProduto) Then
               Call carrega_imagem
            End If
            
            Call Reposicao

        End If
        
        'C�digo de barra
        If Len(Me.txtCodigo_Produto.Text) > 6 Then
            strDigito_Produto_Digitado = Left(txtCodigo_Produto.Text, 1)
            
            If strDigito_Peso_Variavel = strDigito_Produto_Digitado Then
            
                'Produto pes�vel e pre�o variavel
                strCodigo_Produto_Etiqueta = Mid(txtCodigo_Produto.Text, 2, 4)
                strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 7)
                
                If booIntegracao_Retaguarda = True And booPreco_online = True Then
                   strSql = Empty
                   strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco,DFPath_imagem_TBProduto " & _
                            "FROM TBProduto " & _
                            "INNER JOIN TBITENS_TABELA_PRECO " & _
                            "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                            "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & strEmpresa_Operador & ") " & _
                            "AND TBProduto.IXCodigo_TBProduto = " & strCodigo_Produto_Etiqueta & " " & _
                            "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                   Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstInf_Produtos, "Otica", Me)

                   If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                      MsgBox "C�digo de barra n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                      txtCodigo_Produto.Text = Empty
                      txtCodigo_Produto.SetFocus
                      Me.txtPreco_Unitario.TabStop = False
                      Me.txtQuantidade_Produto.TabStop = False
                      Set rstInf_Produtos = Nothing
                      Exit Sub
                   End If
                    
                   If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
                      MsgBox "Pre�o do Item n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                      txtCodigo_Produto.Text = Empty
                      txtCodigo_Produto.SetFocus
                      Me.txtPreco_Unitario.TabStop = False
                      Me.txtQuantidade_Produto.TabStop = False
                      Set rstInf_Produtos = Nothing
                      Exit Sub
                   End If
                   
                   txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
                   If frmTela_Venda.booEncerrante = False Then
                      txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
                   End If
                   txtQuantidade_Produto.Text = 1
                
                Else
                   strSql = Empty
                   strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,DFPreco_venda_TBProduto,DFPath_imagem_TBProduto " & _
                            "FROM TBProduto " & _
                            "WHERE TBProduto.IXCodigo_TBProduto = " & strCodigo_Produto_Etiqueta & " " & _
                            "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                   Call Movimentacoes.Select_geral(strSql, "BDPDV", rstInf_Produtos, "PDV", Me)

                   If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                      MsgBox "C�digo de barra n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                      txtCodigo_Produto.Text = Empty
                      txtCodigo_Produto.SetFocus
                      Me.txtPreco_Unitario.TabStop = False
                      Me.txtQuantidade_Produto.TabStop = False
                      Set rstInf_Produtos = Nothing
                      Exit Sub
                   End If
                    
                   If rstInf_Produtos!DFPreco_venda_TBProduto <= 0 Or IsNull(rstInf_Produtos!DFPreco_venda_TBProduto) Then
                      MsgBox "Pre�o do Item n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                      txtCodigo_Produto.Text = Empty
                      txtCodigo_Produto.SetFocus
                      Me.txtPreco_Unitario.TabStop = False
                      Me.txtQuantidade_Produto.TabStop = False
                      Set rstInf_Produtos = Nothing
                      Exit Sub
                   End If
                   
                   txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
                   If frmTela_Venda.booEncerrante = False Then
                      txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_venda_TBProduto
                   End If
                   txtQuantidade_Produto.Text = 1
                   
                End If
            
                Call Verifica_ICMS
                
                If booICMS_cadastrado = False And frmTela_Venda.intImpressoes_suportadas <> 1 Then
                    MsgBox "Aliquota de ICMS ou codigo ref. para impressora n�o cadastrada.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    Me.txtDescricao_Produto.Text = Empty
                    Me.txtQuantidade_Produto.Text = Empty
                    Me.txtPreco_Total.Text = Empty
                    Me.txtPreco_Unitario.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Me.txtPreco_Unitario.TabStop = False
                    Me.txtQuantidade_Produto.TabStop = False
                    Set rstInf_Produtos = Nothing
                    Exit Sub
                End If
                                
                If Not IsNull(rstInf_Produtos!DFPath_imagem_TBProduto) Then
                   Call carrega_imagem
                End If
            
                If strPreco_Peso_Parametro = 0 Then
                   strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 5)
                   strDecimal = Mid(txtCodigo_Produto.Text, 11, 2)
                   strPreco_Peso = strPreco_Peso & "," & strDecimal
                   If strCasas_Decimais = 2 Then
                      strPreco_Peso = Format(strPreco_Peso, "#,###0.00")
                   End If
                   If strCasas_Decimais = 3 Then
                      strPreco_Peso = Format(strPreco_Peso, "#,###0.000")
                   End If
                   strQuantidade = CDbl(strPreco_Peso) / CDbl(txtPreco_Unitario.Text)
                   strQuantidade = Format(strQuantidade, "#,###0.000")
                   txtQuantidade_Produto.Text = strQuantidade
                   txtPreco_Unitario.Text = CDbl(txtPreco_Unitario.Text)
                   strTotal = CDbl(txtPreco_Unitario.Text) * CDbl(strQuantidade)
                   If strCasas_Decimais = 2 Then
                      strTotal = Format(strTotal, "#,###0.00")
                   End If
                   If strCasas_Decimais = 3 Then
                      strTotal = Format(strTotal, "#,###0.000")
                   End If
                   txtPreco_Total.Text = strTotal
                Else
                   If strCasas_Decimais = 2 Then
                      strPreco_Peso = Format(strPreco_Peso, "#,###0.00")
                   End If
                   If strCasas_Decimais = 3 Then
                      strPreco_Peso = Format(strPreco_Peso, "#,###0.000")
                   End If
                   strTotal = strPreco_Peso * strPreco_Tabela
                   txtQuantidade_Produto.Text = strPreco_Peso
                   If strCasas_Decimais = 2 Then
                      txtPreco_Unitario.Text = Format(strPreco_Tabela, "#,###0.00")
                      txtPreco_Total.Text = Format(strTotal, "#,###0.00")
                   End If
                   If strCasas_Decimais = 3 Then
                      txtPreco_Unitario.Text = Format(strPreco_Tabela, "#,###0.000")
                      txtPreco_Total.Text = Format(strTotal, "#,###0.000")
                   End If
                End If
                
                Call Reposicao
              
            Else
                Dim rstID_Codautomacao As New ADODB.Recordset
                
                'Produto n�o pes�vel e pre�o n�o variavel
                'Query para verificar a validade do cod. de automa��o
                strSql = Empty
                strSql = "SELECT TBCodigo_barras.FKId_TBProduto " & _
                         "FROM TBCodigo_barras " & _
                         "WHERE IXCodigo_TBCodigo_barras = " & txtCodigo_Produto.Text & " "

                If booIntegracao_Retaguarda = True Then
                    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstID_Codautomacao, "Otica", Me
                Else
                    Movimentacoes.Select_geral strSql, "BDPDV", rstID_Codautomacao, "PDV", Me
                End If
                
                If rstID_Codautomacao.BOF = True And rstID_Codautomacao.EOF = True Then
                    MsgBox "C�digo de automa��o n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Me.txtPreco_Unitario.TabStop = False
                    Me.txtQuantidade_Produto.TabStop = False
                    Set rstID_Codautomacao = Nothing
                    Exit Sub
                Else
                    strID_Produto = rstID_Codautomacao!FKId_TBProduto
                End If
                
                Set rstID_Codautomacao = Nothing
                
                If booIntegracao_Retaguarda = True And booPreco_online = True Then
                
                    strSql = Empty
                    strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco, " & _
                             "TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                             "FROM TBItens_tabela_preco " & _
                             "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                             "WHERE FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & Me.strEmpresa_Operador & ") AND " & _
                             "FKId_TBProduto = " & strID_Produto & " " & _
                             "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstInf_Produtos, "Otica", Me)
    
                    If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                       MsgBox "C�digo de barra n�o cadastrado ou pre�o na tabela vigente n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                       txtCodigo_Produto.Text = Empty
                       txtCodigo_Produto.SetFocus
                       Me.txtPreco_Unitario.TabStop = False
                       Me.txtQuantidade_Produto.TabStop = False
                       Set rstInf_Produtos = Nothing
                       Exit Sub
                    End If
            
                    If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
                       MsgBox "Pre�o do Item n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                       txtCodigo_Produto.Text = Empty
                       txtCodigo_Produto.SetFocus
                       Me.txtPreco_Unitario.TabStop = False
                       Me.txtQuantidade_Produto.TabStop = False
                       Set rstInf_Produtos = Nothing
                       Exit Sub
                    End If
                    
                    txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
                    If frmTela_Venda.booEncerrante = False Then
                       txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
                    End If
                    txtQuantidade_Produto.Text = 1
                
                Else
                    strSql = Empty
                    strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,DFPreco_venda_TBProduto,DFPath_imagem_TBProduto " & _
                             "FROM TBProduto " & _
                             "WHERE TBProduto.IXCodigo_TBProduto = " & strCodigo_Produto_Etiqueta & " " & _
                             "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""

                    Call Movimentacoes.Select_geral(strSql, "BDPDV", rstInf_Produtos, "PDV", Me)

                    If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                       MsgBox "C�digo de barra n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                       txtCodigo_Produto.Text = Empty
                       txtCodigo_Produto.SetFocus
                       Me.txtPreco_Unitario.TabStop = False
                       Me.txtQuantidade_Produto.TabStop = False
                       Set rstInf_Produtos = Nothing
                       Exit Sub
                    End If
            
                    If rstInf_Produtos!DFPreco_venda_TBProduto <= 0 Or IsNull(rstInf_Produtos!DFPreco_venda_TBProduto) Then
                       MsgBox "Pre�o do Item n�o cadastrado.Verifique!", vbCritical, "Only Tech"
                       txtCodigo_Produto.Text = Empty
                       txtCodigo_Produto.SetFocus
                       Me.txtPreco_Unitario.TabStop = False
                       Me.txtQuantidade_Produto.TabStop = False
                       Set rstInf_Produtos = Nothing
                       Exit Sub
                    End If
                    
                    txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
                    If frmTela_Venda.booEncerrante = False Then
                       txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_venda_TBProduto
                    End If
                    txtQuantidade_Produto.Text = 1
                    
                End If
                
                Call Verifica_ICMS
                
                If booICMS_cadastrado = False And frmTela_Venda.intImpressoes_suportadas <> 1 Then
                    MsgBox "Aliquota de ICMS ou codigo ref. para impressora n�o cadastrada.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    Me.txtDescricao_Produto.Text = Empty
                    Me.txtQuantidade_Produto.Text = Empty
                    Me.txtPreco_Total.Text = Empty
                    Me.txtPreco_Unitario.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Me.txtPreco_Unitario.TabStop = False
                    Me.txtQuantidade_Produto.TabStop = False
                    Set rstInf_Produtos = Nothing
                    Exit Sub
                End If
                
                If Not IsNull(rstInf_Produtos!DFPath_imagem_TBProduto) Then
                   Call carrega_imagem
                End If
                
                Call Reposicao
            End If
        End If
    Else
        If booAcionado_Fechamento_cupom = False And booConsulta = False Then
           frmTela_Venda.txtCodigo_Produto.SetFocus
        End If
        booAcionado_Fechamento_cupom = True
    End If
    
    'Quantidade
    If txtQuantidade_Produto.TabStop = True Then
        If Me.txtCodigo_Produto.Text <> " " And Me.txtCodigo_Produto.Text <> "" Then
           On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
           frmTela_Venda.txtQuantidade_Produto.SetFocus
        Else
           Me.txtCodigo_Produto.SetFocus
        End If
    End If
    
    'Pre�o Unit�rio
    If Me.txtPreco_Unitario.TabStop = True Then
        If Me.txtCodigo_Produto.Text <> " " And Me.txtCodigo_Produto.Text <> "" Then
           Me.txtPreco_Unitario.SetFocus
           On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
        Else
           Me.txtCodigo_Produto.SetFocus
        End If
    End If
    
End Sub

Private Sub txtDescricao_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Unitario_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Unitario_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 44 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Unitario_LostFocus()
     If txtPreco_Unitario.Text <> "" Then
        If txtPreco_Unitario.Text = "," Then
           txtPreco_Unitario.Text = 0
        End If
        If txtPreco_Unitario.Text = 0 Then
           txtPreco_Unitario.Text = Empty
           txtPreco_Unitario.SetFocus
        Else
            If strCasas_Decimais = 2 Then
               txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
            End If
            If strCasas_Decimais = 3 Then
               txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.000")
            End If
            Call Processando_item
            txtCodigo_Produto.SetFocus
        End If
     End If
End Sub

Public Function Limpa_Tela()
    txtCodigo_Produto.Text = Empty
    txtDescricao_Produto.Text = Empty
    txtPreco_Total.Text = Empty
    txtPreco_Unitario = Empty
    txtQuantidade_Produto.Text = Empty
    booComanda = False
    imgProduto.Picture = LoadPicture("")
    Me.txtStatus.Text = "Aguardando Venda...."
End Function

Private Function Processando_item()

     'Calculando item
     If Not txtPreco_Unitario.Text = "" And Not txtQuantidade_Produto.Text = "" Then
        If strCasas_Decimais = 2 Then
           txtPreco_Total = Format(CDbl(txtPreco_Unitario.Text) * CDbl(txtQuantidade_Produto.Text), "#,###0.00")
        End If
        If strCasas_Decimais = 3 Then
           txtPreco_Total = Format(CDbl(txtPreco_Unitario.Text) * CDbl(txtQuantidade_Produto.Text), "#,###0.000")
        End If
     Else
        Exit Function
     End If
    
     Dim strCodigo_Produto As String
     Dim strDescricao_Produto As String * 29
     Dim strAliquota As String
     'Dim strTipo_quantidade As String * 1
     Dim strQuantiade As String * 7
     Dim strQuantiade_imp As String * 7
     'Dim strCasas_Decimais As String * 1
     Dim strValor_Unitario As String
     Dim strValor_Unitario_imp As String
     'Dim strTipo_desconto As String * 1
     Dim strValor_desconto As String * 8
     
     booEncerrante = False
     
     strCodigo_Produto = Me.txtCodigo_Produto.Text
     strDescricao_Produto = Me.txtDescricao_Produto.Text
     strQuantiade = Me.txtQuantidade_Produto.Text
     strValor_Unitario = txtPreco_Unitario.Text
     strValor_Unitario_imp = txtPreco_Unitario.Text
     strValor_desconto = "0,00"

     strAliquota = intAliquota_ICMS
     
     If strTipo_quantidade = "F" Then
        strQuantiade_imp = Format(Me.txtQuantidade_Produto.Text, "#,###0.00")
     Else
        strQuantiade_imp = Me.txtQuantidade_Produto.Text
     End If
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'Imprime item a item
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'ECF
     If frmTela_Venda.booCupom_fiscal = True Then
        Dim strComando As String
        
        'VERIFICA SE PRIMEIRO ITEM PARA ABRIR O CUPOM
        If HflexGrid.Rows < 3 And frmTela_Venda.intImpressoes_suportadas <> 1 Then
           Comandos_impressoras_fiscais.Abre_Cupom (strImpresora)
        End If
                
        If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
           Call Comandos_impressoras_fiscais.Vende_Item(strImpresora, strCodigo_Produto, strDescricao_Produto, strQuantiade_imp, strValor_Unitario_imp, strAliquota_imp_fisc, CInt(strCasas_Decimais), strTipo_desconto, CDbl(strValor_desconto), strTipo_quantidade, booGaveta_integrada)
        End If
        
        If frmTela_Venda.booInterrompe_venda = True Then
           Call Limpa_Tela
           Set rstInf_Produtos = Nothing
           Me.txtCodigo_Produto.SetFocus
           Exit Function
        End If
     End If
    
     'Verificando se existe o item no cupom....se existir ele adiciona sua respectiva quantidade e pre�o no item j� cadastrado
     If HflexGrid.Rows > 3 Then
        Dim intContador_item As Integer
        intContador_item = 1
        
        Me.HflexGrid.Col = 1
        Me.HflexGrid.Row = intContador_item + 3
        
        Do While Me.HflexGrid.Rows - 3 > intContador_item
           Me.HflexGrid.Row = intContador_item + 2
           If Format(txtCodigo_Produto, "0000000000000") = Me.HflexGrid.Text Then
              Call Adiciona_Item_Existente
              Exit Function
           End If
           intContador_item = intContador_item + 1
           Me.HflexGrid.Row = intContador_item
        Loop
     End If
     
     'Montando dysplay de itens de cupom
     HflexGrid.Cols = 6
     HflexGrid.ColWidth(0) = 0
     HflexGrid.Rows = HflexGrid.Rows + 2
     
     If HflexGrid.Rows = 4 Then
        'Cabe�alho 1
        HflexGrid.Row = 0
        HflexGrid.Col = 1
        HflexGrid.FixedAlignment(1) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = "C�digo"
        HflexGrid.Col = 2
        HflexGrid.Text = "Descri��o"
        HflexGrid.Col = 3
        HflexGrid.Text = "Qtd."
        HflexGrid.Col = 4
        HflexGrid.Text = "X"
        HflexGrid.Col = 5
        HflexGrid.Text = "Vlr.Unit."
        'Cabe�alho 2
        HflexGrid.Row = 1
        HflexGrid.Col = 1
        HflexGrid.FixedAlignment(1) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = ""
        HflexGrid.Col = 2
        HflexGrid.Text = ""
        HflexGrid.Col = 3
        HflexGrid.Text = ""
        HflexGrid.Col = 4
        HflexGrid.Text = ""
        HflexGrid.Col = 5
        HflexGrid.CellAlignment = 7
        HflexGrid.CellFontBold = True
        HflexGrid.Text = "T.Item"
        'Separador
        HflexGrid.Row = 2
        HflexGrid.RowHeight(2) = 100
        HflexGrid.Col = 1
        HflexGrid.FixedAlignment(1) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = "------------------------------------"
        HflexGrid.Col = 2
        HflexGrid.Text = "--------------------------------------------------"
        HflexGrid.Col = 3
        HflexGrid.Text = "----------------"
        HflexGrid.Col = 4
        HflexGrid.Text = "--------"
        HflexGrid.Col = 5
        HflexGrid.Text = "--------------"
     Else
        HflexGrid.Rows = HflexGrid.Rows - 1
     End If
     
     'Detalhe 1
     HflexGrid.Row = HflexGrid.Rows - 1
     HflexGrid.Col = 1
     HflexGrid.Font.Name = "Tahoma"
     HflexGrid.Text = Format(txtCodigo_Produto, "0000000000000")
     HflexGrid.Col = 2
     HflexGrid.Text = strDescricao_Produto
     HflexGrid.Col = 3
     HflexGrid.Text = strQuantiade
     HflexGrid.Col = 4
     HflexGrid.Text = "X"
     HflexGrid.Col = 5
     HflexGrid.Text = strValor_Unitario
     HflexGrid.Rows = HflexGrid.Rows + 1
     
     'Detalhe 2
     HflexGrid.Row = HflexGrid.Rows - 1
     HflexGrid.Col = 5
     HflexGrid.CellFontBold = True
     HflexGrid.CellFontSize = 6
     
     If strCasas_Decimais = 2 Then
        HflexGrid.Text = Format(CDbl(strQuantiade) * CDbl(strValor_Unitario), "#,###0.00")
     End If
     
     If strCasas_Decimais = 3 Then
        HflexGrid.Text = Format(CDbl(strQuantiade) * CDbl(strValor_Unitario), "#,###0.000")
     End If
     
     'Formatando Colunas
     HflexGrid.ColWidth(1) = 1100
     HflexGrid.ColWidth(2) = 2000
     HflexGrid.ColWidth(3) = 350
     HflexGrid.ColWidth(4) = 150
     HflexGrid.ColWidth(5) = 650
     
     Me.HflexGrid.SetFocus
     Me.HflexGrid.TopRow = Me.HflexGrid.Rows - 2
     
     If txtPreco_total_cupom.Text = "" Then txtPreco_total_cupom.Text = 0
     If txtPreco_Total.Text = "" Then txtPreco_Total.Text = 0
     
     If strCasas_Decimais = 2 Then
        txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     End If
     
     If strCasas_Decimais = 3 Then
        txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.000")
     End If
     
     Call Limpa_Tela
     
     Call Grava_Processa_Cupom
     
     Set rstInf_Produtos = Nothing
     
     txtQuantidade_Produto.TabStop = False
     txtPreco_Unitario.TabStop = False
     txtCodigo_Produto.SetFocus
    
End Function

Private Function Reposicao()

    If strCasas_Decimais = 2 Then
       txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
       txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.00")
    End If
    If strCasas_Decimais = 3 Then
       txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.000")
       txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.000")
    End If
    
    'Verificando se vai passar pela quantidade e no antes de calcular
    If Me.txtQuantidade_Produto.TabStop = False And Me.txtPreco_Unitario.TabStop = False Then
        Call Processando_item
    Else
        Set rstInf_Produtos = Nothing
    End If
    
End Function

Private Function carrega_imagem()

    On Error GoTo Erro_imagem
    
    Me.imgProduto.Picture = LoadPicture(rstInf_Produtos!DFPath_imagem_TBProduto)
        
Fim_Imagem:
    
    Exit Function
    
Erro_imagem:

    If Err.Number <> 76 Then
       If booIntegracao_Retaguarda = True Then
          Erro.Erro Me, "Otica"
       Else
          Erro.Erro Me, "PDV"
       End If
    Else
       GoTo Fim_Imagem
    End If

End Function

Private Sub txtQuantidade_Produto_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 44 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtQuantidade_Produto_LostFocus()
    If Me.txtQuantidade_Produto.TabStop = True Then
        If txtQuantidade_Produto.Text = "," Then
           txtQuantidade_Produto.Text = 0
        End If
        If txtQuantidade_Produto.Text = 0 Then
           txtQuantidade_Produto.Text = Empty
           txtQuantidade_Produto.SetFocus
        Else
            If strCasas_Decimais = 2 Then
               txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
               txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.00")
            End If
            If strCasas_Decimais = 3 Then
               txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.000")
               txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.000")
            End If
            Call Processando_item
        End If
    End If
End Sub
Private Function Finaliza_Cupom()
    frmFechamento_Cupom.Show 1
End Function
Private Function Verifica_ICMS()

    Dim strST As String
    Dim strST2 As String
    Dim dblAliquota_icms As Double
    Dim dblTotal_Icms As Double
    Dim rstUF As New ADODB.Recordset
    Dim strUF_Emitente As String
    Dim intIDCfo As Long
    
    'Verifica a uf do emitente
    strSql = Empty
    strSql = "SELECT TBCidade_otica.DFUf_TBCidade_otica FROM TBEmpresa " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBEmpresa.Fkid_TBCidade_otica  = TBCidade_otica.pkid_TBCidade_otica " & _
             "WHERE TBEmpresa.PKCodigo_TBempresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    If booIntegracao_Retaguarda = True Then
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstUF, "Otica", Me)
    Else
       Call Movimentacoes.Select_geral(strSql, "BDPDV", rstUF, "PDV", Me)
    End If
    
    strUF_Emitente = rstUF!DFUf_TBCidade_otica
    
    Set rstUF = Nothing
     
    'Calculando a parte do ICMS relacionado ao Item
    'Concatenando o valor da Situa��o Tribut�ria que est� no cadastro de produto
    strST = rstInf_Produtos!DFCst1_TBProduto
    strST2 = rstInf_Produtos!DFCst2_TBProduto
    
    'ICMS E ST
    'Verifica se a ST for 030 ou 060 o valor da aliquota e o valor de ICMS � 0;
    'E Grava na tabela CFO_Pedido mais uma CFO para este pedido
    If strST = "030" Or strST = "060" Then
    
       dblAliquota_icms = 0
       dblTotal_Icms = 0
       
       Dim rstVerifica_Estado_ST As New ADODB.Recordset
       Dim rstCFO_ST As New ADODB.Recordset
       
       strSql = Empty
       strSql = "SELECT TBCidade_otica.DFUf_TBCidade_otica " & _
                "FROM TBEmpresa " & _
                "INNER JOIN TBCidade_otica " & _
                "ON TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                "WHERE PKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                
       If booIntegracao_Retaguarda = True Then
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVerifica_Estado_ST, "Otica", Me)
       Else
          Call Movimentacoes.Select_geral(strSql, "BDPDV", rstVerifica_Estado_ST, "PDV", Me)
       End If
       
       If rstVerifica_Estado_ST!DFUf_TBCidade_otica = strUF_Emitente Then
          'Localizando no parametro o proximo cfo de substitui��o para dentro do estado
          strSql = Empty
          strSql = "SELECT DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
                   "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
          If booIntegracao_Retaguarda = True Then
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCFO_ST, "Otica", Me
          Else
             Movimentacoes.Select_geral strSql, "BDPDV", rstCFO_ST, "PDV", Me
          End If
       Else
          'Localizando no parametro o proximo cfo de substitui��o para dentro do estado
          strSql = Empty
          strSql = "SELECT DFProximo_cfop_venda_fora_estado_substituicao_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
                   "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
          If booIntegracao_Retaguarda = True Then
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCFO_ST, "Otica", Me
          Else
             Movimentacoes.Select_geral strSql, "BDPDV", rstCFO_ST, "PDV", Me
          End If
       End If
       
       'Localizando o ID do CFO
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'Lembrar Marcos para fazer teste caso o produto nao                             '
       'esteja cadastrado no estado para ICMS(Giordano).                               '
       'altera��o feita na busca do ID do CFO (ERRO de passagem de valor para a funcao)'
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstCFO_ST.BOF = True And rstCFO_ST.EOF = True Then
          MsgBox "Verifique se o CFO na tabela de par�metros fiscais est� preenchida corretamente!", vbCritical, "Only Tech"
       End If
       
       If booIntegracao_Retaguarda = True Then
          intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", rstCFO_ST.Fields("CFO"), "TBCFOP", "Otica", Me, "BDRetaguarda")
       Else
          intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", rstCFO_ST.Fields("CFO"), "TBCFOP", "PDV", Me, "BDPDV")
       End If
       
       If rstCFO_ST.BOF = True And rstCFO_ST.EOF = True Then
          MsgBox "Verifique se o CFO na tabela de par�metros fiscais est� preenchida corretamente!", vbCritical, "Only Tech"
       End If
       
       Set rstVerifica_Estado_ST = Nothing
       Set rstCFO_ST = Nothing
    Else
        Dim rstVerifica_Estado_ICMS As New ADODB.Recordset
        'Query para pegar ICMS do item
        strSql = Empty
        strSql = "SELECT " & _
                 "DFPercentual_icms_saida_juridica_TBEstado_icms,DFTributacao_impressora_fiscal_TBEstado_icms," & _
                 "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                 "FROM TBEstado_icms " & _
                 "WHERE FKId_TBProduto = " & rstInf_Produtos!PKId_TBProduto & " " & _
                 "AND DFUf_TBEstado_icms = '" & strUF_Emitente & "'"
                 
        If booIntegracao_Retaguarda = True Then
           Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
        Else
           Movimentacoes.Select_geral strSql, "BDPDV", rstVerifica_Estado_ICMS, "PDV", Me
        End If
        
        If rstVerifica_Estado_ICMS.BOF = True And rstVerifica_Estado_ICMS.EOF = True Then
           Set rstVerifica_Estado_ICMS = Nothing
           'Query para pegar ICMS do item, com estado **
           strSql = Empty
           strSql = "SELECT " & _
                    "DFPercentual_icms_saida_juridica_TBEstado_icms,DFTributacao_impressora_fiscal_TBEstado_icms," & _
                    "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                    "FROM TBEstado_icms " & _
                    "WHERE FKId_TBProduto = " & rstInf_Produtos!PKId_TBProduto & " " & _
                    "AND DFUf_TBEstado_icms = '**' "
                    
           If booIntegracao_Retaguarda = True Then
              Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
           Else
              Movimentacoes.Select_geral strSql, "BDPDV", rstVerifica_Estado_ICMS, "PDV", Me
           End If
        End If
                
        If rstVerifica_Estado_ICMS.BOF = True And rstVerifica_Estado_ICMS.EOF = True Then
           booICMS_cadastrado = False
           intAliquota_ICMS = Empty
           strAliquota_imp_fisc = Empty
        Else
           intAliquota_ICMS = rstVerifica_Estado_ICMS!DFPercentual_icms_saida_fisica_TBEstado_icms
           If IsNull(rstVerifica_Estado_ICMS!DFTributacao_impressora_fiscal_TBEstado_icms) = True Or rstVerifica_Estado_ICMS!DFTributacao_impressora_fiscal_TBEstado_icms = "" Then
              booICMS_cadastrado = False
           Else
              strAliquota_imp_fisc = Trim(rstVerifica_Estado_ICMS!DFTributacao_impressora_fiscal_TBEstado_icms)
           End If
        End If
    End If

End Function

Private Function Adiciona_Item_Existente()

     HflexGrid.Col = 3
     HflexGrid.Text = CDbl(HflexGrid.Text) + CDbl(Me.txtQuantidade_Produto.Text)
     
     HflexGrid.Row = HflexGrid.Row + 1
     HflexGrid.Col = 5
     
     If strCasas_Decimais = 2 Then
        HflexGrid.Text = Format(CDbl(HflexGrid.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     End If
     
     If strCasas_Decimais = 3 Then
        HflexGrid.Text = Format(CDbl(HflexGrid.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.000")
     End If
     
     If strCasas_Decimais = 2 Then
        Me.txtPreco_total_cupom.Text = Format(CDbl(Me.txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     End If
     
     If strCasas_Decimais = 3 Then
        Me.txtPreco_total_cupom.Text = Format(CDbl(Me.txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.000")
     End If
     
     Me.HflexGrid.SetFocus
     Me.HflexGrid.Row = Me.HflexGrid.Rows - 1
     
     Call Limpa_Tela
     
     Set rstInf_Produtos = Nothing
     
     txtQuantidade_Produto.TabStop = False
     txtPreco_Unitario.TabStop = False
     txtCodigo_Produto.SetFocus

End Function

Private Function Abre_Cupom()
        
    'Retorno = Bematech_FI_AbreCupom("08706114709")
     
    'Fun��o que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Abertura de Cupom")
    
End Function
Private Function Cancela_item_cupom()
    Call Comandos_impressoras_fiscais.Cancela_item(strImpresora)
End Function

Private Function Cancela_cupom()
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strSem_Produto As String
    Dim lngID_Nota As Long
    Dim CNconexao_cancelamento As New DLLConexao_Sistema.conexao
    Dim ocorrencia As New DLLSistema.Ocorrencia_de_produto
    Dim estoque As New DLLSistema.estoque
    Dim rstNumero_orcamento As New ADODB.Recordset
    Dim Nota As String
    Dim Serie As String
    Dim lngID_Cupom As Long
    Dim lngID_Cupom_local As Long
    Dim rstCupom_processado As New ADODB.Recordset
    
    booErro_processamento_impressora = False
    
    'Envia comando para a impressora
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
       Call Comandos_impressoras_fiscais.Cancela_cupom(strImpresora)
       If booErro_processamento_impressora = True Then
          Exit Function
       End If
       
    End If
    
    strSql = Empty
    strSql = "SELECT * FROM TBProcessamento_Cupom"
    Movimentacoes.Select_geral strSql, "BDPDV", rstCupom_processado, "PDV", Me
    
    'Verifica se existe cupom n�o gravado
    'Pois se esta recordset estiver realmente vazia, representa q ele ter� q ir cancelar o ultimo gravado no banco
    'Caso contr�rio...limpar� somente a tela e eliminar� o mesmo da tabela e nem entrar� no IF abaixo
    '                                              ||
    '                                              ||
    '                                              ||
    '                                              ||
    '                                             \  /
    '                                              \/
    If rstCupom_processado.BOF = True And rstCupom_processado.EOF = True Then
        'Verifica se cupom ou or�amento
        If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
           'Pr�ximo Or�amento
            strSql = Empty
            strSql = "SELECT DFProximo_cupom_TBParametros_ecf,DFProximo_serie_cupom_TBParametros_ecf FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
        Else
            'Pr�ximo Or�amento
            strSql = Empty
            strSql = "SELECT DFProximo_orcamento_balcao_TBParametros_ecf,DFProximo_serie_orcamento_balcao_TBParametros_ecf FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
        End If
        
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           Movimentacoes.Select_geral strSql, "BDRetaguarda", rstNumero_orcamento, "Otica", Me
        Else
           Movimentacoes.Select_geral strSql, "BDPDV", rstNumero_orcamento, "PDV", Me
        End If
        
        If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
            Nota = CInt(rstNumero_orcamento!DFProximo_cupom_TBParametros_ecf) - 1
            Serie = rstNumero_orcamento!DFProximo_serie_cupom_TBParametros_ecf
        Else
            Nota = CInt(rstNumero_orcamento!DFProximo_orcamento_balcao_TBParametros_ecf) - 1
            Serie = rstNumero_orcamento!DFProximo_serie_orcamento_balcao_TBParametros_ecf
        End If
        
        Set rstNumero_orcamento = Nothing
        
        'Indicando o banco � conectar-se
        CNconexao_cancelamento.Initial_Catalog = "BDRetaguarda"
        
        'Estabelecendo conex�o com o banco
        CNconexao_cancelamento.Abrir_conexao ("Otica")
        
        'Indica o inicio da transa��o junto o banco
        CNconexao_cancelamento.CNconexao.BeginTrans
        
        If booIntegracao_Retaguarda = True Then
        
            lngID_Nota = Funcoes_Gerais.Localiza_ID("PKId_TBnota_saida", "DFNumero_TBNota_saida", Nota, "TBNota_Saida", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", CInt(strEmpresa_Operador), "DFSerie_TBNota_saida", Serie)
            lngID_Cupom = Funcoes_Gerais.Localiza_ID("PKId_TBCupom", "DFNumero_TBCupom", Nota, "TBCupom", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", CInt(strEmpresa_Operador), "DFSerie_TBCupom", Serie)
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                                         C O N S U L T A S
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim rstTitulos_Receber As New ADODB.Recordset
            Dim rstBusca_Quantidade As New ADODB.Recordset
            Dim rstBusca_Produto As New ADODB.Recordset
            Dim rstInf_corpo_nota As New ADODB.Recordset
            Dim rstVendedor As New ADODB.Recordset
            
            strSql = Empty
            strSql = "SELECT PKId_TBTitulo_receber FROM TBTitulo_receber " & _
                     "WHERE DFNumero_documento_TBTitulo_receber = " & Nota & " AND " & _
                     "DFSerie_TBTitulo_receber = '" & Serie & "' " & _
                     "AND FKCodigo_TBEmpresa = " & strEmpresa_Operador & ""
            Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTitulos_Receber, "Otica", Me
            
            strSql = Empty
            strSql = "SELECT * FROM TBItens_nota_saida " & _
                     "WHERE TBItens_nota_saida.FKId_TBnota_saida = " & lngID_Nota & ""
            Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Quantidade, "Otica", Me)
            
            If rstBusca_Quantidade.RecordCount = 0 Then
               strSem_Produto = "0"
            Else
               strSem_Produto = "1"
            End If
                
            strSql = Empty
            strSql = "SELECT * FROM TBProduto WHERE IXCodigo_TBEmpresa = " & strEmpresa_Operador & ""
            Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Produto, "Otica", Me)
                   
            strSql = Empty
            strSql = "SELECT * FROM TBNota_saida WHERE DFNumero_TBNota_saida = " & Nota & " AND DFSerie_TBNota_saida = '" & Serie & "' AND FKCodigo_TBEmpresa = " & Me.strEmpresa_Operador & ""
            Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstInf_corpo_nota, "Otica", Me)
            
            strSql = Empty
            strSql = "SELECT * FROM TBVendedor WHERE IXCodigo_TBEmpresa = " & strEmpresa_Operador & ""
            Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVendedor, "Otica", Me)
                   
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                                         G R A V A � � E S
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            On Error GoTo Erro_Transacao
            
            'Alterando para cancelado o cupom
            strSql = Empty
            strSql = "UPDATE TBCupom " & _
                     "SET DFCancelado_TBCupom = '1', " & _
                     "DFUsuario_cancelamento_TBCupom = '" & strOperador & "', " & _
                     "DFMotivo_cancelamento_TBCupom = 'Cancelamento de Cupom no PDV.: " & Me.strPDV & "'" & _
                     "WHERE PKId_TBCupom = " & lngID_Cupom & " "
            'Gravando
            CNconexao_cancelamento.CNconexao.Execute strSql
            
            'Alterando para cancelada a Nota de saida
            strSql = Empty
            strSql = "UPDATE TBNota_Saida " & _
                     "SET DFCancelado_TBNota_saida = '1', " & _
                     "DFUsuario_cancelamento_TBNota_saida = '" & strOperador & "', " & _
                     "DFMotivo_cancelamento_TBNota_saida = 'Cancelamento de Cupom no PDV.: " & Me.strPDV & "'" & _
                     "WHERE PKId_TBNota_saida = " & lngID_Nota & " "
            'Gravando
            CNconexao_cancelamento.CNconexao.Execute strSql
            
            'Matando titulos nota_saida
            strSql = Empty
            strSql = "DELETE FROM TBTitulo_nota_saida WHERE FKId_TBNota_saida = " & lngID_Nota & ""
            CNconexao_cancelamento.CNconexao.Execute strSql
            
            strSql = "DELETE FROM TBTitulo_recebido " & _
                     "WHERE FKId_TBTitulo_receber = " & rstTitulos_Receber!PKId_TBTitulo_receber & " "
            'Gravando
            CNconexao_cancelamento.CNconexao.Execute strSql
            
            rstTitulos_Receber.MoveNext
               
            strSql = Empty
            strSql = "DELETE FROM TBTITULO_RECEBER " & _
                     "WHERE DFNumero_documento_TBTitulo_receber = " & Nota & " AND " & _
                     "DFSerie_TBTitulo_receber = '" & Serie & "'" & _
                     "AND FKCodigo_TBEmpresa = " & strEmpresa_Operador & ""
            CNconexao_cancelamento.CNconexao.Execute strSql
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'LENDO A TABELA DE ITENS NOTA E BUSCANDO A QUANTIDADE QUE SER� ADICIONADA '
            'AO ESTOQUE.                                                              '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            rstBusca_Quantidade.MoveFirst
            
            Do While rstBusca_Quantidade.EOF = False
                rstBusca_Produto.MoveFirst
                rstBusca_Produto.Find ("PKId_TBProduto = " & rstBusca_Quantidade.Fields("FKId_TBProduto") & "")
                
                rstVendedor.MoveFirst
                rstVendedor.Find ("PKId_TBVendedor = " & rstInf_corpo_nota.Fields("FKId_TBVendedor") & "")
                
                ocorrencia.Data_Movimento = Date
                ocorrencia.Estoque_Anterior = rstBusca_Produto.Fields("DFEstoque_atual_TBProduto")
                ocorrencia.Estoque_Atual = CDbl(rstBusca_Produto.Fields("DFEstoque_atual_TBProduto")) + CDbl(rstBusca_Quantidade.Fields("DFQuantidade_baixa_estoque_TBItens_nota_saida"))
                ocorrencia.Hora_Movimento = Format(Now, "hh:mm:ss")
                ocorrencia.ID_Produto = rstBusca_Produto.Fields("PKId_TBProduto")
                ocorrencia.Observacao = "Canc. N.F.Sa�da N�:" & Nota & "-" & Serie & "-Adi��o de Estoque"
                ocorrencia.Programa = "Canc.N.F.Sa�da - PDV"
                ocorrencia.Quantidade_Movimentada = CDbl(rstBusca_Quantidade.Fields("DFQuantidade_baixa_estoque_TBItens_nota_saida"))
                ocorrencia.Usuario = strOperador
                ocorrencia.Gravar "Otica", False
                'Manipula��o do Estoque.
                estoque.ID_Produto = rstBusca_Produto.Fields("PKId_TBProduto")
                estoque.Quantidade_Menor_Unidade_Item = CDbl(rstBusca_Quantidade.Fields("DFQuantidade_baixa_estoque_TBItens_nota_saida"))
                estoque.Quantidade_Antes_Atualizar_Estoque = rstBusca_Produto.Fields("DFEstoque_atual_TBProduto")
                estoque.Adicionar_Estoque "Otica", True, CNconexao_cancelamento
                
                rstBusca_Quantidade.MoveNext
            Loop
            
            Set rstBusca_Quantidade = Nothing
            Set rstBusca_Produto = Nothing
        Else
            lngID_Cupom_local = Funcoes_Gerais.Localiza_ID("PKId_TBCupom", "DFNumero_TBCupom", Nota, "TBCupom", "Otica", Me, "BDPDV", "FKCodigo_TBEmpresa", CInt(strEmpresa_Operador), "DFSerie_TBCupom", Serie)
            'Alterando para cancelado o cupom
            strSql = Empty
            strSql = "UPDATE TBCupom " & _
                     "SET DFCancelado_TBCupom = '1', " & _
                     "DFUsuario_cancelamento_TBCupom = '" & strOperador & "', " & _
                     "DFMotivo_cancelamento_TBCupom = 'Cancelamento de Cupom no PDV.: & " & Me.strPDV & "'" & _
                     "WHERE PKId_TBNota_saida = " & lngID_Cupom_local & " "
            
            'COMITANDO A TRANSA��O
            CNconexao_cancelamento.CNconexao.Execute strSql
        End If
        
        'COMITANDO A TRANSA��O
        CNconexao_cancelamento.CNconexao.CommitTrans
        
        'Fechando a conex�o
        CNconexao_cancelamento.CNconexao.Close
    End If
    
    Set rstCupom_processado = Nothing
    
    'Limpando o cupom em processamento
    Call Limpa_Processa_Cupom
    
    Exit Function
    
Erro_Transacao:

    'ROOLBACK NA TRANSA��O
    CNconexao_cancelamento.CNconexao.RollbackTrans
    
    MsgBox Err.Description, vbCritical, "Only Tech"
    
    Exit Function
    
End Function
Public Function Grava_Processa_Cupom()

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Rotina gravada somente no pdv local
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo Erro
    
    Dim CNconexao_Processa_cupom As New DLLConexao_Sistema.conexao

    strSql = Empty
    strSql = "INSERT INTO TBProcessamento_Cupom(DFData_TBProcessamento_Cupom,DFHora_TBProcessamento_Cupom) " & _
             "VALUES ('" & Format(Now, "YYYYMMDD") & "','" & Format(Now, "HH:MM:SS") & "')"
    
    'Indicando o banco � conectar-se
    CNconexao_Processa_cupom.Initial_Catalog = "BDPDV"
    
    'Estabelecendo conex�o com o banco
    CNconexao_Processa_cupom.Abrir_conexao ("PDV")
    
    'Indica o inicio da transa��o junto o banco
    CNconexao_Processa_cupom.CNconexao.BeginTrans
    
    On Error GoTo Erro_Transacao
    
    CNconexao_Processa_cupom.CNconexao.Execute strSql
    
    'Indica o inicio da transa��o junto o banco
    CNconexao_Processa_cupom.CNconexao.CommitTrans
     
    Set CNconexao_Processa_cupom = Nothing
    
    Exit Function
    
Erro:

    MsgBox Err.Description, vbCritical, "Only Tech"
    
    Exit Function

Erro_Transacao:

    'ROOLBACK NA TRANSA��O
    CNconexao_Processa_cupom.CNconexao.RollbackTrans
    
    MsgBox Err.Description, vbCritical, "Only Tech"
    
    Exit Function

End Function
Public Function Limpa_Processa_Cupom()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Rotina gravada somente no pdv local
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo Erro
    
    Dim CNconexao_Processa_cupom As New DLLConexao_Sistema.conexao

    strSql = Empty
    strSql = "DELETE FROM TBProcessamento_Cupom"
    
    'Indicando o banco � conectar-se
    CNconexao_Processa_cupom.Initial_Catalog = "BDPDV"
    
    'Estabelecendo conex�o com o banco
    CNconexao_Processa_cupom.Abrir_conexao ("PDV")
    
    'Indica o inicio da transa��o junto o banco
    CNconexao_Processa_cupom.CNconexao.BeginTrans
    
    On Error GoTo Erro_Transacao
    
    CNconexao_Processa_cupom.CNconexao.Execute strSql
    
    'Indica o inicio da transa��o junto o banco
    CNconexao_Processa_cupom.CNconexao.CommitTrans
     
    Set CNconexao_Processa_cupom = Nothing
    
    Exit Function
    
Erro:

    MsgBox Err.Description, vbCritical, "Only Tech"
    
    Exit Function

Erro_Transacao:

    'ROOLBACK NA TRANSA��O
    CNconexao_Processa_cupom.CNconexao.RollbackTrans
    
    MsgBox Err.Description, vbCritical, "Only Tech"
    
    Exit Function



End Function

