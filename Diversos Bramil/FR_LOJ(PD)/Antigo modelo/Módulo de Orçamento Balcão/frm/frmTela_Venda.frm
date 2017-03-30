VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
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
      Left            =   11160
      Top             =   8520
   End
   Begin VB.TextBox txtNumero_loja 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      BackColor       =   &H00C0C0C0&
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
      BackColor       =   &H00C0C0C0&
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
      BackColor       =   &H00C0C0C0&
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
      BackColor       =   &H00C0C0C0&
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
      BackColor       =   &H00C0C0C0&
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
      Width           =   11565
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
      Height          =   3645
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   930
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
      Height          =   5055
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   4575
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
      Width           =   3645
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   8070
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   3675
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
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
      Height          =   915
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   6330
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
      Caption         =   "Código"
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
      Caption         =   "Descrição"
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
      Caption         =   "Preço Unitário"
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
      Caption         =   "Preço Total"
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
' Módulo.................: Frente de Loja                                                 '
' Objetivo...............: Tela de Vendas                                                 '
' Data de Criação........: 04/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String
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
Dim booICMS_cadastrado As Boolean
Dim rstParametro_ecf As New ADODB.Recordset

Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intRetorno As Integer
    
    'Verifica se foi precionado F2 e abre tela de consulta
    If KeyCode = 113 Then
        Me.txtStatus.Text = "Consultando o Item..."
        frmConsulta_Produto.Show (1)
    End If
    
    'Verifica se foi precionado F4 e Finaliza a Compra
    If KeyCode = 115 Then
        booAcionado_Fechamento_cupom = True
        Me.txtStatus.Text = "Finalizando a compra..."
        Call Finaliza_Cupom
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F5 e Cancela o último item
    
    If KeyCode = 116 Then
        booAcionado_Fechamento_cupom = True
        Me.txtStatus.Text = "Cancelando último item..."
        
        If Me.HflexGrid.Rows - 3 > 0 Then
            
            'Subtraindo totalizador..........
            txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) - CDbl(Me.HflexGrid.Text), "#,###0.00")
            
            Me.HflexGrid.RemoveItem (Me.HflexGrid.Rows - 1)
         
            If Me.HflexGrid.Rows - 2 > 0 Then Me.HflexGrid.RemoveItem (Me.HflexGrid.Rows - 1)
        End If
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F6 e Cancela o cupom
    If KeyCode = 117 Then
        booAcionado_Fechamento_cupom = True
        Me.txtStatus.Text = "Cancelando o cupom..."
        'Verifica se confirma a operação
        intRetorno = MsgBox("Se confirmar a operação este cupom será vcancelado.Confirma a operação?", vbYesNo, "Only Tech")
        If intRetorno = 7 Then
           Exit Sub
        End If
        
        Call Limpa_Tela
        
        Me.HflexGrid.Clear
        Me.HflexGrid.ClearStructure
        Me.txtPreco_total_cupom.Text = ""
        Me.HflexGrid.Rows = 2
        Me.txtCodigo_Produto.SetFocus
        
    End If
    
    'Verifica se foi precionado F7 e Finaliza o Operador
    If KeyCode = 118 Then
        booAcionado_Fechamento_cupom = True
        Me.txtStatus.Text = "Finalizando o Operador..."
        frmFechamento_Operador.Show (1)
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F8 e Aciona Sangria
    If KeyCode = 119 Then
        booAcionado_Fechamento_cupom = True
        Me.txtStatus.Text = "Acionando a Sangria..."
        frmSangria.Show (1)
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F9 e libera a digitação da quantidade
    If KeyCode = 120 Then
        'Libera para digitação da quantidade
        Me.txtStatus.Text = "Digite a quantidade..."
        txtQuantidade_Produto.TabStop = True
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F10 e libera a digitação do preço
    If KeyCode = 121 Then
        'Libera para digitação da qunantidade
        Me.txtStatus.Text = "Digite o valor do preço..."
        Me.txtPreco_Unitario.TabStop = True
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi precionado F12 e abre a tela de comanda
    If KeyCode = 123 Then
        booAcionado_Fechamento_cupom = True
        Me.txtStatus.Text = "Comanda..."
        frmComanda.Show (1)
        txtCodigo_Produto.SetFocus
    End If
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    
    'Indica se para o programa se a impressora está lacrada ou não
    booImpressora_lacrada = False
    
    'Indica se este cupom recebeu carga ou não de uma comanda
    booComanda = False
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    strSQL = Empty
    strSQL = "SELECT * FROM TBEmpresa Where PKCodigo_TBEmpresa = " & strEmpresa_Operador & ""
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstEmpresa, "Otica", Me
    
    txtNumero_check_out.Text = strPDV
    txtNumero_Nome_Operadora.Text = "Operador: " & strOperador
    txtVersao_software.Text = "Versão 1.0"
    txtNumero_loja.Text = "Loja: " & rstEmpresa!PKCodigo_TBempresa
    imgLogo_Empresa.Picture = LoadPicture(rstEmpresa!DFPath_logomarca_TBEmpresa)
    
    '-------------------------------------------------------------------------------------------------------
    'Abrindo Impressora não fiscal
    '-------------------------------------------------------------------------------------------------------
    Dim intPorta As Integer
    Dim strComunica As String
    
    ' Fecha a porta que está aberta
    intPorta = FechaPorta()
    If intPorta <= 0 Then
       MsgBox "Problemas ao Fechar a Porta de Comunicação com a imp. não fiscal.Reinicie a aplicação", vbCritical, "Only Tech"
    End If

    ' Abre a porta de comunicacao com imp. não fiscal
    intPorta = IniciaPorta("LPT1")
    If intPorta <= 0 Then
       MsgBox "Problemas ao Abrir a Porta de Comunicação com a imp. não fiscal.Reinicie a aplicação", vbCritical, "Only Tech"
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
    End
End Sub

Private Sub Timer1_Timer()
    txtData_Hora.Text = "Data: " & Format(Now, "DD/MM/YYYY") & " - " & Format(Now, "hh:mm:ss")
End Sub

Private Sub txtCodigo_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Produto_LostFocus()

    Dim rstBusca_Paramentros As New ADODB.Recordset
    Dim strPreco_Peso_Parametro As String
    Dim strDigito_Produto_Digitado As String
    Dim strCodigo_Produto_Etiqueta As String
    Dim strPreco_Peso As String
    Dim strDecimal As String
    Dim strQuantidade As String
    Dim strTotal As String
    Dim strPreco_Tabela As String
    Dim strID_Produto As String
    
    booICMS_cadastrado = True
    
    strSQL = Empty
    strSQL = "SELECT * FROM TBParametros_ecf WHERE FKCodigo_TBEmpresa = " & strEmpresa_Operador & ""
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstBusca_Paramentros, "Otica", Me)
    
    strDigito_Peso_Variavel = rstBusca_Paramentros.Fields("DFCodigo_inicial_peso_variavel_TBParametros_ecf")
    
    If rstBusca_Paramentros.Fields("DFPreco_peso_balanca_TBParametros_ecf") = False Then
       strPreco_Peso_Parametro = 0
    Else
       strPreco_Peso_Parametro = 1
    End If
    
    Set rstBusca_Paramentros = Nothing
    
    If txtCodigo_Produto.Text <> "" Then
        Me.txtStatus.Text = "Processando o item..."
        'Código Interno
        If Len(Me.txtCodigo_Produto.Text) < 7 Then
            strSQL = Empty
            strSQL = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco,DFPath_imagem_TBProduto " & _
                     "FROM TBProduto " & _
                     "INNER JOIN TBITENS_TABELA_PRECO " & _
                     "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                     "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
                     "AND TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & ""
                     
            Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstInf_Produtos, "Otica", Me
            
            If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                MsgBox "Código Interno não cadastrado.Verifique!", vbCritical, "Only Tech"
                txtCodigo_Produto.Text = Empty
                txtCodigo_Produto.SetFocus
                Set rstInf_Produtos = Nothing
                Exit Sub
            End If
            
            If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
                MsgBox "Preço do Item não cadastrado.Verifique!", vbCritical, "Only Tech"
                txtCodigo_Produto.Text = Empty
                txtCodigo_Produto.SetFocus
                Set rstInf_Produtos = Nothing
                Exit Sub
            End If
            
            Call Verifica_ICMS
            
            If booICMS_cadastrado = False Then
                MsgBox "Aliquota de ICMS não cadastrada.Verifique!", vbCritical, "Only Tech"
                txtCodigo_Produto.Text = Empty
                txtCodigo_Produto.SetFocus
                Set rstInf_Produtos = Nothing
                Exit Sub
            End If
            
            txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
            txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
            txtQuantidade_Produto.Text = 1
            
            If Not IsNull(rstInf_Produtos!DFPath_imagem_TBProduto) Then
               Call Carrega_imagem
            End If
            
            Call Reposicao

        End If
        
        'Código de barra
        If Len(Me.txtCodigo_Produto.Text) > 6 Then
            strDigito_Produto_Digitado = Left(txtCodigo_Produto.Text, 1)
            
            If strDigito_Peso_Variavel = strDigito_Produto_Digitado Then
                'Produto pesável e preço variavel
                strCodigo_Produto_Etiqueta = Mid(txtCodigo_Produto.Text, 2, 4)
                strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 7)
              
                strSQL = Empty
                strSQL = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco,DFPath_imagem_TBProduto " & _
                         "FROM TBProduto " & _
                         "INNER JOIN TBITENS_TABELA_PRECO " & _
                         "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                         "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
                         "AND TBProduto.IXCodigo_TBProduto = " & strCodigo_Produto_Etiqueta & ""
                Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstInf_Produtos, "Otica", Me)
               
                If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                   MsgBox "Código de barra não cadastrado.Verifique!", vbCritical, "Only Tech"
                   txtCodigo_Produto.Text = Empty
                   txtCodigo_Produto.SetFocus
                   Set rstInf_Produtos = Nothing
                   Exit Sub
                End If
              
                If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
                  MsgBox "Preço do Item não cadastrado.Verifique!", vbCritical, "Only Tech"
                  txtCodigo_Produto.Text = Empty
                  txtCodigo_Produto.SetFocus
                  Set rstInf_Produtos = Nothing
                  Exit Sub
                End If
                
                Call Verifica_ICMS
                
                If booICMS_cadastrado = False Then
                    MsgBox "Aliquota de ICMS não cadastrada.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Set rstInf_Produtos = Nothing
                    Exit Sub
                End If
                
                txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
                txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
                txtQuantidade_Produto.Text = 1
                
                If Not IsNull(rstInf_Produtos!DFPath_imagem_TBProduto) Then
                   Call Carrega_imagem
                End If
            
                If strPreco_Peso_Parametro = 0 Then
                   strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 5)
                   strDecimal = Mid(txtCodigo_Produto.Text, 11, 2)
                   strPreco_Peso = strPreco_Peso & "," & strDecimal
                   strPreco_Peso = Format(strPreco_Peso, "#,###0.00")
                   strQuantidade = CDbl(strPreco_Peso) / CDbl(txtPreco_Unitario.Text)
                   strQuantidade = Format(strQuantidade, "#,###0.000")
                   txtQuantidade_Produto.Text = strQuantidade
                   txtPreco_Unitario.Text = CDbl(txtPreco_Unitario.Text)
                   strTotal = CDbl(txtPreco_Unitario.Text) * CDbl(strQuantidade)
                   strTotal = Format(strTotal, "#,###0.00")
                   txtPreco_Total.Text = strTotal
                Else
                   strPreco_Peso = Format(strPreco_Peso, "#,###0.000")
                   strTotal = strPreco_Peso * strPreco_Tabela
                   txtQuantidade_Produto.Text = strPreco_Peso
                   txtPreco_Unitario.Text = Format(strPreco_Tabela, "#,###0.00")
                   txtPreco_Total.Text = Format(strTotal, "#,###0.00")
                End If
                
                Call Reposicao
              
            Else
                Dim rstID_Codautomacao As New ADODB.Recordset
                
                'Produto não pesável e preço não variavel
                'Query para verificar a validade do cod. de automação
                strSQL = Empty
                strSQL = "SELECT TBCodigo_barras.FKId_TBProduto " & _
                         "FROM TBCodigo_barras " & _
                         "WHERE IXCodigo_TBCodigo_barras = " & txtCodigo_Produto.Text & " "
                         
                Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstID_Codautomacao, "Otica", Me
                
                If rstID_Codautomacao.BOF = True And rstID_Codautomacao.EOF = True Then
                    MsgBox "Código de automação não cadastrado.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Set rstID_Codautomacao = Nothing
                    Exit Sub
                Else
                    strID_Produto = rstID_Codautomacao!PKId_TBProduto
                End If
                
                Set rstID_Codautomacao = Nothing
                
                strSQL = Empty
                strSQL = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco, " & _
                         "TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                         "FROM TBItens_tabela_preco " & _
                         "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                         "WHERE FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) AND " & _
                         "FKId_TBProduto = " & strID_Produto & ""
                         
                Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstInf_Produtos, "Otica", Me)
                
                If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
                    MsgBox "Código de barra não cadastrado.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Set rstInf_Produtos = Nothing
                    Exit Sub
                End If
            
                If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
                    MsgBox "Preço do Item não cadastrado.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Set rstInf_Produtos = Nothing
                    Exit Sub
                End If
                
                Call Verifica_ICMS
                
                If booICMS_cadastrado = False Then
                    MsgBox "Aliquota de ICMS não cadastrada.Verifique!", vbCritical, "Only Tech"
                    txtCodigo_Produto.Text = Empty
                    txtCodigo_Produto.SetFocus
                    Set rstInf_Produtos = Nothing
                    Exit Sub
                End If
                    
                txtDescricao_Produto.Text = rstInf_Produtos!DFDescricao_resumida_TBProduto
                txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
                txtQuantidade_Produto.Text = 1
                
                If Not IsNull(rstInf_Produtos!DFPath_imagem_TBProduto) Then
                   Call Carrega_imagem
                End If
                
                Call Reposicao
            End If
        End If
    Else
    
        If booAcionado_Fechamento_cupom = False Then
           Me.txtCodigo_Produto.SetFocus
        End If
        
        booAcionado_Fechamento_cupom = True
    End If
    
    'Quantidade
    If txtQuantidade_Produto.TabStop = True Then
        If Me.txtCodigo_Produto.Text <> " " And Me.txtCodigo_Produto.Text <> "" Then
           Me.txtQuantidade_Produto.SetFocus
           On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
        Else
           Me.txtCodigo_Produto.SetFocus
        End If
    End If
    'Preço Unitário
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

Private Sub txtPreco_Unitario_LostFocus()
     If txtPreco_Unitario.Text <> "" Then
        txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
        Call Processando_item
        txtCodigo_Produto.SetFocus
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

     
     'Acessando o parametro para o ECF
     strSQL = Empty
     strSQL = "SELECT * FROM TBPARAMETROS_ECF"
     Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstParametro_ecf, "Otica", Me

     'Calculando item
     If Not txtPreco_Unitario.Text = "" And Not txtQuantidade_Produto.Text = "" Then
        txtPreco_Total = Format(CDbl(txtPreco_Unitario.Text) * CDbl(txtQuantidade_Produto.Text), "#,###0.00")
     End If
     
     'Verificando se existe o item no cupom....se existir ele adiciona sua respectiva quantidade e preço no item já cadastrado
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
     
     Dim strCodigo_Produto As String * 14
     Dim strDescricao_Produto As String
     Dim strAliquota As String
     Dim strTipo_quantidade As String * 1
     Dim strQuantiade As String
     Dim strCasas_Decimais As String * 1
     Dim strValor_Unitario As String
     Dim strTipo_desconto As String * 1
     Dim strValor_desconto As String * 8
     
     strCodigo_Produto = Me.txtCodigo_Produto.Text
     strDescricao_Produto = Me.txtDescricao_Produto.Text
     strQuantiade = Me.txtQuantidade_Produto.Text
     strValor_Unitario = txtPreco_Unitario.Text
     strValor_desconto = "0,00"
     
     'Montando dysplay de itens de cupom
     HflexGrid.Cols = 6
     HflexGrid.ColWidth(0) = 0
     HflexGrid.Rows = HflexGrid.Rows + 2
     
     If HflexGrid.Rows = 4 Then
        'Cabeçalho 1
        HflexGrid.Row = 0
        HflexGrid.Col = 1
        HflexGrid.FixedAlignment(1) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = "Código"
        HflexGrid.Col = 2
        HflexGrid.Text = "Descrição"
        HflexGrid.Col = 3
        HflexGrid.Text = "Qtd."
        HflexGrid.Col = 4
        HflexGrid.Text = "X"
        HflexGrid.Col = 5
        HflexGrid.Text = "Vlr.Unit."
        'Cabeçalho 2
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
     HflexGrid.Text = Format(CDbl(strQuantiade) * CDbl(strValor_Unitario), "#,###0.00")
     
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
     
     txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     
     Call Limpa_Tela
     
     Set rstInf_Produtos = Nothing
     Set rstParametro_ecf = Nothing
     
     txtQuantidade_Produto.TabStop = False
     txtPreco_Unitario.TabStop = False
     txtCodigo_Produto.SetFocus
    
End Function

Private Function Reposicao()

    txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
    txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.00")
    
    'Verificando se vai passar pela quantidade e no antes de calcular
    If Me.txtQuantidade_Produto.TabStop = False And Me.txtPreco_Unitario.TabStop = False Then
        Call Processando_item
    End If
    
End Function

Private Function Carrega_imagem()

    On Error GoTo Erro_imagem
    
    Me.imgProduto.Picture = LoadPicture(rstInf_Produtos!DFPath_imagem_TBProduto)
        
Fim_Imagem:
    
    Exit Function
    
Erro_imagem:

    If Err.Number <> 76 Then
       Erro.Erro Me, "Otica"
    Else
       GoTo Fim_Imagem
    End If

End Function

Private Sub txtQuantidade_Produto_LostFocus()
    If Me.txtQuantidade_Produto.TabStop = True Then
        txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
        txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.00")
        Call Processando_item
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
    strSQL = Empty
    strSQL = "SELECT TBCidade_otica.DFUf_TBCidade_otica FROM TBEmpresa " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBEmpresa.Fkid_TBCidade_otica  = TBCidade_otica.pkid_TBCidade_otica " & _
             "WHERE TBEmpresa.PKCodigo_TBempresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstUF, "Otica", Me)
    
    strUF_Emitente = rstUF!DFUf_TBCidade_otica
    
    Set rstUF = Nothing
     
    'Calculando a parte do ICMS relacionado ao Item
    'Concatenando o valor da Situação Tributária que está no cadastro de produto
    strST = rstInf_Produtos!DFCst1_TBProduto
    strST2 = rstInf_Produtos!DFCst2_TBProduto
    
    'ICMS E ST
    'Verifica se a ST for 030 ou 060 o valor da aliquota e o valor de ICMS é 0;
    'E Grava na tabela CFO_Pedido mais uma CFO para este pedido
    If strST = "030" Or strST = "060" Then
    
       dblAliquota_icms = 0
       dblTotal_Icms = 0
       
       Dim rstVerifica_Estado_ST As New ADODB.Recordset
       Dim rstCFO_ST As New ADODB.Recordset
       
       strSQL = Empty
       strSQL = "SELECT TBCidade_otica.DFUf_TBCidade_otica " & _
                "FROM TBEmpresa " & _
                "INNER JOIN TBCidade_otica " & _
                "ON TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                "WHERE PKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                
       Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstVerifica_Estado_ST, "Otica", Me)
       
       If rstVerifica_Estado_ST!DFUf_TBCidade_otica = strUF_Emitente Then
          'Localizando no parametro o proximo cfo de substituição para dentro do estado
          strSQL = Empty
          strSQL = "SELECT DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
                   "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
          
          Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstCFO_ST, "Otica", Me
       Else
          'Localizando no parametro o proximo cfo de substituição para dentro do estado
          strSQL = Empty
          strSQL = "SELECT DFProximo_cfop_venda_fora_estado_substituicao_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
                   "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
          Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstCFO_ST, "Otica", Me
       End If
       
       'Localizando o ID do CFO
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'Lembrar Marcos para fazer teste caso o produto nao                             '
       'esteja cadastrado no estado para ICMS(Giordano).                               '
       'alteração feita na busca do ID do CFO (ERRO de passagem de valor para a funcao)'
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstCFO_ST.BOF = True And rstCFO_ST.EOF = True Then
          MsgBox "Verifique se o CFO na tabela de parâmetros fiscais está preenchida corretamente!", vbCritical, "Only Tech"
       End If
         
       intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", rstCFO_ST.Fields("CFO"), "TBCFOP", "Otica", Me, "BDRetaguarda")
            
       If rstCFO_ST.BOF = True And rstCFO_ST.EOF = True Then
          MsgBox "Verifique se o CFO na tabela de parâmetros fiscais está preenchida corretamente!", vbCritical, "Only Tech"
       End If
       
       Set rstVerifica_Estado_ST = Nothing
       Set rstCFO_ST = Nothing
    Else
        Dim rstVerifica_Estado_ICMS As New ADODB.Recordset
        'Query para pegar ICMS do item
        strSQL = Empty
        strSQL = "SELECT " & _
                 "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                 "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                 "FROM TBEstado_icms " & _
                 "WHERE FKId_TBProduto = " & rstInf_Produtos!PKId_TBProduto & " " & _
                 "AND DFUf_TBEstado_icms = '" & strUF_Emitente & "'"
                 
        Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
        
        If rstVerifica_Estado_ICMS.BOF = True And rstVerifica_Estado_ICMS.EOF = True Then
           Set rstVerifica_Estado_ICMS = Nothing
           'Query para pegar ICMS do item, com estado **
           strSQL = Empty
           strSQL = "SELECT " & _
                    "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                    "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                    "FROM TBEstado_icms " & _
                    "WHERE FKId_TBProduto = " & rstInf_Produtos!PKId_TBProduto & " " & _
                    "AND DFUf_TBEstado_icms = '**' "
                          
           Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
        End If
        
        If rstVerifica_Estado_ICMS.BOF = True And rstVerifica_Estado_ICMS.EOF = True Then
           booICMS_cadastrado = False
        End If
        
    End If

End Function

Private Function Adiciona_Item_Existente()

     HflexGrid.Col = 3
     HflexGrid.Text = CDbl(HflexGrid.Text) + CDbl(Me.txtQuantidade_Produto.Text)
     
     HflexGrid.Row = HflexGrid.Row + 1
     HflexGrid.Col = 5
     
     HflexGrid.Text = Format(CDbl(HflexGrid.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     
     Me.txtPreco_total_cupom.Text = Format(CDbl(Me.txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     
     Me.HflexGrid.SetFocus
     Me.HflexGrid.Row = Me.HflexGrid.Rows - 1
     
     Call Limpa_Tela
     
     Set rstInf_Produtos = Nothing
     Set rstParametro_ecf = Nothing
     
     txtQuantidade_Produto.TabStop = False
     txtPreco_Unitario.TabStop = False
     txtCodigo_Produto.SetFocus

End Function
