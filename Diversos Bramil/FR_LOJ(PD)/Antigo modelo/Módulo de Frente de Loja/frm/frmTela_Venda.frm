VERSION 5.00
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11490
      Top             =   7800
   End
   Begin VB.TextBox txtStatus 
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
      Left            =   30
      MaxLength       =   40
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8610
      Width           =   10875
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
      Left            =   8010
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
      Top             =   6450
      Width           =   2325
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
      Left            =   7890
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
      Left            =   3900
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6450
      Width           =   3165
   End
   Begin VB.ListBox lstItens_Cupom 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      ItemData        =   "frmTela_Venda.frx":340E
      Left            =   7350
      List            =   "frmTela_Venda.frx":3410
      TabIndex        =   21
      Top             =   750
      Width           =   3945
   End
   Begin VB.Image imgInd_pouco_papel 
      Height          =   255
      Left            =   11550
      Picture         =   "frmTela_Venda.frx":3412
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   10980
      Shape           =   3  'Circle
      Top             =   8640
      Width           =   225
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   11400
      Top             =   8610
      Width           =   525
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   10890
      Top             =   8610
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   7320
      X2              =   11400
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Image imgLogo_Empresa 
      Height          =   2055
      Left            =   540
      Picture         =   "frmTela_Venda.frx":371C
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
      Top             =   510
      Width           =   4275
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   5115
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4275
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   7770
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
      Left            =   7770
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
      Left            =   3720
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
      Width           =   2535
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
      Left            =   7680
      Shape           =   4  'Rounded Rectangle
      Top             =   7410
      Width           =   3705
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   7650
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
      Left            =   3630
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
      Width           =   2535
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
      Left            =   3270
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
      Left            =   3780
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
      Left            =   7800
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
      Left            =   6810
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
      Left            =   7350
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
Dim strSql As String
Dim rstInf_Produtos As New ADODB.Recordset
Dim booImpressora_lacrada As Boolean
Dim rstEmpresa As New ADODB.Recordset
Public strOperador As String
Public strEmpresa_Operador As String
Public booInterrompe_venda As Boolean
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Verifica se foi preciona do F2 e abre tela de consulta
    If KeyCode = 113 Then
        frmConsulta_Produto.Show (1)
    End If
    
    'Verifica se foi preciona do F3 e abre novo cupom
    If KeyCode = 114 Then
        Call Abre_cupom
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi preciona do F4 e Finaliza a Compra
    If KeyCode = 115 Then
        Call Finaliza_Cupom
        txtCodigo_Produto.SetFocus
    End If
    
    'Verifica se foi preciona do F6 e Cancela
    If KeyCode = 117 Then
        Call Cancela_cupom
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
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    strSql = Empty
    strSql = "SELECT * FROM TBEmpresa Where PKCodigo_TBEmpresa = " & strEmpresa_Operador & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstEmpresa, "Otica", Me
    
    
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    
    LocalRetorno = LeParametrosIni("Sistema", "Retorno")
    
    If LocalRetorno = "-2" Then
        LocalRetorno = "0" 'devolve o retorno na variavel
    Else
        LocalRetorno = Left(LocalRetorno, 1)
    End If
    
    Retorno = Bematech_FI_AbrePortaSerial
   
    If Retorno = -4 Or Retorno = -5 Then
        MsgBox "Erro ao acessar a porta de comunicação com a impressora.Verifique! A aplicação está imposibilitada de ser iniciada", vbCritical, "Only Tech"
        End
    End If
    
    '--- Verificações de periféricos e componentes ---------------------------------------------------------
    
    'Verificar se impressora está ligada.
    Retorno = Bematech_FI_VerificaImpressoraLigada()
    If Retorno = -6 Then
       MsgBox "A Impressora se encontra DESLIGADA.Verifique! A aplicação está imposibilitada de ser iniciada", vbInformation + vbOKOnly, "Atenção"
       End
    End If
    
    'Verifica se a impressora está online ou em intervenção
    Dim strModo As String
    
    strModo = Space(1)
    
    Retorno = Bematech_FI_VerificaModoOperacao(strModo)
    
    If Not strModo = "1" Then
       MsgBox "A Impressora se encontra em Intervenção Técnica.Verifique! A aplicação está imposibilitada de ser iniciada", vbInformation + vbOKOnly, "Atenção"
       Call VerificaRetornoImpressora("", "", "Modo Operação")
    End If
    
    Dim strRetorno_status As String
    Dim strValor_retorno As String
    
    'Verificando a bobina de papel
    strRetorno_status = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
    strValor_retorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
    
    'Verificando se a bobina está acabando
    If (ST1 >= 64) Then
        imgInd_pouco_papel.Visible = True
    End If
    
    If (ST1 >= 128) Then
        MsgBox "Impressora sem bobina.Troque antes de iniciar a venda.", vbInformation, "Only Tech"
    End If
    
    '-------------------------------------------------------------------------------------------------------
    'Informações pertinentes à lei
    txtNumero_check_out.Text = "ECF: " & "001"
    txtNumero_Nome_Operadora.Text = "Operador: " & strOperador
    txtVersao_software.Text = "Versão 1.0"
    txtNumero_loja.Text = "Loja: " & rstEmpresa!PKCodigo_TBEmpresa
    imgLogo_Empresa.Picture = LoadPicture(rstEmpresa!DFPath_logomarca_TBEmpresa)
    
    Set rstEmpresa = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    Retorno = Bematech_FI_FechaPortaSerial()
    Call VerificaRetornoImpressora("", "", "BemaFI32")
    End
End Sub

Private Sub Timer1_Timer()
    txtData_Hora.Text = "Data: " & Format(Now, "DD/MM/YYYY") & " - " & Format(Now, "hh:mm:ss")
End Sub

Private Sub txtCodigo_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Produto_LostFocus()
    'Verifica Código interno
    If Len(Me.txtCodigo_Produto.Text) < 7 And txtCodigo_Produto.Text <> "" Then
        
        strSql = Empty
        strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco,DFPath_imagem_TBProduto " & _
                 "FROM TBProduto " & _
                 "INNER JOIN TBITENS_TABELA_PRECO " & _
                 "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                 "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
                 "AND TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & ""
                 
        Movimentacoes.Select_geral strSql, "BDRetaguarda", rstInf_Produtos, "Otica", Me
        
        If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
            MsgBox "Código Interno não cadastrado.Verifique!", vbCritical, "Only Tech"
            txtCodigo_Produto.Text = Empty
            txtCodigo_Produto.SetFocus
            Set rstInf_Produtos = Nothing
            Exit Sub
        End If
        
        If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Then
            MsgBox "Preço do Item não cadastrado.Verifique!", vbCritical, "Only Tech"
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
    'Verifica Código de Barra
    If Len(Me.txtCodigo_Produto.Text) > 7 And txtCodigo_Produto.Text <> "" Then
    End If
    
    'Verifica Etiqueta peso
    If Len(Me.txtCodigo_Produto.Text) < 7 And txtCodigo_Produto.Text <> "" Then
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
        Call Processando_item
        txtCodigo_Produto.SetFocus
     End If
End Sub

Private Function Limpa_tela()
    txtCodigo_Produto.Text = Empty
    txtDescricao_Produto.Text = Empty
    txtPreco_Total.Text = Empty
    txtPreco_Unitario = Empty
    txtQuantidade_Produto.Text = Empty
    imgProduto.Picture = LoadPicture("")
End Function

Private Function Processando_item()

     Dim rstParametro_ecf As New ADODB.Recordset
     
     'Acessando o parametro para o ECF
     strSql = Empty
     strSql = "SELECT * FROM TBPARAMETROS_ECF"
     Movimentacoes.Select_geral strSql, "BDRetaguarda", rstParametro_ecf, "Otica", Me

     'Calculando item
     If Not txtPreco_Unitario.Text = "" And Not txtQuantidade_Produto.Text = "" Then
        txtPreco_Total = Format(CDbl(txtPreco_Unitario.Text) * CDbl(txtQuantidade_Produto.Text), "#,###0.00")
     End If
     
     Dim strCodigo_Produto As String
     Dim strDescricao_Produto As String * 29
     Dim strAliquota As String
     Dim strTipo_quantidade As String * 1
     Dim strQuantiade As String * 7
     Dim strCasas_Decimais As String * 1
     Dim strValor_Unitario As String * 8
     Dim strTipo_desconto As String * 1
     Dim strValor_desconto As String * 8
     
     strCodigo_Produto = Me.txtCodigo_Produto.Text
     strDescricao_Produto = Me.txtDescricao_Produto.Text
     
     '--- Aliquotas --------------------------------------------------------------------------------------
     
     Dim rstAliqota As New ADODB.Recordset
     
     'Query para localizar a aliquota do item dentro da UF da empresa
     strSql = "SELECT DFPercentual_icms_saida_fisica_TBEstado_icms FROM TBEstado_icms " & _
              "INNER JOIN TBPRODUTO " & _
              "ON TBEstado_icms.FKId_TBProduto = TBPRODUTO.PKId_TBProduto " & _
              "WHERE DFUf_TBEstado_icms  = (SELECT DFUf_TBCidade_otica FROM TBEMPRESA INNER JOIN TBCidade_otica ON TBEMPRESA.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica ) " & _
              "AND TBPRODUTO.IXCodigo_TBProduto = " & txtCodigo_Produto.Text & " "
              
     Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAliqota, "Otica", Me
     
     If rstAliqota.BOF = True And rstAliqota.EOF = True Then
     
         Set rstAliqota = Nothing
         
        'Query para localizar a aliquota do item dentro da UF "**"
         strSql = "SELECT DFPercentual_icms_saida_juridica_TBEstado_icms FROM TBEstado_icms " & _
                  "INNER JOIN TBPRODUTO " & _
                  "ON TBEstado_icms.FKId_TBProduto = TBPRODUTO.PKId_TBProduto " & _
                  "WHERE DFUf_TBEstado_icms  = '**' " & _
                  "AND TBPRODUTO.IXCodigo_TBProduto = " & txtCodigo_Produto.Text & " "
              
         Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAliqota, "Otica", Me
     End If
     
     If rstAliqota.BOF = True And rstAliqota.EOF = True Then
        MsgBox "Este item não possui aliquota de ICMS cadastrada.Verifique!", vbCritical, "Only Tech"
        Call Limpa_tela
        Set rstAliqota = Nothing
        Exit Function
     End If
     
     If booImpressora_lacrada = False Then
        strAliquota = "1200"
     Else
        strAliquota = rstAliqota!DFPercentual_icms_saida_juridica_TBEstado_icms & "00"
     End If
     
     Set rstAliqota = Nothing
     
     '------------------------------------------------------------------------------------------------------
     
     'Verifica se existe reg. no parâmetro
     If rstParametro_ecf.EOF = True And rstParametro_ecf.BOF = True Then
        MsgBox "Verifique as informações contidas no parâmetro do concentrador!Item impossibilitado de ser incluido neste cupom", vbCritical, "Only Tech"
        Call Limpa_tela
        Exit Function
     End If
     
     strTipo_quantidade = rstParametro_ecf!DFTipo_quantidade_TBParametros_ecf
     strQuantiade = Me.txtQuantidade_Produto.Text
     strCasas_Decimais = rstParametro_ecf!DFNumero_decimais_TBParametros_ecf
     strValor_Unitario = Me.txtPreco_Unitario.Text
     strTipo_desconto = rstParametro_ecf!DFTipo_desconto_TBParametros_ecf
     strValor_desconto = "0,00"
          
     Retorno = Bematech_FI_VendeItem(strCodigo_Produto, strDescricao_Produto, strAliquota, strTipo_quantidade, strQuantiade, strCasas_Decimais, strValor_Unitario, strTipo_desconto, strValor_desconto)

     'Função que analisa o retorno da impressora
     Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
     
     'Verifica retorno da impressora e interrompe a venda
     If booInterrompe_venda = True Then
        Call Limpa_tela
        Me.txtCodigo_Produto.SetFocus
        Exit Function
     End If
     
     lstItens_Cupom.AddItem txtCodigo_Produto
     
     If txtPreco_total_cupom.Text = "" Then txtPreco_total_cupom.Text = 0
     If txtPreco_Total.Text = "" Then txtPreco_Total.Text = 0
     
     txtPreco_total_cupom.Text = Format(CDbl(txtPreco_total_cupom.Text) + CDbl(Me.txtPreco_Total.Text), "#,###0.00")
     
     Call Limpa_tela
     
     Set rstInf_Produtos = Nothing
     Set rstParametro_ecf = Nothing
     
     txtCodigo_Produto.SetFocus
    
End Function

Private Sub txtQuantidade_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Abre_cupom()
        
    Retorno = Bematech_FI_AbreCupom("08706114709")
     
    'Função que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Abertura de Cupom")
    
End Function

Private Function Cancela_cupom()

    Retorno = Bematech_FI_CancelaCupom()
    
    'Função que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
    
End Function

Private Function Finaliza_Cupom()

End Function

Private Function Reposicao()

    txtPreco_Unitario.Text = Format(txtPreco_Unitario, "#,###0.00")
    txtPreco_Total.Text = Format(txtPreco_Total, "#,###0.00")
    
    Call Processando_item
    
End Function

Private Function Carrega_imagem()

    On Error GoTo Erro_imagem
    
    Me.imgProduto.Picture = LoadPicture(rstInf_Produtos!DFPath_imagem_TBProduto)
        
Fim_Imagem:
    
    Exit Function
    
Erro_imagem:

    If Err.Number <> 76 Then
       erro.erro Me, "Otica"
    Else
       GoTo Fim_Imagem
    End If

End Function
