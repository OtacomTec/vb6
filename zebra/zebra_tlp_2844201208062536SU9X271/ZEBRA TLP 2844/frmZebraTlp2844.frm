VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{65968E46-BFE1-4C60-83AA-C79112B8F133}#2.0#0"; "Buttom.ocx"
Begin VB.Form frmZebraTlp2844 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZEBRA TLP 2844"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmZebraTlp2844.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   9105
      Left            =   0
      ScaleHeight     =   9045
      ScaleWidth      =   13485
      TabIndex        =   8
      Top             =   0
      Width           =   13545
      Begin VB.TextBox txtNumCompra 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   6105
         MaxLength       =   7
         TabIndex        =   1
         Top             =   127
         Width           =   1875
      End
      Begin VB.TextBox txtNumNota 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   9660
         MaxLength       =   15
         TabIndex        =   2
         Top             =   127
         Width           =   1575
      End
      Begin VB.TextBox txtValUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   11310
         TabIndex        =   6
         Top             =   1050
         Width           =   1215
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   8970
         TabIndex        =   5
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5370
         MaxLength       =   30
         TabIndex        =   4
         ToolTipText     =   "Caso não saiba o código do produto, informar o nome aqui e prescionar ENTER."
         Top             =   1050
         Width           =   2625
      End
      Begin VB.TextBox txtCodPrd 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2190
         TabIndex        =   3
         Top             =   1050
         Width           =   1995
      End
      Begin VB.PictureBox Picture4 
         Height          =   375
         Left            =   12570
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   1050
         Width           =   375
         Begin VB.CommandButton cmdAddPrd 
            Caption         =   "+"
            Height          =   315
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Adciona produtos na listagem."
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   12990
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   1050
         Width           =   375
         Begin VB.CommandButton cmdRemPrd 
            Caption         =   "-"
            Height          =   315
            Left            =   0
            TabIndex        =   10
            ToolTipText     =   "Remove produtos da listagem"
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.ComboBox cboLoja 
         Height          =   345
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   1995
      End
      Begin TabDlg.SSTab sstProduto 
         Height          =   4260
         Left            =   900
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   7514
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "PRODUTOS"
         TabPicture(0)   =   "frmZebraTlp2844.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdTabelaConsultaPrd"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdFecharPesqPrd"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Picture9"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin VB.PictureBox Picture9 
            Height          =   435
            Left            =   7950
            ScaleHeight     =   375
            ScaleWidth      =   1545
            TabIndex        =   14
            Top             =   3720
            Width           =   1605
            Begin VB.CommandButton cmdSelecionarPrd 
               BackColor       =   &H00C0FFFF&
               Caption         =   "&SELECIONAR"
               Height          =   375
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.CommandButton cmdFecharPesqPrd 
            Caption         =   "X"
            Height          =   285
            Left            =   9210
            TabIndex        =   13
            Top             =   0
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grdTabelaConsultaPrd 
            Height          =   3225
            Left            =   90
            TabIndex        =   16
            Top             =   420
            Width           =   9435
            _ExtentX        =   16642
            _ExtentY        =   5689
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   -2147483630
            BackColorFixed  =   16744576
            BackColorSel    =   8421504
            ForeColorSel    =   16777215
            BackColorBkg    =   14737632
            GridColor       =   0
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdTabelaPrd 
         Height          =   7035
         Left            =   120
         TabIndex        =   17
         Top             =   1500
         Width           =   13245
         _ExtentX        =   23363
         _ExtentY        =   12409
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BackColorFixed  =   16744576
         BackColorSel    =   8421504
         ForeColorSel    =   16777215
         BackColorBkg    =   14737632
         GridColor       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Buttom.Buttom_Mega cmdLocalizar 
         Height          =   495
         Left            =   11310
         TabIndex        =   18
         ToolTipText     =   "Realiza o pre-enchimento automatico dos produtos para imprimir etiqueta apartir do numero da compra ou numero da nota."
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "LOCALIZAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmZebraTlp2844.frx":0CE6
         PICN            =   "frmZebraTlp2844.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   60
         Picture         =   "frmZebraTlp2844.frx":19DC
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NÚMERO DA COMPRA:"
         Height          =   225
         Left            =   4260
         TabIndex        =   28
         Top             =   180
         Width           =   1785
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO DA NOTA:"
         Height          =   225
         Left            =   8070
         TabIndex        =   27
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO UNIT:"
         Height          =   225
         Left            =   10110
         TabIndex        =   26
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTD:"
         Height          =   225
         Left            =   8070
         TabIndex        =   25
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO:"
         Height          =   225
         Left            =   4260
         TabIndex        =   24
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓDIGO:"
         Height          =   225
         Left            =   1410
         TabIndex        =   23
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1410
         TabIndex        =   22
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTIDADE TOTAL..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   8670
         Width           =   2565
      End
      Begin VB.Label lblQtdTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   8670
         Width           =   540
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1410
         X2              =   12870
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOJA..:"
         Height          =   225
         Left            =   1410
         TabIndex        =   19
         Top             =   180
         Width           =   735
      End
   End
   Begin Buttom.Buttom_Mega cmdFechar 
      Height          =   615
      Left            =   11670
      TabIndex        =   29
      Top             =   9120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&FECHAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmZebraTlp2844.frx":3C16
      PICN            =   "frmZebraTlp2844.frx":3C32
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Buttom.Buttom_Mega cmdImprimir 
      Height          =   615
      Left            =   9810
      TabIndex        =   30
      Top             =   9120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&IMPRIMIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmZebraTlp2844.frx":490C
      PICN            =   "frmZebraTlp2844.frx":4928
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Buttom.Buttom_Mega cmdConfigurar 
      Height          =   615
      Left            =   3720
      TabIndex        =   31
      Top             =   9120
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&CONFIGURAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmZebraTlp2844.frx":5602
      PICN            =   "frmZebraTlp2844.frx":561E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmZebraTlp2844"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    On Error GoTo erro
    
    VBA.DoEvents
    AplicaBorda Me
    
    Call CarregaComboEspecial_Loja_e_Filiais
    
    If cboLoja.ListCount > 0 Then cboLoja.ListIndex = 0
    
    KeyPreview = True
    
    Call sbMontaGrid
    Call sbMontaGridConsultaProdutos
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then VBA.SendKeys ("{TAB}")
    If KeyAscii > 0 Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub cmdConfigurar_Click()
    'frmArgoxConfig.Show vbModal
End Sub

Private Sub cmdRemPrd_Click()
    
    On Error GoTo erro
    
    With grdTabelaPrd
        If .Rows = 1 Then
            VBA.MsgBox "Não existem mais produtos para remover.", vbExclamation, gblStrNomeModulo
        ElseIf .Rows = 2 Then
            'Diminui as quantidade e o valor total da nota
            lblQtdTotal.Caption = VBA.Format(fcVarChekValor(lblQtdTotal.Caption) - fcVarChekValor(.TextMatrix(.Row, 2)), String(4, "0"))
            '---------------------------------------------
            .Rows = .Rows - 1
        Else
            'Diminui as quantidade e o valor total da nota
            lblQtdTotal.Caption = VBA.Format(fcVarChekValor(lblQtdTotal.Caption) - fcVarChekValor(.TextMatrix(.Row, 2)), String(4, "0"))
            '---------------------------------------------
            .RemoveItem (.Row)
        End If
    End With
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number, Me.Name & ".cmdRemPrd_Click")
End Sub

Private Sub cmdLocalizar_Click()
    
    On Error GoTo erro
    
    Dim rsTabela                As Recordset
    Dim rsTabelaCmpPrd          As Recordset
    Dim clsCompras              As New clsCompras
    Dim clsComprasPrd           As New clsComprasPrd
    
    grdTabelaPrd.Rows = 1
    
    If VBA.Trim(txtNumCompra.Text) = "" And VBA.Trim(txtNumNota.Text) = "" Then
        VBA.MsgBox "Informe o Numero da Compra ou Numero da Nota Fiscal, antes de prosseguir.", vbExclamation, gblStrNomeModulo
        txtNumCompra.SetFocus
        GoTo SairSub
    End If
    
    Set rsTabela = clsCompras.Compras_Consultar(txtNumCompra.Text, , , , , , , , , txtNumNota.Text)
    Do While Not rsTabela.EOF
        Set rsTabelaCmpPrd = clsComprasPrd.ComprasProdutos_Consultar(rsTabela!cmpcCodigo)
        Do While Not rsTabelaCmpPrd.EOF
            Call sbPreencheGrid(rsTabelaCmpPrd!prdcCodigo, _
                                fcRetornaDadosDoProduto(rsTabelaCmpPrd!prdcCodigo, "NOME"), _
                                fcVarChekValor(rsTabelaCmpPrd!cmpnQtd), _
                                fcRetornaDadosDoProduto(rsTabelaCmpPrd!prdcCodigo, "VL_UNIT"))
            rsTabelaCmpPrd.MoveNext
        Loop
        rsTabelaCmpPrd.Close
        rsTabela.MoveNext
    Loop
    rsTabela.Close
    
SairSub:
    Set rsTabela = Nothing
    Set rsTabelaCmpPrd = Nothing
    Set clsCompras = Nothing
    Set clsComprasPrd = Nothing
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number, Me.Name & ".")
End Sub

Private Sub cmdAddPrd_Click()
        
    On Error GoTo erro
    
    If VBA.Trim(txtCodPrd.Text) = "" Then
        VBA.MsgBox "Informe o produto antes de adcionar.", vbExclamation, gblStrNomeModulo
        txtCodPrd.SetFocus
        GoTo SairSub
    End If
        
    'Caso o estabelecimento possua controle de dt de validade dos produtos
    'e o produto sendo inserido tb possua dt de validade o sistema ira
    'gravar o numero do lote e dt de vencimento, em duas colunas ocultas.
    Call sbPreencheGrid(txtCodPrd.Text, _
                        txtDescricao.Text, _
                        txtQtd.Text, _
                        txtValUnit.Text)
    
SairSub:
    txtCodPrd.Text = ""
    txtDescricao.Text = ""
    txtValUnit.Text = ""
    txtQtd.Text = ""
    txtCodPrd.SetFocus
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub sbPreencheGrid(ByVal CodigoPrd As String, _
                           ByVal Descricao As String, _
                           ByVal Qtd As String, _
                           ByVal PrecoUnit As String)
    
    On Error GoTo erro
    
    DoEvents
    With grdTabelaPrd
        .AddItem " ", 1
        .TextMatrix(1, 0) = CodigoPrd
        .TextMatrix(1, 1) = Descricao
        .TextMatrix(1, 2) = Qtd
        .TextMatrix(1, 3) = VBA.Format(fcVarChekValor(PrecoUnit), "Currency")
    End With
    DoEvents
    
    With grdTabelaPrd
        If .Rows > 1 Then
            .Col = 0
            .Sort = 5
        End If
    End With
    
    lblQtdTotal.Caption = VBA.Format(fcVarChekValor(lblQtdTotal.Caption) + fcVarChekValor(Qtd), String(4, "0"))
        
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If VBA.Trim(txtCodPrd.Text) = "" And VBA.Trim(txtDescricao.Text) <> "" Then
            Call sbLocalizarProdutoPeloNome(txtDescricao.Text)
        End If
    End If
End Sub

Private Sub sbLocalizarProdutoPeloNome(ByVal Nome As String)
    
    On Error GoTo erro
        
    Dim rsTabela    As Recordset
    Dim clsProdutos As New clsProdutos
    
    Set rsTabela = clsProdutos.Produtos_Consultar(, Nome)
    If Not rsTabela.EOF Then
        rsTabela.MoveLast
        If rsTabela.RecordCount = 1 Then
            txtCodPrd.Text = rsTabela!prdcCodigo
            txtDescricao.Text = rsTabela!prdcDescricao
            txtValUnit.Text = VBA.Format(rsTabela!prdnValorAvista, "Currency")
            txtQtd.Text = fcVarChekValor(rsTabela!prdnEstoqueAtual)
        Else
            grdTabelaConsultaPrd.Rows = 1
            rsTabela.MoveFirst
            sstProduto.Visible = True
            Do While Not rsTabela.EOF
                DoEvents
                With grdTabelaConsultaPrd
                    .AddItem " ", 1
                    .TextMatrix(1, 0) = rsTabela!prdcCodigo
                    .TextMatrix(1, 1) = rsTabela!prdcDescricao
                    .TextMatrix(1, 2) = fcVarChekValor(rsTabela!prdnEstoqueAtual)
                    .TextMatrix(1, 3) = VBA.Format(rsTabela!prdnValorAvista, "Currency")
                End With
                DoEvents
                rsTabela.MoveNext
            Loop
            grdTabelaConsultaPrd.SetFocus
        End If
    End If
        
    rsTabela.Close
    Set rsTabela = Nothing
    Set clsProdutos = Nothing
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub txtCodPrd_LostFocus()
    
    On Error GoTo erro
    
    If VBA.Trim(txtCodPrd.Text) <> "" Then
        Call sbLocalizarProdutoPorParteDoCodigo(txtCodPrd.Text)
    Else
        txtDescricao.Text = ""
        txtValUnit.Text = ""
    End If
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number, Me.Name & ".txtCodPrd_LostFocus")
End Sub

Private Sub sbLocalizarProdutoPorParteDoCodigo(ByVal ParteDoCodigo As String)
    
    On Error GoTo erro
        
    Dim rsTabela        As Recordset
    Dim clsProdutos     As New clsProdutos
    
    If VBA.Trim(txtCodPrd.Text) <> "" Then
        'Pesquisa primeiro pelo codigo de forma integra caso encontre considera o primeiro
        'caso contrario busca pelo codigo em suas partes
        Set rsTabela = clsProdutos.Produtos_Consultar(ParteDoCodigo)
        If Not rsTabela.EOF Then
            txtCodPrd.Text = rsTabela!prdcCodigo
            txtDescricao.Text = rsTabela!prdcDescricao
            txtValUnit.Text = VBA.Format(rsTabela!prdnValorAvista, "Currency")
            txtQtd.Text = fcVarChekValor(rsTabela!prdnEstoqueAtual)
        Else
            Set rsTabela = clsProdutos.Produtos_Consultar(, , , ParteDoCodigo)
            If Not rsTabela.EOF Then
                rsTabela.MoveLast
                If rsTabela.RecordCount = 1 Then
                    txtCodPrd.Text = rsTabela!prdcCodigo
                    txtDescricao.Text = rsTabela!prdcDescricao
                    txtValUnit.Text = VBA.Format(rsTabela!prdnValorAvista, "Currency")
                    txtQtd.Text = fcVarChekValor(rsTabela!prdnEstoqueAtual)
                Else
                    grdTabelaConsultaPrd.Rows = 1
                    rsTabela.MoveFirst
                    sstProduto.Visible = True
                    Do While Not rsTabela.EOF
                        DoEvents
                        With grdTabelaConsultaPrd
                            .AddItem " ", 1
                            .TextMatrix(1, 0) = rsTabela!prdcCodigo
                            .TextMatrix(1, 1) = rsTabela!prdcDescricao
                            .TextMatrix(1, 2) = fcVarChekValor(rsTabela!prdnEstoqueAtual)
                            .TextMatrix(1, 3) = VBA.Format(rsTabela!prdnValorAvista, "Currency")
                        End With
                        DoEvents
                        rsTabela.MoveNext
                    Loop
                    grdTabelaConsultaPrd.SetFocus
                End If
            Else
                VBA.MsgBox "Produto não encontrado e ou não cadastrado.", vbExclamation, gblStrNomeModulo
            End If
        End If
        rsTabela.Close
    End If
    
    Set rsTabela = Nothing
    Set clsProdutos = Nothing
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub sbMontaGrid()

    On Error GoTo erro
    
    'DESCRICAO DAS COLUNAS
    With grdTabelaPrd
        .Cols = 4
        .TextMatrix(0, 0) = "CÓDIGO"
        .TextMatrix(0, 1) = "DESCRIÇÃO"
        .TextMatrix(0, 2) = "QUANTIDADE"
        .TextMatrix(0, 3) = "PREÇO UNITARIO"
    End With
    'LARGURA DAS COLUNAS
    With grdTabelaPrd
        .ColWidth(0) = 2300
        .ColWidth(1) = 6300
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
    End With
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub cmdFecharPesqPrd_Click()
    sstProduto.Visible = False
    txtCodPrd.SetFocus
End Sub

Private Sub grdTabelaConsultaPrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSelecionarPrd_Click
    End If
End Sub

Private Sub cmdSelecionarPrd_Click()

    With grdTabelaConsultaPrd
        If .Rows > 1 Then
            txtCodPrd.Text = .TextMatrix(.Row, 0)
            txtDescricao.Text = .TextMatrix(.Row, 1)
            txtQtd.Text = .TextMatrix(.Row, 2)
            txtValUnit.Text = VBA.Format(fcRetornaDadosDoProduto(txtCodPrd.Text, "VL_UNIT"), "Currency")
        End If
    End With
    sstProduto.Visible = False
    Call txtCodPrd_LostFocus
    txtDescricao.SetFocus

End Sub

Private Sub sbMontaGridConsultaProdutos()

    On Error GoTo erro
    
    'DESCRICAO DAS COLUNAS
    With grdTabelaConsultaPrd
        .Cols = 4
        .TextMatrix(0, 0) = "CÓDIGO"
        .TextMatrix(0, 1) = "DESCRIÇÃO"
        .TextMatrix(0, 2) = "ESTOQUE"
        .TextMatrix(0, 3) = "VALOR VENDA"
    End With
    'LARGURA DAS COLUNAS
    With grdTabelaConsultaPrd
        .ColWidth(0) = 1500
        .ColWidth(1) = 4700
        .ColWidth(2) = 1200
        .ColWidth(3) = 1400
    End With
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Private Sub txtNumCompra_LostFocus()
    txtNumCompra.Text = VBA.Format(txtNumCompra.Text, VBA.String(7, "0"))
End Sub

Private Sub txtQtd_GotFocus()
    VBA.SendKeys ("{HOME}+{END}")
End Sub

Private Sub cmdFechar_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdImprimir_Click()
    
    On Error GoTo erro
   
   
    Dim Imp_Encontrada      As String
    Dim Impressora          As Printer
    Dim i                   As Integer
    Dim J                   As Integer
    
    Dim lhPrinter           As Long
    Dim lReturn             As Long
    Dim lpcWritten          As Long
    Dim lDoc                As Long
    Dim sWrittenData        As String
    Dim MyDocInfo           As DOCINFO
    Dim intTotalEtq         As Integer
   
    'Verifica se existe alguma impressora zebra instalada
     For Each Impressora In Printers
         Set Printer = Impressora
         Imp_Encontrada = Impressora.DeviceName
         If VBA.Right(Impressora.DeviceName, 8) = "TLP 2844" Then   'Nome da Impressora instalada no Windows
            Imp_Encontrada = Impressora.DeviceName
            Exit For
         End If
     Next
    
    'Se não Tiver
    If Imp_Encontrada = "" Then
       MsgBox "Não Existe a Etiquetadora Zebra Instalada no Windows, Verifique!", vbCritical, App.Title
       Exit Sub
    End If
   
    'Verifica se o arquivo temporario de etiquetas existe
    If VBA.Dir("c:\prnspl.prn") <> "" Then
        VBA.Kill "c:\prnspl.prn"
    End If
    
    'Cria arquivo temporario com as etiquetas
    Open "c:\prnspl.prn" For Append As #1
        
        With grdTabelaPrd
            For i = 1 To .Rows - 1
                
                'Ex.: 10 copias sendo 3 etiquetas por linha
                intTotalEtq = VBA.CLng(.TextMatrix(i, 2))
                For J = 1 To VBA.CLng(.TextMatrix(i, 2))
                    
                    If intTotalEtq >= 3 Then
                        
                        Print #1, "N"   'Limpa a memoria da impressora a cada nova impressao
                        Print #1, "D10" 'Determina o fator de escuridao da etiqueta
                        
                        'Primeira Etiqueta
                        Print #1, "A0,0,0,2,1,1,N," & VBA.Chr(34) & VBA.Left(cboLoja.Text, 14) & VBA.Chr(34)                    'Nome da loja
                        Print #1, "B0,20,0,1,1,1,40,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                         'Codigo de barras
                        Print #1, "A0,65,0,2,1,1,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                            'Codigo do produto
                        Print #1, "A0,90,0,1,1,1,N," & VBA.Chr(34) & VBA.Left(.TextMatrix(i, 1), 20) & VBA.Chr(34)              'Nome do produto
                        Print #1, "A0,110,0,1,1,1,N," & VBA.Chr(34) & VBA.Format(.TextMatrix(i, 3), "Currency") & VBA.Chr(34)   'Preço
                    
                        'Segunda Etiqueta
                        Print #1, "A210,000,0,2,1,1,N," & VBA.Chr(34) & VBA.Left(cboLoja.Text, 14) & VBA.Chr(34)                'Nome da loja
                        Print #1, "B210,20,0,1,1,1,40,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                       'Codigo de barras
                        Print #1, "A210,65,0,2,1,1,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                          'Codigo do produto
                        Print #1, "A210,90,0,1,1,1,N," & VBA.Chr(34) & VBA.Left(.TextMatrix(i, 1), 20) & VBA.Chr(34)            'Nome do produto
                        Print #1, "A210,110,0,1,1,1,N," & VBA.Chr(34) & VBA.Format(.TextMatrix(i, 3), "Currency") & VBA.Chr(34) 'Preço
                        
                        'Terceira Etiqueta
                        Print #1, "A450,000,0,2,1,1,N," & VBA.Chr(34) & VBA.Left(cboLoja.Text, 14) & VBA.Chr(34)                'Nome da loja
                        Print #1, "B450,20,0,1,1,1,40,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                       'Codigo de barras
                        Print #1, "A450,65,0,2,1,1,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                          'Codigo do produto
                        Print #1, "A450,90,0,1,1,1,N," & VBA.Chr(34) & VBA.Left(.TextMatrix(i, 1), 20) & VBA.Chr(34)            'Nome do produto
                        Print #1, "A450,110,0,1,1,1,N," & VBA.Chr(34) & VBA.Format(.TextMatrix(i, 3), "Currency") & VBA.Chr(34) 'Preço
                        Print #1, "P1"
                        
                        intTotalEtq = intTotalEtq - 3
                    ElseIf intTotalEtq >= 2 Then
                    
                        Print #1, "N"   'Limpa a memoria da impressora a cada nova impressao
                        Print #1, "D10" 'Determina o fator de escuridao da etiqueta
                        
                        'Primeira Etiqueta
                        Print #1, "A0,0,0,2,1,1,N," & VBA.Chr(34) & VBA.Left(cboLoja.Text, 14) & VBA.Chr(34)                    'Nome da loja
                        Print #1, "B0,20,0,1,1,1,40,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                         'Codigo de barras
                        Print #1, "A0,65,0,2,1,1,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                            'Codigo do produto
                        Print #1, "A0,90,0,1,1,1,N," & VBA.Chr(34) & VBA.Left(.TextMatrix(i, 1), 20) & VBA.Chr(34)              'Nome do produto
                        Print #1, "A0,110,0,1,1,1,N," & VBA.Chr(34) & VBA.Format(.TextMatrix(i, 3), "Currency") & VBA.Chr(34)   'Preço
                    
                        'Segunda Etiqueta
                        Print #1, "A210,000,0,2,1,1,N," & VBA.Chr(34) & VBA.Left(cboLoja.Text, 14) & VBA.Chr(34)                'Nome da loja
                        Print #1, "B210,20,0,1,1,1,40,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                       'Codigo de barras
                        Print #1, "A210,65,0,2,1,1,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                          'Codigo do produto
                        Print #1, "A210,90,0,1,1,1,N," & VBA.Chr(34) & VBA.Left(.TextMatrix(i, 1), 20) & VBA.Chr(34)            'Nome do produto
                        Print #1, "A210,110,0,1,1,1,N," & VBA.Chr(34) & VBA.Format(.TextMatrix(i, 3), "Currency") & VBA.Chr(34) 'Preço
                        Print #1, "P1"
                        
                        intTotalEtq = intTotalEtq - 2
                    ElseIf intTotalEtq >= 1 Then
                    
                        Print #1, "N"   'Limpa a memoria da impressora a cada nova impressao
                        Print #1, "D10" 'Determina o fator de escuridao da etiqueta
                    
                        'Primeira Etiqueta
                        Print #1, "A0,0,0,2,1,1,N," & VBA.Chr(34) & VBA.Left(cboLoja.Text, 14) & VBA.Chr(34)                    'Nome da loja
                        Print #1, "B0,20,0,1,1,1,40,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                         'Codigo de barras
                        Print #1, "A0,65,0,2,1,1,N," & VBA.Chr(34) & .TextMatrix(i, 0) & VBA.Chr(34)                            'Codigo do produto
                        Print #1, "A0,90,0,1,1,1,N," & VBA.Chr(34) & VBA.Left(.TextMatrix(i, 1), 20) & VBA.Chr(34)              'Nome do produto
                        Print #1, "A0,110,0,1,1,1,N," & VBA.Chr(34) & VBA.Format(.TextMatrix(i, 3), "Currency") & VBA.Chr(34)   'Preço
                        Print #1, "P1"
                        
                        intTotalEtq = intTotalEtq - 2
                    End If
                    
                Next
            Next
        End With
    Close #1
    
    'Abre conexcao com a impressora, e verifica se o retorno do status
    lReturn = OpenPrinter(Imp_Encontrada, lhPrinter, 0)
    If lReturn = 0 Then
        VBA.MsgBox "A impressora não suporta a instrução 'OpenPrinter', favor contactar a ADM.", vbExclamation, gblStrNomeModulo
        Exit Sub
    End If
    
    MyDocInfo.pDocName = "prnspl.prn"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    
    'Inicia o processo de impressão na impressora Zebra TLP 2844
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    
    sWrittenData = AbreArquivo
    
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
        
    
    VBA.MsgBox "Impressão concluida com sucesso.", vbInformation, gblStrNomeModulo
    
'################### EXEMPLO INICIAL DO CODIGO ###################
''Primeira Etiqueta
'Print #1, "A0,000,0,2,1,1,N," & VBA.Chr(34) & "NOME LOJA" & VBA.Chr(34)
'Print #1, "B0,20,0,1,1,1,40,B," & VBA.Chr(34) & "A123456" & VBA.Chr(34)
'Print #1, "A0,90,0,1,1,1,N," & VBA.Chr(34) & "NOME PRODUTO" & VBA.Chr(34)
'Print #1, "A0,110,0,1,1,1,N," & VBA.Chr(34) & "PRECO" & VBA.Chr(34)
''Print #1, "P1"
'
''Segunda Etiqueta
'Print #1, "A210,000,0,2,1,1,N," & VBA.Chr(34) & "NOME LOJA" & VBA.Chr(34)
'Print #1, "B210,20,0,1,1,1,40,B," & VBA.Chr(34) & "A123456" & VBA.Chr(34)
'Print #1, "A210,90,0,1,1,1,N," & VBA.Chr(34) & "NOME PRODUTO" & VBA.Chr(34)
'Print #1, "A210,110,0,1,1,1,N," & VBA.Chr(34) & "PRECO" & VBA.Chr(34)
''Print #1, "P1"
'
''Terceira Etiqueta
'Print #1, "A450,000,0,2,1,1,N," & VBA.Chr(34) & "NOME LOJA" & VBA.Chr(34)
'Print #1, "B450,20,0,1,1,1,40,B," & VBA.Chr(34) & "A123456" & VBA.Chr(34)
'Print #1, "A450,90,0,1,1,1,N," & VBA.Chr(34) & "NOME PRODUTO" & VBA.Chr(34)
'Print #1, "A450,110,0,1,1,1,N," & VBA.Chr(34) & "PRECO" & VBA.Chr(34)
'Print #1, "P1"
    
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number, Me.Name & ".cmdImprimir_Click")
End Sub

Private Function AbreArquivo() As String
    
    Dim strLinha    As String
    Dim strString   As String

    Open "c:\prnspl.prn" For Input As #1
        strLinha = Input(LOF(1), 1)
    Close #1
    
    AbreArquivo = strLinha
    
End Function

Private Sub CarregaComboEspecial_Loja_e_Filiais()

    On Error GoTo erro
    
    Dim rsTabela                As Recordset
    Dim clsEstabelecimento      As New clsEstabelecimento
    Dim clsFiliais              As New clsFiliais
    
    cboLoja.Clear
    
    Set rsTabela = clsEstabelecimento.EstabelecimentoProdutos_Consultar("001")
    Do While Not rsTabela.EOF
        cboLoja.AddItem rsTabela!estcNomeFantasia
        rsTabela.MoveNext
    Loop
    rsTabela.Close
    
    
    Set rsTabela = clsFiliais.Filiais_Consultar
    Do While Not rsTabela.EOF
        cboLoja.AddItem rsTabela!filcNome
        rsTabela.MoveNext
    Loop
    rsTabela.Close
    
    
    Set rsTabela = Nothing
    Set clsEstabelecimento = Nothing
    Set clsFiliais = Nothing
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number, Me.Name & ".CarregaComboEspecial_Loja_e_Filiais")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dbBanco.Close
    Set dbBanco = Nothing
    End
End Sub
