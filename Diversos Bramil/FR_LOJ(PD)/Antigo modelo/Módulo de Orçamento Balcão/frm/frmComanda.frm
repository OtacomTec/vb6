VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmComanda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comanda"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComanda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   8460
      ScaleHeight     =   2625
      ScaleWidth      =   405
      TabIndex        =   8
      Top             =   4860
      Width           =   405
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   2520
      ScaleHeight     =   645
      ScaleWidth      =   75
      TabIndex        =   14
      Top             =   3210
      Width           =   75
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   2550
      ScaleHeight     =   105
      ScaleWidth      =   6135
      TabIndex        =   13
      Top             =   3720
      Width           =   6135
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   2550
      ScaleHeight     =   105
      ScaleWidth      =   6135
      TabIndex        =   12
      Top             =   3150
      Width           =   6135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   8580
      ScaleHeight     =   645
      ScaleWidth      =   285
      TabIndex        =   11
      Top             =   3210
      Width           =   285
   End
   Begin VB.CommandButton cmdAtualizar 
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
      Left            =   7290
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancelar"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1740
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
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
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1740
      Width           =   1635
   End
   Begin VB.TextBox txtCodigo_Vendedor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   660
      Left            =   390
      TabIndex        =   1
      Top             =   3180
      Width           =   1815
   End
   Begin VB.TextBox txtNumero_Comanda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   660
      Left            =   390
      TabIndex        =   0
      Top             =   1590
      Width           =   3465
   End
   Begin MSDataListLib.DataCombo dtcVendedor 
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3210
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1085
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      BackColor       =   8454143
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgComanda 
      DragMode        =   1  'Automatic
      Height          =   2565
      Left            =   390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4920
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   4524
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Itens da Comanda"
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
      Left            =   390
      TabIndex        =   15
      Top             =   4290
      Width           =   2610
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   7140
      Shape           =   4  'Rounded Rectangle
      Top             =   7980
      Width           =   1665
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   3060
      Width           =   1995
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Comandas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   2220
      TabIndex        =   7
      Top             =   90
      Width           =   2730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   4860
      X2              =   0
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   7170
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   5190
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   3060
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vendedor"
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
      Left            =   390
      TabIndex        =   4
      Top             =   2670
      Width           =   1365
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   1470
      Width           =   3675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "N° Comanda.:"
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
      Left            =   390
      TabIndex        =   3
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1530
      Width           =   3645
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   2370
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   6465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   8655
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   4860
      Width           =   8475
   End
End
Attribute VB_Name = "frmComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String

Private Sub cmdAtualizar_Click()

     Dim strCodigo_Produto As String * 14
     Dim strDescricao_Produto As String
     Dim strAliquota As String
     Dim strTipo_quantidade As String * 1
     Dim strQuantiade As Double
     Dim strCasas_Decimais As String * 1
     Dim strValor_Unitario As String
     Dim strTipo_desconto As String * 1
     Dim strValor_desconto As String * 8
     
     'Adicionando os itens da comanda no cupom
     intIndice_itens = hfgComanda.Rows
     intCont_Itens = 1
    
     Do While intCont_Itens < hfgComanda.Rows
     
        hfgComanda.Row = intCont_Itens
        hfgComanda.Col = 1
        
        strCodigo_Produto = Me.hfgComanda.Text
        
        hfgComanda.Col = 2
        strDescricao_Produto = Me.hfgComanda.Text
        
        hfgComanda.Col = 3
        strQuantiade = Me.hfgComanda.Text
        
        hfgComanda.Col = 4
        strValor_Unitario = Me.hfgComanda.Text
        
        'Montando dysplay de itens de cupom
        frmTela_Venda.HflexGrid.Cols = 6
        frmTela_Venda.HflexGrid.ColWidth(0) = 0
        frmTela_Venda.HflexGrid.Rows = frmTela_Venda.HflexGrid.Rows + 2
        
        If frmTela_Venda.HflexGrid.Rows = 4 Then
           'Cabeçalho 1
           frmTela_Venda.HflexGrid.Row = 0
           frmTela_Venda.HflexGrid.Col = 1
           frmTela_Venda.HflexGrid.FixedAlignment(1) = 2
           frmTela_Venda.HflexGrid.Font.Name = "Tahoma"
           frmTela_Venda.HflexGrid.Text = "Código"
           frmTela_Venda.HflexGrid.Col = 2
           frmTela_Venda.HflexGrid.Text = "Descrição"
           frmTela_Venda.HflexGrid.Col = 3
           frmTela_Venda.HflexGrid.Text = "Qtd."
           frmTela_Venda.HflexGrid.Col = 4
           frmTela_Venda.HflexGrid.Text = "X"
           frmTela_Venda.HflexGrid.Col = 5
           frmTela_Venda.HflexGrid.Text = "Vlr.Unit."
           'Cabeçalho 2
           frmTela_Venda.HflexGrid.Row = 1
           frmTela_Venda.HflexGrid.Col = 1
           frmTela_Venda.HflexGrid.FixedAlignment(1) = 2
           frmTela_Venda.HflexGrid.Font.Name = "Tahoma"
           frmTela_Venda.HflexGrid.Text = ""
           frmTela_Venda.HflexGrid.Col = 2
           frmTela_Venda.HflexGrid.Text = ""
           frmTela_Venda.HflexGrid.Col = 3
           frmTela_Venda.HflexGrid.Text = ""
           frmTela_Venda.HflexGrid.Col = 4
           frmTela_Venda.HflexGrid.Text = ""
           frmTela_Venda.HflexGrid.Col = 5
           frmTela_Venda.HflexGrid.CellAlignment = 7
           frmTela_Venda.HflexGrid.CellFontBold = True
           frmTela_Venda.HflexGrid.Text = "T.Item"
           'Separador
           frmTela_Venda.HflexGrid.Row = 2
           frmTela_Venda.HflexGrid.RowHeight(2) = 100
           frmTela_Venda.HflexGrid.Col = 1
           frmTela_Venda.HflexGrid.FixedAlignment(1) = 2
           frmTela_Venda.HflexGrid.Font.Name = "Tahoma"
           frmTela_Venda.HflexGrid.Text = "------------------------------------"
           frmTela_Venda.HflexGrid.Col = 2
           frmTela_Venda.HflexGrid.Text = "--------------------------------------------------"
           frmTela_Venda.HflexGrid.Col = 3
           frmTela_Venda.HflexGrid.Text = "----------------"
           frmTela_Venda.HflexGrid.Col = 4
           frmTela_Venda.HflexGrid.Text = "--------"
           frmTela_Venda.HflexGrid.Col = 5
           frmTela_Venda.HflexGrid.Text = "--------------"
        Else
           frmTela_Venda.HflexGrid.Rows = frmTela_Venda.HflexGrid.Rows - 1
        End If
        
        'Detalhe 1
        frmTela_Venda.HflexGrid.Row = frmTela_Venda.HflexGrid.Rows - 1
        frmTela_Venda.HflexGrid.Col = 1
        frmTela_Venda.HflexGrid.Font.Name = "Tahoma"
        frmTela_Venda.HflexGrid.Text = Format(strCodigo_Produto, "0000000000000")
        frmTela_Venda.HflexGrid.Col = 2
        frmTela_Venda.HflexGrid.Text = strDescricao_Produto
        frmTela_Venda.HflexGrid.Col = 3
        frmTela_Venda.HflexGrid.Text = strQuantiade
        frmTela_Venda.HflexGrid.Col = 4
        frmTela_Venda.HflexGrid.Text = "X"
        frmTela_Venda.HflexGrid.Col = 5
        frmTela_Venda.HflexGrid.Text = strValor_Unitario
        frmTela_Venda.HflexGrid.Rows = frmTela_Venda.HflexGrid.Rows + 1
        
        'Detalhe 2
        frmTela_Venda.HflexGrid.Row = frmTela_Venda.HflexGrid.Rows - 1
        frmTela_Venda.HflexGrid.Col = 5
        frmTela_Venda.HflexGrid.CellFontBold = True
        frmTela_Venda.HflexGrid.CellFontSize = 6
        frmTela_Venda.HflexGrid.Text = Format(CDbl(strQuantiade) * CDbl(strValor_Unitario), "#,###0.00")
        
        'Formatando Colunas
        frmTela_Venda.HflexGrid.ColWidth(1) = 1100
        frmTela_Venda.HflexGrid.ColWidth(2) = 2000
        frmTela_Venda.HflexGrid.ColWidth(3) = 350
        frmTela_Venda.HflexGrid.ColWidth(4) = 150
        frmTela_Venda.HflexGrid.ColWidth(5) = 650
        
        intCont_Itens = intCont_Itens + 1
        
        If frmTela_Venda.txtPreco_total_cupom.Text = "" Or frmTela_Venda.txtPreco_total_cupom.Text = Empty Then
           frmTela_Venda.txtPreco_total_cupom.Text = 0
        End If
        
        'Totalizando os itens
        frmTela_Venda.txtPreco_total_cupom.Text = Format(CDbl(frmTela_Venda.txtPreco_total_cupom.Text) + (CDbl(strQuantiade) * CDbl(strValor_Unitario)), "#,###0.00")
        
        frmTela_Venda.HflexGrid.TopRow = frmTela_Venda.HflexGrid.Rows - 2
     Loop
     
     Unload Me
     
     frmTela_Venda.HflexGrid.SetFocus

     If frmTela_Venda.txtPreco_total_cupom.Text = "" Then frmTela_Venda.txtPreco_total_cupom.Text = 0
     If frmTela_Venda.txtPreco_Total.Text = "" Then frmTela_Venda.txtPreco_Total.Text = 0
     
     Call frmTela_Venda.Limpa_Tela
     
     'Informando ao cupom inf. ref. à comanda
     frmTela_Venda.booComanda = True
     
     Set rstInf_Produtos = Nothing
     Set rstParametro_ecf = Nothing
     
     frmTela_Venda.txtQuantidade_Produto.TabStop = False
     frmTela_Venda.txtPreco_Unitario.TabStop = False
     frmTela_Venda.txtCodigo_Produto.SetFocus
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim strNomes As String
    Dim strTamanho As String
    Dim rstComanda_Vendedor As New ADODB.Recordset
    
    'Verificar se existe a comanda para este vendedor
    strSQL = Empty
    strSQL = "SELECT TBComanda.DFNumero_cupom_TBComanda,TBComanda.PKCodigo_TBComanda," & _
             "TBVendedor.IXCodigo_TBVendedor," & _
             "TBVendedor.DFNome_TBVendedor," & _
             "TBComanda.DFData_lancamento_TBComanda," & _
             "TBComanda.DFHora_abertura_TBComanda," & _
             "TBComanda.DFNumero_cupom_TBComanda " & _
             "FROM TBComanda " & _
             "INNER JOIN TBVendedor ON TBComanda.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
             "WHERE IXCodigo_TBVendedor = " & Me.txtCodigo_Vendedor.Text & " " & _
             "AND TBComanda.PKCodigo_TBComanda = " & Me.txtNumero_Comanda.Text & ""
             
         
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstComanda_Vendedor, "Otica", Me
    
    If rstComanda_Vendedor.EOF = True And rstComanda_Vendedor.BOF = True Then
       MsgBox "Não consta esta comanda para este.Verifique!", vbCritical, "Only Tech"
       Me.txtCodigo_Vendedor.Text = Empty
       Me.txtNumero_Comanda.Text = Empty
       Me.hfgComanda.Clear
       Me.dtcVendedor.Text = Empty
       Me.txtNumero_Comanda.SetFocus
       Exit Sub
    End If
    
    If rstComanda_Vendedor!DFNumero_cupom_TBComanda <> "" Or Not IsNull(rstComanda_Vendedor!DFNumero_cupom_TBComanda) Then
       MsgBox "Esta comanda já foi fechada.Verifique!", vbCritical, "Only Tech"
       Me.txtCodigo_Vendedor.Text = Empty
       Me.txtNumero_Comanda.Text = Empty
       Me.hfgComanda.Clear
       Me.dtcVendedor.Text = Empty
       Me.txtNumero_Comanda.SetFocus
       Exit Sub
    End If
    
    
    Me.hfgComanda.Clear
    
    strSQL = Empty
    strSQL = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,DFQuantidade_TBItens_comanda,DFPreco_TBItens_comanda,DFValor_total_TBItens_comanda " & _
             "FROM TBItens_comanda " & _
             "INNER JOIN TBProduto " & _
             "ON TBItens_comanda.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
             "WHERE FKCodigo_TBComanda = " & Me.txtNumero_Comanda.Text & " "
             
    strNomes = " ,Produto,    Quant.,  Preço Unit.,  Total Item"
    
    strTamanho = "500,3000,1000,1200,1200"
    
    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgComanda, strTamanho, strNomes, "BDRetaguarda", "Otica", Me, "S"
    
    'Informações para o cupom da comanda
    frmTela_Venda.strNumero_Comanda = Me.txtNumero_Comanda
    frmTela_Venda.strVendedor_Comanda = Me.txtCodigo_Vendedor.Text
    
    Me.cmdAtualizar.SetFocus
                              
End Sub

Private Sub dtcVendedor_GotFocus()
    If txtCodigo_Vendedor.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcVendedor)
    End If
End Sub

Private Sub dtcVendedor_LostFocus()
    If dtcVendedor.Text <> Empty Then
       txtCodigo_Vendedor.Text = dtcVendedor.BoundText
    End If
End Sub

Private Sub Form_Load()
    
    strSQL = "SELECT TBVendedor.IXCodigo_TBVendedor,TBVendedor.DFNome_TBVendedor FROM TBVendedor"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSQL, "BDRetaguarda", "Otica", Me
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCodigo_Vendedor_LostFocus()
    If txtCodigo_Vendedor.Text <> Empty Then
       dtcVendedor.BoundText = txtCodigo_Vendedor.Text
    End If
End Sub
