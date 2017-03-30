VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFechamento_Cupom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Fechamento do  cupom"
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
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   165
      Left            =   11670
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   225
      lastProp        =   500
      _cx             =   397
      _cy             =   291
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5820
      ScaleHeight     =   615
      ScaleWidth      =   285
      TabIndex        =   13
      Top             =   3510
      Width           =   285
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   2610
      ScaleHeight     =   645
      ScaleWidth      =   45
      TabIndex        =   12
      Top             =   3510
      Width           =   45
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   2610
      ScaleHeight     =   75
      ScaleWidth      =   3465
      TabIndex        =   11
      Top             =   4050
      Width           =   3465
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   2610
      ScaleHeight     =   75
      ScaleWidth      =   3465
      TabIndex        =   10
      Top             =   3450
      Width           =   3465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   10680
      ScaleHeight     =   3885
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   2370
      Width           =   255
   End
   Begin VB.TextBox txtTroco 
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
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5670
      Width           =   3375
   End
   Begin VB.TextBox txtTotal_Cupom 
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
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2430
      Width           =   3405
   End
   Begin VB.TextBox txtValor_pago 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   4590
      Width           =   3375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HflexGrid 
      DragMode        =   1  'Automatic
      Height          =   3825
      Left            =   6660
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2430
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6747
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
         Size            =   9
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
   Begin MSDataListLib.DataCombo dtcFinalizadora_cupom 
      Height          =   570
      Left            =   2610
      TabIndex        =   0
      ToolTipText     =   "Finalizadora"
      Top             =   3510
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.Shape Shape17 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   3645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   4860
      X2              =   0
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Finalizando Cupom"
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
      Left            =   1890
      TabIndex        =   14
      Top             =   540
      Width           =   4995
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Finalizadora:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   540
      TabIndex        =   9
      Top             =   3870
      Width           =   1845
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   3675
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   4245
      Left            =   6540
      Shape           =   4  'Rounded Rectangle
      Top             =   2190
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Troco..........:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   540
      TabIndex        =   6
      Top             =   6000
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   3645
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valor Pago..:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   540
      TabIndex        =   5
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total...........:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   540
      TabIndex        =   4
      Top             =   2790
      Width           =   1845
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   2250
      Width           =   3675
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   2310
      Width           =   3645
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   3645
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   3645
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   4185
      Left            =   6450
      Shape           =   4  'Rounded Rectangle
      Top             =   2370
      Width           =   4425
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   3420
      Width           =   3645
   End
End
Attribute VB_Name = "frmFechamento_Cupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Dim dblValor_pago As Double
Dim CNconexao As New DLLConexao_Sistema.conexao
Dim dblTotal_item As Double
Dim lngID_Numero_Nota As Long
Dim lngVendedor As Long
Dim lngPlano_pagamento As Long
Dim strNumero_Nota As String
Dim strSerie_nota As String
'Recordsets
Dim rstFinalizadora As New ADODB.Recordset
Dim rstProdutos As New ADODB.Recordset
Dim rstTabela As New ADODB.Recordset
Dim rstNumero_orcamento As New ADODB.Recordset
Dim Relatorio As New CRAXDRT.Report
Dim Aplicacao As New CRAXDRT.Application
Dim ocorrencia As New DLLSistema.Ocorrencia_de_produto
Dim estoque As New DLLSistema.estoque
Dim Tabela As String

Private Sub dtcFinalizadora_cupom_GotFocus()
     Call Movimentacoes.Verifica_DataCombo(Me.dtcFinalizadora_cupom)
     Me.dtcFinalizadora_cupom.SetFocus
End Sub

Private Sub dtcFinalizadora_cupom_LostFocus()
     If Me.dtcFinalizadora_cupom.Text = "" Then Me.dtcFinalizadora_cupom.SetFocus
End Sub

Private Sub Form_Load()

    Me.txtTotal_Cupom.Text = frmTela_Venda.txtPreco_total_cupom.Text
    
    dblValor_pago = 0
    
    'Carregando a combo de finalizadora
    strSQL = Empty
    strSQL = "SELECT IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora FROM TBFinalizadora WHERE DFControle_venda_TBFinalizadora = 1"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_cupom, strSQL, "BDRetaguarda", "Otica", Me
        
 End Sub

Private Sub Form_Unload(Cancel As Integer)

    
    CNconexao.Fechar_conexao
    
End Sub

Private Sub txtValor_pago_LostFocus()

    If txtValor_pago.Text = "" Or IsNull(txtValor_pago) Then txtValor_pago.Text = 0
    
    txtValor_pago.Text = Format(txtValor_pago, "#,###0.00")
    dblValor_pago = CDbl(txtValor_pago.Text) + dblValor_pago
    
    If dtcFinalizadora_cupom.Text <> "" Then
        If CDbl(Me.txtValor_pago.Text) > 0 Then
            If CDbl(dblValor_pago) - (Me.txtTotal_Cupom.Text) > 0 Then
                Me.txtTroco.Text = Format(CDbl(dblValor_pago) - (Me.txtTotal_Cupom.Text), "#,###0.00")
            Else
                If CDbl(Me.txtValor_pago.Text) - (Me.txtTotal_Cupom.Text) < 0 Then
                   txtTroco.Text = Format(0, "#,###0.00")
                End If
            End If
            Call Fecha_Cupom
        Else
           Me.dtcFinalizadora_cupom.Text = ""
           Me.dtcFinalizadora_cupom.SetFocus
        End If
    Else
        Me.dtcFinalizadora_cupom.SetFocus
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Function Fecha_Cupom()

     'Montando as finalizadoras deste cupom
     HflexGrid.Cols = 3
     HflexGrid.ColWidth(0) = 0
     HflexGrid.Rows = HflexGrid.Rows + 2
     
     If HflexGrid.Rows = 4 Then
        'Cabeçalho 1
        HflexGrid.Row = 0
        HflexGrid.Col = 1
        HflexGrid.FixedAlignment(1) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = "Finalizadora"
        HflexGrid.Col = 2
        HflexGrid.Text = "Vlr.Finalizadora"
        'Separador
        HflexGrid.Row = 1
        HflexGrid.RowHeight(2) = 100
        HflexGrid.Col = 1
        HflexGrid.FixedAlignment(1) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = "--------------------------------------------------------------"
        HflexGrid.Col = 2
        HflexGrid.Text = "--------------------------------------------------------------"
     Else
        HflexGrid.Rows = HflexGrid.Rows - 1
     End If
     
     'Detalhe 1
     HflexGrid.Row = HflexGrid.Rows - 1
     HflexGrid.Col = 1
     HflexGrid.Font.Name = "Tahoma"
     HflexGrid.Text = Me.dtcFinalizadora_cupom.Text
     HflexGrid.Col = 2
     HflexGrid.Text = Me.txtValor_pago.Text
     
     'Formatando Colunas
     HflexGrid.ColWidth(1) = 2000
     HflexGrid.ColWidth(2) = 2000
     
     Me.HflexGrid.SetFocus
     Me.HflexGrid.TopRow = Me.HflexGrid.Rows - 2
     
     If dblValor_pago < Me.txtTotal_Cupom.Text Then
        Me.dtcFinalizadora_cupom.Text = Empty
        txtValor_pago.Text = ""
        Me.dtcFinalizadora_cupom.SetFocus
     Else
        MsgBox "Aguarde a impressão do orçamento!", vbInformation, "Only Tech"
        Call Grava_cupom
        frmTela_Venda.Limpa_Tela
        frmTela_Venda.HflexGrid.Clear
        frmTela_Venda.HflexGrid.Rows = 2
        frmTela_Venda.txtPreco_total_cupom.Text = Empty
        Unload Me
     End If

End Function

Public Function Grava_cupom()
    
    'Abrindo uma conexão nova
    CNconexao.Banco = "BDRetaguarda"
    CNconexao.Abrir_conexao "Otica"
    
    'SELECTS -----------------------------------------------------------------------------------------------
    'Finalizadoras
    strSQL = Empty
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstFinalizadora, "Otica", Me
    
    'Próxima Tabela
    strSQL = Empty
    strSQL = "SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBParametros_venda WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstTabela, "Otica", Me
    
    If rstTabela.BOF = True And rstTabela.EOF = True Then
       MsgBox "Não existe tabela vigente cadastrada no sistema.Venda impossibilitada de ser concluida.Verifique!", vbCritical, , "Only Tech"
       Set rstTabela = Nothing
       Set rstFinalizadora = Nothing
       Exit Function
    End If
    
    'Próximo Orçamento
    strSQL = Empty
    strSQL = "SELECT DFProximo_orcamento_balcao_TBParametros_ecf,DFProximo_serie_orcamento_balcao_TBParametros_ecf FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstNumero_orcamento, "Otica", Me
    
    If rstNumero_orcamento.BOF = True And rstNumero_orcamento.EOF = True Then
       MsgBox "Número do próximo orçamento balcão não cadastrado.Venda impossibilitada de ser concluida.Verifique!", vbCritical, , "Only Tech"
       Set rstTabela = Nothing
       Set rstFinalizadora = Nothing
       Set rstNumero_orcamento = Nothing
       Exit Function
    End If
    
    'Produto
    strSQL = Empty
    strSQL = "SELECT * FROM TBProduto WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " ORDER BY IXCodigo_TBProduto ASC"
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstProdutos, "Otica", Me
    
    If frmTela_Venda.booComanda = False Then
        lngVendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", "9999", "TBVendedor", "Otica", Me, "BDRetaguarda")
    Else
        lngVendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", frmTela_Venda.strVendedor_Comanda, "TBVendedor", "Otica", Me, "BDRetaguarda")
    End If
    
    lngPlano_pagamento = Funcoes_Gerais.Localiza_ID("PKId_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", "9999", "TBPlano_pagamento", "Otica", Me, "BDRetaguarda")
    
    'Abrindo Transações
    CNconexao.CNconexao.BeginTrans
    
    Call Grava_Corpo_Nota
    
    'Comitando Transações
    CNconexao.CNconexao.CommitTrans
    
    strNumero_Nota = rstNumero_orcamento!DFProximo_orcamento_balcao_TBParametros_ecf
    strSerie_nota = rstNumero_orcamento!DFProximo_serie_orcamento_balcao_TBParametros_ecf
       
    'Consultando o Id da nota gravada.
    lngID_Numero_Nota = Funcoes_Gerais.Localiza_ID("PKId_TBNota_saida", "DFNumero_TBNota_saida", strNumero_Nota, "TBNota_saida", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", frmTela_Venda.strEmpresa_Operador, "DFSerie_TBNota_saida", strSerie_nota)
    
    'Reabrindo Transações
    CNconexao.CNconexao.BeginTrans
    
    Call Grava_Finalizadoras
    Call Grava_Itens_Nota
    
    If frmTela_Venda.booComanda = True Then
        'Verifica se é cupom abastecido de comanda e finaliza a mesma
        Call Fecha_comanda
    End If
    
    'Atualizando o parâmetro ECF
    strSQL = Empty
    strSQL = "UPDATE TBParametros_ecf SET DFProximo_orcamento_balcao_TBParametros_ecf = " & rstNumero_orcamento!DFProximo_orcamento_balcao_TBParametros_ecf + 1 & " " & _
             "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    CNconexao.CNconexao.Execute strSQL
     
    'Comitando Transações
    CNconexao.CNconexao.CommitTrans
    
    Call Imprime_Cupom_nao_fiscal
    
    Set rstFinalizadora = Nothing
    Set rstNumero_orcamento = Nothing
    Set rstTabela = Nothing
    Set rstProdutos = Nothing
    
End Function
Private Function Grava_Finalizadoras()

    Dim intCont_Finalizadora As Integer
    Dim intIndice As Integer
    Dim strID_Finalizadora As String
    Dim strObservacao As String
    Dim dblValor_pago As Double
    
    'On Error GoTo Erro
    
    'Gravando as operações de acordo com as finalizadoras utilizadas neste cupom
    intCont_Finalizadora = Me.HflexGrid.Rows
    intIndice = 1
    
    Do While intIndice + 2 < Me.HflexGrid.Rows
        
        Me.HflexGrid.Row = intIndice + 2
        Me.HflexGrid.Col = 1
        
        rstFinalizadora.MoveFirst
        rstFinalizadora.Find ("DFDescricao_TBFinalizadora = '" & Me.HflexGrid.Text & "'")
        
        strFinalizadora = rstFinalizadora!PKId_TBFinalizadora
        strObservacao = Me.HflexGrid.Text
        
        Me.HflexGrid.Col = 2
        dblValor_pago = Me.HflexGrid.Text
       
        strCampos = "FKCodigo_TBPdv,FKId_TBFinalizadora,FKCodigo_TBOperadores_ecf,DFData_TBOperacao_caixa," & _
                    "DFHora_TBOperacao_caixa,DFValor_TBOperacao_caixa,DFTipo_operacao_TBOperacao_caixa,DFStatus_aberto_fechado_TBOperacao_caixa," & _
                    "DFObservacao_TBOperacao_caixa"
                  
        strValores = "" & frmTela_Venda.txtNumero_check_out & "," & _
                     "" & strFinalizadora & "," & _
                     "" & frmTela_Venda.strCodigo_Operador & "," & _
                     "'" & Format(Now, "YYYYMMDD") & "'," & _
                     "'" & Format(Now, "hh:mm:ss") & "'," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblValor_pago) & "," & _
                     "1," & _
                     "0, '" & strObservacao & "'"
        
        'Gravando Operações
        strSQL = Empty
        strSQL = "INSERT INTO TBoperacao_caixa ( " & strCampos & ") VALUES ( " & strValores & ")"
        
        CNconexao.CNconexao.Execute strSQL
        
        intIndice = intIndice + 1
        
    Loop
    
    Exit Function
    
Erro:

    'Rollback na transação
    CNconexao.CNconexao.RollbackTrans
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando as finalizadoras.Verifique"
    
    Exit Function
    
End Function
Private Function Grava_Corpo_Nota()

    On Error GoTo Erro
    
    strSQL = Empty
    strSQL = "INSERT INTO TBNota_Saida(" & _
             "FKCodigo_TBEmpresa, " & _
             "FKCodigo_TBTabela_preco, " & _
             "FKId_TBVendedor," & _
             "FKId_TBPlano_pagamento," & _
             "FKCodigo_TBTransportadora," & _
             "DFTipo_operacao_TBNota_Saida," & _
             "DFEmitente_TBNota_Saida," & _
             "DFTotal_itens_TBNota_Saida," & _
             "DFTotal_Nota_TBNota_Saida," & _
             "DFTotal_Nota_tabela_TBNota_Saida," & _
             "DFDesconto_especial_TBNota_Saida," & _
             "DFDesconto_indenizacao_TBNota_Saida," & _
             "DFPrevisao_TBNota_Saida," & _
             "DFValor_ipi_TBNota_Saida," & _
             "DFDespesas_acessorias_TBNota_Saida," & _
             "DFNumero_TBNota_saida," & _
             "DFSerie_TBNota_saida," & _
             "DFData_Emissao_TBNota_saida," & _
             "DFData_Saida_TBNota_saida," & _
             "DFFaturista_TBNota_saida," & _
             "DFDigitador_TBNota_saida, "
     strSQL = strSQL & "DFTotal_custo_medio_TBNota_saida," & _
             "DFTotal_custo_real_TBNota_saida," & _
             "DFTotal_custo_contabil_TBNota_saida," & _
             "DFNumero_pedido_TBNota_saida," & _
             "DFTotal_descontos_itens_TBNota_Saida," & _
             "DFTotal_peso_liquido_TBNota_Saida," & _
             "DFTotal_peso_bruto_TBNota_Saida," & _
             "DFTipo_emitente_TBNota_Saida," & _
             "DFCancelado_TBNota_saida," & _
             "DFIntegrado_fiscal_TBNota_saida," & _
             "DFMotivo_cancelamento_TBNota_saida," & _
             "DFUsuario_cancelamento_TBNota_saida, " & _
             "DFObservacao_TBNota_saida) "
     strSQL = strSQL & "VALUES (" & _
             "" & frmTela_Venda.strEmpresa_Operador & "," & _
             "" & rstTabela.Fields("DFNumero_tabela_vigente_TBParametros_venda") & "," & _
             "" & lngVendedor & "," & _
             "" & lngPlano_pagamento & "," & _
             "9999," & _
             "1," & _
             "9999," & _
             "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
             "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
             "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
             "0," & _
             "0," & _
             "1," & _
             "0,"
     strSQL = strSQL + "0," & _
             "" & rstNumero_orcamento.Fields!DFProximo_orcamento_balcao_TBParametros_ecf & "," & _
             "'" & rstNumero_orcamento.Fields!DFProximo_serie_orcamento_balcao_TBParametros_ecf & "'," & _
             "'" & Format(Now, "YYYYMMDD") & "'," & _
             "'" & Format(Now, "YYYYMMDD") & "'," & _
             "'" & frmTela_Venda.strOperador & "'," & _
             "'" & frmTela_Venda.strOperador & "'," & _
             "0," & _
             "0," & _
             "0," & _
             "1," & _
             "0," & _
             "0," & _
             "0," & _
             "0," & _
             "0," & _
             "''," & _
             "''," & _
             "''," & _
             "'Nota gerada apartir da emissão de um cupom fiscal no módulo de varejo')"
             
    CNconexao.CNconexao.Execute strSQL
    
    Exit Function
    
Erro:

    'Rollback na transação
    CNconexao.CNconexao.RollbackTrans
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando o corpo da nota.Verifique"
    
    Exit Function
    
End Function
Private Function Grava_Itens_Nota()
    
    Dim intCont_Itens As Integer
    Dim intIndice_itens As Integer
    Dim strObservacao As String
    Dim dblValor_pago As Double
    Dim rstUF As New ADODB.Recordset
    
    Dim dblQuantidade As Double
    Dim dblValor_Unitario As Double
    Dim dblValor_Total As Double
    Dim strCodigo_item As Double
    Dim strDescricao As String
    
    'Verifica a uf do emitente
    strSQL = Empty
    strSQL = "SELECT TBCidade_otica.DFUf_TBCidade_otica FROM TBEmpresa " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBEmpresa.Fkid_TBCidade_otica  = TBCidade_otica.pkid_TBCidade_otica " & _
             "WHERE TBEmpresa.PKCodigo_TBempresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstUF, "Otica", Me)
    
    strUF_Emitente = rstUF!DFUf_TBCidade_otica
    
    Set rstUF = Nothing
    
    '--CFO-------------------------------------------------------------------------------------------------
    
    Dim rstCFO As New ADODB.Recordset
    Dim strCodigo_cfo As String
    
    'Verifica O CFO no parametro
    strSQL = Empty
    strSQL = "SELECT DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
             "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstCFO, "Otica", Me)
    
    intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCFOP", "DFCodigo_TBCfop", rstCFO!CFO, "TBCFOP", "Otica", Me, "BDRetaguarda")
    
    Set rstCFO = Nothing
    
    'On Error GoTo Erro
    
    'Gravando as operações de acordo com as finalizadoras utilizadas neste cupom
    intIndice_itens = frmTela_Venda.HflexGrid.Rows
    intCont_Itens = 1
    
    Do While intCont_Itens + 2 < frmTela_Venda.HflexGrid.Rows
        frmTela_Venda.HflexGrid.Row = intCont_Itens + 2
        
        'Movendo as inf. das linhas
        frmTela_Venda.HflexGrid.Col = 1
        strCodigo_item = frmTela_Venda.HflexGrid.Text
        
        'Localizando e convertendo todos os casos para Código Interno
        If Len(CStr(strCodigo_item)) > 6 Then
        
            Dim strDigito_Produto_Digitado As String
            Dim strCodigo_Produto_Etiqueta As String
            
            strDigito_Produto_Digitado = Left(strCodigo_item, 1)
            
            If frmTela_Venda.strDigito_Peso_Variavel = strDigito_Produto_Digitado Then
               'Produto pesável e preço variavel
               strCodigo_Produto_Etiqueta = Mid(strCodigo_item, 2, 4)
               strCodigo_item = strCodigo_Produto_Etiqueta
            Else
               Dim rstCodigo_Interno As New ADODB.Recordset
               
               'Produto não pesável e preço não variavel
               strID_Produto = Funcoes_Gerais.Localiza_ID("FKId_TBProduto", "IXCodigo_TBCodigo_barras", frmTela_Venda.HflexGrid.Text, "TBCodigo_barras", "Otica", Me, "BDRetaguarda")
               strSQL = Empty
               strSQL = "SELECT IXCodigo_TBProduto FROM TBProduto WHERE PKId_TBProduto = " & strID_Produto & ""
               Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstCodigo_Interno, "Otica", Me
               strCodigo_item = rstCodigo_Interno!IXCodigo_TBProduto
               
               Set rstCodigo_Interno = Nothing
            End If
        End If
        
        frmTela_Venda.HflexGrid.Col = 2
        strDescricao = frmTela_Venda.HflexGrid.Text
                
        frmTela_Venda.HflexGrid.Col = 3
        dblQuantidade = frmTela_Venda.HflexGrid.Text
        
        frmTela_Venda.HflexGrid.Col = 5
        dblValor_Unitario = frmTela_Venda.HflexGrid.Text
        
        'Total do Item
        frmTela_Venda.HflexGrid.Row = frmTela_Venda.HflexGrid.Row + 1
        frmTela_Venda.HflexGrid.Col = 5
        dblValor_Total = frmTela_Venda.HflexGrid.Text
        frmTela_Venda.HflexGrid.Row = frmTela_Venda.HflexGrid.Row - 1
        '------------------------------------------------------------
        
        rstProdutos.MoveFirst
        rstProdutos.Find ("IXCodigo_TBProduto = " & strCodigo_item & "")
           
        'Custos --------------------------------------------------------------------------------------------
        Dim dblCusto_Real As Double
        Dim dblCusto_Contabil As Double
        Dim dblCusto_Medio As Double
        
        'calculando o total do item
        dblTotal_item = dblValor_Unitario * dblQuantidade
        dblCusto_Real = CDbl(rstProdutos!DFCusto_real_TBProduto)
        dblCusto_Contabil = CDbl(rstProdutos!DFCusto_contabil_TBProduto)
        dblCusto_Medio = CDbl(rstProdutos!DFCusto_medio_TBProduto)
        'End If
            
        'multiplicando pela quantidade
        dblCusto_Real = dblCusto_Real * dblQuantidade
        dblCusto_Contabil = dblCusto_Contabil * dblQuantidade
        dblCusto_Medio = dblCusto_Medio * dblQuantidade
            
        'Multiplicar pela quantidade vendida
        dblCusto_Real = Format(dblCusto_Real, "#,###0.00")
        dblCusto_Contabil = Format(dblCusto_Contabil, "#,###0.00")
        dblCusto_Medio = Format(dblCusto_Medio, "#,###0.00")
        
        'Calculando o Peso dos Itens
        dblPeso_Liquido_item = CDbl(rstProdutos!DFPeso_liquido_TBProduto) * dblQuantidade
        dblPeso_Bruto_item = CDbl(rstProdutos!DFPeso_bruto_TBProduto) * dblQuantidade
        
        'Impostos ------------------------------------------------------------------------------------------
        Dim strST As String
        Dim strST2 As String
        Dim dblAliquota_icms As Double
        Dim dblTotal_Icms As Double
         
        'Calculando a parte do ICMS relacionado ao Item
        'Concatenando o valor da Situação Tributária que está no cadastro de produto
        strST = rstProdutos!DFCst1_TBProduto
        strST2 = rstProdutos!DFCst2_TBProduto
        
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
                     "WHERE FKId_TBProduto = " & rstProdutos!PKId_TBProduto & " " & _
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
                        "WHERE FKId_TBProduto = " & rstProdutos!PKId_TBProduto & " " & _
                        "AND DFUf_TBEstado_icms = '**' "
                              
               Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
            End If
            
            dblAliquota_icms = rstVerifica_Estado_ICMS!DFPercentual_icms_saida_fisica_TBEstado_icms
            
            'Calculando o total de ICMS do item
            'ALTERADO EM 01/07/2004  dblTotal_icms = (dblTotal_Praticado * dblAliquota_icms) / 100
            dblTotal_Icms = (dblTotal_item * dblAliquota_icms) / 100
            'Mata a recordset
            Set rstVerifica_Estado_ICMS = Nothing
        End If
        '---------------------------------------------------------------------------------------------------
        strSQL = Empty
        strSQL = "INSERT INTO TBItens_nota_saida(" & _
                 "FKId_TBNota_Saida," & _
                 "FKId_TBProduto," & _
                 "FKId_TBCfop," & _
                 "DFCst1_TBItens_nota_saida," & _
                 "DFCst2_TBItens_nota_saida," & _
                 "DFQuantidade_TBItens_nota_saida," & _
                 "DFTipo_preco_TBItens_nota_saida," & _
                 "DFPreco_tabela_TBItens_nota_saida," & _
                 "DFPercentual_desconto_TBItens_nota_saida," & _
                 "DFPreco_praticado_TBItens_nota_saida," & _
                 "DFValor_total_tabela_TBItens_nota_saida," & _
                 "DFValor_total_praticado_TBItens_nota_saida," & _
                 "DFPercentual_icms_TBItens_nota_saida," & _
                 "DFValor_total_icms_TBItens_nota_saida," & _
                 "DFUnidade_TBItens_nota_saida," & _
                 "DFPeso_liquido_TBItens_nota_saida," & _
                 "DFPeso_bruto_TBItens_nota_saida," & _
                 "DFQuantidade_baixa_estoque_TBItens_nota_saida," & _
                 "DFDivisor_baixa_estoque_TBItens_nota_saida," & _
                 "DFValor_total_item_TBItens_nota_saida,"
        strSQL = strSQL + "DFCusto_medio_TBItens_nota_saida," & _
                 "DFCusto_real_TBItens_nota_saida," & _
                 "DFCusto_contabil_TBItens_nota_saida," & _
                 "FkId_TBVendedor) " & _
                 "VALUES(" & _
                 "" & lngID_Numero_Nota & "," & _
                 "" & rstProdutos!PKId_TBProduto & "," & _
                 "" & intIDCfo & "," & _
                 "'" & strST & "'," & _
                 "'" & strST2 & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade) & "," & _
                 "1," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblValor_Unitario) & "," & _
                 "0," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblValor_Unitario) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblValor_Total) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblValor_Total) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Icms) & "," & _
                 "'" & rstProdutos!DFUnidade_venda_TBProduto & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Liquido_item) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Bruto_item) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade) & "," & _
                 "" & rstProdutos!DFFator_venda_TBProduto & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblTotal_item) & ","
        strSQL = strSQL + "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Medio) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Real) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Contabil) & "," & _
                 "" & lngVendedor & ")"
                 
        CNconexao.CNconexao.Execute strSQL
        
        'Manipulando Estoque - Gravando ocorrência
        ocorrencia.Data_Movimento = Date
        ocorrencia.Estoque_Anterior = rstProdutos!DFEstoque_atual_TBProduto
        ocorrencia.Estoque_Atual = CDbl(rstProdutos!DFEstoque_atual_TBProduto) - dblQuantidade
        ocorrencia.Hora_Movimento = Format(Now, "hh:mm:ss")
        ocorrencia.ID_Produto = rstProdutos!PKId_TBProduto
        ocorrencia.Observacao = "Inclusão de Item no Orcamento Balcão Nº:" & rstNumero_orcamento.Fields!DFProximo_orcamento_balcao_TBParametros_ecf & "- Baixa de Estoque"
        ocorrencia.Programa = "Emissão de orçamento balcão"
        ocorrencia.Quantidade_Movimentada = dblQuantidade
        ocorrencia.Usuario = frmTela_Venda.strOperador
        ocorrencia.Gravar "Otica", True, CNconexao
        
        'Manipulando Estoque - Baixando Estoque
        estoque.ID_Produto = rstProdutos!PKId_TBProduto
        estoque.Quantidade_Menor_Unidade_Item = dblQuantidade
        estoque.Quantidade_Antes_Atualizar_Estoque = rstProdutos!DFEstoque_atual_TBProduto
        estoque.Subtrair_Estoque "Otica", True, CNconexao
        
        intCont_Itens = intCont_Itens + 2

    Loop
    
    Exit Function
    
Erro:

    'Rollback na transação
    CNconexao.CNconexao.RollbackTrans
    
    Call Controle_Transacional_manual
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando os itens do cupom.Verifique"
    
    Exit Function
        
End Function

Private Function Imprime_Cupom_nao_fiscal()

    Dim rstImprime As New ADODB.Recordset
    
    'Inf. Cupons
    strSQL = Empty
    strSQL = "SELECT TBITENS_NOTA_SAIDA.*,TBNOTA_SAIDA.DFNumero_TBNota_saida,TBNOTA_SAIDA.DFSerie_TBNota_saida,TBNOTA_SAIDA.DFData_Saida_TBNota_saida," & _
             "TBEmpresa.DFRazao_Social_TBEmpresa,TBEmpresa.DFEndereco_TBEmpresa,TBEmpresa.DFBairro_TBEmpresa,TBEmpresa.DFCep_TBEmpresa," & _
             "TBEmpresa.FKId_TBCidade_otica,TBEmpresa.DFCnpj_TBEmpresa,TBEmpresa.DFInscricao_estadual_TBEmpresa," & _
             "TBEmpresa.DFTelefone_TBEmpresa , TBEmpresa.DFFax_TBEmpresa,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBCidade_otica.DFNome_TBCidade_otica,TBCidade_otica.DFPais_TBCidade_otica,TBNota_saida.DFTotal_nota_TBNota_saida " & _
             "FROM TBITENS_NOTA_SAIDA,TBNota_saida,TBEmpresa,TBCidade_otica,TBProduto " & _
             "WHERE FKId_TBNota_saida = " & lngID_Numero_Nota & " " & _
             "AND TBITENS_NOTA_SAIDA.FKId_TBNota_saida = TBNota_saida.PKId_TBNota_saida " & _
             "AND TBNOTA_SAIDA.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
             "AND TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
             "AND TBITENS_NOTA_SAIDA.FKId_TBProduto = TBProduto.PKId_TBProduto "

    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstImprime, "Otica", Me
    
    'Cabeçalho
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Empresa
    strLinha_Impressao = rstImprime!DFRazao_Social_TBEmpresa
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    'N ° Orçamento
    strLinha_Impressao = "ORÇAMENTO: " & rstImprime!DFNumero_TBNota_saida
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    'Cabeçalho
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Cabeçalho 1
    strLinha_Impressao = "CODIGO(INTERNO)             PRODUTO"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Cabeçalho 2
    strLinha_Impressao = "  QUANTIDADE   X  VLR.UNIT.   TOTAL ITEM"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    Dim dblTotal_Cupom As Double
    dblTotal_Cupom = Empty
    
    Do While rstImprime.EOF = False And rstImprime.BOF = False
       Dim strDescricao_Produto As String * 40
       Dim strCodigo_Produto As String * 10
       Dim strQuantidade As String * 8
       Dim strPreco_Unitario As String * 10
       Dim strPreco_Total_Item As String * 15
       
       strCodigo_Produto = rstImprime!IXCodigo_TBProduto
       strDescricao_Produto = rstImprime!DFDescricao_resumida_TBProduto
       strQuantidade = Format(rstImprime!DFQuantidade_TBItens_nota_saida, "#,###0.00")
       strPreco_Unitario = Format(rstImprime!DFPreco_praticado_TBItens_nota_saida, "#,###0.00")
       strPreco_Total_Item = Format(rstImprime!DFValor_total_praticado_TBItens_nota_saida, "#,###0.00")
             
       'Linha 1
       strLinha_Impressao = strCodigo_Produto & " " & strDescricao_Produto
       sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
       iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
       
       'Linha 2
       strLinha_Impressao = "      " & strQuantidade & " X " & strPreco_Unitario & " = " & strPreco_Total_Item
       sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
       iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
       
       dblTotal_Cupom = dblTotal_Cupom + rstImprime!DFValor_total_praticado_TBItens_nota_saida
       
       rstImprime.MoveNext
    Loop
    
    Set rstImprime = Nothing
    
    'Salto
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    'Rodapé Total
    strLinha_Impressao = "    Total.: " & Format(dblTotal_Cupom, "#,###0.00")
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 1, 1)
    
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Rodapé - Mensagem
    strLinha_Impressao = "Obrigado pela preferência.Volte Sempre!"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 1)
    
    'Salto
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    'Rodapé - Mensagem 2
    strLinha_Impressao = "Este documento não tem validade fiscal"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 1)
    
    'Salto
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
        
    'Impressão com CRYSTAL -------------------------------------------------------------------------------------
    
''    Dim intTamanho_string As Integer
''    Dim inttamanho_From As Integer
''    Dim strCaminho As String
''    Dim strSql_antes_from As String
''    Dim strSql_pos_from As String
''    Dim strRemontada_sql As String
''    Dim strNome_cliente As String
''    Dim adrImprime As New ADODB.Recordset
''
''    On Error GoTo Erro
''
''    'Inserindo a hora no nome da tabela
''    Tabela = "TBTEMP_RELATORIO" & time
''    Tabela = Replace(Tabela, ":", "_")
''
''    'Montando a nova string  de SQL com o INTO para criação da tabela temporária
''    intTamanho_string = Len(strSQL)
''    inttamanho_From = InStr(1, strSQL, "FROM")
''    strSql_antes_from = Mid(strSQL, 1, inttamanho_From - 1)
''    strSql_pos_from = Mid(strSQL, inttamanho_From, intTamanho_string)
''    strRemontada_sql = strSql_antes_from + "INTO " & Tabela & " " + strSql_pos_from
''
''    On Error GoTo Erro
''
''    CNconexao.CNconexao.Execute strRemontada_sql
''
''    'Abrindo a recordset com as informações da tabela temporaria
''    adrImprime.Open "SELECT * FROM " & Tabela & "", CNconexao.CNconexao, adOpenKeyset, adLockOptimistic
''
''    strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptEmissao_cupom_balcao.rpt"
''
''    DoEvents
''
''    Set Relatorio = Aplicacao.OpenReport(strCaminho)
''
''    Relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
''    Relatorio.DiscardSavedData
''
''    'Indica que a impresão é direta para a impressora
''    Relatorio.PrintOut False
''
''    crvFiltrar.ReportSource = Relatorio
''    crvFiltrar.Refresh
''    crvFiltrar.ViewReport
''
''    Set adrImprime = Nothing
''    Set Aplicacao = Nothing
''
    Exit Function
    
Erro:
'    If Err.Number = -2147206461 Then
'       MsgBox "Arquivo do relatório não encontrado, verifique! A APLICAÇÃO SERÁ REINICIADA.", vbCritical, "Only Tech"
'       End
'    End If
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando os itens do cupom.Verifique"
    
    MsgBox "Verifique, pois todas as gravações forma concluídas com sucesso,Reimprime e cancele este cupom de n° - " & strNumero_Nota & ".Verifique!", vbCritical, "Only Tech"
    
    Exit Function
    
End Function

Function Fecha_comanda()

    On Error GoTo Erro
    
    'Fechando a comnada
    strSQL = Empty
    strSQL = "UPDATE TBComanda SET DFNumero_cupom_TBComanda = " & rstNumero_orcamento.Fields!DFProximo_orcamento_balcao_TBParametros_ecf & " " & _
             "WHERE PKCodigo_TBComanda = " & frmTela_Venda.strNumero_Comanda & ""
             
    CNconexao.CNconexao.Execute strSQL
        
    Exit Function
    
Erro:

    CNconexao.CNconexao.RollbackTrans
    
    Call Controle_Transacional_manual
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando os itens do cupom.Verifique"
    
    Exit Function
    
End Function

Private Function Controle_Transacional_manual()

    On Error GoTo Erro
    
    strSQL = Empty
    strSQL = "DELETE FROM TBNOTA_SAIDA WHERE PKId_TBNota_saida = " & lngID_Numero_Nota & ""
    
    CNconexao.CNconexao.Execute strSQL
    
    Exit Function
    
Erro:

    CNconexao.CNconexao.RollbackTrans
    
    MsgBox Err.Number & "-" & Err.Description & "-" & " Transacional Manual.Verifique"
    
    Exit Function
    
End Function
