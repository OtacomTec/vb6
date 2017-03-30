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
   ClientWidth     =   11985
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
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   2560
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3210
      Width           =   3375
   End
   Begin VB.TextBox txtFalta 
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
      Left            =   2560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4290
      Width           =   3375
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11520
      Picture         =   "frmFechamento_Cupom.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   30
      Width           =   435
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   345
      Left            =   10440
      TabIndex        =   14
      Top             =   330
      Visible         =   0   'False
      Width           =   375
      lastProp        =   500
      _cx             =   661
      _cy             =   609
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
      Left            =   5790
      ScaleHeight     =   615
      ScaleWidth      =   285
      TabIndex        =   12
      Top             =   5430
      Width           =   285
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   2550
      ScaleHeight     =   645
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   5430
      Width           =   75
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   2610
      ScaleHeight     =   75
      ScaleWidth      =   3465
      TabIndex        =   10
      Top             =   5970
      Width           =   3465
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   2610
      ScaleHeight     =   75
      ScaleWidth      =   3465
      TabIndex        =   9
      Top             =   5370
      Width           =   3465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   10680
      ScaleHeight     =   4815
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   2130
      Width           =   255
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
      Left            =   2560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2190
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
      Left            =   2560
      TabIndex        =   2
      Top             =   6510
      Width           =   3375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HflexGrid 
      DragMode        =   1  'Automatic
      Height          =   4755
      Left            =   6660
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2190
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   8387
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
   Begin MSDataListLib.DataCombo dtcFinalizadora_cupom 
      Height          =   570
      Left            =   2560
      TabIndex        =   0
      ToolTipText     =   "Finalizadora"
      Top             =   5430
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
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   3645
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   3645
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Falta............:"
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
      Left            =   390
      TabIndex        =   16
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   3645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   4890
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
      TabIndex        =   13
      Top             =   540
      Width           =   4995
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Finalizadora.:"
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
      Left            =   390
      TabIndex        =   8
      Top             =   5790
      Width           =   1995
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   3675
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   5175
      Left            =   6540
      Shape           =   4  'Rounded Rectangle
      Top             =   2010
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Troco...........:"
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
      Left            =   390
      TabIndex        =   5
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   2550
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   3645
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valor Pago...:"
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
      Left            =   390
      TabIndex        =   4
      Top             =   6840
      Width           =   1995
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Cupom:"
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
      Left            =   390
      TabIndex        =   3
      Top             =   2550
      Width           =   1995
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   2010
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
      Top             =   2070
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
      Top             =   6480
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
      Top             =   3240
      Width           =   3645
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   5205
      Left            =   6450
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   4395
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   5340
      Width           =   3645
   End
End
Attribute VB_Name = "frmFechamento_Cupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim dblValor_pago As Double
'Conexões-----------------------------------------------------
Dim CNconexao As New DLLConexao_Sistema.conexao
Dim CNconexao_concentrador As New DLLConexao_Sistema.conexao
Dim CNconexao_local_pdv As New DLLConexao_Sistema.conexao
'-------------------------------------------------------------
Dim dblTotal_item As Double
Dim lngID_Numero_Nota As Long
Dim lngID_Cupom As Long
Dim lngID_Cupom_Concentrador As Long
Dim lngVendedor As Long
Dim lngPlano_pagamento As Long
Dim strNumero_Nota As String
Dim strSerie_nota As String
'Recordsets
Dim rstFinalizadora As New ADODB.Recordset
Dim rstFinalizadora_Retaguarda As New ADODB.Recordset
Dim rstProdutos As New ADODB.Recordset
Dim rstTabela As New ADODB.Recordset
Dim rstNumero_orcamento As New ADODB.Recordset
Dim Relatorio As New CRAXDRT.Report
Dim Aplicacao As New CRAXDRT.Application
Dim ocorrencia As New DLLSistema.Ocorrencia_de_produto
Dim estoque As New DLLSistema.estoque
Dim intIP_Concentrador As Long
Dim strDescr_Finalizadora As String * 15
Public lngCodigo_vendedor As Long
Dim Tabela As String
Dim intPrevisao As Integer
Dim strCod_Finalizadora As String
Public Cod_Cliente As Long
Dim lngID_Cupom_local As Long
Public booCliente_Acionado As Boolean
Public dblLimite As Double
Public dblPago_Afinidade As Double
Public lngCliente_Afinidade As Double
Public intDia_Vencimento_Cliente As Integer
Dim booImprime_direto As Boolean

Private Sub dtcFinalizadora_cupom_GotFocus()
     Call Movimentacoes.Verifica_DataCombo(Me.dtcFinalizadora_cupom)
     Me.dtcFinalizadora_cupom.SetFocus
End Sub

Private Sub dtcFinalizadora_cupom_LostFocus()
     If Me.dtcFinalizadora_cupom.Text = "" And booCliente_Acionado = False Then
        Me.dtcFinalizadora_cupom.SetFocus
     Else
       If booCliente_Acionado = False Then
       
          Dim rstTroco As New ADODB.Recordset
          Dim booTroco As Boolean
          
          strSql = Empty
          strSql = "SELECT DFTroco_TBFinalizadora FROM TBFinalizadora WHERE DFControle_venda_TBFinalizadora = 1 AND IXCodigo_TBFinalizadora = " & Me.dtcFinalizadora_cupom.BoundText & ""
          Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTroco, "Otica", Me
          
          If rstTroco.BOF = True And rstTroco.EOF = True Then
             booTroco = True
          Else
             booTroco = rstTroco!DFTroco_TBFinalizadora
          End If
          If booTroco = False Then
             Me.txtValor_pago.Text = Me.txtTotal_Cupom.Text
             Call txtValor_pago_LostFocus
          End If
          
          Set rstTroco = Nothing
          
       End If
     End If
End Sub

Private Sub Form_Load()
    Me.txtTotal_Cupom.Text = frmTela_Venda.txtPreco_total_cupom.Text
    
    booCliente_Acionado = False
    
    dblValor_pago = 0
    
    'Carregando a combo de finalizadora
    strSql = Empty
    strSql = "SELECT IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora FROM TBFinalizadora WHERE DFControle_venda_TBFinalizadora = 1"
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_cupom, strSql, "BDRetaguarda", "Otica", Me, "IXCodigo_TBFinalizadora"
    Else
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_cupom, strSql, "BDPDV", "PDV", Me
    End If
 End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If frmTela_Venda.booIntegracao_Retaguarda = True Then
'       CNconexao.CNconexao.Close
'    End If
'    CNconexao_local_pdv.Fechar_conexao
End Sub

Private Sub txtValor_pago_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtValor_pago_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Public Sub txtValor_pago_LostFocus()
    
    If txtValor_pago.Text = "" Or IsNull(txtValor_pago) Then txtValor_pago.Text = 0
    If txtValor_pago.Text = "," Then txtValor_pago.Text = 0: Me.dtcFinalizadora_cupom.Text = ""
    
    txtValor_pago.Text = Format(txtValor_pago, "#,###0.00")
    dblValor_pago = CDbl(txtValor_pago.Text) + dblValor_pago
    
    If dtcFinalizadora_cupom.Text <> "" Then
        If CDbl(Me.txtValor_pago.Text) > 0 Then
            If CDbl(dblValor_pago) - (Me.txtTotal_Cupom.Text) > 0 Then
                Me.txtTroco.Text = Format(CDbl(dblValor_pago) - (Me.txtTotal_Cupom.Text), "#,###0.00")
                Me.txtFalta.Text = Format(0, "#,###0.00")
            Else
                If CDbl(Me.txtValor_pago.Text) - (Me.txtTotal_Cupom.Text) < 0 Then
                   txtTroco.Text = Format(0, "#,###0.00")
                   Me.txtFalta.Text = Format(CDbl(Me.txtTotal_Cupom.Text) - (dblValor_pago), "#,###0.00")
                Else
                   txtFalta.Text = Format(0, "#,###0.00")
                End If
            End If
            Call Fecha_Cupom
        Else
           If frmCliente.Enabled = False Then
              Me.dtcFinalizadora_cupom.Text = ""
              Me.dtcFinalizadora_cupom.SetFocus
           Else
              Me.txtValor_pago.SetFocus
           End If
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
    'Aceita o cartão do cliente
    If KeyAscii = 3 Then
       booCliente_Acionado = True
       frmCliente.Show (1)
    End If
    
    'Habilita a volta a tela dos itens com esc
    If KeyAscii = 27 Then
        Unload Me
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
     
     'Total por finalizadora
     HflexGrid.Col = 2
     
     HflexGrid.Row = 3
     If HflexGrid.Text <> "" Then
        Do While HflexGrid.Text <> ""
           dblFinalizafora_paga = CDbl(HflexGrid.Text) + dblFinalizafora_paga
           HflexGrid.Row = HflexGrid.Row + 1
        Loop
        
        If (CDbl(dblFinalizafora_paga) + CDbl(Me.txtValor_pago.Text)) > CDbl(txtTotal_Cupom.Text) Then
           HflexGrid.Text = CDbl(txtTotal_Cupom.Text) - CDbl(dblFinalizafora_paga)
        Else
           HflexGrid.Text = Me.txtValor_pago.Text
        End If
     Else
        If CDbl(Me.txtValor_pago.Text) > CDbl(txtTotal_Cupom.Text) Then
           HflexGrid.Text = CDbl(Me.txtValor_pago.Text) - (CDbl(Me.txtValor_pago.Text) - CDbl(txtTotal_Cupom.Text))
        Else
           HflexGrid.Text = Me.txtValor_pago.Text
        End If
     End If
     
     HflexGrid.Text = Format(HflexGrid.Text, "#,###0.00")
     
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
        If frmTela_Venda.booComissao_vendedor = True And booCliente_Acionado = False Then
           frmVendedor.Show (1)
        Else
           lngCodigo_vendedor = 9999
        End If
        
        Call Grava_cupom
        
        'Trava a venda ate a gaveta estar fechada
''        If frmTela_Venda.booGaveta_integrada = True Then
''           Dim booAberta As Boolean
''
''           booAberta = Comandos_impressoras_fiscais.Status_gaveta_aberta(frmTela_Venda.strImpresora)
''
''           Do While booAberta = True
''              booAberta = Comandos_impressoras_fiscais.Status_gaveta_aberta(frmTela_Venda.strImpresora)
''           Loop
''        End If

        frmTela_Venda.Limpa_Tela
        frmTela_Venda.HflexGrid.Clear
        frmTela_Venda.HflexGrid.Rows = 2
        frmTela_Venda.txtPreco_total_cupom.Text = Empty
        
        'Fechando as conexões
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           Set CNconexao = Nothing
        End If
        
        Set CNconexao_local_pdv = Nothing
        
        'limpando o cupom em processamento
        frmTela_Venda.Limpa_Processa_Cupom
        
        Unload Me
        
     End If
     
End Function

Public Function Grava_cupom()

    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       'Abrindo uma conexão nova com o Retaguarda
       If CNconexao.CNconexao <> "" Then
          Set CNconexao = Nothing
          CNconexao.Banco = "BDRetaguarda"
          CNconexao.Abrir_conexao "Otica"
       Else
          CNconexao.Banco = "BDRetaguarda"
          CNconexao.Abrir_conexao "Otica"
       End If
    End If
    
 '   If frmTela_Venda.booIntegracao_Retaguarda = False Then
       'Abrindo uma conexão nova com o Retaguarda
       If CNconexao_local_pdv.CNconexao <> "" Then
          Set CNconexao_local_pdv = Nothing
          CNconexao_local_pdv.Banco = "BDPDV"
          CNconexao_local_pdv.Abrir_conexao "PDV"
       Else
         CNconexao_local_pdv.Banco = "BDPDV"
         CNconexao_local_pdv.Abrir_conexao "PDV"
       End If
'    End If
    'Abrindo uma conexão nova com o pdv(Banco Local)
    
    'SELECTS -----------------------------------------------------------------------------------------------
    'Finalizadoras
    strSql = Empty
    strSql = "SELECT * FROM TBFinalizadora ORDER BY IXCodigo_TBFinalizadora"
    
    On Error GoTo Correcao_bug_vb
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstFinalizadora_Retaguarda, "Otica", Me
    Else
       Movimentacoes.Select_geral strSql, "BDPDV", rstFinalizadora_Retaguarda, "PDV", Me
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se a finalizadora é obrigada a enviar direto para CF
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rstFinalizadora_Retaguarda.MoveFirst
    rstFinalizadora_Retaguarda.Find ("IXCodigo_TBFinalizadora = " & Me.dtcFinalizadora_cupom.BoundText & "")

    If IsNull(rstFinalizadora_Retaguarda!DFImprime_direto_TBFinalizadora) = True Then
       booImprime_direto = False
    End If
    
    If rstFinalizadora_Retaguarda!DFImprime_direto_TBFinalizadora = True Then
       booImprime_direto = True
    End If
    
    If rstFinalizadora_Retaguarda!DFImprime_direto_TBFinalizadora = False Then
       booImprime_direto = False
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
        'Próxima Tabela
        strSql = Empty
        strSql = "SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBParametros_venda WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
         
        Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTabela, "Otica", Me
    
        If rstTabela.BOF = True And rstTabela.EOF = True Then
           MsgBox "Não existe tabela vigente cadastrada no sistema.Venda impossibilitada de ser concluida.Verifique!", vbCritical, , "Only Tech"
           Set rstTabela = Nothing
           Set rstFinalizadora = Nothing
           Exit Function
        End If
    End If
    
    'Verifica se cupom ou orçamento
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 And booImprime_direto = True Then
       'Próximo Orçamento
        strSql = Empty
        strSql = "SELECT DFProximo_cupom_TBParametros_ecf,DFProximo_serie_cupom_TBParametros_ecf FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    Else
        'Próximo Orçamento
        strSql = Empty
        strSql = "SELECT DFProximo_orcamento_balcao_TBParametros_ecf,DFProximo_serie_orcamento_balcao_TBParametros_ecf FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    End If
    
    'Força quando a finalizadora marcar para imprimir direto, que o mesmo seja CF
    If booImprime_direto = True Then
        'Próximo Orçamento
        strSql = Empty
        strSql = "SELECT DFProximo_cupom_TBParametros_ecf,DFProximo_serie_cupom_TBParametros_ecf FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    End If
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstNumero_orcamento, "Otica", Me
    Else
       Movimentacoes.Select_geral strSql, "BDPDV", rstNumero_orcamento, "PDV", Me
    End If
    
    If rstNumero_orcamento.BOF = True And rstNumero_orcamento.EOF = True Then
       MsgBox "Número do próximo orçamento balcão não cadastrado.Venda impossibilitada de ser concluida.Verifique!", vbCritical, , "Only Tech"
       Set rstTabela = Nothing
       Set rstFinalizadora_Retaguarda = Nothing
       Set rstNumero_orcamento = Nothing
       Exit Function
    End If
    
    On Error GoTo Erro
    
    'Produto
    strSql = Empty
    strSql = "SELECT * FROM TBProduto WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " ORDER BY IXCodigo_TBProduto ASC"
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstProdutos, "Otica", Me
    Else
       Movimentacoes.Select_geral strSql, "BDPDV", rstProdutos, "PDV", Me
    End If
    
    If frmTela_Venda.booComanda = False Then
        lngVendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", "" & lngCodigo_vendedor & "", "TBVendedor", "Otica", Me, "BDRetaguarda")
        lngPlano_pagamento = Funcoes_Gerais.Localiza_ID("PKId_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", "9999", "TBPlano_pagamento", "Otica", Me, "BDRetaguarda")
    Else
        lngVendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", "" & lngCodigo_vendedor & "", "TBVendedor", "Otica", Me, "BDRetaguarda")
        lngPlano_pagamento = Funcoes_Gerais.Localiza_ID("PKId_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", "9999", "TBPlano_pagamento", "Otica", Me, "BDRetaguarda")
    End If
    
    'Abrindo Transações
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.BeginTrans
    End If
    
    '------------------------------------------------------
    'Verificar se cupom ou orçamento
    '------------------------------------------------------
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 And booImprime_direto = False Then
       strNumero_Nota = rstNumero_orcamento!DFProximo_cupom_TBParametros_ecf
       strSerie_nota = rstNumero_orcamento!DFProximo_serie_cupom_TBParametros_ecf
    Else
       'Força quando a finalizadora marcar para imprimir direto, que o mesmo seja CF
       If booImprime_direto = False Then
          strNumero_Nota = rstNumero_orcamento!DFProximo_orcamento_balcao_TBParametros_ecf
          strSerie_nota = rstNumero_orcamento!DFProximo_serie_orcamento_balcao_TBParametros_ecf
       Else
          strNumero_Nota = rstNumero_orcamento!DFProximo_cupom_TBParametros_ecf
          strSerie_nota = rstNumero_orcamento!DFProximo_serie_cupom_TBParametros_ecf
       End If
    End If
        
    CNconexao_local_pdv.CNconexao.BeginTrans
    
    Call Grava_Corpo_Nota
        
    'Comitando Transações
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Dim rstID_Titulo As New ADODB.Recordset
       'Debitando a compra do seu limite
       If Me.booCliente_Acionado = True Then
          strSql = Empty
          strSql = "UPDATE TBCliente " & _
                   "SET DFLimite_credito_TBCliente = " & Funcoes_Gerais.Grava_Moeda(dblLimite) & " " & _
                   "WHERE IXCodigo_TBCliente  = " & Me.lngCliente_Afinidade & " " & _
                   "AND IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
          CNconexao.CNconexao.Execute strSql
       End If
       
       CNconexao.CNconexao.CommitTrans
       'Baixando titulo gerado
       
       strSql = Empty
       strSql = "SELECT max(PKId_TBTitulo_receber) as ID_Titulo FROM TBTitulo_receber WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstID_Titulo, "Otica", Me
       
       If Me.booCliente_Acionado = False Then
          strSql = Empty
          strSql = "INSERT INTO TBTitulo_recebido(FKId_TBTitulo_receber,DFValor_TBTitulo_recebido,DFData_recebimento_TBTitulo_recebido,DFUsuario_TBTitulo_recebido,DFValor_diferença_TBTitulo_recebido) " & _
                   "VALUES (" & rstID_Titulo!ID_Titulo & "," & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & ",'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "','" & frmTela_Venda.strOperador & "'," & 0 & ")"
          CNconexao.CNconexao.Execute strSql
       End If
       
       Set rstID_Titulo = Nothing
       
    End If
    
    CNconexao_local_pdv.CNconexao.CommitTrans
    
    'Consultando o Id da nota gravada.
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       lngID_Numero_Nota = Funcoes_Gerais.Localiza_ID("PKId_TBNota_saida", "DFNumero_TBNota_saida", strNumero_Nota, "TBNota_saida", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", frmTela_Venda.strEmpresa_Operador, "DFSerie_TBNota_saida", strSerie_nota)
       lngID_Cupom = Funcoes_Gerais.Localiza_ID("PKId_TBCupom", "DFNumero_TBCupom", strNumero_Nota, "TBCupom", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", frmTela_Venda.strEmpresa_Operador, "DFSerie_TBCupom", strSerie_nota)
    Else
       'Ajustar para cupom
       lngID_Numero_Nota = Funcoes_Gerais.Localiza_ID("PKId_TBNota_saida", "DFNumero_TBNota_saida", strNumero_Nota, "TBNota_saida", "PDV", Me, "BDPDV", "FKCodigo_TBEmpresa", frmTela_Venda.strEmpresa_Operador, "DFSerie_TBNota_saida", strSerie_nota)
    End If
    
    'Consultando o Id do cupom no pdv local
    lngID_Cupom_local = Funcoes_Gerais.Localiza_ID("PKId_TBCupom", "DFNumero_TBCupom", strNumero_Nota, "TBCupom", "PDV", Me, "BDPDV", "FKCodigo_TBEmpresa", frmTela_Venda.strEmpresa_Operador, "DFSerie_TBCupom", strSerie_nota)
    
    'Reabrindo Transações
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.BeginTrans
    End If
    
    CNconexao_local_pdv.CNconexao.BeginTrans
    
    Call Grava_Finalizadoras(strNumero_Nota, strSerie_nota)
    Call Grava_Itens_Nota
    
    If frmTela_Venda.booComanda = True Then
        'Verifica se é cupom abastecido de comanda e finaliza a mesma
        Call Fecha_comanda
    End If
    
    'Atualizando o parâmetro ECF
    '------------------------------------------------------
    'Verificar o parametro do com ou sem cupom
    '------------------------------------------------------
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 And booImprime_direto = False Then
        strSql = Empty
        strSql = "UPDATE TBParametros_ecf SET DFProximo_cupom_TBParametros_ecf = " & rstNumero_orcamento!DFProximo_cupom_TBParametros_ecf + 1 & " " & _
                 "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    Else
        'Força quando a finalizadora marcar para imprimir direto, que o mesmo seja CF
        If booImprime_direto = False Then
           strSql = Empty
           strSql = "UPDATE TBParametros_ecf SET DFProximo_orcamento_balcao_TBParametros_ecf = " & rstNumero_orcamento!DFProximo_orcamento_balcao_TBParametros_ecf + 1 & " " & _
                    "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
        Else
            strSql = Empty
            strSql = "UPDATE TBParametros_ecf SET DFProximo_cupom_TBParametros_ecf = " & rstNumero_orcamento!DFProximo_cupom_TBParametros_ecf + 1 & " " & _
                     "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
        End If
    End If
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.Execute strSql
    Else
       CNconexao_local_pdv.CNconexao.Execute strSql
    End If
    
    'Comitando Transações
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.CommitTrans
    End If
    
    CNconexao_local_pdv.CNconexao.CommitTrans
        
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
        'Fechando o cupom
        Call Comandos_impressoras_fiscais.Fecha_Cupom(frmTela_Venda.strImpresora, strDescr_Finalizadora, "Obrigado e volte sempre!", Me.txtTotal_Cupom.Text, strCod_Finalizadora, "", Me.txtValor_pago.Text)
        
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           Call Gravar_Impostos_Nota(lngID_Numero_Nota)
        End If
        Call Gravar_Impostos_Cupom(lngID_Cupom)
        
        If booCliente_Acionado = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
           'AFINIDADE
           Dim rstCliente_afinidade As New ADODB.Recordset
           Dim strNome_Cliente As String
           
           strNome_Cliente = ""
           
           strSql = "SELECT DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBCliente  = " & Me.lngCliente_Afinidade & " "
           Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCliente_afinidade, "Otica", Me
           
           If rstCliente_afinidade.BOF = True And rstCliente_afinidade.EOF = True Then
              strNome_Cliente = ""
           End If
           
           If IsNull(rstCliente_afinidade!DFNome_TBCliente) = True Then
              strNome_Cliente = ""
           End If
           
           strNome_Cliente = rstCliente_afinidade!DFNome_TBCliente
           
           Call Comandos_impressoras_fiscais.Imprime_afinidade_vinculada(strDescr_Finalizadora, Me.txtTotal_Cupom.Text, dblLimite, strNome_Cliente, frmTela_Venda.strImpresora)
           
           Set rstCliente_afinidade = Nothing
           
        End If
        
    Else
       If booImprime_direto = False Then
          Call Imprime_Cupom_nao_fiscal
       Else
          Call Imprime_Cupom_fiscal
       End If
    End If
   
    Set rstFinalizadora = Nothing
    Set rstFinalizadora_Retaguarda = Nothing
    Set rstNumero_orcamento = Nothing
    Set rstTabela = Nothing
    Set rstProdutos = Nothing
    
    Exit Function
    
    'Trecho construido para corrigir bugs de controles do VB
Erro:

   MsgBox "ERRO AO GRAVAR CUPOM " & Err.Number & " - " & Err.Description, vbCritical, "Only Tech"
    
   Exit Function

Correcao_bug_vb:

   If Err.Number <> 91 Then
      MsgBox "ERRO AO GRAVAR CUPOM " & Err.Number & " - " & Err.Description, vbCritical, "Only Tech"
   End If
   
   Exit Function
   
End Function
Private Function Grava_Finalizadoras(Nota As String, Serie As String)

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
        
        If Me.HflexGrid.Row = 3 Then
           strDescr_Finalizadora = Me.HflexGrid.Text
        Else
           strDescr_Finalizadora = strDescr_Finalizadora & "," & Me.HflexGrid.Text
        End If
        
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           rstFinalizadora_Retaguarda.MoveFirst
           rstFinalizadora_Retaguarda.Find ("DFDescricao_TBFinalizadora = '" & Me.HflexGrid.Text & "'")
           strFinalizadora = rstFinalizadora_Retaguarda!PKId_TBFinalizadora
           If frmTela_Venda.intImpressoes_suportadas <> 1 Then
              strCod_Finalizadora = rstFinalizadora_Retaguarda!DFCodificacao_impressora_fiscal_TBFinalizadora
           End If
        Else
           rstFinalizadora.MoveFirst
           rstFinalizadora.Find ("DFDescricao_TBFinalizadora = '" & Me.HflexGrid.Text & "'")
           strFinalizadora = rstFinalizadora!PKId_TBFinalizadora
           If frmTela_Venda.intImpressoes_suportadas <> 1 Then
              strCod_Finalizadora = rstFinalizadora_Retaguarda!DFCodificacao_impressora_fiscal_TBFinalizadora
           End If
        End If
        
        strObservacao = "Cupom.: " & Nota & " Serie.: " & Serie
        
        Me.HflexGrid.Col = 2
        dblValor_pago = Me.HflexGrid.Text
       
        strCampos = "FKCodigo_TBPdv,FKId_TBFinalizadora,FKCodigo_TBOperadores_ecf,DFData_TBOperacao_caixa," & _
                    "DFHora_TBOperacao_caixa,DFValor_TBOperacao_caixa,DFTipo_operacao_TBOperacao_caixa,DFStatus_aberto_fechado_TBOperacao_caixa," & _
                    "DFObservacao_TBOperacao_caixa,FKCodigo_TBEmpresa,DFNumero_Cupom_TBOperacao_caixa"
                  
        strValores = "" & frmTela_Venda.txtNumero_check_out & "," & _
                     "" & strFinalizadora & "," & _
                     "" & frmTela_Venda.strCodigo_Operador & "," & _
                     "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
                     "'" & Format(Now, "hh:mm:ss") & "'," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblValor_pago) & "," & _
                     "1," & _
                     "0, '" & strObservacao & "'," & frmTela_Venda.strEmpresa_Operador & "," & Nota & ""
        
        'Gravando Operações
        strSql = Empty
        strSql = "INSERT INTO TBoperacao_caixa ( " & strCampos & ") VALUES ( " & strValores & ")"
        
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           CNconexao.CNconexao.Execute strSql
        End If
        
        CNconexao_local_pdv.CNconexao.Execute strSql
        
        intIndice = intIndice + 1
        
    Loop
    
    Exit Function
    
Erro:

    'Rollback na transação
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.RollbackTrans
    End If
    
    CNconexao_local_pdv.CNconexao.RollbackTrans
       
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando as finalizadoras.Verifique"
    
    Exit Function
    
End Function

Private Function Grava_Corpo_Nota()

    Dim intIDCliente_titulo As Long
    
    On Error GoTo Erro
    
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 1 Then
       intPrevisao = 0
    Else
       intPrevisao = 1
    End If
    'Força para se CF, quando finalizadora tiver parametrizado para forçar
    If booImprime_direto = True Then
       intPrevisao = 0
    End If
    
    If Cod_Cliente = 0 Or IsNull(Cod_Cliente) Then
       Cod_Cliente = 9999
    End If
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
        
        'Gravando o corpo da nota
        strSql = Empty
        strSql = "INSERT INTO TBNota_Saida(" & _
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
         strSql = strSql & "DFTotal_custo_medio_TBNota_saida," & _
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
         strSql = strSql & "VALUES (" & _
                 "" & frmTela_Venda.strEmpresa_Operador & "," & _
                 "" & rstTabela.Fields("DFNumero_tabela_vigente_TBParametros_venda") & "," & _
                 "" & lngVendedor & "," & _
                 "" & lngPlano_pagamento & "," & _
                 "9999," & _
                 "1," & _
                 "" & Cod_Cliente & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "0," & _
                 "0," & _
                 "'" & intPrevisao & "'," & _
                 "0,"
         strSql = strSql + "0," & _
                 "" & strNumero_Nota & "," & _
                 "'" & strSerie_nota & "'," & _
                 "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
                 "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
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
                 "'Nota gerada apartir da emissão de um cupom fiscal no módulo de PDV - " & frmTela_Venda.txtNumero_check_out.Text & "')"
                 
        CNconexao.CNconexao.Execute strSql
        
        'Gravando no titulo receber
        If Cod_Cliente = 0 Or IsNull(Cod_Cliente) Then
           intIDCliente_titulo = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", "9999", "TBCliente", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", CInt(frmTela_Venda.strEmpresa_Operador))
        Else
           intIDCliente_titulo = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", CStr(Cod_Cliente), "TBCliente", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", CInt(frmTela_Venda.strEmpresa_Operador))
        End If
        
        intIDPlano_titulo = Funcoes_Gerais.Localiza_ID("PKId_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", "9999", "TBPlano_pagamento", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", CInt(frmTela_Venda.strEmpresa_Operador))
        
        Dim strVencimento As Date
    
        'Verificando quando cliente acionado e lança os vencimentos de acordo com a afinidade
        If booCliente_Acionado = True Then
           If Format(frmTela_Venda.dtpData_operacao, "DD") <= intDia_Vencimento_Cliente Then
              strVencimento = intDia_Vencimento_Cliente & "/" & Format(CDate(frmTela_Venda.dtpData_operacao), "M") & "/" & Format(CDate(frmTela_Venda.dtpData_operacao), "YYYY")
           Else
              strVencimento = intDia_Vencimento_Cliente & "/" & Format(CDate(frmTela_Venda.dtpData_operacao), "M") + 1 & "/" & Format(CDate(frmTela_Venda.dtpData_operacao), "YYYY")
           End If
        Else
           strVencimento = frmTela_Venda.dtpData_operacao
        End If
        

        strSql = Empty
        strSql = "INSERT INTO TBTITULO_RECEBER (" & _
                 "FKCodigo_TBEmpresa," & _
                 "FKId_TBVendedor," & _
                 "FKId_TBCliente," & _
                 "FKId_TBPlano_pagamento," & _
                 "DFTipo_documento_TBTitulo_receber," & _
                 "DFNumero_documento_TBTitulo_receber," & _
                 "DFLetra_TBTitulo_receber," & _
                 "DFData_emissao_TBTitulo_receber," & _
                 "DFData_vencimento_TBTitulo_receber," & _
                 "DFValor_TBTitulo_receber," & _
                 "DFObervacao_TBTitulo_receber," & _
                 "DFNumero_gerado_TBTitulo_receber," & _
                 "DFLetra_gerada_TBTitulo_receber," & _
                 "DFNosso_numero_TBTitulo_receber," & _
                 "DFNosso_numero_digito_TBTitulo_receber," & _
                 "DFCarteira_TBTitulo_receber," & _
                 "FKCodigo_TBBancos," & _
                 "DFPrevisao_TBTitulo_receber," & _
                 "DFSerie_TBTitulo_receber,DFJa_impresso_TBTitulo_receber,DFIntegrado_banco_TBTitulo_receber) " & _
                 "VALUES("
        strSql = strSql + "" & frmTela_Venda.strEmpresa_Operador & "," & _
                 "" & lngVendedor & "," & _
                 "" & intIDCliente_titulo & "," & _
                 "" & intIDPlano_titulo & "," & _
                 "1," & _
                 "" & strNumero_Nota & "," & _
                 "'A'," & _
                 "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
                 "'" & Format(strVencimento, "YYYYMMDD") & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "'Titulo gerado automaticamente pela movimentação de Emissão de Cupom no PDV - " & frmTela_Venda.txtNumero_check_out.Text & " . Cupom N°: " & strNumero_Nota & " - Série: " & strSerie_nota & "'," & _
                 "0," & _
                 "' '," & _
                 "''," & _
                 "''," & _
                 "''," & _
                 "0," & _
                 "'" & intPrevisao & "'," & _
                 "'" & strSerie_nota & "'," & 0 & "," & 0 & ")"
                 
        CNconexao.CNconexao.Execute strSql
        
        'Gravando o corpo do cupom, com inf pertinentes e integradas no retaguarda.ex Vendedor
        strSql = Empty
        strSql = "INSERT INTO TBCupom(" & _
                 "FKCodigo_TBEmpresa," & _
                 "FKId_TBVendedor," & _
                 "PKCodigo_TBPdv," & _
                 "DFTipo_operacao_TBCupom," & _
                 "DFNumero_TBCupom," & _
                 "DFSerie_TBCupom," & _
                 "DFEmitente_TBCupom," & _
                 "DFTotal_itens_TBCupom," & _
                 "DFTotal_cupom_TBCupom," & _
                 "DFTotal_cupom_tabela_TBCupom," & _
                 "DFData_Saida_TBCupom," & _
                 "DFHora_Saida," & _
                 "FKCodigo_TBOperadores_ecf," & _
                 "DFPrevisao_TBCupom," & _
                 "DFCancelado_TBCupom," & _
                 "DFMotivo_cancelamento_TBCupom," & _
                 "DFUsuario_cancelamento_TBCupom," & _
                 "DFIntegrado_fiscal_TBCupom," & _
                 "DFBase_calculo_subst_tributaria_TBCupom," & _
                 "DFValor_subst_tributaria_TBCupom," & _
                 "DFObservacao_TBCupom," & _
                 "DFCupom_Registrado_TBCupom) "
        strSql = strSql + "VALUES (" & _
                 "" & frmTela_Venda.strEmpresa_Operador & "," & _
                 "" & lngVendedor & "," & _
                 "" & frmTela_Venda.txtNumero_check_out.Text & "," & _
                 "1," & _
                 "" & strNumero_Nota & "," & _
                 "'" & strSerie_nota & "'," & _
                 "" & Cod_Cliente & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
                 "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
                 "'" & Format(Now, "HH:MM:SS") & "'," & _
                 "" & frmTela_Venda.strCodigo_Operador & "," & _
                 "'" & intPrevisao & "'," & _
                 "''," & _
                 "''," & _
                 "''," & _
                 "''," & _
                 "0," & _
                 "0," & _
                 "'Cupom gerado apartir da emissão de um cupom fiscal no módulo de PDV - " & frmTela_Venda.txtNumero_check_out.Text & " . Cupom N°: " & strNumero_Nota & " - Série: " & strSerie_nota & "'," & _
                 "'0')"
                 
        CNconexao.CNconexao.Execute strSql
        
    End If
    
    'Gravando no pdv LOCAL
    Dim intIDvendedor As Long
    
    intIDvendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", "" & lngCodigo_vendedor & "", "TBVendedor", "PDV", Me, "BDPDV", "IXCodigo_TBEmpresa", CInt(frmTela_Venda.strEmpresa_Operador))
        
    'Gravando o corpo do cupom
    strSql = Empty
    strSql = "INSERT INTO TBCupom(" & _
             "FKCodigo_TBEmpresa," & _
             "FKId_TBVendedor," & _
             "PKCodigo_TBPdv," & _
             "DFTipo_operacao_TBCupom," & _
             "DFNumero_TBCupom," & _
             "DFSerie_TBCupom," & _
             "DFEmitente_TBCupom," & _
             "DFTotal_itens_TBCupom," & _
             "DFTotal_cupom_TBCupom," & _
             "DFTotal_cupom_tabela_TBCupom," & _
             "DFData_Saida_TBCupom," & _
             "DFHora_Saida," & _
             "FKCodigo_TBOperadores_ecf," & _
             "DFPrevisao_TBCupom," & _
             "DFCancelado_TBCupom," & _
             "DFMotivo_cancelamento_TBCupom," & _
             "DFUsuario_cancelamento_TBCupom," & _
             "DFIntegrado_fiscal_TBCupom," & _
             "DFBase_calculo_subst_tributaria_TBCupom," & _
             "DFValor_subst_tributaria_TBCupom," & _
             "DFObservacao_TBCupom," & _
             "DFCupom_Registrado_TBCupom) "
    strSql = strSql + "VALUES (" & _
             "" & frmTela_Venda.strEmpresa_Operador & "," & _
             "" & intIDvendedor & "," & _
             "" & frmTela_Venda.txtNumero_check_out.Text & "," & _
             "1," & _
             "" & strNumero_Nota & "," & _
             "'" & strSerie_nota & "'," & _
             "" & Cod_Cliente & "," & _
             "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
             "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
             "" & Funcoes_Gerais.Grava_Moeda(frmTela_Venda.txtPreco_total_cupom.Text) & "," & _
             "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
             "'" & Format(Now, "HH:MM:SS") & "'," & _
             "" & frmTela_Venda.strCodigo_Operador & "," & _
             "'" & intPrevisao & "'," & _
             "''," & _
             "''," & _
             "''," & _
             "''," & _
             "0," & _
             "0," & _
             "'Cupom gerado apartir da emissão de um cupom fiscal no módulo de PDV - " & frmTela_Venda.txtNumero_check_out.Text & " . Cupom N°: " & strNumero_Nota & " - Série: " & strSerie_nota & "'," & _
             "'0')"
    
    CNconexao_local_pdv.CNconexao.Execute strSql
       
    Exit Function
    
Erro:

    'Rollback na transação
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.RollbackTrans
    End If
    
    CNconexao_local_pdv.CNconexao.RollbackTrans
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando o corpo do cupom!Verifique.", vbCritical, "Only Tech"
    
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
    strSql = Empty
    strSql = "SELECT TBCidade_otica.DFUf_TBCidade_otica FROM TBEmpresa " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBEmpresa.Fkid_TBCidade_otica  = TBCidade_otica.pkid_TBCidade_otica " & _
             "WHERE TBEmpresa.PKCodigo_TBempresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstUF, "Otica", Me)
    Else
       Call Movimentacoes.Select_geral(strSql, "BDPDV", rstUF, "PDV", Me)
    End If
    
    strUF_Emitente = rstUF!DFUf_TBCidade_otica
    
    Set rstUF = Nothing
    
    '--CFO-------------------------------------------------------------------------------------------------
    
    Dim rstCFO As New ADODB.Recordset
    Dim strCodigo_cfo As String
    
    'Verifica O CFO no parametro
    strSql = Empty
    strSql = "SELECT DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
             "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
             
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstCFO, "Otica", Me)
    Else
       Call Movimentacoes.Select_geral(strSql, "BDPDV", rstCFO, "PDV", Me)
    End If
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCFOP", "DFCodigo_TBCfop", rstCFO!CFO, "TBCFOP", "Otica", Me, "BDRetaguarda")
    Else
       intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCFOP", "DFCodigo_TBCfop", rstCFO!CFO, "TBCFOP", "PDV", Me, "BDPDV")
    End If
    
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
               If frmTela_Venda.booIntegracao_Retaguarda = True Then
                  strSql = "SELECT FKId_TBProduto FROM TBCodigo_barras " & _
                  "INNER JOIN TBProduto ON TBCodigo_barras.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                  "WHERE IXCodigo_TBCodigo_barras = " & frmTela_Venda.HflexGrid.Text & " " & _
                  "AND TBProduto.IXCodigo_TBEmpresa = " & CInt(frmTela_Venda.strEmpresa_Operador) & ""
                  
                  Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCodigo_Interno, "Otica", Me
                  
                  If Not IsNull(rstCodigo_Interno.Fields("FKId_TBProduto")) And rstCodigo_Interno.RecordCount <> 0 Then
                     strID_Produto = rstCodigo_Interno.Fields("FKId_TBProduto")
                  Else
                     MsgBox "Produto Inválido para venda.", vbCritical, "Only Tech"
                  End If
               Else
                  'strID_Produto = Funcoes_Gerais.Localiza_ID("FKId_TBProduto", "IXCodigo_TBCodigo_barras", frmTela_Venda.HflexGrid.Text, "TBCodigo_barras", "PDV", Me, "BDPDV", "IXCodigo_TBEmpresa", CInt(frmTela_Venda.strEmpresa_Operador))
                  
                  strSql = "SELECT FKId_TBProduto FROM TBCodigo_barras " & _
                           "INNER JOIN TBProduto ON TBCodigo_barras.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                           "WHERE IXCodigo_TBCodigo_barras = " & frmTela_Venda.HflexGrid.Text & " " & _
                           "AND IXCodigo_TBEmpresa = " & CInt(frmTela_Venda.strEmpresa_Operador) & " "
                  
                  Movimentacoes.Select_geral strSql, "BDPDV", rstCodigo_Interno, "PDV", Me
                  
                  If Not IsNull(rstCodigo_Interno.Fields("FKId_TBProduto")) And rstCodigo_Interno.RecordCount <> 0 Then
                     strID_Produto = rstCodigo_Interno.Fields("FKId_TBProduto")
                  Else
                     MsgBox "Produto Inválido para venda.", vbCritical, "Only Tech"
                  End If
               End If
               Set rstCodigo_Interno = Nothing
               
               strSql = Empty
               strSql = "SELECT IXCodigo_TBProduto FROM TBProduto WHERE PKId_TBProduto = " & strID_Produto & " "

               If frmTela_Venda.booIntegracao_Retaguarda = True Then
                  Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCodigo_Interno, "Otica", Me
               Else
                  Movimentacoes.Select_geral strSql, "BDPDV", rstCodigo_Interno, "PDV", Me
               End If
               
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
           
           strSql = Empty
           strSql = "SELECT TBCidade_otica.DFUf_TBCidade_otica " & _
                    "FROM TBEmpresa " & _
                    "INNER JOIN TBCidade_otica " & _
                    "ON TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                    "WHERE PKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                    
           If frmTela_Venda.booIntegracao_Retaguarda = True Then
              Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVerifica_Estado_ST, "Otica", Me)
           Else
              Call Movimentacoes.Select_geral(strSql, "BDPDV", rstVerifica_Estado_ST, "PDV", Me)
           End If
           
           If rstVerifica_Estado_ST!DFUf_TBCidade_otica = strUF_Emitente Then
              'Localizando no parametro o proximo cfo de substituição para dentro do estado
              strSql = Empty
              strSql = "SELECT DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
                       "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
                       
              If frmTela_Venda.booIntegracao_Retaguarda = True Then
                 Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCFO_ST, "Otica", Me
              Else
                 Movimentacoes.Select_geral strSql, "BDPDV", rstCFO_ST, "PDV", Me
              End If
           Else
              'Localizando no parametro o proximo cfo de substituição para dentro do estado
              strSql = Empty
              strSql = "SELECT DFProximo_cfop_venda_fora_estado_substituicao_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
                       "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
                       
              If frmTela_Venda.booIntegracao_Retaguarda = True Then
                 Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCFO_ST, "Otica", Me
              Else
                 Movimentacoes.Select_geral strSql, "BDPDV", rstCFO_ST, "PDV", Me
              End If
              
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
           
           If frmTela_Venda.booIntegracao_Retaguarda = True Then
              intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", rstCFO_ST.Fields("CFO"), "TBCFOP", "Otica", Me, "BDRetaguarda")
           Else
              intIDCfo = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", rstCFO_ST.Fields("CFO"), "TBCFOP", "PDV", Me, "BDPDV")
           End If
                
           If rstCFO_ST.BOF = True And rstCFO_ST.EOF = True Then
              MsgBox "Verifique se o CFO na tabela de parâmetros fiscais está preenchida corretamente!", vbCritical, "Only Tech"
           End If
           
           Set rstVerifica_Estado_ST = Nothing
           Set rstCFO_ST = Nothing
        Else
            Dim rstVerifica_Estado_ICMS As New ADODB.Recordset
            'Query para pegar ICMS do item
            strSql = Empty
            strSql = "SELECT " & _
                     "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                     "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                     "FROM TBEstado_icms " & _
                     "WHERE FKId_TBProduto = " & rstProdutos!PKId_TBProduto & " " & _
                     "AND DFUf_TBEstado_icms = '" & strUF_Emitente & "'"
                     
            If frmTela_Venda.booIntegracao_Retaguarda = True Then
               Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
            Else
               Movimentacoes.Select_geral strSql, "BDPDV", rstVerifica_Estado_ICMS, "PDV", Me
            End If
            
            If rstVerifica_Estado_ICMS.BOF = True And rstVerifica_Estado_ICMS.EOF = True Then
               Set rstVerifica_Estado_ICMS = Nothing
               'Query para pegar ICMS do item, com estado **
               strSql = Empty
               strSql = "SELECT " & _
                        "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                        "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                        "FROM TBEstado_icms " & _
                        "WHERE FKId_TBProduto = " & rstProdutos!PKId_TBProduto & " " & _
                        "AND DFUf_TBEstado_icms = '**' "
               If frmTela_Venda.booIntegracao_Retaguarda = True Then
                  Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_Estado_ICMS, "Otica", Me
               Else
                  Movimentacoes.Select_geral strSql, "BDPDV", rstVerifica_Estado_ICMS, "PDV", Me
               End If
            End If
            
            dblAliquota_icms = rstVerifica_Estado_ICMS!DFPercentual_icms_saida_fisica_TBEstado_icms
            
            'Calculando o total de ICMS do item
            'ALTERADO EM 01/07/2004  dblTotal_icms = (dblTotal_Praticado * dblAliquota_icms) / 100
            dblTotal_Icms = (dblTotal_item * dblAliquota_icms) / 100
            'Mata a recordset
            Set rstVerifica_Estado_ICMS = Nothing
        End If
        '---------------------------------------------------------------------------------------------------
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
            strSql = Empty
            strSql = "INSERT INTO TBItens_nota_saida(" & _
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
            strSql = strSql + "DFCusto_medio_TBItens_nota_saida," & _
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
                     "" & Funcoes_Gerais.Grava_Moeda(dblAliquota_icms) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Icms) & "," & _
                     "'" & rstProdutos!DFUnidade_venda_TBProduto & "'," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Liquido_item) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Bruto_item) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade) & "," & _
                     "" & rstProdutos!DFFator_venda_TBProduto & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblTotal_item) & ","
            strSql = strSql + "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Medio) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Real) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Contabil) & "," & _
                     "" & lngVendedor & ")"
                     
            CNconexao.CNconexao.Execute strSql
            
            strSql = Empty
            strSql = "INSERT INTO TBItens_cupom(" & _
                     "FKId_TBCupom," & _
                     "DFCodigo_TBProduto," & _
                     "DFCst1_TBItens_cupom," & _
                     "DFCst2_TBItens_cupom," & _
                     "DFQuantidade_TBItens_cupom," & _
                     "DFTipo_preco_TBItens_cupom," & _
                     "DFPreco_tabela_TBItens_cupom," & _
                     "DFPercentual_desconto_TBItens_cupom," & _
                     "DFPreco_praticado_TBItens_cupom," & _
                     "DFValor_total_tabela_TBItens_cupom," & _
                     "DFValor_total_praticado_TBItens_cupom," & _
                     "DFPercentual_icms_TBItens_cupom," & _
                     "DFValor_total_icms_TBItens_cupom," & _
                     "DFUnidade_TBItens_cupom," & _
                     "DFCusto_real_TBItens_cupom," & _
                     "DFCusto_contabil_TBItens_cupom," & _
                     "DFCusto_medio_TBItens_cupom," & _
                     "DFPeso_liquido_TBItens_cupom," & _
                     "DFPeso_bruto_TBItens_cupom," & _
                     "DFQuantidade_baixa_estoque_TBItens_cupom," & _
                     "DFValor_total_item_TBItens_cupom," & _
                     "DFDivisor_baixa_estouqe_TBItens_cupom," & _
                     "DFItens_cupom_Registrado_TBItens_cupom) "
            strSql = strSql + "VALUES(" & _
                     "" & lngID_Cupom & "," & _
                     "" & strCodigo_item & "," & _
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
                     "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Real) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Contabil) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Medio) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Liquido_item) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Bruto_item) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade) & "," & _
                     "" & Funcoes_Gerais.Grava_Moeda(dblTotal_item) & "," & rstProdutos!DFFator_venda_TBProduto & "," & _
                     "1)"
                     
            CNconexao.CNconexao.Execute strSql
        
            'Manipulando Estoque - Gravando ocorrência
            ocorrencia.Data_Movimento = Date
            ocorrencia.Estoque_Anterior = rstProdutos!DFEstoque_atual_TBProduto
            ocorrencia.Estoque_Atual = CDbl(rstProdutos!DFEstoque_atual_TBProduto) - dblQuantidade
            ocorrencia.Hora_Movimento = Format(Now, "hh:mm:ss")
            ocorrencia.ID_Produto = rstProdutos!PKId_TBProduto
            ocorrencia.Observacao = "Inclusão de Item no Orcamento Balcão Nº:" & strNumero_Nota & "- Baixa de Estoque"
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
        End If
        
        strSql = Empty
        strSql = "INSERT INTO TBItens_cupom(" & _
                 "FKId_TBCupom," & _
                 "DFCodigo_TBProduto," & _
                 "DFCst1_TBItens_cupom," & _
                 "DFCst2_TBItens_cupom," & _
                 "DFQuantidade_TBItens_cupom," & _
                 "DFTipo_preco_TBItens_cupom," & _
                 "DFPreco_tabela_TBItens_cupom," & _
                 "DFPercentual_desconto_TBItens_cupom," & _
                 "DFPreco_praticado_TBItens_cupom," & _
                 "DFValor_total_tabela_TBItens_cupom," & _
                 "DFValor_total_praticado_TBItens_cupom," & _
                 "DFPercentual_icms_TBItens_cupom," & _
                 "DFValor_total_icms_TBItens_cupom," & _
                 "DFUnidade_TBItens_cupom," & _
                 "DFCusto_real_TBItens_cupom," & _
                 "DFCusto_contabil_TBItens_cupom," & _
                 "DFCusto_medio_TBItens_cupom," & _
                 "DFPeso_liquido_TBItens_cupom," & _
                 "DFPeso_bruto_TBItens_cupom," & _
                 "DFQuantidade_baixa_estoque_TBItens_cupom," & _
                 "DFValor_total_item_TBItens_cupom," & _
                 "DFDivisor_baixa_estouqe_TBItens_cupom," & _
                 "DFItens_cupom_Registrado_TBItens_cupom) "
        strSql = strSql + "VALUES(" & _
                 "" & lngID_Cupom_local & "," & _
                 "" & strCodigo_item & "," & _
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
                 "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Real) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Contabil) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Medio) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Liquido_item) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Bruto_item) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(dblTotal_item) & "," & rstProdutos!DFFator_venda_TBProduto & "," & _
                 "1)"
                 
        CNconexao_local_pdv.CNconexao.Execute strSql
        
    Loop
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
        'Gravando o dado CFO-Substituição na tabela CFO-PEDIDO
        strSql = Empty
        strSql = "INSERT INTO TBCfop_nota_saida(FKId_TBCfop,FKId_TBNota_saida) VALUES ( " & intIDCfo & "," & lngID_Numero_Nota & ")"
        CNconexao.CNconexao.Execute strSql
    End If
        
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
    Dim dblTotal_Cupom As Double
    Dim strDescricao_Produto As String * 40
    Dim strCodigo_Produto As String * 10
    Dim strQuantidade As String * 8
    Dim strPreco_Unitario As String * 10
    Dim strPreco_Total_Item As String * 15
    
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
    
        'Inf. Nota de Saída
        strSql = Empty
        strSql = "SELECT TBITENS_NOTA_SAIDA.DFValor_total_praticado_TBItens_nota_saida AS TOTAL_ITEM,TBITENS_NOTA_SAIDA.DFPreco_praticado_TBItens_nota_saida AS PRECO_PRATICADO,TBITENS_NOTA_SAIDA.DFQuantidade_TBItens_nota_saida AS QUANTIDADE,TBNOTA_SAIDA.DFNumero_TBNota_saida AS NUM_CUPOM,TBNOTA_SAIDA.DFSerie_TBNota_saida,TBNOTA_SAIDA.DFData_Saida_TBNota_saida AS DATA," & _
                 "TBEmpresa.DFRazao_Social_TBEmpresa,TBEmpresa.DFEndereco_TBEmpresa,TBEmpresa.DFBairro_TBEmpresa,TBEmpresa.DFCep_TBEmpresa,TBNOTA_SAIDA.DFEmitente_TBNota_saida, " & _
                 "TBEmpresa.FKId_TBCidade_otica,TBEmpresa.DFCnpj_TBEmpresa,TBEmpresa.DFInscricao_estadual_TBEmpresa," & _
                 "TBEmpresa.DFTelefone_TBEmpresa , TBEmpresa.DFFax_TBEmpresa,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBCidade_otica.DFNome_TBCidade_otica,TBCidade_otica.DFPais_TBCidade_otica,TBNota_saida.DFTotal_nota_TBNota_saida " & _
                 "FROM TBITENS_NOTA_SAIDA,TBNota_saida,TBEmpresa,TBCidade_otica,TBProduto " & _
                 "WHERE FKId_TBNota_saida = " & lngID_Numero_Nota & " " & _
                 "AND TBITENS_NOTA_SAIDA.FKId_TBNota_saida = TBNota_saida.PKId_TBNota_saida " & _
                 "AND TBNOTA_SAIDA.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                 "AND TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                 "AND TBITENS_NOTA_SAIDA.FKId_TBProduto = TBProduto.PKId_TBProduto "
        
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstImprime, "Otica", Me
    Else
       'Inf. Cupons
       strSql = Empty
       strSql = "SELECT TBCUPOM.DFNumero_TBCupom AS NUM_CUPOM,TBCUPOM.DFSerie_TBCupom,TBCUPOM.DFTotal_cupom_TBCupom,TBCUPOM.DFData_Saida_TBCupom AS DATA," & _
                "TBCUPOM.DFHora_Saida,TBCUPOM.DFOperador_TBCupom,TBITENS_CUPOM.DFCodigo_TBProduto,TBITENS_CUPOM.DFQuantidade_TBItens_cupom AS QUANTIDADE," & _
                "TBEmpresa.DFRazao_Social_TBEmpresa,TBEmpresa.DFEndereco_TBEmpresa,TBEmpresa.DFBairro_TBEmpresa,TBEmpresa.DFCep_TBEmpresa,TBNOTA_SAIDA.DFEmitente_TBNota_saida," & _
                "TBEmpresa.FKId_TBCidade_otica,TBEmpresa.DFCnpj_TBEmpresa,TBEmpresa.DFInscricao_estadual_TBEmpresa," & _
                "TBEmpresa.DFTelefone_TBEmpresa , TBEmpresa.DFFax_TBEmpresa,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBCidade_otica.DFNome_TBCidade_otica,TBCidade_otica.DFPais_TBCidade_otica," & _
                "TBITENS_CUPOM.DFPreco_praticado_TBItens_cupom AS PRECO_PRATICADO,TBITENS_CUPOM.DFValor_total_item_TBItens_cupom AS TOTAL_ITEM,TBITENS_CUPOM.DFUnidade_TBItens_cupom " & _
                "FROM TBCUPOM,TBEmpresa,TBCidade_otica,TBProduto " & _
                "WHERE TBCUPOM.PKId_TBCupom = TBITENS_CUPOM.FKID_TBCupom " & _
                "AND TBCUPOM.PKId_TBCupom = " & lngID_Cupom & " " & _
                "AND TBCUPOM.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                "AND TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                "AND TBITENS_CUPOM.FKId_TBProduto = TBProduto.PKId_TBProduto "
                
       Movimentacoes.Select_geral strSql, "BDPDV", rstImprime, "PDV", Me
    End If
    
    If frmTela_Venda.intTipo_imp_orcamento = 0 Then
    
        'Cabeçalho
        strLinha_Impressao = "-----------------------------------------------------------"
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
        
        'Empresa
        strLinha_Impressao = rstImprime!DFRazao_Social_TBEmpresa
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
        
        'N° Orçamento
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           strLinha_Impressao = "ORÇAMENTO: " & rstImprime!NUM_CUPOM & "      " & "DATA - HORA: " & Format(rstImprime!Data, "DD/MM/YYYY") & " - " & Format(Now, "HH:MM:SS")
        Else
           strLinha_Impressao = "ORÇAMENTO: " & rstImprime!NUM_CUPOM & "      " & "DATA - HORA: " & Format(rstImprime!Data, "DD/MM/YYYY") & " - " & Format(rstImprime!DFHora_Saida_TBCupom, "HH:MM:SS")
        End If
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
                
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
                
        dblTotal_Cupom = Empty
        
        Do While rstImprime.EOF = False And rstImprime.BOF = False
           strCodigo_Produto = rstImprime!IXCodigo_TBProduto
           strDescricao_Produto = rstImprime!DFDescricao_resumida_TBProduto
           strQuantidade = Format(rstImprime!Quantidade, "#,###0.00")
           strPreco_Unitario = Format(rstImprime!PRECO_PRATICADO, "#,###0.00")
           strPreco_Total_Item = Format(rstImprime!TOTAL_ITEM, "#,###0.00")
                 
           'Linha 1
           strLinha_Impressao = strCodigo_Produto & " " & strDescricao_Produto
           sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
           iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
           
           'Linha 2
           strLinha_Impressao = "      " & strQuantidade & " X " & strPreco_Unitario & " = " & strPreco_Total_Item
           sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
           iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
           
           dblTotal_Cupom = dblTotal_Cupom + rstImprime!TOTAL_ITEM
           
           rstImprime.MoveNext
        Loop
        
        'Salto
        strLinha_Impressao = "  "
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
        
        'Linha
        strLinha_Impressao = "-----------------------------------------------------------"
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
        
        'Rodapé Total
        strLinha_Impressao = "TOTAL.: " & Format(dblTotal_Cupom, "#,###0.00")
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 1, 1)
        
        strLinha_Impressao = "-----------------------------------------------------------"
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
        
        'Rodapé - Mensagem
        strLinha_Impressao = "           OBRIGADO PELA PREFERENCIA.VOLTE SEMPRE!"
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
        
        'Salto
        strLinha_Impressao = "  "
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 1, 0)
        
        'Rodapé - Mensagem 2
        strLinha_Impressao = "      --- ESTE DOCUMENTO NAO POSSUI VALIDADE FISCAL  --- "
        sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
        iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
        
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
    End If
    
    'Impressão na porta da impressora
    If frmTela_Venda.intTipo_imp_orcamento = 1 Then
        'Cabeçalho
        strLinha_Impressao = "-----------------------------------------------------------"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, 'salta uma linha
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        'Empresa
        strLinha_Impressao = rstImprime!DFRazao_Social_TBEmpresa
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        'N ° Orçamento
        If frmTela_Venda.booIntegracao_Retaguarda = True Then
           strLinha_Impressao = "LST COMPRA: " & rstImprime!NUM_CUPOM & "      " & "DATA - HORA: " & Format(rstImprime!Data, "DD/MM/YYYY") & " - " & Format(Now, "HH:MM:SS")
        Else
           strLinha_Impressao = "LST COMPRA: " & rstImprime!NUM_CUPOM & "      " & "DATA - HORA: " & Format(rstImprime!Data, "DD/MM/YYYY") & " - " & Format(rstImprime!DFHora_Saida_TBCupom, "HH:MM:SS")
        End If
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                
        'Cabeçalho
        strLinha_Impressao = "-----------------------------------------------------------"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        'Cabeçalho 1
        strLinha_Impressao = "CODIGO(INTERNO)             PRODUTO"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                
        'Cabeçalho 2
        strLinha_Impressao = "  QUANTIDADE   X  VLR.UNIT.   TOTAL ITEM"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        strLinha_Impressao = "-----------------------------------------------------------"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        Do While rstImprime.EOF = False And rstImprime.BOF = False
           
           strCodigo_Produto = rstImprime!IXCodigo_TBProduto
           strDescricao_Produto = rstImprime!DFDescricao_resumida_TBProduto
           strQuantidade = Format(rstImprime!Quantidade, "#,###0.00")
           strPreco_Unitario = Format(rstImprime!PRECO_PRATICADO, "#,###0.00")
           strPreco_Total_Item = Format(rstImprime!TOTAL_ITEM, "#,###0.00")
                 
           'Linha 1
           strLinha_Impressao = strCodigo_Produto & " " & strDescricao_Produto
           'Mandando o comando direto na porta da impressora
           Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
           Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
           Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
           'Linha 2
           strLinha_Impressao = "      " & strQuantidade & " X " & strPreco_Unitario & " = " & strPreco_Total_Item
           'Mandando o comando direto na porta da impressora
           Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
           Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
           Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
           
           dblTotal_Cupom = dblTotal_Cupom + rstImprime!TOTAL_ITEM
           
           rstImprime.MoveNext
        Loop
        
        strLinha_Impressao = "-----------------------------------------------------------"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        'Rodapé Total
        strLinha_Impressao = "TOTAL.: " & Format(dblTotal_Cupom, "#,###0.00")
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        strLinha_Impressao = "-----------------------------------------------------------"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        
        'Rodapé - Mensagem
        strLinha_Impressao = "          OBRIGADO PELA PREFERENCIA.VOLTE SEMPRE!"
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, 'salta uma linha
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                
        'Rodapé - Mensagem 2
        strLinha_Impressao = "     --- ESTE DOCUMENTO NAO POSSUI VALIDADE FISCAL  --- "
        'Mandando o comando direto na porta da impressora
        Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
        Print #1, 'salta uma linha
        Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
    End If
    
    'AFINIDADE
    If Me.booCliente_Acionado = True Then
    
        Dim rstAfinidade As New ADODB.Recordset
        
        rstImprime.MoveFirst
        
        strSql = Empty
        strSql = "SELECT * FROM TBCLIENTE WHERE IXCODIGO_TBCLIENTE = " & rstImprime!DFEmitente_TBNota_saida & " " & _
                 "AND IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & " "
        Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAfinidade, "Otica", Me
        
        If frmTela_Venda.intTipo_imp_orcamento = 0 Then
            'Cabeçalho
            strLinha_Impressao = "-----------------------------------------------------------"
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
            
            'Empresa
            strLinha_Impressao = rstImprime!DFRazao_Social_TBEmpresa
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
            
            'N ° Orçamento
            If frmTela_Venda.booIntegracao_Retaguarda = True Then
               strLinha_Impressao = "LST.COMPRA-AFINIDADE: " & rstImprime!NUM_CUPOM & " DATA - HORA: " & Format(rstImprime!Data, "DD/MM/YYYY") & " - " & Format(Now, "HH:MM:SS")
            Else
               strLinha_Impressao = "LST.COMPRA-AFINIDADE: " & rstImprime!NUM_CUPOM & " DATA - HORA: " & Format(rstImprime!Data, "DD/MM/YYYY") & " - " & Format(Now, "HH:MM:SS")
            End If
            
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
                    
            'CLIENTE
            strLinha_Impressao = "-----------------------------------------------------------"
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
                        
            strLinha_Impressao = "CLIENTE..: " & rstAfinidade!IXCodigo_TBCliente & " - " & rstAfinidade!DFNome_TBCliente
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
            
            strLinha_Impressao = "ENDERECO.: " & rstAfinidade!DFEndereco_TBCliente & "," & " & rstAfinidade!DFNumero_TBCliente & "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
            
            strLinha_Impressao = "DIA VENC.: " & rstAfinidade!DFDia_vencimento_TBCliente & "- SALDO RESTANTE.:" & " & rstAfinidade!DFLimite_credito_TBCliente & "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
            
            'Cabeçalho
            strLinha_Impressao = "-----------------------------------------------------------"
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
            
            'Salto
            strLinha_Impressao = "  "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
            
            'Total
            strLinha_Impressao = "TOTAL.: " & Format(Me.dblPago_Afinidade, "#,###0.00")
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 1, 1)
            
            'Salto
            strLinha_Impressao = "  "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
            
            'Empresa
            strLinha_Impressao = "RECONHECO E PAGAREI A DIVIDA AQUI REPRESENTADA"
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
            
            'Salto
            strLinha_Impressao = "  "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
            
            'Assinatura
            strLinha_Impressao = "          ----------------------------------                "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
            
            'Nome do Cliente
            strLinha_Impressao = "        " & rstAfinidade!DFNome_TBCliente
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
                    
            'Salto
            strLinha_Impressao = "  "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
            
            strLinha_Impressao = "-----------------------------------------------------------"
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
            
            'Rodapé - Mensagem
            strLinha_Impressao = "         OBRIGADO PELA PREFERENCIA.VOLTE SEMPRE!"
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 1)
            
            'Salto
            strLinha_Impressao = "  "
            sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
            iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
            
            'Rodapé - Mensagem 2
            strLinha_Impressao = "    --- ESTE DOCUMENTO NAO POSSUI VALIDADE FISCAL  --- "
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
        End If
        
        'Impressão na porta da impressora
        If frmTela_Venda.intTipo_imp_orcamento = 1 Then
            'Cabeçalho
            strLinha_Impressao = "-----------------------------------------------------------"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, 'salta uma linha
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Empresa
            strLinha_Impressao = rstImprime!DFRazao_Social_TBEmpresa
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'N ° Orçamento
            If frmTela_Venda.booIntegracao_Retaguarda = True Then
               strLinha_Impressao = "LST.COMPRA-AFINIDADE: " & rstImprime!NUM_CUPOM & ""
            Else
               strLinha_Impressao = "LST.COMPRA-AFINIDADE: " & rstImprime!NUM_CUPOM & ""
            End If
            
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                    
            'CLIENTE
            strLinha_Impressao = "-----------------------------------------------------------"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                        
            strLinha_Impressao = "CLIENTE..: " & rstAfinidade!IXCodigo_TBCliente & " - " & rstAfinidade!DFNome_TBCliente
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            strLinha_Impressao = "ENDERECO.: " & rstAfinidade!DFEndereco_TBCliente & "," & rstAfinidade!DFNumero_TBCliente
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            strLinha_Impressao = "DIA VENC.: (" & rstAfinidade!DFDia_vencimento_TBCliente & ")  - SALDO RESTANTE.: " & rstAfinidade!DFLimite_credito_TBCliente
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                    
            'Cabeçalho
            strLinha_Impressao = "-----------------------------------------------------------"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Rodapé Total
            strLinha_Impressao = "TOTAL.: " & Format(Me.dblPago_Afinidade, "#,###0.00")
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Cabeçalho
            strLinha_Impressao = "-----------------------------------------------------------"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Print #1,
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Reconhecimento
            strLinha_Impressao = "     RECONHECO E PAGAREI A DIVIDA AQUI REPRESENTADA"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Print #1,
            Print #1,
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Assinatura
            strLinha_Impressao = "          ----------------------------------                "
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                                            
            'Nome do Cliente
            strLinha_Impressao = "          " & rstAfinidade!DFNome_TBCliente
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Print #1,
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
                    
            'Cabeçalho
            strLinha_Impressao = "-----------------------------------------------------------"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Print #1,
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Rodapé - Mensagem
            strLinha_Impressao = "        OBRIGADO PELA PREFERENCIA.VOLTE SEMPRE!"
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
            
            'Rodapé - Mensagem 2
            strLinha_Impressao = "  --- ESTE DOCUMENTO NAO POSSUI VALIDADE FISCAL  --- "
            'Mandando o comando direto na porta da impressora
            Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (frmTela_Venda.strCaminho_impComum)
            Print #1, 'salta uma linha
            Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'Imprime caracter Comprimido para matriciais Epson
            
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
        End If
        
        Set rstAfinidade = Nothing
        
    End If

    Set rstImprime = Nothing
    
    Exit Function
    
Erro:
    If Err.Number = -2147206461 Then
       MsgBox "Arquivo do relatório não encontrado, verifique! A APLICAÇÃO SERÁ REINICIADA.", vbCritical, "Only Tech"
       End
    End If
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando os itens do cupom.Verifique"
    
    MsgBox "Verifique, pois todas as gravações forma concluídas com sucesso,Reimprime e cancele este cupom de n° - " & strNumero_Nota & ".Verifique!", vbCritical, "Only Tech"
    
    Exit Function
    
End Function

Function Fecha_comanda()

    On Error GoTo Erro
    'Fechando a comnada
    strSql = Empty
    strSql = "UPDATE TBComanda SET DFNumero_cupom_TBComanda = " & strNumero_Nota & " " & _
             "WHERE PKCodigo_TBComanda = " & frmTela_Venda.strNumero_Comanda & ""
             
    CNconexao.CNconexao.Execute strSql

Exit Function
    
Erro:

    CNconexao.CNconexao.RollbackTrans
    
    Call Controle_Transacional_manual
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando os itens do cupom.Verifique"
    
    Exit Function
    
End Function

Private Function Controle_Transacional_manual()

    On Error GoTo Erro
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
        strSql = Empty
        strSql = "DELETE FROM TBNOTA_SAIDA WHERE PKId_TBNota_saida = " & lngID_Numero_Nota & ""
        
        CNconexao.CNconexao.Execute strSql
        
        strSql = Empty
        strSql = "DELETE FROM TBTITULO_RECEBER WHERE PKId_TBTitulo_receber = " & lngID_Titulo_receber & ""
        
        CNconexao.CNconexao.Execute strSql
        
    End If
    
    strSql = Empty
    strSql = "DELETE FROM TBCupom WHERE PKId_TBCupom = " & lngID_Cupom & ""
    
    CNconexao_local_pdv.CNconexao.Execute strSql
    
    Exit Function
    
Erro:

    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       CNconexao.CNconexao.RollbackTrans
    End If
    
    CNconexao_local_pdv.CNconexao.RollbackTrans
    
    MsgBox Err.Number & "-" & Err.Description & "-" & " Transacional Manual.Verifique"
    
    Exit Function
    
End Function
Private Function Grava_Concentrador()
    'Abrindo uma conexão nova com o concentrador
    CNconexao_concentrador.Data_Source = frmTela_Venda.intIP_Concentrador
    CNconexao_concentrador.Banco = "BDPDV"
    CNconexao_concentrador.Abrir_conexao "PDV"
End Function
Private Function Imprime_Cupom_fiscal()

    Dim rstImprime As New ADODB.Recordset
    Dim strCodigo_Produto As String
    Dim strDescricao_Produto As String * 29
    Dim strAliquota As String
    Dim strTipo_quantidade As String * 1
    Dim strQuantiade As String * 7
    Dim strCasas_Decimais As String * 1
    Dim strValor_Unitario As String
    Dim strValor_Unitario_imp As String
    Dim strTipo_desconto As String * 1
    Dim strValor_desconto As String * 8
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
    
        'Inf. Nota de Saída
        strSql = Empty
        strSql = "SELECT TBITENS_NOTA_SAIDA.DFValor_total_praticado_TBItens_nota_saida AS TOTAL_ITEM,TBITENS_NOTA_SAIDA.DFPreco_praticado_TBItens_nota_saida AS PRECO_PRATICADO,TBITENS_NOTA_SAIDA.DFQuantidade_TBItens_nota_saida AS QUANTIDADE,TBNOTA_SAIDA.DFNumero_TBNota_saida AS NUM_CUPOM,TBNOTA_SAIDA.DFSerie_TBNota_saida,TBNOTA_SAIDA.DFData_Saida_TBNota_saida," & _
                 "TBEmpresa.DFRazao_Social_TBEmpresa,TBEmpresa.DFEndereco_TBEmpresa,TBEmpresa.DFBairro_TBEmpresa,TBEmpresa.DFCep_TBEmpresa," & _
                 "TBEmpresa.FKId_TBCidade_otica,TBEmpresa.DFCnpj_TBEmpresa,TBEmpresa.DFInscricao_estadual_TBEmpresa," & _
                 "TBEmpresa.DFTelefone_TBEmpresa , TBEmpresa.DFFax_TBEmpresa,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBCidade_otica.DFNome_TBCidade_otica,TBCidade_otica.DFPais_TBCidade_otica,TBNota_saida.DFTotal_nota_TBNota_saida " & _
                 "FROM TBITENS_NOTA_SAIDA,TBNota_saida,TBEmpresa,TBCidade_otica,TBProduto " & _
                 "WHERE FKId_TBNota_saida = " & lngID_Numero_Nota & " " & _
                 "AND TBITENS_NOTA_SAIDA.FKId_TBNota_saida = TBNota_saida.PKId_TBNota_saida " & _
                 "AND TBNOTA_SAIDA.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                 "AND TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                 "AND TBITENS_NOTA_SAIDA.FKId_TBProduto = TBProduto.PKId_TBProduto "
        
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstImprime, "Otica", Me
    Else
       'Inf. Cupons
       strSql = Empty
       strSql = "SELECT TBCUPOM.DFNumero_TBCupom AS NUM_CUPOM,TBCUPOM.DFSerie_TBCupom,TBCUPOM.DFTotal_cupom_TBCupom,TBCUPOM.DFData_Saida_TBCupom," & _
                "TBCUPOM.DFHora_Saida,TBCUPOM.DFOperador_TBCupom,TBITENS_CUPOM.DFCodigo_TBProduto,TBITENS_CUPOM.DFQuantidade_TBItens_cupom AS QUANTIDADE," & _
                "TBEmpresa.DFRazao_Social_TBEmpresa,TBEmpresa.DFEndereco_TBEmpresa,TBEmpresa.DFBairro_TBEmpresa,TBEmpresa.DFCep_TBEmpresa," & _
                "TBEmpresa.FKId_TBCidade_otica,TBEmpresa.DFCnpj_TBEmpresa,TBEmpresa.DFInscricao_estadual_TBEmpresa," & _
                "TBEmpresa.DFTelefone_TBEmpresa , TBEmpresa.DFFax_TBEmpresa,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBCidade_otica.DFNome_TBCidade_otica,TBCidade_otica.DFPais_TBCidade_otica," & _
                "TBITENS_CUPOM.DFPreco_praticado_TBItens_cupom AS PRECO_PRATICADO,TBITENS_CUPOM.DFValor_total_item_TBItens_cupom AS TOTAL_ITEM,TBITENS_CUPOM.DFUnidade_TBItens_cupom " & _
                "FROM TBCUPOM,TBEmpresa,TBCidade_otica,TBProduto " & _
                "WHERE TBCUPOM.PKId_TBCupom = TBITENS_CUPOM.FKID_TBCupom " & _
                "AND TBCUPOM.PKId_TBCupom = " & lngID_Cupom & " " & _
                "AND TBCUPOM.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                "AND TBEmpresa.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
                "AND TBITENS_CUPOM.FKId_TBProduto = TBProduto.PKId_TBProduto "
       'aqui  --   time out
       Movimentacoes.Select_geral strSql, "BDPDV", rstImprime, "PDV", Me
    End If
    
    Comandos_impressoras_fiscais.Abre_Cupom (frmTela_Venda.strImpresora)
    
    Do While rstImprime.EOF = False And rstImprime.BOF = False
    
       strCodigo_Produto = rstImprime!IXCodigo_TBProduto
       strDescricao_Produto = rstImprime!DFDescricao_resumida_TBProduto
       strQuantidade = rstImprime!Quantidade
       strPreco_Unitario = Format(rstImprime!PRECO_PRATICADO, "#,###0.00")
       strPreco_Total_Item = Format(rstImprime!TOTAL_ITEM, "#,###0.00")
       strValor_desconto = "0,00"
       
       '--- Aliquotas --------------------------------------------------------------------------------------
       Dim rstAliqota As New ADODB.Recordset
        
       'Query para localizar a aliquota do item dentro da UF da empresa
       strSql = "SELECT DFPercentual_icms_saida_fisica_TBEstado_icms,DFTributacao_impressora_fiscal_TBEstado_icms FROM TBEstado_icms " & _
                "INNER JOIN TBPRODUTO " & _
                "ON TBEstado_icms.FKId_TBProduto = TBPRODUTO.PKId_TBProduto " & _
                "WHERE DFUf_TBEstado_icms  = (SELECT MAX(DFUf_TBCidade_otica) FROM TBEMPRESA INNER JOIN TBCidade_otica ON TBEMPRESA.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica ) " & _
                "AND TBPRODUTO.IXCodigo_TBProduto = " & rstImprime!IXCodigo_TBProduto & " "
                 
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAliqota, "Otica", Me
        
       If rstAliqota.BOF = True And rstAliqota.EOF = True Then
          Set rstAliqota = Nothing
          'Query para localizar a aliquota do item dentro da UF "**"
           strSql = "SELECT DFPercentual_icms_saida_fisica_TBEstado_icms,DFTributacao_impressora_fiscal_TBEstado_icms FROM TBEstado_icms " & _
                    "INNER JOIN TBPRODUTO " & _
                    "ON TBEstado_icms.FKId_TBProduto = TBPRODUTO.PKId_TBProduto " & _
                    "WHERE DFUf_TBEstado_icms  = '**' " & _
                    "AND TBPRODUTO.IXCodigo_TBProduto = " & rstImprime!IXCodigo_TBProduto & " "
                 
           Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAliqota, "Otica", Me
       End If
       
       strAliquota = Trim(rstAliqota!DFTributacao_impressora_fiscal_TBEstado_icms)
       
       strTipo_quantidade = frmTela_Venda.strTipo_quantidade
       
       If strTipo_quantidade = "F" Then
          strQuantiade = Format(strQuantidade, "#,###0.00")
       Else
          strQuantiade = strQuantidade
       End If
     
       strCasas_Decimais = frmTela_Venda.strCasas_Decimais
       strTipo_desconto = frmTela_Venda.strTipo_desconto
       
       '------------------------------------------------------------------------------------------------------
       'ECF

        If frmTela_Venda.strImpresora = "Bematech" Then
           Retorno = Bematech_FI_VendeItem(strCodigo_Produto, strDescricao_Produto, strAliquota, strTipo_quantidade, strQuantiade, 2, strPreco_Unitario, strTipo_desconto, strValor_desconto)

           'Função que analisa o retorno da impressora
           Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")

           'Verifica retorno da impressora e interrompe a venda
           If booInterrompe_venda = True Then
              MsgBox "ERRO ao imprimir os item do cupom " & strCodigo_Produto & ".Verifique!", vbCritical, "Onlytech"
              Exit Function
           End If
        End If
        
       Set rstAliqota = Nothing
       rstImprime.MoveNext
       
    Loop
    
    Set rstImprime = Nothing
    
    'Fechando o cupom
    If frmTela_Venda.strImpresora = "Bematech" Then
'       Retorno = Bematech_FI_FechaCupomResumido(strDescr_Finalizadora, "Obrigado e volte sempre!")
       'Função que analisa o retorno da impressora
       'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
       Call Comandos_impressoras_fiscais.Fecha_Cupom(frmTela_Venda.strImpresora, strDescr_Finalizadora, "Obrigado e volte sempre!", Me.txtTotal_Cupom.Text, strCod_Finalizadora, "", Me.txtValor_pago.Text)
       If Retorno = 1 Then
          'Fechando as conexões
          If frmTela_Venda.booIntegracao_Retaguarda = True Then
             CNconexao.CNconexao.Close
          End If
          CNconexao_local_pdv.Fechar_conexao
          Unload Me
       End If
     End If
  
    Exit Function
    
Erro:
    
    MsgBox Err.Number & "-" & Err.Description & "-" & "Gravando os itens do cupom.Verifique"
    
    MsgBox "Verifique, pois todas as gravações forma concluídas com sucesso,Reimprime e cancele este cupom de n° - " & strNumero_Nota & ".Verifique!", vbCritical, "Only Tech"
    
    Exit Function
    
End Function
Private Function Gravar_Impostos_Nota(ID_Nota)
        
    Dim rstImpostos As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT " & _
             "DFPercentual_icms_TBItens_nota_saida," & _
             "SUM(DFValor_total_praticado_TBItens_nota_saida) AS BASE_CALCULO," & _
             "SUM(DFValor_total_icms_TBItens_nota_saida) As TOTAL_ICMS " & _
             "FROM TBItens_nota_saida " & _
             "WHERE FKId_TBNota_saida = " & ID_Nota & " " & _
             "GROUP BY DFPercentual_icms_TBItens_nota_saida"
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstImpostos, "Otica", Me
    
    rstImpostos.MoveFirst
    
    'Reiniciando + 1 transação
    CNconexao.CNconexao.BeginTrans
    
    Do While Not rstImpostos.EOF
          'Gravando a tabela de Impostos_nota
          strSql = Empty
          strSql = "INSERT INTO TBImpostos_nota_saida (" & _
                   "FKId_TBNota_saida,DFAliquota_TBImpostos_nota_saida,DFBase_calculo_TBImpostos_nota_saida,DFValor_TBImpostos_nota_saida )" & _
                   "VALUES(" & _
                   "" & ID_Nota & "," & Funcoes_Gerais.Grava_Moeda(rstImpostos!DFPercentual_icms_TBItens_nota_saida) & "," & Funcoes_Gerais.Grava_Moeda(rstImpostos!BASE_CALCULO) & "," & Funcoes_Gerais.Grava_Moeda(rstImpostos!TOTAL_ICMS) & ") "
                   
        CNconexao.CNconexao.Execute strSql
          
        rstImpostos.MoveNext
    Loop
    
    'Comitando a gravação dos  registros na tabela de titulos_nota
    CNconexao.CNconexao.CommitTrans
    
    Set rstImpostos = Nothing

End Function
Private Function Gravar_Impostos_Cupom(ID_Nota)
        
    Dim rstImpostos As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT " & _
             "DFPercentual_icms_TBItens_cupom," & _
             "SUM(DFValor_total_praticado_TBItens_cupom) AS BASE_CALCULO," & _
             "SUM(DFValor_total_icms_TBItens_cupom) As TOTAL_ICMS " & _
             "FROM TBItens_cupom " & _
             "WHERE FKId_TBCupom = " & ID_Nota & " " & _
             "GROUP BY DFPercentual_icms_TBItens_cupom"
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstImpostos, "Otica", Me
    
    rstImpostos.MoveFirst
       
    'Reiniciando + 1 transação
    CNconexao.CNconexao.BeginTrans
    
    Do While Not rstImpostos.EOF
          'Gravando a tabela de Impostos_nota
          strSql = Empty
          strSql = "INSERT INTO TBImpostos_cupom (" & _
                   "FKId_TBNota_saida,DFAliquota_TBImpostos_cupom,DFBase_calculo_TBImpostos_cupom,DFValor_TBImpostos_cupom )" & _
                   "VALUES(" & _
                   "" & ID_Nota & "," & Funcoes_Gerais.Grava_Moeda(rstImpostos!DFPercentual_icms_TBItens_cupom) & "," & Funcoes_Gerais.Grava_Moeda(rstImpostos!BASE_CALCULO) & "," & Funcoes_Gerais.Grava_Moeda(rstImpostos!TOTAL_ICMS) & ") "
                   
        CNconexao.CNconexao.Execute strSql
          
        rstImpostos.MoveNext
    Loop
    
    'Comitando a gravação dos  registros na tabela de titulos_nota
    CNconexao.CNconexao.CommitTrans
    
    Set rstImpostos = Nothing

End Function
