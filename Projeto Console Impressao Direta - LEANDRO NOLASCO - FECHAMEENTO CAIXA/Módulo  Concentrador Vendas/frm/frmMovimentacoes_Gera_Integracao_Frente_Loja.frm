VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovimentacoes_Gera_Integracao_Frente_Loja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gera Integração Frente Loja - Exportação"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":1782
   ScaleHeight     =   4470
   ScaleWidth      =   8205
   Begin AutoCompletar.CbCompleta cbbEmpresa_Recebimento 
      Height          =   360
      Left            =   5970
      TabIndex        =   0
      Top             =   600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCaminho 
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   1290
      Width           =   2835
   End
   Begin VB.CommandButton cmdCaminho 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3000
      Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":1AC4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Localize o caminho onde o arquivo será salvo"
      Top             =   1300
      Width           =   405
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   10350
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":1E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":2168
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":2482
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":2BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja.frx":2ED0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1(0)"
      HotImageList    =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair - CTRL+S"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcCodigo_empresa 
      Height          =   360
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      BackColor       =   -2147483639
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgOpcoes_Exportacao 
      Height          =   2655
      Left            =   90
      TabIndex        =   3
      Top             =   1740
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin AutoCompletar.CbCompleta cbbFabricante_Ecf 
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   1290
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicial 
      Height          =   360
      Left            =   4980
      TabIndex        =   11
      Top             =   1290
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   8388608
      Format          =   49741825
      CurrentDate     =   37881
   End
   Begin MSComCtl2.DTPicker dtpFinal 
      Height          =   360
      Left            =   6750
      TabIndex        =   12
      Top             =   1290
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   8388608
      Format          =   49741825
      CurrentDate     =   37881
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   240
      Left            =   6420
      TabIndex        =   14
      Top             =   1410
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   240
      Left            =   5010
      TabIndex        =   13
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante Ecf"
      Height          =   240
      Left            =   3480
      TabIndex        =   10
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa Recebimento"
      Height          =   240
      Left            =   5970
      TabIndex        =   9
      Top             =   360
      Width           =   1920
   End
   Begin VB.Label lblCaminho 
      AutoSize        =   -1  'True
      Caption         =   "Caminho do Arquivo"
      Height          =   240
      Left            =   90
      TabIndex        =   8
      Top             =   1050
      Width           =   1725
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "frmMovimentacoes_Gera_Integracao_Frente_Loja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Movimentação Gera Integração Frente de Loja                    '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rafael de Oliveira Gomes                                       '
' Data de Criação........: 19/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strsql As String
Dim strDestino As String
Dim intConta_Sequencial As Integer
Dim I As Integer
Dim NumArq As String
Dim NumArq2 As String
Dim strIntegracao As String
Dim log As New DLLSystemManager.log
Dim rstBusca_Cliente As New ADODB.Recordset
Dim rstBusca_Finalizadora As New ADODB.Recordset
Dim rstBusca_Abrir_Gaveta As New ADODB.Recordset
Dim rstBusca_Operador_Ecf As New ADODB.Recordset
Dim rstBusca_Produto As New ADODB.Recordset
Dim rstBusca_Tabela_Vigente As New ADODB.Recordset
Dim rstBusca_Codigo_Barras As New ADODB.Recordset
Dim rstBusca_Composicao As New ADODB.Recordset
'RETIRADO TEMPORARIAMENTE PARA DEFINIÇÃO DE COMO LOGICA SERA IMPLEMENTADA FUTURAMENTE''''''''''''''''''''
'Dim rstBusca_Promocoes As New ADODB.Recordset                                                          '
'Dim rstBusca_Familia As New ADODB.Recordset                                                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rstBusca_Secao As New ADODB.Recordset
Dim rstBusca_Estado_Icms As New ADODB.Recordset
Option Explicit

Private Sub cmdCaminho_Click()
    Unload frmMovimentacoes_Gera_Integracao_Frente_Loja_Caminho
    frmAguarde.Show
    DoEvents
    frmMovimentacoes_Gera_Integracao_Frente_Loja_Caminho.Show
    Unload frmAguarde
End Sub

Private Sub dtcCodigo_empresa_LostFocus()
    dtcCodigo_empresa.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'HABILITA A TROCA DE CAMPOS PELO ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'TECLAS DE ATALHO DA TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: Call Gravar   'CTRL+G
                       Case 67: Call Cancelar 'CTRL+C
                       Case 83: Unload Me     'CTRL+S
                End Select
    End Select

    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcCodigo_empresa, "Otica", "BDRetaguarda", Me
End Sub

Private Sub hfgOpcoes_Exportacao_Click()
    'VERIFICANDO SE O USUARIO CLICOU EM LINHA NÃO PERMITIDA
    If hfgOpcoes_Exportacao.Row = 0 Then Exit Sub
    
    'MARCAÇÃO DE SIM / NÃO - CONFORME O CLICK DO USUARIO - BUSCA DE INFORMAÇÕES SOBRE O NUMERO DE REGISTROS A GERAR
    If hfgOpcoes_Exportacao.Col = 2 Then
       If hfgOpcoes_Exportacao.Text = "X" Then
          hfgOpcoes_Exportacao.Text = Empty
          hfgOpcoes_Exportacao.Col = 3
          hfgOpcoes_Exportacao.Text = Empty
          hfgOpcoes_Exportacao.Col = 2
       Else
          hfgOpcoes_Exportacao.CellFontBold = True
          hfgOpcoes_Exportacao.CellForeColor = &HC00000
          hfgOpcoes_Exportacao.Text = "X"
          
          'BUSCANDO CLIENTES A GERAR
          If hfgOpcoes_Exportacao.Row = 1 Then
             Dim rstCliente As New ADODB.Recordset
             
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT PKId_TBCliente FROM TBCliente " & _
                      "WHERE IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                      "AND DFData_cadastro_TBCliente >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                      "AND DFData_cadastro_TBCliente <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "
                      
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstCliente, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstCliente.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
             
             Unload frmAguarde
          'BUSCANDO FINALIZADORAS A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 2 Then
             Dim rstFinalizadora As New ADODB.Recordset
             
             frmAguarde.Show
             DoEvents
                        
             strsql = Empty
             strsql = "SELECT IXCodigo_TBFinalizadora FROM TBFinalizadora"
          
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstFinalizadora, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstFinalizadora.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
             
             Unload frmAguarde
          'BUSCANDO OPERADORES ECF A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 3 Then
             Dim rstOperadores_Ecf As New ADODB.Recordset
                        
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT PKCodigo_TBOperadores_ecf FROM TBOperadores_Ecf"
          
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstOperadores_Ecf, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstOperadores_Ecf.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
             
             Unload frmAguarde
          'BUSCANDO PRODUTOS A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 4 Then
             Dim rstProduto As New ADODB.Recordset
             Dim rstTabela_Vigente As New ADODB.Recordset
             
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
                      "FROM TBParametros_venda " & _
                      "WHERE IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "
                        
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstTabela_Vigente, "Otica", Me
        
             If rstTabela_Vigente.RecordCount <> 0 Then
                If Not IsNull(rstTabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda) Then
                   strsql = Empty
                   strsql = "SELECT PKId_TBProduto " & _
                            "FROM TBProduto, TBItens_tabela_preco, TBEstado_icms " & _
                            "WHERE TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                            "AND TBEstado_icms.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                            "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                            "AND FKCodigo_TBTabela_preco = '" & rstTabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda & "' " & _
                            "AND DFData_cadastro_TBProduto >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                            "AND DFData_cadastro_TBProduto <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "
                Else
                   strsql = Empty
                   strsql = "SELECT PKId_TBProduto " & _
                            "FROM TBProduto,TBEstado_icms " & _
                            "WHERE TBEstado_icms.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                            "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                            "AND DFData_cadastro_TBProduto >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                            "AND DFData_cadastro_TBProduto <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "
                End If
             End If

             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstProduto, "Otica", Me

             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstProduto.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2

             Unload frmAguarde
          'BUSCANDO CODIGO DE BARRAS A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 5 Then
             Dim rstCodigo_Barras As New ADODB.Recordset
             Dim rstTabela_Vigente2 As New ADODB.Recordset
             
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
                      "FROM TBParametros_venda " & _
                      "WHERE IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "
                        
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstTabela_Vigente2, "Otica", Me
             
             If rstTabela_Vigente2.RecordCount <> 0 Then
                If Not IsNull(rstTabela_Vigente2!DFNumero_tabela_vigente_TBParametros_venda) Then
                   strsql = Empty
                   strsql = "SELECT PKId_TBCodigo_barras " & _
                            "FROM TBProduto, TBCodigo_barras, TBItens_tabela_preco " & _
                            "WHERE TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                            "AND TBCodigo_barras.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                            "AND FKCodigo_TBTabela_preco = '" & rstTabela_Vigente2!DFNumero_tabela_vigente_TBParametros_venda & "' "
                Else
                   strsql = Empty
                   strsql = "SELECT PKId_TBCodigo_barras " & _
                            "FROM TBProduto, TBCodigo_barras " & _
                            "WHERE TBCodigo_barras.FKId_TBProduto = TBProduto.PKId_TBProduto "
                End If
             End If
             
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstCodigo_Barras, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstCodigo_Barras.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
             
             Unload frmAguarde
          'BUSCANDO COMPOSIÇÃO A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 6 Then
             Dim rstComposicao_produto As New ADODB.Recordset
                        
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT PKId_TBComposicao_produto " & _
                      "FROM TBComposicao_produto, TBProduto " & _
                      "WHERE TBProduto.PKId_TBProduto = TBComposicao_produto.FKId_TBProduto "
                      
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstComposicao_produto, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstComposicao_produto.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
          
             Unload frmAguarde
          'BUSCANDO PRODUTO PROMOÇÃO A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 7 Then
          '''''''''''''''''''''''''''''''''''''''''''''''''''''
          '  Dim rstPromocao_produto As New ADODB.Recordset   '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''
             frmAguarde.Show
             DoEvents

          '''''''''''''''''''''''''''''''''''''''''''''''''''''
          '   strsql = Empty                                  '
          '   strsql = "SELECT * FROM TBPromocao_produto"     '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''

          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '   Movimentacoes.Select_geral strsql, "BDRetaguarda", rstPromocao_produto, "Otica", Me   '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = 0 & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2

             Unload frmAguarde
          'BUSCANDO CORRELATO A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 8 Then
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = "1" & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
          'BUSCANDO SEÇÃO A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 9 Then
             Dim rstSecao As New ADODB.Recordset
                        
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT PKCodigo_TBSecao FROM TBSecao"
          
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstSecao, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstSecao.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
          
             Unload frmAguarde
          'BUSCANDO ESTADOS ICMS A GERAR
          ElseIf hfgOpcoes_Exportacao.Row = 10 Then
             Dim rstEstado_Icms As New ADODB.Recordset
                        
             frmAguarde.Show
             DoEvents
             
             strsql = Empty
             strsql = "SELECT PKId_TBEstado_icms " & _
                      "FROM TBEstado_Icms , TBProduto " & _
                      "WHERE TBProduto.PKId_TBProduto = TBEstado_Icms.FKId_TBProduto " & _
                      "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "
                      
             Movimentacoes.Select_geral strsql, "BDRetaguarda", rstEstado_Icms, "Otica", Me
          
             hfgOpcoes_Exportacao.Col = 3
             hfgOpcoes_Exportacao.Text = rstEstado_Icms.RecordCount & " registro(s) à gerar"
             hfgOpcoes_Exportacao.Col = 2
          
             Unload frmAguarde
          End If
          
          Set rstCliente = Nothing
          Set rstFinalizadora = Nothing
          Set rstTabela_Vigente = Nothing
          Set rstOperadores_Ecf = Nothing
          Set rstProduto = Nothing
          Set rstCodigo_Barras = Nothing
          Set rstComposicao_produto = Nothing
       ''''''''''''''''''''''''''''''''''''''''''
       '   Set rstPromocao_produto = Nothing    '
       '   Set rstFamilia = Nothing             '
       ''''''''''''''''''''''''''''''''''''''''''
          Set rstSecao = Nothing
          Set rstEstado_Icms = Nothing
       End If
    End If
End Sub

Private Sub hfgOpcoes_Exportacao_KeyPress(KeyAscii As Integer)
    'MARCANDO GRID COM ESPAÇO
    If KeyAscii = 32 Then
       Call hfgOpcoes_Exportacao_Click
    End If
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Gravar
           Case 2: Call Cancelar
           Case 4: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    
    'INFORMAÇÕES CONSTANTES PARA O LOG
    log.Usuario = MDIPrincipal.ocxUsuario.Nome
    log.Programa = "Movimentacao Gera Integração Frente de Loja"
    log.Estacao = MDIPrincipal.ocxUsuario.Estacao
    
    'INFORMAÇÕES VARIAVEIS PARA O LOG
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.ocxUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando a Movimentacao Gera Integração Frente de Loja"
    'GRAVANDO O LOG
    log.Gravar_log "Otica", Me
    
    'MONTANDO DATA COMBO DA EMPRESA
    strsql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcCodigo_empresa, strsql, "BDRetaguarda", "Otica", Me
    
    dtcCodigo_empresa.BoundText = MDIPrincipal.ocxUsuario.Empresa

    dtpInicial.Value = Date
    dtpFinal.Value = Date

    cbbEmpresa_Recebimento.Clear
    cbbEmpresa_Recebimento.AddItem ("Only Tech")
    cbbEmpresa_Recebimento.AddItem ("Fantastsoft")
    
    cbbFabricante_Ecf.Clear
    cbbFabricante_Ecf.AddItem ("Afrac")
    cbbFabricante_Ecf.AddItem ("Bematech")
    cbbFabricante_Ecf.AddItem ("Daruma")
    cbbFabricante_Ecf.AddItem ("Sweda")
    cbbFabricante_Ecf.AddItem ("Yanco")
              
    Call Monta_Opcoes_Exportacao
        
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Descricao = "Finalizando a Movimentação Gera Integração Frente de Loja"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'GRAVANDO O LOG
    log.Gravar_log "Otica", Me
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Function Gravar()
    On Error GoTo Erro
    
    'VERIFICANDO SE O CAMINHO DO ARQUIVO FOI CARREGADO
    If txtCaminho.Text = Empty Then
       MsgBox "Caminho para geração do arquivo não informado. Verifique!", vbInformation, "Only Tech"
       txtCaminho.SetFocus
       Exit Function
    End If
    
    'VERIFICANDO SE A EMPRESA DE RECEBIMENTO DO ARQUIVO FOI CARREGADA
    If cbbEmpresa_Recebimento.Text = Empty Then
       MsgBox "Empresa de recebimento para geração do arquivo não informado. Verifique!", vbInformation, "Only Tech"
       cbbEmpresa_Recebimento.SetFocus
       Exit Function
    End If
    
    'VERIFICANDO SE O FABRICANTE ECF DO ARQUIVO FOI CARREGADO
    If cbbFabricante_Ecf.Text = Empty Then
       MsgBox "Fabricante Ecf para geração do arquivo não informado. Verifique!", vbInformation, "Only Tech"
       cbbFabricante_Ecf.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                                                                                                   '
    '      SELECIONANDO ARQUIVOS GERADOS ANTERIORMENTE E MOVENDO PARA PASTA DE ARQUIVOS GERADOS         '
    '                                                                                                   '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim strExtensao As String
    Dim strArquivos As String
    Dim strOrigem As String
    Dim strDestino As String
    Dim strNome_Arquivo As String
    Dim strData As String
    Dim strHora As String
    Dim strCaminho_Destino As String
    
    strArquivos = Dir$(txtCaminho.Text & "\*.*")
    
    While strArquivos <> Empty
        strNome_Arquivo = strArquivos
    
        strOrigem = txtCaminho.Text & "\" & strNome_Arquivo & ""
        
        strData = Format(Now, "YYYYMMDD")
        strHora = Format(Now, "HHMMSS")
        
        strCaminho_Destino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
        
        strCaminho_Destino = Left(strCaminho_Destino, CDbl(Len(strCaminho_Destino) - 4)) & "\INTEGRAÇÃO\GERADOS"
        
        'CAPTURANDO O ARQUIVO GERADO NO DIA ANTERIOR, MUDANDO CAMINHO DA PASTA DESTINO E RENOMEANDO TXT
        strDestino = strCaminho_Destino & "\" & Left(strNome_Arquivo, CDbl(Len(strNome_Arquivo) - 4)) & "_" & strData & "_" & strHora & ".DAT"
    
        'MOVENDO OS ARQUIVOS GERADOS PARA PASTA CORRESPONDENTE
        FileCopy strOrigem, strDestino
        
        Kill (strOrigem)

        strArquivos = Dir$(txtCaminho.Text & "\*.*")
    Wend

    'BUSCANDO REGISTROS PARA INTEGRAÇÃO DO RETAGUARDA COM  FRENTE DE LOJA
    If cbbEmpresa_Recebimento.Text = "Only Tech" Then
        MsgBox "Sistema não preparado para gerar integração Only Tech Retaguarda - Only Tech Frente de Loja. Verifique!", vbInformation, "Only Tech"
        
        Exit Function
    'EMPRESA DE RECEBIMENTO DO ARQUIVO DE INTEGRAÇÃO DO FRENTE DE LOJA
    ElseIf cbbEmpresa_Recebimento.Text = "Fantastsoft" Then
       'BUSCANDO INFORMAÇÕES DO CLIENTE
       strsql = Empty
       strsql = "SELECT IXCodigo_TBCliente," & _
                "DFCpf_TBCliente," & _
                "DFRegistro_geral_TBCliente," & _
                "DFInscricao_estadual_TBCliente," & _
                "DFNome_TBCliente," & _
                "DFEndereco_TBCliente," & _
                "DFBairro_TBCliente," & _
                "DFNome_TBCidade_otica," & _
                "DFCep_TBCliente," & _
                "DFUf_TBCidade_otica," & _
                "DFLimite_credito_TBCliente," & _
                "DFDia_vencimento_TBCliente," & _
                "DFTolerancia_TBCliente," & _
                "DFNumero_contrato_TBCliente," & _
                "DFTipo_pessoa_TBCliente " & _
                "FROM TBCliente, TBCidade_otica " & _
                "WHERE TBCidade_otica.PKId_TBCidade_otica = TBCliente.FKId_TBCidade_otica " & _
                "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                "AND DFData_cadastro_TBCliente >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND DFData_cadastro_TBCliente <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "

       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Cliente, "Otica", Me
       
       'BUSCANDO INFORMAÇÕES DA FINALIZADORA
       strsql = Empty
       strsql = "SELECT IXCodigo_TBFinalizadora," & _
                "DFDescricao_TBFinalizadora," & _
                "DFCodigo_asc_TBFinalizadora," & _
                "DFTroco_TBFinalizadora " & _
                "FROM TBFinalizadora "

       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Finalizadora, "Otica", Me

       strsql = Empty
       strsql = "SELECT DFGaveta_integrada_TBPdv " & _
                "FROM TBPdv " & _
                "WHERE IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "

       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Abrir_Gaveta, "Otica", Me

       'BUSCANDO INFORMAÇÕES DA OPERADOR ECF
       strsql = Empty
       strsql = "SELECT PKCodigo_TBOperadores_ecf," & _
                "DFNome_TBOperadores_ecf," & _
                "DFSenha_TBOperadores_ecf," & _
                "DFNivel_TBOperadores_ecf " & _
                "FROM TBOperadores_ecf " & _
                "WHERE FKCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "

       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Operador_Ecf, "Otica", Me

       'BUSCANDO INFORMAÇÕES DO PRODUTO
       strsql = Empty
       strsql = "SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
                "FROM TBParametros_venda " & _
                "WHERE IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "
                
       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Tabela_Vigente, "Otica", Me

       If rstBusca_Tabela_Vigente.RecordCount <> 0 Then
          If Not IsNull(rstBusca_Tabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda) Then
             strsql = Empty
             strsql = "SELECT IXCodigo_TBProduto," & _
                      "DFDescricao_TBProduto," & _
                      "DFDescricao_resumida_TBProduto," & _
                      "DFPeso_variavel_TBProduto," & _
                      "DFPreco_avista_TBItens_tabela_preco," & _
                      "DFTributacao_impressora_fiscal_TBEstado_icms," & _
                      "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                      "FKCodigo_TBSecao " & _
                      "FROM TBProduto, TBItens_tabela_preco, TBEstado_icms " & _
                      "WHERE TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                      "AND TBEstado_icms.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                      "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                      "AND FKCodigo_TBTabela_preco = '" & rstBusca_Tabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda & "' " & _
                      "AND DFData_cadastro_TBProduto >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                      "AND DFData_cadastro_TBProduto <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "
          Else
             strsql = Empty
             strsql = "SELECT IXCodigo_TBProduto," & _
                      "DFDescricao_TBProduto," & _
                      "DFDescricao_resumida_TBProduto," & _
                      "DFPeso_variavel_TBProduto," & _
                      "DFTributacao_impressora_fiscal_TBEstado_icms," & _
                      "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                      "FKCodigo_TBSecao " & _
                      "FROM TBProduto,TBEstado_icms " & _
                      "WHERE TBEstado_icms.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                      "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                      "AND DFData_cadastro_TBProduto >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                      "AND DFData_cadastro_TBProduto <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "
          End If
       End If

       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Produto, "Otica", Me
       
       'BUSCANDO INFORMAÇÕES DO CODIGO BARRAS PRODUTO
       If rstBusca_Tabela_Vigente.RecordCount <> 0 Then
          If Not IsNull(rstBusca_Tabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda) Then
             strsql = Empty
             strsql = "SELECT IXCodigo_TBCodigo_barras," & _
                      "IXCodigo_TBProduto," & _
                      "DFDescricao_resumida_TBProduto," & _
                      "DFPreco_avista_TBItens_tabela_preco " & _
                      "FROM TBProduto, TBCodigo_barras, TBItens_tabela_preco " & _
                      "WHERE TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                      "AND TBCodigo_barras.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                      "AND FKCodigo_TBTabela_preco = '" & rstBusca_Tabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda & "' "
          Else
             strsql = Empty
             strsql = "SELECT IXCodigo_TBCodigo_barras," & _
                      "IXCodigo_TBProduto," & _
                      "DFDescricao_resumida_TBProduto " & _
                      "FROM TBProduto, TBCodigo_barras " & _
                      "WHERE TBCodigo_barras.FKId_TBProduto = TBProduto.PKId_TBProduto "
          End If
       End If
       
       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Codigo_Barras, "Otica", Me

       'BUSCANDO INFORMAÇÕES DA COMPOSIÇÃO
       strsql = Empty
       strsql = "SELECT IXCodigo_TBProduto," & _
                "DFCodigo_produto_relacionado_TBComposicao_produto," & _
                "DFQuantidade_baixa_estoque_TBComposicao_produto " & _
                "FROM TBProduto, TBComposicao_produto " & _
                "WHERE TBComposicao_produto.FKId_TBProduto = TBProduto.PKId_TBProduto "
                
       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Composicao, "Otica", Me

''''''''RETIRADO TEMPORARIAMENTE ATÉ Q SE DECIDA COMO FAZER A QUESTÃO DE PRODUTOS EM PROMOÇÃO'''''''''''''''''''
'       BUSCANDO INFORMAÇÕES DA PRODUTO PROMOCAO                                                               '
'       strsql = Empty                                                                                         '
'       strsql = ""                                                                                            '
'                                                                                                              '
'       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Promocoes, "Otica", Me                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
''''''''RETIRADO POIS O SISTEMA NÃO TRABALHA ATE O MOMENTO COM CONCEITO DE FAMILIA DE PRODUTOS''''''''''''''''''
'       BUSCANDO INFORMAÇÕES DA FAMILIA                                                                        '
'       strsql = Empty                                                                                         '
'       strsql = ""                                                                                            '
'                                                                                                              '
'       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Familia, "Otica", Me                       '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       'BUSCANDO INFORMAÇÕES DA SECAO
       strsql = Empty
       strsql = "SELECT PKCodigo_TBSecao," & _
                "DFDescricao_TBsecao " & _
                "FROM TBSecao "

       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Secao, "Otica", Me

       'BUSCANDO INFORMAÇÕES DO ESTADO ICMS
       strsql = Empty
       strsql = "SELECT PKId_TBEstado_icms," & _
                "DFTributacao_impressora_fiscal_TBEstado_icms," & _
                "DFPercentual_icms_saida_juridica_TBEstado_icms " & _
                "FROM TBEstado_icms,TBProduto " & _
                "WHERE TBEstado_icms.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                "AND IXCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' "
                
       Movimentacoes.Select_geral strsql, "BDRetaguarda", rstBusca_Estado_Icms, "Otica", Me
       
       hfgOpcoes_Exportacao.Col = 2
       
       For I = 1 To hfgOpcoes_Exportacao.Rows - 1
          hfgOpcoes_Exportacao.Row = I
          
          If hfgOpcoes_Exportacao.Text = "X" Then
             If hfgOpcoes_Exportacao.Row = 1 Then
                Call Gera_Cliente
             ElseIf hfgOpcoes_Exportacao.Row = 2 Then
                Call Gera_Finalizadora
             ElseIf hfgOpcoes_Exportacao.Row = 3 Then
                Call Gera_Operador_Ecf
             ElseIf hfgOpcoes_Exportacao.Row = 4 Then
                Call Gera_Produto
             ElseIf hfgOpcoes_Exportacao.Row = 5 Then
                Call Gera_Codigo_Barras
             ElseIf hfgOpcoes_Exportacao.Row = 6 Then
                Call Gera_Composicao
             ElseIf hfgOpcoes_Exportacao.Row = 7 Then
                Call Gera_Promocoes
             ElseIf hfgOpcoes_Exportacao.Row = 8 Then
                Call Gera_Familia
             ElseIf hfgOpcoes_Exportacao.Row = 9 Then
                Call Gera_Secao
             ElseIf hfgOpcoes_Exportacao.Row = 10 Then
                Call Gera_Estado_Icms
             End If
          End If
       Next I
    
       Set rstBusca_Cliente = Nothing
       Set rstBusca_Finalizadora = Nothing
       Set rstBusca_Abrir_Gaveta = Nothing
       Set rstBusca_Operador_Ecf = Nothing
       Set rstBusca_Produto = Nothing
       Set rstBusca_Tabela_Vigente = Nothing
       Set rstBusca_Codigo_Barras = Nothing
       Set rstBusca_Composicao = Nothing
       
''''''''RETIRADO TEMPORARIAMENTE PARA DEFINIÇÃO DE COMO LOGICA SERA IMPLEMENTADA FUTURAMENTE''''''''''''''''''''
'       set rstBusca_Promocoes = Nothing                                                                       '
'       set rstBusca_Familia = Nothing                                                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

       Set rstBusca_Secao = Nothing
       Set rstBusca_Estado_Icms = Nothing
    End If
    
    Unload frmAguarde
    
    Call Cancelar
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Gravar")
    
    Exit Function
End Function

Private Function Gera_Cliente()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"

       Open txtCaminho.Text & "\CLIENTES.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
    
       strIntegracao = Empty
       strIntegracao = "CLIENTE    CPFCGC    INSCRICAOIDENTIDADE    NOME    ENDERECO    BAIRRO    CIDADE    CEP    ESTADO    LIMITECREDITO    DIALIMITE    CADCARENCIA    CADPERIODO    TAXAJUROS    CARTAO"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Cliente.RecordCount <> 0 Then
          rstBusca_Cliente.MoveFirst
   
          intConta_Sequencial = rstBusca_Cliente.RecordCount
       
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strCpf_Cgc As String * 14
             Dim strInscricao_Identidade As String * 20
             Dim strNome As String * 40
             Dim strEndereco As String * 40
             Dim strBairro As String * 25
             Dim strCidade As String * 25
             Dim strCep As String * 8
             Dim strEstado As String * 2
             Dim strLimite_Credito As String * 10
             Dim strDia_Limite As String * 6
             Dim strCad_Carencia As String * 6
             Dim strCad_Periodo As String * 6
             Dim strTaxa_Juros As String * 6
             Dim strCartao As String * 20
             
             If Not IsNull(rstBusca_Cliente!IXCodigo_TBCliente) Then
                strCodigo = Trim(rstBusca_Cliente!IXCodigo_TBCliente)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFCpf_TBCliente) Then
                strCpf_Cgc = Format(Trim(Replace(CStr(Replace(Replace(rstBusca_Cliente!DFCpf_TBCliente, "-", ""), "/", "")), ".", "")), "00000000000000")
             Else
                strCpf_Cgc = "Null"
             End If
          
             'VERIFICANDO O TIPO DE PESSOA PARA GERAÇÃO DO ARQUIVO
             If Not IsNull(rstBusca_Cliente!DFTipo_pessoa_TBCliente) Then
                If rstBusca_Cliente!DFTipo_pessoa_TBCliente = True Then
                   If Not IsNull(rstBusca_Cliente!DFInscricao_estadual_TBCliente) Then
                      strInscricao_Identidade = Trim(Replace(CStr(Replace(Replace(rstBusca_Cliente!DFInscricao_estadual_TBCliente, "-", ""), "/", "")), ".", ""))
                   Else
                      strInscricao_Identidade = "Null"
                   End If
                ElseIf rstBusca_Cliente!DFTipo_pessoa_TBCliente = False Then
                   If Not IsNull(rstBusca_Cliente!DFRegistro_geral_TBCliente) Then
                      strInscricao_Identidade = Trim(Replace(CStr(Replace(Replace(rstBusca_Cliente!DFRegistro_geral_TBCliente, "-", ""), "/", "")), ".", ""))
                   Else
                      strInscricao_Identidade = "Null"
                   End If
                Else
                   strInscricao_Identidade = "00000000000000"
                End If
             Else
                strInscricao_Identidade = "00000000000000"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFNome_TBCliente) Then
                strNome = Trim(rstBusca_Cliente!DFNome_TBCliente)
             Else
                strNome = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFEndereco_TBCliente) Then
                strEndereco = Trim(rstBusca_Cliente!DFEndereco_TBCliente)
             Else
                strEndereco = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFBairro_TBCliente) Then
                strBairro = Trim(rstBusca_Cliente!DFBairro_TBCliente)
             Else
                strBairro = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFNome_TBCidade_otica) Then
                strCidade = Trim(rstBusca_Cliente!DFNome_TBCidade_otica)
             Else
                strCidade = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFCep_TBCliente) Then
                strCep = Replace(Trim(rstBusca_Cliente!DFCep_TBCliente), "-", "")
             Else
                strCep = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFUf_TBCidade_otica) Then
                strEstado = Trim(rstBusca_Cliente!DFUf_TBCidade_otica)
             Else
                strEstado = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFLimite_credito_TBCliente) Then
                strLimite_Credito = Format(Trim(rstBusca_Cliente!DFLimite_credito_TBCliente), "#,###0.00")
             Else
                strLimite_Credito = "Null"
             End If
             
             If Not IsNull(rstBusca_Cliente!DFDia_vencimento_TBCliente) Then
                strDia_Limite = Trim(rstBusca_Cliente!DFDia_vencimento_TBCliente)
             Else
                strDia_Limite = "Null"
             End If
   
             If Not IsNull(rstBusca_Cliente!DFTolerancia_TBCliente) Then
                strCad_Carencia = Trim(rstBusca_Cliente!DFTolerancia_TBCliente)
             Else
                strCad_Carencia = "Null"
             End If
             
             strCad_Periodo = "0"
             strTaxa_Juros = "0"
             
             If Not IsNull(rstBusca_Cliente!DFNumero_contrato_TBCliente) Then
                strCartao = Trim(rstBusca_Cliente!DFNumero_contrato_TBCliente)
             Else
                strCartao = "Null"
             End If
   
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strCpf_Cgc) & "    " & Trim(strInscricao_Identidade) & "    " & Trim(strNome) & "    " & Trim(strEndereco) & "    " & Trim(strBairro) & "    " & Trim(strCidade) & "    " & Trim(strCep) & "    " & Trim(strEstado) & "    " & Trim(strLimite_Credito) & "    " & Trim(strDia_Limite) & "    " & Trim(strCad_Carencia) & "    " & Trim(strCad_Periodo) & "    " & Trim(strTaxa_Juros) & "    " & Trim(strCartao)
             
             Print #NumArq, strIntegracao
             
             rstBusca_Cliente.MoveNext
             
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Cliente processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\CLIENTES.DAT", strDestino & "\CLIENTES.DAT"
End Function

Private Function Gera_Finalizadora()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
          
       Open txtCaminho.Text & "\MODALIDADES.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGO    DESCRICAO    TECLA    PERMITIRUSO    PERMITIRTROCO    VALORMAXIMO    SOMENTEVALORTOTAL    ABRIRGAVETA    EMITECONTRAVALE    EMITETROCOCHEQUE    USARCONTACORRENTE    EMISSAODOC    AUTENTICACOES    TIPOTEF"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Finalizadora.RecordCount <> 0 Then
          rstBusca_Finalizadora.MoveFirst
      
          intConta_Sequencial = rstBusca_Finalizadora.RecordCount
      
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strDescricao As String * 20
             Dim strTecla As String * 6
             Dim strPermitir_Uso As String * 6
             Dim strPermitir_Troco As String * 6
             Dim strValor_Maximo As String * 10
             Dim strSomente_Valor_Total As String * 1
             Dim strAbrir_Gaveta As String * 1
             Dim strEmite_Contra_Vale As String * 1
             Dim strEmite_Troco_Cheque As String * 1
             Dim strUsar_Conta_Corrente As String * 1
             Dim strEmissao_Doc As String * 6
             Dim strAutenticacoes As String * 6
             Dim strTipo_Tef As String * 6
             
             If Not IsNull(rstBusca_Finalizadora!IXCodigo_TBFinalizadora) Then
                strCodigo = Trim(rstBusca_Finalizadora!IXCodigo_TBFinalizadora)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Finalizadora!DFDescricao_TBFinalizadora) Then
                strDescricao = Trim(rstBusca_Finalizadora!DFDescricao_TBFinalizadora)
             Else
                 strDescricao = "Null"
             End If
                   
             If Not IsNull(rstBusca_Finalizadora!DFCodigo_asc_TBFinalizadora) Then
                strTecla = Trim(rstBusca_Finalizadora!DFCodigo_asc_TBFinalizadora)
             Else
                strTecla = "Null"
             End If
             
             strPermitir_Uso = "1"
             strPermitir_Troco = "1"
             strValor_Maximo = "0"
             
             If Not IsNull(rstBusca_Finalizadora!DFTroco_TBFinalizadora) Then
                strSomente_Valor_Total = Trim(rstBusca_Finalizadora!DFTroco_TBFinalizadora)
             Else
                strSomente_Valor_Total = "Null"
             End If
             
             If Not IsNull(rstBusca_Abrir_Gaveta!DFGaveta_integrada_TBPdv) Then
                strAbrir_Gaveta = Trim(rstBusca_Abrir_Gaveta!DFGaveta_integrada_TBPdv)
             Else
                strAbrir_Gaveta = "Null"
             End If
              
             strEmite_Contra_Vale = "0"
             strEmite_Troco_Cheque = "0"
             strUsar_Conta_Corrente = "0"
             strEmissao_Doc = "0"
             strAutenticacoes = "0"
             strTipo_Tef = "0"
    
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strDescricao) & "    " & Trim(strTecla) & "    " & Trim(strPermitir_Uso) & "    " & Trim(strPermitir_Troco) & "    " & Trim(strValor_Maximo) & "    " & Trim(strSomente_Valor_Total) & "    " & Trim(strAbrir_Gaveta) & "    " & Trim(strEmite_Contra_Vale) & "    " & Trim(strEmite_Troco_Cheque) & "    " & Trim(strUsar_Conta_Corrente) & "    " & Trim(strEmissao_Doc) & "    " & Trim(strAutenticacoes) & "    " & Trim(strTipo_Tef)
              
             Print #NumArq, strIntegracao
           
             rstBusca_Finalizadora.MoveNext
             
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Finalizadora processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\MODALIDADES.DAT", strDestino & "\MODALIDADES.DAT"
End Function

Private Function Gera_Operador_Ecf()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
       
       Open txtCaminho.Text & "\OPERADORES.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGO    NOME    SENHA    FUNCAO"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Operador_Ecf.RecordCount <> 0 Then
          rstBusca_Operador_Ecf.MoveFirst
   
          intConta_Sequencial = rstBusca_Operador_Ecf.RecordCount
       
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strNome As String * 40
             Dim strSenha As String * 10
             Dim strFuncao As String * 6
             
             If Not IsNull(rstBusca_Operador_Ecf!PKCodigo_TBOperadores_ecf) Then
                strCodigo = Trim(rstBusca_Operador_Ecf!PKCodigo_TBOperadores_ecf)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Operador_Ecf!DFNome_TBOperadores_ecf) Then
                strNome = Trim(rstBusca_Operador_Ecf!DFNome_TBOperadores_ecf)
             Else
                strNome = "Null"
             End If
                   
             If Not IsNull(rstBusca_Operador_Ecf!DFSenha_TBOperadores_ecf) Then
                strSenha = Trim(rstBusca_Operador_Ecf!DFSenha_TBOperadores_ecf)
             Else
                strSenha = "Null"
             End If
             
             strFuncao = "1"
             
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strNome) & "    " & Trim(strSenha) & "    " & Trim(strFuncao)
             
             Print #NumArq, strIntegracao

             rstBusca_Operador_Ecf.MoveNext
              
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Operador Ecf processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\OPERADORES.DAT", strDestino & "\OPERADORES.DAT"
End Function

Private Function Gera_Produto()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
          
       Open txtCaminho.Text & "\PRODUTOS.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
    
       strIntegracao = Empty
       strIntegracao = "CODIGO    CODIGOSWEDA    DESCRICAO    DESCRICAOREDUZIDA    ECFIMP1LINHA    UNIDADE    COMBUSTIVEL    PRECO    TIPOTRIBUTACAO    TRIBUTACAO    FAMILIA    SECAO"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Produto.RecordCount <> 0 Then
          rstBusca_Produto.MoveFirst
       
          intConta_Sequencial = rstBusca_Produto.RecordCount
       
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strCodigo_Sweda As String * 13
             Dim strDescricao As String * 50
             Dim strDescricao_Reduzida As String * 24
             Dim strEcf_Imp_1Linha As String * 1
             Dim strUnidade As String * 2
             Dim strCombustivel As String * 1
             Dim strPreco As String * 10
             Dim strTipo_Tributacao As String * 1
             Dim strTributacao As String * 10
             Dim strFamilia As String * 6
             Dim strSecao As String * 6
             
             If Not IsNull(rstBusca_Produto!IXCodigo_TBProduto) Then
                strCodigo = Trim(rstBusca_Produto!IXCodigo_TBProduto)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Produto!IXCodigo_TBProduto) Then
                strCodigo_Sweda = Trim(rstBusca_Produto!IXCodigo_TBProduto)
             Else
                strCodigo_Sweda = "Null"
             End If
             
             If Not IsNull(rstBusca_Produto!DFDescricao_TBProduto) Then
                strDescricao = Trim(rstBusca_Produto!DFDescricao_TBProduto)
             Else
                strDescricao = "Null"
             End If
             
             If Not IsNull(rstBusca_Produto!DFDescricao_resumida_TBProduto) Then
                strDescricao_Reduzida = Trim(rstBusca_Produto!DFDescricao_resumida_TBProduto)
             Else
                strDescricao_Reduzida = "Null"
             End If
    
             strEcf_Imp_1Linha = "1"
             
             If Not IsNull(rstBusca_Produto!DFPeso_variavel_TBProduto) Then
                If rstBusca_Produto!DFPeso_variavel_TBProduto = True Then
                   strUnidade = "KG"
                Else
                   strUnidade = "SD"
                End If
             Else
                strUnidade = "SD"
             End If
                                        
             strCombustivel = "0"
                                        
             If rstBusca_Tabela_Vigente.RecordCount <> 0 Then
                If Not IsNull(rstBusca_Tabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda) Then
                   If Not IsNull(rstBusca_Produto!DFPreco_avista_TBItens_tabela_preco) Then
                      strPreco = Format(Trim(rstBusca_Produto!DFPreco_avista_TBItens_tabela_preco), "#,###0.00")
                   Else
                      strPreco = "Null"
                   End If
                Else
                  strPreco = "Null"
                End If
             Else
                strPreco = "Null"
             End If
             
             If Not IsNull(rstBusca_Produto!DFTributacao_impressora_fiscal_TBEstado_icms) Then
                If Left(rstBusca_Produto!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "T" Then
                   strTipo_Tributacao = "T"
                ElseIf Left(rstBusca_Produto!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "S" Then
                   strTipo_Tributacao = "S"
                ElseIf Left(rstBusca_Produto!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "F" Then
                   strTipo_Tributacao = "F"
                ElseIf Left(rstBusca_Produto!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "N" Then
                   strTipo_Tributacao = "S"
                ElseIf Left(rstBusca_Produto!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "I" Then
                   strTipo_Tributacao = "I"
                End If
             Else
                strTipo_Tributacao = "Null"
             End If
             
             If Not IsNull(rstBusca_Produto!DFPercentual_icms_saida_juridica_TBEstado_icms) Then
                strTributacao = Trim(rstBusca_Produto!DFPercentual_icms_saida_juridica_TBEstado_icms)
             Else
                strTributacao = "Null"
             End If
             
             strFamilia = "1"
             
             If Not IsNull(rstBusca_Produto!FKCodigo_TBSecao) Then
                strSecao = Trim(rstBusca_Produto!FKCodigo_TBSecao)
             Else
                strSecao = "Null"
             End If
    
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strCodigo_Sweda) & "    " & Trim(strDescricao) & "    " & Trim(strDescricao_Reduzida) & "    " & Trim(strEcf_Imp_1Linha) & "    " & Trim(strUnidade) & "    " & Trim(strCombustivel) & "    " & Trim(strPreco) & "    " & Trim(strTipo_Tributacao) & "    " & Trim(strTributacao) & "    " & Trim(strFamilia) & "    " & Trim(strSecao)
              
             Print #NumArq, strIntegracao
             
             rstBusca_Produto.MoveNext
              
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Produto processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq

       FileCopy txtCaminho.Text & "\PRODUTOS.DAT", strDestino & "\PRODUTOS.DAT"
End Function

Private Function Gera_Codigo_Barras()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
       
       Open txtCaminho.Text & "\PRODUTOSEAN.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGOEAN    CODIGO    DESCRICAOREDUZIDA    PRECO"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Codigo_Barras.RecordCount <> 0 Then
          rstBusca_Codigo_Barras.MoveFirst
       
          intConta_Sequencial = rstBusca_Codigo_Barras.RecordCount
       
          Do While intConta_Sequencial <> 0
             Dim strCodigo_Ean As String * 13
             Dim strCodigo As String * 6
             Dim strDescricao_Reduzida As String * 24
             Dim strPreco As String * 10
             
             If Not IsNull(rstBusca_Codigo_Barras!IXCodigo_TBCodigo_barras) Then
                strCodigo_Ean = Trim(rstBusca_Codigo_Barras!IXCodigo_TBCodigo_barras)
             Else
                strCodigo_Ean = "Null"
             End If
             
             If Not IsNull(rstBusca_Codigo_Barras!IXCodigo_TBProduto) Then
                strCodigo = Trim(rstBusca_Codigo_Barras!IXCodigo_TBProduto)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Codigo_Barras!DFDescricao_resumida_TBProduto) Then
                strDescricao_Reduzida = Trim(rstBusca_Codigo_Barras!DFDescricao_resumida_TBProduto)
             Else
                strDescricao_Reduzida = "Null"
             End If
                                      
             If rstBusca_Tabela_Vigente.RecordCount <> 0 Then
                If Not IsNull(rstBusca_Tabela_Vigente!DFNumero_tabela_vigente_TBParametros_venda) Then
                   If Not IsNull(rstBusca_Codigo_Barras!DFPreco_avista_TBItens_tabela_preco) Then
                      strPreco = Format(Trim(rstBusca_Codigo_Barras!DFPreco_avista_TBItens_tabela_preco), "#,###0.00")
                   Else
                      strPreco = "Null"
                   End If
                Else
                  strPreco = "Null"
                End If
             Else
                strPreco = "Null"
             End If
    
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo_Ean) & "    " & Trim(strCodigo) & "    " & Trim(strDescricao_Reduzida) & "    " & Trim(strPreco)
              
             Print #NumArq, strIntegracao
       
             rstBusca_Codigo_Barras.MoveNext
              
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Produto Código Barras processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\PRODUTOSEAN.DAT", strDestino & "\PRODUTOSEAN.DAT"
End Function

Private Function Gera_Composicao()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile

       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
       
       Open txtCaminho.Text & "\PRODUTOSCOMP.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
    
       strIntegracao = Empty
       strIntegracao = "CODIGO    PRODUTO    QUANTIDADE"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Composicao.RecordCount <> 0 Then
          rstBusca_Composicao.MoveFirst
      
          intConta_Sequencial = rstBusca_Composicao.RecordCount
      
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strProduto As String * 6
             Dim strQuantidade As String * 10
             
             If Not IsNull(rstBusca_Composicao!DFCodigo_produto_relacionado_TBComposicao_produto) Then
                strCodigo = Trim(rstBusca_Composicao!DFCodigo_produto_relacionado_TBComposicao_produto)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Composicao!IXCodigo_TBProduto) Then
                strProduto = Trim(rstBusca_Composicao!IXCodigo_TBProduto)
             Else
                strProduto = "Null"
             End If
             
             If Not IsNull(rstBusca_Composicao!DFQuantidade_baixa_estoque_TBComposicao_produto) Then
                strQuantidade = Format(Trim(rstBusca_Composicao!DFQuantidade_baixa_estoque_TBComposicao_produto), "#,###0.000")
             Else
                strQuantidade = "Null"
             End If
                                      
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strProduto) & "    " & Trim(strQuantidade)
             
             Print #NumArq, strIntegracao
   
             rstBusca_Composicao.MoveNext
             
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Produto Composição processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\PRODUTOSCOMP.DAT", strDestino & "\PRODUTOSCOMP.DAT"
End Function

Private Function Gera_Promocoes()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
       
       Open txtCaminho.Text & "\PRODUTOSPROMO.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGO    CODIGOEAN    DATAINICIAL    DATAFINAL    PRECO"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
       MsgBox "Geração do Arquivo Integração Frente Loja - Produto Promoção processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\PRODUTOSPROMO.DAT", strDestino & "\PRODUTOSCOMP.DAT"
End Function

Private Function Gera_Familia()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
       
       Open txtCaminho.Text & "\PRODUTOSFAMILIA.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGO    DESCRICAO"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Dim strCodigo As String * 6
       Dim strDescricao As String * 40
       
       strCodigo = "1"
       strDescricao = "GERAL"
       
       strIntegracao = Empty
       strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strDescricao)
          
       Print #NumArq, strIntegracao
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Família processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\PRODUTOSFAMILIA.DAT", strDestino & "\PRODUTOSFAMILIA.DAT"
End Function

Private Function Gera_Secao()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
            
       Open txtCaminho.Text & "\SECOES.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGO    DESCRICAO    SECAONIVEL1"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Secao.RecordCount <> 0 Then
          rstBusca_Secao.MoveFirst
       
          intConta_Sequencial = rstBusca_Secao.RecordCount
       
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strDescricao As String * 30
             Dim strSecao_Nivel1 As String * 6
              
             If Not IsNull(rstBusca_Secao!PKCodigo_TBSecao) Then
                strCodigo = Trim(rstBusca_Secao!PKCodigo_TBSecao)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Secao!DFDescricao_TBsecao) Then
                strDescricao = Trim(rstBusca_Secao!DFDescricao_TBsecao)
             Else
                strDescricao = "Null"
             End If
              
             strSecao_Nivel1 = "1"
             
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strDescricao) & "    " & Trim(strSecao_Nivel1)
              
             Print #NumArq, strIntegracao
       
             rstBusca_Secao.MoveNext
             
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Seção processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\PRODUTOSFAMILIA.DAT", strDestino & "\SECOES.DAT"
End Function

Private Function Gera_Estado_Icms()
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                    A B E R T U R A       D O       A R Q.       T E X T O                        '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       'CAPTURANDO ESPAÇO LIVRE NA MEMORIA
       NumArq = FreeFile
       
       strDestino = Empty
       strDestino = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
       strDestino = Left(strDestino, CDbl(Len(strDestino) - 4)) & "\INTEGRAÇÃO"
       
       Open txtCaminho.Text & "\TRIBUTACOES.DAT" For Append As #NumArq
                    
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                                                                                  '
       '             M O N T A G E M    D O S    R E G I S T R O S    D O     L A Y O U T                 '
       '                                                                                                  '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                             C    A    B    E    Ç    A    L    H    O                            '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strIntegracao = Empty
       strIntegracao = "*"
          
       Print #NumArq, strIntegracao
       
       strIntegracao = Empty
       strIntegracao = "CODIGO    TIPO    VALOR    FABRICANTEECF"
          
       Print #NumArq, strIntegracao
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '                                       C    O    R    P    O                                      '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rstBusca_Estado_Icms.RecordCount <> 0 Then
          rstBusca_Estado_Icms.MoveFirst
       
          intConta_Sequencial = rstBusca_Estado_Icms.RecordCount
      
          Do While intConta_Sequencial <> 0
             Dim strCodigo As String * 6
             Dim strTipo As String * 1
             Dim strValor As String * 10
             Dim strFabricante_Ecf As String * 20
             
             If Not IsNull(rstBusca_Estado_Icms!PKId_TBEstado_icms) Then
                strCodigo = Trim(rstBusca_Estado_Icms!PKId_TBEstado_icms)
             Else
                strCodigo = "Null"
             End If
             
             If Not IsNull(rstBusca_Estado_Icms!DFTributacao_impressora_fiscal_TBEstado_icms) Then
                If Left(rstBusca_Estado_Icms!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "T" Then
                   strTipo = "T"
                ElseIf Left(rstBusca_Estado_Icms!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "S" Then
                   strTipo = "S"
                ElseIf Left(rstBusca_Estado_Icms!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "F" Then
                   strTipo = "F"
                ElseIf Left(rstBusca_Estado_Icms!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "N" Then
                   strTipo = "S"
                ElseIf Left(rstBusca_Estado_Icms!DFTributacao_impressora_fiscal_TBEstado_icms, 1) = "I" Then
                   strTipo = "I"
                End If
             Else
                strTipo = "Null"
             End If
             
             If Not IsNull(rstBusca_Estado_Icms!DFPercentual_icms_saida_juridica_TBEstado_icms) Then
                strValor = Trim(rstBusca_Estado_Icms!DFPercentual_icms_saida_juridica_TBEstado_icms)
             Else
                strValor = "Null"
             End If
             
             If cbbFabricante_Ecf.Text <> Empty Then
                strFabricante_Ecf = cbbFabricante_Ecf.Text
             Else
                strFabricante_Ecf = "Null"
             End If
             
             strIntegracao = Empty
             strIntegracao = "" & Trim(strCodigo) & "    " & Trim(strTipo) & "    " & Trim(strValor) & "    " & Trim(strFabricante_Ecf)
             
             Print #NumArq, strIntegracao
             
             rstBusca_Estado_Icms.MoveNext
             
             intConta_Sequencial = intConta_Sequencial - 1
          Loop
       End If
       
       MsgBox "Geração do Arquivo Integração Frente Loja - Estado Icms processado com sucesso!", vbInformation, "Only Tech"
           
       Close #NumArq
       
       FileCopy txtCaminho.Text & "\TRIBUTACOES.DAT", strDestino & "\TRIBUTACOES.DAT"
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    txtCaminho.Text = Empty
    cbbEmpresa_Recebimento.Text = Empty
    cbbFabricante_Ecf.Text = Empty
    
    For I = 1 To hfgOpcoes_Exportacao.Rows - 1
       hfgOpcoes_Exportacao.Row = I
       
       hfgOpcoes_Exportacao.Col = 2
       hfgOpcoes_Exportacao.Text = Empty
       
       hfgOpcoes_Exportacao.Col = 3
       hfgOpcoes_Exportacao.Text = Empty
    Next I
    
    txtCaminho.SetFocus
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Function CopyFile(strOrigem As String, strDestino As String) As Single
    Static Buf$
    Dim BTest!, FSize!
    Dim Chunk%, F1%, F2%

    Const BUFSIZE = 1024
    
    If Dir(strOrigem) = "" Then
       MsgBox "Arquivo não encontrado."
       Exit Function
    End If
    
    On Error Resume Next
    
    If Len(Dir(strDestino)) Then
       Kill strDestino
    End If

    On Error GoTo FileCopyError
    
    F1 = FreeFile
    Open strOrigem For Binary As F1
    
    F2 = FreeFile
    Open strDestino For Binary As F2

    FSize = LOF(F1)
    BTest = FSize - LOF(F2)
    
    Do
       If BTest < BUFSIZE Then
          Chunk = BTest
       Else
          Chunk = BUFSIZE
       End If
       
       Buf = String(Chunk, " ")
       
       Get F1, , Buf
       Put F2, , Buf
       
       BTest = FSize - LOF(F2)
    Loop Until BTest = 0
    
    Close F1
    Close F2
    
    CopyFile = FSize
    
    Exit Function

FileCopyError:
   MsgBox "Erro ao copiar, verifique!"
   
   Close F1
   Close F2
   
   Exit Function
End Function

Private Sub txtCaminho_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Monta_Opcoes_Exportacao()
    Dim I As Integer
    
    'DEFININDO NÚMERO DE LINHAS E COLUNAS
    hfgOpcoes_Exportacao.Col = 0
    hfgOpcoes_Exportacao.Cols = 4
    hfgOpcoes_Exportacao.Rows = 11
    
    'DEFININDO TAMANHO DAS COLUNAS
    hfgOpcoes_Exportacao.ColWidth(0) = 500
    hfgOpcoes_Exportacao.ColWidth(1) = 4000
    hfgOpcoes_Exportacao.ColWidth(2) = 1000
    hfgOpcoes_Exportacao.ColWidth(3) = 2500
    
    'DEFININDO ALINHAMENTO DO CABEÇALHO
    hfgOpcoes_Exportacao.ColAlignmentFixed(1) = 5
    hfgOpcoes_Exportacao.ColAlignmentFixed(2) = 5
    hfgOpcoes_Exportacao.ColAlignmentFixed(3) = 5
    
    'DEFININDO ALINHAMENTO DA COLUNA SIM DO GRID
    hfgOpcoes_Exportacao.ColAlignment(2) = 5
    
    'COLORINDO COLUNA DE INDICE
    I = hfgOpcoes_Exportacao.Rows - 1
    
    Do While I <> 0
       
       hfgOpcoes_Exportacao.Col = 0: hfgOpcoes_Exportacao.Row = I
       
       hfgOpcoes_Exportacao.CellBackColor = &H80FFFF
       
       I = I - 1
    Loop
    
    'DEFININDO CABEÇALHO PADRÃO
    hfgOpcoes_Exportacao.Row = 0
    hfgOpcoes_Exportacao.Col = 3: hfgOpcoes_Exportacao.CellBackColor = &H8000000F: hfgOpcoes_Exportacao.CellFontBold = True: hfgOpcoes_Exportacao.CellFontSize = 10
    hfgOpcoes_Exportacao.Col = 2: hfgOpcoes_Exportacao.CellBackColor = &H8000000F: hfgOpcoes_Exportacao.CellFontBold = True: hfgOpcoes_Exportacao.CellFontSize = 10
    hfgOpcoes_Exportacao.Col = 1: hfgOpcoes_Exportacao.CellBackColor = &H8000000F: hfgOpcoes_Exportacao.CellFontBold = True: hfgOpcoes_Exportacao.CellFontSize = 10
    hfgOpcoes_Exportacao.Col = 0: hfgOpcoes_Exportacao.CellBackColor = &H8000000F: hfgOpcoes_Exportacao.CellFontBold = True: hfgOpcoes_Exportacao.CellFontSize = 10
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "PROGRAMAS PARA EXPORTAÇÃO"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 2)) = "SIM"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 3)) = "RESUMO INTEGRAÇÃO"
    
    hfgOpcoes_Exportacao.Row = 1
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "1"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "CLIENTE"
    
    hfgOpcoes_Exportacao.Row = 2
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "2"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "FINALIZADORA"
    
    hfgOpcoes_Exportacao.Row = 3
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "3"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "OPERADOR ECF"
    
    hfgOpcoes_Exportacao.Row = 4
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "4"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "PRODUTO"
    
    hfgOpcoes_Exportacao.Row = 5
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "5"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "   CÓDIGO DE BARRAS"
    
    hfgOpcoes_Exportacao.Row = 6
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "6"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "   COMPOSIÇÃO"
    
    hfgOpcoes_Exportacao.Row = 7
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "7"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "   PROMOÇÕES"
    
    hfgOpcoes_Exportacao.Row = 8
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "8"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "FAMÍLIA"
    
    hfgOpcoes_Exportacao.Row = 9
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "9"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "SEÇÃO"
    
    hfgOpcoes_Exportacao.Row = 10
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 0)) = "10"
    hfgOpcoes_Exportacao.TextArray((hfgOpcoes_Exportacao.Row * hfgOpcoes_Exportacao.Cols + hfgOpcoes_Exportacao.Col + 1)) = "ESTADOS ICMS"
End Function
