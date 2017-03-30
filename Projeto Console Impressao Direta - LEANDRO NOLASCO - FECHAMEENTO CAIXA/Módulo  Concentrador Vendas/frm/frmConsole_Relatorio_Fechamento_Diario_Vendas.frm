VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmConsole_Relatorio_Fechamento_Diario_Vendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Fechamento Diário de Vendas"
   ClientHeight    =   7830
   ClientLeft      =   330
   ClientTop       =   540
   ClientWidth     =   12600
   Icon            =   "frmConsole_Relatorio_Fechamento_Diario_Vendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdConfiguracao_relatorio 
      Height          =   795
      Left            =   11640
      Picture         =   "frmConsole_Relatorio_Fechamento_Diario_Vendas.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configuração do relatório"
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdConfiguracao_impressora 
      Height          =   795
      Left            =   10680
      Picture         =   "frmConsole_Relatorio_Fechamento_Diario_Vendas.frx":344C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Configuração da Impressora"
      Top             =   60
      Width           =   915
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   6915
      Left            =   0
      TabIndex        =   1
      Top             =   900
      Width           =   12585
      lastProp        =   500
      _cx             =   22199
      _cy             =   12197
      DisplayGroupTree=   0   'False
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Only Tech Solutions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   5520
   End
End
Attribute VB_Name = "frmConsole_Relatorio_Fechamento_Diario_Vendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Padronizar a interface de visualizaçõ dos Rels                 '
' Data de Criação........: 03/05/2004                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim adrImprime As New ADODB.Recordset
Dim Aplicacao As New CRAXDRT.Application
Dim Relatorio As New CRAXDRT.Report
Dim conexao_relatorio As New DLLConexao_Sistema.Conexao
Public Tabela As String

Private Sub cmdConfiguracao_impressora_Click()
    Relatorio.PrinterSetup Me.hWnd
End Sub

Private Sub Form_Load()

    Dim intTamanho_string As Integer
    Dim inttamanho_From As Integer
    Dim strCaminho As String
    Dim strSql_antes_from As String
    Dim strSql_pos_from As String
    Dim strRemontada_sql As String
    Dim strNome_cliente As String
    
    On Error GoTo erro
        
    'Inserindo a hora no nome da tabela
    Tabela = "TBTEMP_RELATORIO" & time
    Tabela = Replace(Tabela, ":", "_")
    
    'Montando a nova string  de SQL com o INTO para criação da tabela temporária
    intTamanho_string = Len(frmRelatorio_Fechamento_Diario_Vendas.strSql)
    inttamanho_From = InStr(1, frmRelatorio_Fechamento_Diario_Vendas.strSql, "FROM")
    strSql_antes_from = Mid(frmRelatorio_Fechamento_Diario_Vendas.strSql, 1, inttamanho_From - 1)
    strSql_pos_from = Mid(frmRelatorio_Fechamento_Diario_Vendas.strSql, inttamanho_From, intTamanho_string)
    strRemontada_sql = strSql_antes_from + "INTO " & Tabela & " " + strSql_pos_from
    
    On Error GoTo erro
    
    'Indicando o banco à conectar-se
    
    conexao_relatorio.Initial_Catalog = "BDRetaguarda"
    
    'Estabelecendo conexão com o banco
    conexao_relatorio.Abrir_conexao ("Otica")
    
    conexao_relatorio.CNconexao.Execute strRemontada_sql
    
    'Abrindo a recordset com as informações da tabela temporaria
    adrImprime.Open "SELECT * FROM " & Tabela & "", conexao_relatorio.CNconexao, adOpenKeyset, adLockOptimistic
        
    If frmRelatorio_Fechamento_Diario_Vendas.optSintetico = True Then
       If frmRelatorio_Fechamento_Diario_Vendas.optCliente.Value = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_cliente.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFaixa_Horaria = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_faixa_horaria.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optData = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_data.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optPDV = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_pdv.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_vendedor.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optProduto = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_produto.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFinalizadora = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_sintetico_finalizadora.rpt"
       End If
    End If
    
    If frmRelatorio_Fechamento_Diario_Vendas.optGrafico = True Then
       If frmRelatorio_Fechamento_Diario_Vendas.optCliente.Value = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_grafico_cliente.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFaixa_Horaria = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_grafico_faixa_horaria.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optData = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_grafico_data.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_grafico_vendedor.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optProduto = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_grafico_produto.rpt"
       End If
    End If

    If frmRelatorio_Fechamento_Diario_Vendas.optAnalitico = True Then
       If frmRelatorio_Fechamento_Diario_Vendas.optProduto.Value = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_analitico_produto.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_analitico_vendedor.rpt"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFinalizadora = True Then
           strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptCurva_fechamento_venda_analitico_finalizadora.rpt"
       End If
    End If
    
    Set Relatorio = Aplicacao.OpenReport(strCaminho)

    Relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
    Relatorio.DiscardSavedData
    
    strNome_cliente = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me)
    
    'Passano ao Form a Empresa a qual se destina o rel --- Cliente -----
     Relatorio.FormulaFields.GetItemByName("Cliente").Text = "'" + strNome_cliente + "'"
     
    If frmRelatorio_Fechamento_Diario_Vendas.optSintetico = True Then
       If frmRelatorio_Fechamento_Diario_Vendas.optCliente.Value = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Cliente - Sintético" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFaixa_Horaria = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "Faixa Horária - Sintético" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optData = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Data - Sintético" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optPDV = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por PDV - Sintético" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Vendedor - Sintético" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Finalizadora - Sintético" + "'"
       End If
    End If

    If frmRelatorio_Fechamento_Diario_Vendas.optGrafico = True Then
       If frmRelatorio_Fechamento_Diario_Vendas.optCliente.Value = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Cliente - Gráfico" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFaixa_Horaria = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Faixa Horária - Gráfico" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optData = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Faixa Data - Gráfico" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Vendedor - Gráfico" + "'"
       End If
    End If
    
    If frmRelatorio_Fechamento_Diario_Vendas.optAnalitico = True Then
       If frmRelatorio_Fechamento_Diario_Vendas.optProduto.Value = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Produto - Analitico" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optVendedor = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Vendedor - Analitico" + "'"
       ElseIf frmRelatorio_Fechamento_Diario_Vendas.optFinalizadora = True Then
           Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Finalizadora - Analitico" + "'"
       End If
    End If
    
    crvFiltrar.ReportSource = Relatorio
    crvFiltrar.Refresh
    crvFiltrar.ViewReport
    
    Set adrImprime = Nothing
    Set Aplicacao = Nothing
        
    Exit Sub
    
erro:

    If Err.Number = -2147206461 Then
       MsgBox "Arquivo do relatório não encontrado, verifique! A APLICAÇÃO SERÁ REINICIADA.", vbCritical, "Only Tech"
       End
    End If
    
    Call erro.erro(Me, "Otica", "load")
    
    Exit Sub
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    conexao_relatorio.CNconexao.Execute "DROP TABLE " & Tabela & " "
    
    Set Relatorio = Nothing
    
    'Fecha a conexão com o Banco
    conexao_relatorio.Fechar_conexao
End Sub
