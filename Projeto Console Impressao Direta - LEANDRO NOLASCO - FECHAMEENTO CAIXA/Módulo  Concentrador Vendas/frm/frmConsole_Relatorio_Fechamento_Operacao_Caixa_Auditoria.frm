VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmConsole_Relatorio_Fechamento_Operacao_Caixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relat�rio de fechamento de operacao de caixa"
   ClientHeight    =   7830
   ClientLeft      =   330
   ClientTop       =   540
   ClientWidth     =   12600
   Icon            =   "frmConsole_Relatorio_Fechamento_Operacao_Caixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdConfiguracao_relatorio 
      Height          =   795
      Left            =   11640
      Picture         =   "frmConsole_Relatorio_Fechamento_Operacao_Caixa.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configura��o do relat�rio"
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdConfiguracao_impressora 
      Height          =   795
      Left            =   10680
      Picture         =   "frmConsole_Relatorio_Fechamento_Operacao_Caixa.frx":344C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Configura��o da Impressora"
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
Attribute VB_Name = "frmConsole_Relatorio_Fechamento_Operacao_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' M�dulo.................: Concentrador Vendas                                            '
' Objetivo...............: Padronizar a interface de visualiza�� dos Rels                 '
' Data de Cria��o........: 03/05/2004                                                     '
' Equipe Respons�vel.....: Only Tech Solutions                                            '
' �ltima Manuten��o......:                                                                '
' Desenvolvedor..........:                                                                '
' Data �ltima manuten��o.:   /  /                                                         '
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
    
    On Error GoTo Erro
        
    'Inserindo a hora no nome da tabela
    Tabela = "TBTEMP_RELATORIO" & time
    Tabela = Replace(Tabela, ":", "_")
    
    'Montando a nova string  de SQL com o INTO para cria��o da tabela tempor�ria
    intTamanho_string = Len(frmRelatorio_Operacao_Caixa.strsql)
    inttamanho_From = InStr(1, frmRelatorio_Operacao_Caixa.strsql, "FROM")
    strSql_antes_from = Mid(frmRelatorio_Operacao_Caixa.strsql, 1, inttamanho_From - 1)
    strSql_pos_from = Mid(frmRelatorio_Operacao_Caixa.strsql, inttamanho_From, intTamanho_string)
    strRemontada_sql = strSql_antes_from + "INTO " & Tabela & " " + strSql_pos_from
    
    On Error GoTo Erro
    
    'Indicando o banco � conectar-se
    
    conexao_relatorio.Initial_Catalog = "BDRetaguarda"
    
    'Estabelecendo conex�o com o banco
    conexao_relatorio.Abrir_conexao ("Otica")
    
    conexao_relatorio.CNconexao.Execute strRemontada_sql
    
    'Abrindo a recordset com as informa��es da tabela temporaria
    adrImprime.Open "SELECT * FROM " & Tabela & "", conexao_relatorio.CNconexao, adOpenKeyset, adLockOptimistic
    
    'Setando o relatorio a ser chamado
    If frmRelatorio_Operacao_Caixa.optAnalitico.Value = True Then
       If frmRelatorio_Operacao_Caixa.optPDV.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_pdv.rpt"
       ElseIf frmRelatorio_Operacao_Caixa.optFinalizadora.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_finalizadora.rpt"
       ElseIf frmRelatorio_Operacao_Caixa.optData.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_dia.rpt"
       ElseIf frmRelatorio_Operacao_Caixa.optOperador.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_operador.rpt"
       ElseIf frmRelatorio_Operacao_Caixa.optFaixa_Horaria.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_dia_faixa_horaria.rpt"
       End If
    Else
       If frmRelatorio_Operacao_Caixa.optFinalizadora_Periodo.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_finalizadora_periodo_sintetico.rpt"
       ElseIf frmRelatorio_Operacao_Caixa.optFinalizadora.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptFechamento_operacao_caixa_finalizadora_sintetico.rpt"
       End If
    End If
    
    DoEvents
    
    Set Relatorio = Aplicacao.OpenReport(strCaminho)

    Relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
    Relatorio.DiscardSavedData
    
    strNome_cliente = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me)
    
    'Passano ao Form a Empresa a qual se destina o rel --- Cliente -----
    Relatorio.FormulaFields.GetItemByName("Cliente").Text = "'" + strNome_cliente + "'"
    
    'Descritivo do relatorio
    If frmRelatorio_Operacao_Caixa.optPDV.Value = True Then
       Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por PDV" + "'"
    ElseIf frmRelatorio_Operacao_Caixa.optFinalizadora.Value = True Then
       Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Finalizadora" + "'"
    ElseIf frmRelatorio_Operacao_Caixa.optData.Value = True Then
       Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Data" + "'"
    ElseIf frmRelatorio_Operacao_Caixa.optOperador.Value = True Then
       Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Operador" + "'"
    ElseIf frmRelatorio_Operacao_Caixa.optFaixa_Horaria.Value = True Then
       Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Faixa Hor�ria" + "'"
    ElseIf frmRelatorio_Operacao_Caixa.optFinalizadora_Periodo.Value = True Then
       Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" + "por Finalizadora Per�odo" + "'"
       Relatorio.FormulaFields.GetItemByName("Periodo").Text = "'" & frmRelatorio_Operacao_Caixa.dtpInicial.Value & "  a  " & frmRelatorio_Operacao_Caixa.dtpFinal.Value & "'"
    End If

    'Repassano o per�odo para o cabe�alho do rel.
    
    crvFiltrar.ReportSource = Relatorio
    crvFiltrar.Refresh
    crvFiltrar.ViewReport
    
    Set adrImprime = Nothing
    Set Aplicacao = Nothing
        
    Exit Sub
    
Erro:

    If Err.Number = -2147206461 Then
       MsgBox "Arquivo do relat�rio n�o encontrado, verifique! A APLICA��O SER� REINICIADA.", vbCritical, "Only Tech"
       End
    End If
    
    Call Erro.Erro(Me, "Otica", "load")
    
    Exit Sub
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    conexao_relatorio.CNconexao.Execute "DROP TABLE " & Tabela & " "
    
    Set Relatorio = Nothing
    
    'Fecha a conex�o com o Banco
    conexao_relatorio.Fechar_conexao
End Sub
