VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmConsole_Relatorio_Pendencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Pendências"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsole_Relatorio_Pendencias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12630
   Begin VB.CommandButton cmdConfiguracao_impressora 
      Height          =   795
      Left            =   10710
      Picture         =   "frmConsole_Relatorio_Pendencias.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Configuração da Impressora"
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdConfiguracao_relatorio 
      Height          =   795
      Left            =   11670
      Picture         =   "frmConsole_Relatorio_Pendencias.frx":3994
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Configuração do relatório"
      Top             =   0
      Width           =   915
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   6915
      Left            =   30
      TabIndex        =   2
      Top             =   840
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
      TabIndex        =   3
      Top             =   30
      Width           =   5520
   End
End
Attribute VB_Name = "frmConsole_Relatorio_Pendencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                     '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Console Relatório Pendências                                   '
' Data de Criação........: 27/04/2006                                                     '
' Equipe Responsável.....: Jones Sá Peixoto,Marcos Baião,Alex Baião,Rafael Gomes, Rodrigo '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim adrImprime As New ADODB.Recordset
Dim Aplicacao As New CRAXDRT.Application
Dim Relatorio As New CRAXDRT.Report
Dim conexao_relatorio As New DLLConexao_Sistema.conexao
Public Tabela As String

Private Sub cmdConfiguracao_impressora_Click()
    Relatorio.PrinterSetup Me.hwnd
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
    
    'Montando a nova string  de SQL com o INTO para criação da tabela temporária
    intTamanho_string = Len(frmRelatorio_Pendencias.strSql)
    inttamanho_From = InStr(1, frmRelatorio_Pendencias.strSql, "FROM")
    strSql_antes_from = Mid(frmRelatorio_Pendencias.strSql, 1, inttamanho_From - 1)
    strSql_pos_from = Mid(frmRelatorio_Pendencias.strSql, inttamanho_From, intTamanho_string)
    strRemontada_sql = strSql_antes_from + "INTO " & Tabela & " " + strSql_pos_from
    
    On Error GoTo Erro
    
   'Indicando o banco à conectar-se
   conexao_relatorio.Initial_Catalog = "BDRetaguarda"
        
   'Estabelecendo conexão com o banco
   conexao_relatorio.Abrir_conexao ("Otica")
    
    conexao_relatorio.CNConexao.Execute strRemontada_sql
    
    'Abrindo a recordset com as informações da tabela temporaria
    adrImprime.Open "SELECT * FROM " & Tabela & "", conexao_relatorio.CNConexao, adOpenKeyset, adLockOptimistic
    
    If frmRelatorio_Pendencias.optAnalitico.Value = True Then
    
       If frmRelatorio_Pendencias.optFuncionario.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Funcionario_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optCliente.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Cliente_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optMenu.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Menu_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optPrioridade.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Prioridade_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optTipo_Servico.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Tipo_Servico_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optStatus.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Status_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optPrograma.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Programa_Analitico.rpt"
       ElseIf frmRelatorio_Pendencias.optData.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Data_Analitico.rpt"
       End If
       
    ElseIf frmRelatorio_Pendencias.optSintetico Then
    
       If frmRelatorio_Pendencias.optFuncionario.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Funcionario_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optCliente.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Cliente_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optMenu.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Menu_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optPrioridade.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Prioridade_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optTipo_Servico.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Tipo_Servico_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optStatus.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Status_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optPrograma.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Programa_Sintetico.rpt"
       ElseIf frmRelatorio_Pendencias.optData.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Data_Sintetico.rpt"
       End If
       
    Else 'GRÁFICO
    
       If frmRelatorio_Pendencias.optFuncionario.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Funcionario_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optCliente.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Cliente_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optMenu.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Menu_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optPrioridade.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Prioridade_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optTipo_Servico.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Tipo_Servico_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optStatus.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Status_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optPrograma.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Programa_Grafico.rpt"
       ElseIf frmRelatorio_Pendencias.optData.Value = True Then
          strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptRelatorio_Pendencias_Data_Grafico.rpt"
       End If
    
    End If
        
    Set Relatorio = Aplicacao.OpenReport(strCaminho)
    
    Relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
    Relatorio.DiscardSavedData
    
    strNome_cliente = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me)
    
    'Passano ao Form a Empresa a qual se destina o rel --- Cliente -----
    Relatorio.FormulaFields.GetItemByName("Cliente").Text = "'" + strNome_cliente + "'"

    'Repassando o período para o cabeçalho do rel.--- Tipo Relatório -----
    Relatorio.FormulaFields.GetItemByName("Tipo_relatorio").Text = "'" & "Relatório de Pendências referente ao período de " & frmRelatorio_Pendencias.dtpPeriodo_Lancamento_Inicio.Value & " a " & frmRelatorio_Pendencias.dtpPeriodo_Lancamento_Fim.Value & " " & "'"
    
    crvFiltrar.ReportSource = Relatorio
    crvFiltrar.Refresh
    crvFiltrar.ViewReport
    
    Set adrImprime = Nothing
    Set Aplicacao = Nothing
        
    Exit Sub
Erro:
    If Err.Number = -2147206461 Then
       MsgBox "Arquivo do relatório não encontrado, verifique! A APLICAÇÃO SERÁ REINICIADA.", vbCritical, "Only Tech"
       End
    End If
    Call Erro.Erro(Me, "Otica", "load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    conexao_relatorio.CNConexao.Execute "DROP TABLE " & Tabela & " "
    
    Set Relatorio = Nothing
    
    frmRelatorio_Pendencias.optOrdenar_Alfabetico.Value = True
    frmRelatorio_Pendencias.optFuncionario.Value = True
    frmRelatorio_Pendencias.optAnalitico.Value = True
    
    'Fecha a conexão com o Banco
    conexao_relatorio.Fechar_conexao
End Sub








