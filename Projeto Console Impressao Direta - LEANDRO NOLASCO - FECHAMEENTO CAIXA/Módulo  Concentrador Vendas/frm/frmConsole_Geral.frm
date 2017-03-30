VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmConsole_Geral 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listagem"
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
   Icon            =   "frmConsole_Geral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12630
   Begin VB.CommandButton cmdConfiguracao_relatorio 
      Height          =   795
      Left            =   11670
      Picture         =   "frmConsole_Geral.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Configuração do relatório"
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdConfiguracao_impressora 
      Height          =   795
      Left            =   10710
      Picture         =   "frmConsole_Geral.frx":3994
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Configuração da Impressora"
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
Attribute VB_Name = "frmConsole_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                     '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Padronizar a interface de visualização dos Rels - console único'
' Data de Criação........: 05/07/2006                                                     '
' Equipe Responsável.....: Jones Sá Peixoto,Marcos Baião,Alex Baião,Rafael Gomes, Rodrigo '
' Desenvolvedor..........: Leandro Nolasco Ferreira                                                 '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim adrImprime As New ADODB.Recordset
Dim Aplicacao As New CRAXDRT.Application
Dim Relatorio As New CRAXDRT.Report

Dim conexao_relatorio As New DLLConexao_Sistema.Conexao
Dim Tabela As String

Private Type tParametros
    strSQL As String
    strNome_Rpt As String
    strNomes_Formulas As String
    strValores_Formulas As String
    booPastaNF As Boolean
    strSQLSubRpts As String
    strArqsSubRpts As String
    strAlias_Tabelas_Subrelatorios As String
End Type

Dim regParam As tParametros

Dim arrSubRelatorios() As String
Dim arrSQL_SubRelatorio() As String
Dim arrAlias_Tabelas() As String

Public Sub setParametros(ByVal strArg As String, ByVal strRpt As String, Optional ByVal strNome_Formulas_Rpt As String = "", Optional ByVal strValores_Formulas_Rpt As String = "", Optional ByVal PastaNF As Boolean = False, Optional ByVal SQLs_SubRelatorios As String = "", Optional ByVal Arquivos_Subrelatorios As String = "", Optional ByVal Apelidos_Tabelas_SubRelatorios As String = "")
    regParam.strSQL = strArg
    regParam.strNome_Rpt = strRpt
    regParam.strNomes_Formulas = strNome_Formulas_Rpt
    regParam.strValores_Formulas = strValores_Formulas_Rpt
    regParam.booPastaNF = PastaNF
    regParam.strSQLSubRpts = SQLs_SubRelatorios
    regParam.strArqsSubRpts = Arquivos_Subrelatorios
    regParam.strAlias_Tabelas_Subrelatorios = Apelidos_Tabelas_SubRelatorios
End Sub

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
    
    Dim lonIdx As Long
    Dim booAbrir_Subrelatorios As Boolean
    
    ReDim arrSubRelatorios(0) As String
    ReDim arrSQL_SubRelatorio(0) As String
    ReDim arrAlias_Tabelas(0) As String
    
    On Error GoTo Erro
    
    booAbrir_Subrelatorios = False
    
    'Inserindo a hora no nome da tabela
    Tabela = "TBTEMP_RELATORIO" & time
    Tabela = Replace(Tabela, ":", "_")

    'Montando a nova string  de SQL com o INTO para criação da tabela temporária
    intTamanho_string = Len(regParam.strSQL)
    inttamanho_From = InStr(1, regParam.strSQL, "FROM")
    strSql_antes_from = Mid(regParam.strSQL, 1, inttamanho_From - 1)
    strSql_pos_from = Mid(regParam.strSQL, inttamanho_From, intTamanho_string)
    strRemontada_sql = strSql_antes_from + "INTO " & Tabela & " " + strSql_pos_from
    
    On Error GoTo Erro
    
    'Indicando o banco à conectar-se
    conexao_relatorio.Initial_Catalog = "BDRetaguarda"
    
    'Estabelecendo conexão com o banco
    Call conexao_relatorio.Abrir_conexao("Otica")
    
    conexao_relatorio.CNconexao.Execute strRemontada_sql
    
    'Abrindo a recordset com as informações da tabela temporaria
    adrImprime.Open "SELECT * FROM " & Tabela & "", conexao_relatorio.CNconexao, adOpenKeyset, adLockOptimistic
    
    If regParam.booPastaNF Then
        strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me, "NF") & "\" & regParam.strNome_Rpt
    Else
        strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\" & regParam.strNome_Rpt
    End If
        
    Set Relatorio = Aplicacao.OpenReport(strCaminho)
    
    Relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
    
    'Subrelatorios
    'Call setSubRelatorios(arrSubRelatorios, arrSQL_SubRelatorio, arrAlias_Tabelas)
    'Call ReLocateDatabase(Relatorio)
    
    'Separa subrelatorios
    'Verifica se tem que abrir subrelatórios
    For lonIdx = LBound(arrSubRelatorios) To UBound(arrSubRelatorios)
        If arrSubRelatorios(lonIdx) <> Empty Then
            booAbrir_Subrelatorios = True
            Exit For
        End If
    Next lonIdx
    
    'Abre cada um subrelatorio
    If booAbrir_Subrelatorios Then
        For lonIdx = LBound(arrSubRelatorios) To UBound(arrSubRelatorios)
            Call Abrir_SubRelatorios(conexao_relatorio, arrSubRelatorios(lonIdx), arrSQL_SubRelatorio(lonIdx), arrAlias_Tabelas(lonIdx))
            Relatorio.OpenSubreport arrSubRelatorios(lonIdx)
        Next lonIdx
    End If
    
    strNome_cliente = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me)
    
    'Passano fórmulas ao relatório
    'Call setValor_Formulas(Relatorio)
    
    Relatorio.DiscardSavedData
    'Relatorio.Database.Verify
    
    
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
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call Apaga_Tabelas
    
    Set Relatorio = Nothing
    
    'Fecha a conexão com o Banco
    conexao_relatorio.Fechar_conexao
End Sub

Private Sub setValor_Formulas(ByRef objRelatorio As CRAXDRT.Report)
    
    On Error GoTo Erro
    
    Dim arrNomes_Formulas() As String
    Dim arrValores_Formulas() As String
    
    Dim intX As Long
    
    arrNomes_Formulas = Split(regParam.strNomes_Formulas, ";")
    arrValores_Formulas = Split(regParam.strValores_Formulas, ";")
    
    For intX = LBound(arrNomes_Formulas) To UBound(arrNomes_Formulas)
        objRelatorio.FormulaFields.GetItemByName(arrNomes_Formulas(intX)).Text = "'" & arrValores_Formulas(intX) & "'"
    Next
    
    Exit Sub
    
Erro:
    If Err.Number = -2147206461 Then
       MsgBox "Arquivo do relatório não encontrado, verifique! A APLICAÇÃO SERÁ REINICIADA.", vbCritical, "Only Tech"
       End
    End If
    Call Erro.Erro(Me, "Otica", "load")
 End Sub

Private Sub setSubRelatorios(ByRef arrSubRpt() As String, ByRef arrSQL_SubRpt() As String, ByRef arrApelidos() As String)

    arrSubRpt = Split(regParam.strArqsSubRpts, ";")
    arrSQL_SubRpt = Split(regParam.strSQLSubRpts, ";")
    arrAlias_Tabelas = Split(regParam.strAlias_Tabelas_Subrelatorios, ";")
    
End Sub

Private Sub Abrir_SubRelatorios(ByRef objCon As Conexao, ByVal ArqSubrpt As String, ByVal SQLSubRpt As String, ByVal strAlias_Tabela_SubRpt As String)

    Dim adrAux As New ADODB.Recordset
    Dim SubRelatorio As New CRAXDRT.Report
    Dim strCaminho As String

    'Abrindo a recordset com as informações da tabela temporaria
    adrAux.Open SQLSubRpt, objCon.CNconexao, adOpenKeyset, adLockOptimistic
    'adrAux.Open SQLSubRpt & " " & strAlias_Tabela_SubRpt, objCon.CNconexao, adOpenKeyset, adLockOptimistic
    
    If regParam.booPastaNF Then
        strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me, "NF") & "\" & ArqSubrpt
    Else
        strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\" & ArqSubrpt
    End If

    Set SubRelatorio = Aplicacao.OpenReport(strCaminho)
    
    SubRelatorio.Database.Tables.Item(1).SetDataSource adrAux, 3
    SubRelatorio.Database.Verify

    Set SubRelatorio = Nothing

End Sub

Private Sub Apaga_Tabelas()

    Dim lonIdx As Long

    conexao_relatorio.CNconexao.Execute "DROP TABLE " & Tabela
    
    For lonIdx = LBound(arrSQL_SubRelatorio) To UBound(arrSQL_SubRelatorio)
        conexao_relatorio.CNconexao.Execute "DROP TABLE " & Mid(arrSQL_SubRelatorio(lonIdx), InStr(arrSQL_SubRelatorio(lonIdx), "FROM ") + Len("FROM "))
    Next lonIdx

End Sub

Private Sub ReLocateDatabase(crxReport As CRAXDRT.Report)
   
   On Error GoTo Erro
   
   'http://support.businessobjects.com/downloads/service_packs/crystal_reports_en.asp
   
   Dim crxDatabaseTable As CRAXDRT.DatabaseTable
   Dim crxSection As CRAXDRT.Section
   Dim crxTextObject As CRAXDRT.TextObject
   Dim crObject As Object
   Dim crxSubReport As CRAXDRT.Report
   Dim crxSubDatabaseTable As CRAXDRT.DatabaseTable
   Dim crxSubSection As CRAXDRT.Section
   Dim crxSubTextObject As CRAXDRT.TextObject
   Dim crxSubObject As Object
   Dim crxConProp As CRAXDRT.ConnectionProperty

   Dim gCurMdb As String
   
   Dim I As Integer
   
'http://groups.google.com.br/group/microsoft.public.vb.crystal/browse_thread/thread/c87990a8f3303842/7b1183f418c506d2%237b1183f418c506d2
'   gCurMdb = "BDRetaguarda"
'
'
'
'   'relocate databases for all main report files
'   For Each crxDatabaseTable In crxReport.Database.Tables
'      crxDatabaseTable.ConnectionProperties.Item("Data Source") = "ONLYTECH-08"
'      crxDatabaseTable.ConnectionProperties.Item("Initial Catalog") = "BDRetaguarda"
'   Next
'
'
'   For i = 1 To crxDatabaseTable.ConnectionProperties.Count
'       Debug.Print crxConProp.Name
'   Next
'
'
'   'relocate data bases for all subreport files
'   For Each crxSection In crxReport.Sections
'      For Each crObject In crxSection.ReportObjects
'         If crObject.Kind = crSubreportObject Then
'            Set crxSubReport = crxReport.OpenSubreport(crObject.SubreportName)
'            For Each crxSubDatabaseTable In crxSubReport.Database.Tables
'               crxSubDatabaseTable.ConnectionProperties.Item("DatabaseName") = gCurMdb
'            Next
'         End If
'      Next
'   Next


'http://groups.google.com.br/group/microsoft.public.vb.crystal/browse_thread/thread/7c58667ed1992aee/f8632135dc95a778?lnk=gst&q=crxDatabaseTable.ConnectionProperties.Add+%22ServerName%22&rnum=2#f8632135dc95a778

    crxDatabaseTable.ConnectionProperties.DeleteAll
    crxDatabaseTable.ConnectionProperties.Add "Integrated Security", False
    crxDatabaseTable.ConnectionProperties.Add "Data Source", "ONLYTECH-08"
    crxDatabaseTable.ConnectionProperties.Add "Server", "ONLYTECH-08"
    crxDatabaseTable.ConnectionProperties.Add "Server Name", "ONLYTECH-08"
    crxDatabaseTable.ConnectionProperties.Add "ServerName", "ONLYTECH-08"
    crxDatabaseTable.ConnectionProperties.Add "Initial Catalog", "BDRetaguarda"
    crxDatabaseTable.ConnectionProperties.Add "Database", "BDRetaguarda"
    crxDatabaseTable.ConnectionProperties.Add "Database Name", "BDRetaguarda"
    crxDatabaseTable.ConnectionProperties.Add "DatabaseName", "BDRetaguarda"
    crxDatabaseTable.ConnectionProperties.Add "User ID", "sa"
    crxDatabaseTable.ConnectionProperties.Add "UserID", "sa"
    crxDatabaseTable.ConnectionProperties.Add "Password", Empty
    crxDatabaseTable.ConnectionProperties.Add "pwd", Empty
    crxDatabaseTable.ConnectionProperties.Add "PreQEServerName", "ONLYTECH-08"
    crxDatabaseTable.ConnectionProperties.Add "PreQEDatabaseName", "ONLYTECH-08"
    crxDatabaseTable.ConnectionProperties.Add "PreQEServerType", "OLE DB"

   Exit Sub
   
Erro:
    Call Erro.Erro(Me, "Otica")
    Exit Sub
    Resume
End Sub


