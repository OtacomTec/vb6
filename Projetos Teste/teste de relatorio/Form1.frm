VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9150
      Top             =   720
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":00A8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   5655
      Left            =   -30
      TabIndex        =   0
      Top             =   60
      Width           =   11295
      lastProp        =   500
      _cx             =   19923
      _cy             =   9975
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adrImprime As New ADODB.Recordset
Dim aplicacao As New CRAXDDRT.Application
Dim relatorio As New CRAXDDRT.Report
Dim strSQl As String
Public TABELA As String

Private Sub Form_Load()
    Dim conexao As New ADODB.Connection
    
    'sql
    'conexao.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDHomologacao;Data Source=LOGICX-SERVER"
    
    'access em s:
    'conexao.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=S:\Sistemas\Mercedes\BD\BDHomologacao.mdb;Persist Security Info=False"
    
    'access em c:
    'conexao.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Administrador\Meus documentos\Projetos\Mercedes\BDHomologacao.mdb;Persist Security Info=False"
    
    TABELA = "TBTEMP_ESMALTE" & Time
    TABELA = Replace(TABELA, ":", "_")
    'Criação SQL da tabela temporaria
    strSQl = ""
    strSQl = "SELECT TBFornecedor.*, TBEsmalte.*, TBHomologacao_Esmalte.* INTO " & TABELA & " " & _
             "FROM TBFornecedor INNER JOIN (TBEsmalte INNER JOIN TBHomologacao_Esmalte " & _
             "ON TBEsmalte.PKCodigo_TBEsmalte = TBHomologacao_Esmalte.FKCodigo_TBEsmalte) " & _
             "ON TBFornecedor.PKCodigo_TBFornecedor = TBHomologacao_Esmalte.FKCodigo_TBFornecedor " & _
             "WHERE TBEsmalte.PKCodigo_TBEsmalte = 111  "
             
    'Gerar a impressão
    conexao.Execute strSQl
    
    adrImprime.Open "SELECT * FROM " & TABELA & "", conexao, adOpenKeyset, adLockOptimistic
    
    Set relatorio = aplicacao.OpenReport("C:\Mercedes\rptEsmalte.rpt")
    relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
    relatorio.DiscardSavedData
    
    crvFiltrar.ReportSource = relatorio
    crvFiltrar.Refresh
    crvFiltrar.ViewReport
    
    Set adrImprime = Nothing
    Set aplicacao = Nothing
    Set relatorio = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Eliminação da tabela no banco
End Sub




