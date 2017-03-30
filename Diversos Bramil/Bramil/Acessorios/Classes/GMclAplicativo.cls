VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclAplicativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarCódigoDaEmpresa_i As Integer 'local copy
Private mvarTextoCabeçalhoFormulários As String 'local copy
Private mvarCabeçalhoRelatórios_str As String 'local copy
Private mvarUFdaEmpresa_str As String 'local copy
Private mvarNomeReduzidoNovell_bo As Boolean 'local copy
Private mvarPastaAcessórios_str As String 'local copy
Private mvarMixDaEmpresa_b As Byte 'local copy
Private Enum GM_Cript
    GMenCript_DeCriptar = 0
    GMenCript_EnCriptar = 1
End Enum
'local variable(s) to hold property value(s)
Private mvarSenhaBDConfus As String 'local copy
Public Property Let SenhaBDConfus(ByVal vData As String)
    mvarSenhaBDConfus = vData
End Property

Public Property Get SenhaBDConfus() As String
    SenhaBDConfus = mvarSenhaBDConfus
End Property

'Dim MyDebugA As MyDebug
Public Property Get MixDaEmpresa_b() As Byte
    MixDaEmpresa_b = mvarMixDaEmpresa_b
End Property

Public Property Get PastaAcessórios_str() As String
    PastaAcessórios_str = mvarPastaAcessórios_str
End Property

Public Property Get NomeReduzidoNovell_bo() As Boolean
    NomeReduzidoNovell_bo = mvarNomeReduzidoNovell_bo
End Property

Public Property Get UFdaEmpresa_str() As String
    UFdaEmpresa_str = mvarUFdaEmpresa_str
End Property

Public Property Get CabeçalhoRelatórios_str() As String
    CabeçalhoRelatórios_str = mvarCabeçalhoRelatórios_str
End Property

Public Property Get CabeçalhoFormulários_str() As String
    CabeçalhoFormulários_str = mvarCabeçalhoFormulários_str
End Property
Public Property Let CódigoDaEmpresa_i(ByVal vData As Integer)
    mvarCódigoDaEmpresa_i = vData
End Property

Public Property Get CódigoDaEmpresa_i() As Integer
    CódigoDaEmpresa_i = mvarCódigoDaEmpresa_i
End Property

Function ExtrairConfigurações()

    Dim lstrSql As String
      'pega o código da loja
    lstrSql = "Select * FROM tParametros"
    linha = 1
    Set prsSeleção = pfrsSelecao(pdbConfus, lstrSql)
    'pfboQuery pdbConfus, lstrSql, prsSeleção, "S"
    
    If prsSeleção.EOF Then
        MsgBox "Locação tabela GMS002 nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Function
    End If
    mvarCódigoDaEmpresa_i = prsSeleção.Fields("iCodEmpresatPar")              'Código da Empresa
    mvarCabeçalhoFormulários_str = prsSeleção.Fields("strNomeIdentFormtPar")  'cabeçalho dos Formulários
    mvarCabeçalhoRelatórios_str = prsSeleção.Fields("strNomeIdentReporttPar") 'Cabeçalho dos Relatórios
    mvarUFdaEmpresa_str = prsSeleção.Fields("strEstadoEmpresatPar")           'UF da empresa
    mvarNomeReduzidoNovell_bo = prsSeleção.Fields("boNomeExecReduzidotPar")   'Nome Reduzido Novell
    mvarPastaAcessórios_str = prsSeleção.Fields("strLocacaoAcessoriostPar")   'pasta Acessórios
    mvarMixDaEmpresa_b = prsSeleção.Fields("bIndMixEmpresatPar")        'Mix da Empres
End Function

Private Sub Class_Initialize()
    ExtrairConfigurações
End Sub



