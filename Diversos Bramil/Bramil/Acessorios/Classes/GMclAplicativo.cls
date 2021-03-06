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
Private mvarC�digoDaEmpresa_i As Integer 'local copy
Private mvarTextoCabe�alhoFormul�rios As String 'local copy
Private mvarCabe�alhoRelat�rios_str As String 'local copy
Private mvarUFdaEmpresa_str As String 'local copy
Private mvarNomeReduzidoNovell_bo As Boolean 'local copy
Private mvarPastaAcess�rios_str As String 'local copy
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

Public Property Get PastaAcess�rios_str() As String
    PastaAcess�rios_str = mvarPastaAcess�rios_str
End Property

Public Property Get NomeReduzidoNovell_bo() As Boolean
    NomeReduzidoNovell_bo = mvarNomeReduzidoNovell_bo
End Property

Public Property Get UFdaEmpresa_str() As String
    UFdaEmpresa_str = mvarUFdaEmpresa_str
End Property

Public Property Get Cabe�alhoRelat�rios_str() As String
    Cabe�alhoRelat�rios_str = mvarCabe�alhoRelat�rios_str
End Property

Public Property Get Cabe�alhoFormul�rios_str() As String
    Cabe�alhoFormul�rios_str = mvarCabe�alhoFormul�rios_str
End Property
Public Property Let C�digoDaEmpresa_i(ByVal vData As Integer)
    mvarC�digoDaEmpresa_i = vData
End Property

Public Property Get C�digoDaEmpresa_i() As Integer
    C�digoDaEmpresa_i = mvarC�digoDaEmpresa_i
End Property

Function ExtrairConfigura��es()

    Dim lstrSql As String
      'pega o c�digo da loja
    lstrSql = "Select * FROM tParametros"
    linha = 1
    Set prsSele��o = pfrsSelecao(pdbConfus, lstrSql)
    'pfboQuery pdbConfus, lstrSql, prsSele��o, "S"
    
    If prsSele��o.EOF Then
        MsgBox "Loca��o tabela GMS002 nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Function
    End If
    mvarC�digoDaEmpresa_i = prsSele��o.Fields("iCodEmpresatPar")              'C�digo da Empresa
    mvarCabe�alhoFormul�rios_str = prsSele��o.Fields("strNomeIdentFormtPar")  'cabe�alho dos Formul�rios
    mvarCabe�alhoRelat�rios_str = prsSele��o.Fields("strNomeIdentReporttPar") 'Cabe�alho dos Relat�rios
    mvarUFdaEmpresa_str = prsSele��o.Fields("strEstadoEmpresatPar")           'UF da empresa
    mvarNomeReduzidoNovell_bo = prsSele��o.Fields("boNomeExecReduzidotPar")   'Nome Reduzido Novell
    mvarPastaAcess�rios_str = prsSele��o.Fields("strLocacaoAcessoriostPar")   'pasta Acess�rios
    mvarMixDaEmpresa_b = prsSele��o.Fields("bIndMixEmpresatPar")        'Mix da Empres
End Function

Private Sub Class_Initialize()
    ExtrairConfigura��es
End Sub



