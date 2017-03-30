VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
      'Constant Declaration
Private Const ODBC_ADD_DSN = 1        ' Adiciona origem de dados
Private Const ODBC_CONFIG_DSN = 2     ' Configura (edita) origem de dados
Private Const ODBC_REMOVE_DSN = 3     ' Remove origem de dados
Private Const vbAPINull As Long = 0&  ' Ponteiro NULL

Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" (phenv&) As Integer

      'Function Declare
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
       (ByVal hwndParent As Long, ByVal fRequest As Long, _
        ByVal lpszDriver As String, ByVal lpszAttributes As String) _
        As Long
        
Private Declare Function SQLDrivers Lib "ODBC32.DLL" _
       (ByVal henv&, ByVal fDirection%, ByVal szDriverDesc$, _
        ByVal cbDriverDescMax%, pcbDriverDesc&, _
        ByVal szDriverAttr$, ByVal cbDrvrAttrMax%, pcbDrvrAttr&) _
        As Integer  'Pega os drives ODBC instalados
        Private Const SQL_ERROR = -1
        Private Const SQL_INVALID_HANDLE = -2
        Private Const SQL_NO_DATA_FOUND = 100
        Private Const SQL_SUCCESS As Long = 0
        Private Const SQL_SUCCESS_WITH_INFO = 1
        
        Private ghEnv As Long
        Private ghDbc As Long
        Private ghStmt As Long
        Private Const NULL_POINTER = 0
                   
'api para buscar DSN
'Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Const BRM_MSG1 = "Ocorreu o erro número "

Private mvarAtributos As String
Private mvarDSNInstalados() As Variant
Private mvarDRVdosDSNInstalados() As Variant
Private mvarTotalDNS As Integer
Private mvarDriversInstalados() As Variant 'local copy
Private mvarErro As String 'local copy
Private mvarDriversAtributos() As Variant 'local copy
Private mvarMensagemDeErro As String 'local copy
Public Property Get MensagemDeErro() As String
    MensagemDeErro = mvarMensagemDeErro
End Property

Public Property Get DriversAtributos() As Variant
        DriversAtributos = mvarDriversAtributos
End Property

Public Property Get DriversInstalados() As Variant
        DriversInstalados = mvarDriversInstalados
End Property

Public Property Get TotalDNS() As Integer
    TotalDNS = mvarTotalDNS
End Property

Public Property Get DRVdosDSNInstalados() As Variant
        DRVdosDSNInstalados = mvarDRVdosDSNInstalados
End Property

Public Property Get DSNInstalados() As Variant
        DSNInstalados = mvarDSNInstalados
End Property

Public Property Get Atributos() As String
    Atributos = mvarAtributos
End Property
Function ExisteODBCSistema(NomeDoODBC As String) As Boolean
    Dim lstrODBC As Variant
    Dim lclReg As GMclRegedit ' clRegedit
    Set lclReg = New GMclRegedit
    Dim i
    
    lstrODBC = lclReg.ListarChaves("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI")
    
    For i = 1 To UBound(lstrODBC)
        If NomeDoODBC = (lstrODBC(i)) Then
            ExisteODBCSistema = True
            Exit For
        End If
    Next i
End Function

Function CriarODBCSistema(NomeDSN As String, _
                         driver As String, _
                         Optional Descrição As String, _
                         Optional BaseDeDados As String = "", _
                         Optional Usuario As String = "", _
                         Optional Senha As String = " ", _
                         Optional SomenteLeitura As String = "00", _
                         Optional ModoExclusivo As String = "00", _
                         Optional SafeTransactions As String = "0", _
                         Optional TamanhoBuffer As String = "2048", _
                         Optional TempoLimitePágina As String = "5", _
                         Optional ImplicitCommitSync As String = "", _
                         Optional Threads As String = "3", _
                         Optional UserCommitSync As String = "Yes") As Boolean
    
    On Error GoTo ErroCls
    Err.Clear
    If Trim(NomeDSN) = "" _
        Or Trim(driver) = "" Then Err.Raise 449 'erro Argumento não opcional
    
    Dim lstrDriveID As String
    Dim lstrFIL As String
    Dim ModoExclusivoX
    
    Dim lclReg As GMclRegedit 'clRegedit
    Set lclReg = New GMclRegedit ' clRegedit
    
        
    
    lstrDriveID = "25"
    lstrFIL = "MS Access;"
    
    If Not lclReg.WinRegCriarChave("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN) Then GoTo ErroCls
    
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                           "Description", Descrição) Then GoTo ErroCls
                                    
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                           "DBQ", BaseDeDados) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                           "Driver", driver) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                           "FIL", lstrFIL) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                           "PWD", Senha) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                           "UID", Usuario) Then GoTo ErroCls
    
    'If Not lclReg.WinRegAdicionarBinário("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                         "Exclusive", CDbl(ModoExclusivo)) Then GoTo ErroCls
    'If Not lclReg.WinRegAdicionarBinário("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                         "ReadOnly", SomenteLeitura) Then GoTo ErroCls
    
    If Not lclReg.WinRegAdicionarDWORD("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                       "DriverID", CDbl(lstrDriveID)) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarDWORD("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN, _
                                       "Safe Transactions", CDbl(SafeTransactions)) Then GoTo ErroCls
    
    If Not lclReg.WinRegCriarChave("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines") Then GoTo ErroCls
    If Not lclReg.WinRegCriarChave("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines\Jet") Then GoTo ErroCls
    
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines\Jet", _
                                           "ImplicitCommitSync", ImplicitCommitSync) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines\Jet", _
                                           "UserCommitSync", UserCommitSync) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarDWORD("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines\Jet", _
                                       "MaxBufferSize", CDbl(TamanhoBuffer)) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarDWORD("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines\Jet", _
                                       "PageTimeout", CDbl(TempoLimitePágina)) Then GoTo ErroCls
    If Not lclReg.WinRegAdicionarDWORD("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\" & NomeDSN & "\Engines\Jet", _
                                       "Threads", CDbl(Threads)) Then GoTo ErroCls
                                    
    If Not lclReg.WinRegAdicionarSequência("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI\ODBC Data Sources", _
                                            NomeDSN, driver) Then GoTo ErroCls
             
ErroCls:
    If Err.Number = 0 And lclReg.MsgErro = "" Then
        CriarODBCSistema = True
        mvarMensagemDeErro = ""
        DSN_Instalados
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.CriarODBCSistema - " & lclReg.MsgErro
    End If

      
End Function


Function RemoveDSN(NomeDSN As String) As Boolean
      Dim intRet As Long
      Dim strDriver As String
      Dim strAttributes As String
      Dim driver
      
      driver = IIf(IsNull(NomeDSN), "Microsoft Access Driver (*.mdb)", driver)
      
      strAttributes = "DSN=" & NomeDSN & Chr$(0)
      'To show dialog, use Form1.Hwnd instead of vbAPINull.
      intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)
      
      If intRet Then RemoveDSN = True: DSN_Instalados
End Function

Function BuscaDSN(NomeDSN As String) As Boolean
    BuscaDSN = False
    Dim i
    
    For i = 0 To UBound(DSNInstalados)
        If DSNInstalados(i) = NomeDSN Then BuscaDSN = True
    Next i
End Function
Private Sub DSN_Instalados()
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long     'handle to the environment
    
    ReDim mvarDSNInstalados(0)
    ReDim mvarDRVdosDSNInstalados(0)
    
    i = 0
    If SQLAllocEnv(lHenv) <> -1 Then
        Do While i = SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            If i = 0 Then
                ReDim Preserve mvarDSNInstalados(UBound(mvarDSNInstalados) + 1)
                ReDim Preserve mvarDRVdosDSNInstalados(UBound(mvarDRVdosDSNInstalados) + 1)
                mvarDSNInstalados(UBound(mvarDSNInstalados)) = VBA.Left(sDSNItem, iDSNLen)
                mvarDRVdosDSNInstalados(UBound(mvarDRVdosDSNInstalados)) = VBA.Left(sDRVItem, iDRVLen)
            End If
        Loop
    End If
    mvarTotalDNS = UBound(mvarDSNInstalados)
End Sub

Private Sub Class_Initialize()
    DSN_Instalados
    DRV_Instalados
End Sub
Private Sub DRV_Instalados()
    Dim szDriverDesc As String * 512
    Dim cbDriverDescMax As Integer
    Dim pcbDriverDesc As Long
    Dim szDriverAttributes As String * 2048
    Dim cbDrvrAttrMax As Integer
    Dim pcbDrvrAttr As Long
    Dim result As Integer
    Dim p
    
    ReDim mvarDriversInstalados(0)
    ReDim mvarDriversAtributos(0)
    
    cbDriverDescMax = 512
    cbDrvrAttrMax = 2048
    'result = SQL_SUCCESS
    
    result = SQLAllocEnv(ghEnv)
    Do While result <> SQL_NO_DATA_FOUND
        result = SQLDrivers(ghEnv, SQL_FETCH_NEXT, szDriverDesc, cbDriverDescMax, pcbDriverDesc, szDriverAttributes, cbDrvrAttrMax, pcbDrvrAttr)
        If result = SQL_ERROR Then
            mvarErro = "Erro buscando lista de Drives ODBC"
            'ErrResult = ODBCError("Env", ghEnv, 0, 0, result, "Error getting list of registered drivers.")
            Exit Sub
        End If
        
        If result = SQL_NO_DATA_FOUND Then
            Exit Do
        End If
        ReDim Preserve mvarDriversInstalados(UBound(mvarDriversInstalados) + 1)
        ReDim Preserve mvarDriversAtributos(UBound(mvarDriversAtributos) + 1)
           
        mvarDriversInstalados(UBound(mvarDriversInstalados)) = Mid(szDriverDesc, 1, InStr(szDriverDesc, Chr(0)) - 1)
        mvarDriversAtributos(UBound(mvarDriversAtributos)) = Trim(Replace(szDriverAttributes, Chr(0), " "))  'Mid(szDriverAttributes, 1, InStr(szDriverAttributes, Chr(0)) - 1)
    Loop

End Sub

Private Function ExcluirODBC(NomeDSN As String, driver As String)
      Dim intRet As Long
            
      Dim strDriver As String
      Dim strAttributes As String

      strDriver = "SQL Server"
      strAttributes = "DSN=DSN_TEMP" & Chr$(0)
      intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, _
      strDriver, strAttributes)
      If intRet Then
          MsgBox "DSN Deleted"
      Else
          MsgBox "Delete Failed"
      End If
End Function
