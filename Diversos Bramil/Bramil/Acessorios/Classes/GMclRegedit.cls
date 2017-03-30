VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclRegedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'CLASSE DESENVOLVIDA POR Luis Henrique Borges
'
Private Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
 
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long

'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Dim hKey As Long, ChavePrincipal As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

'This constant determins wether or not to display error messages to the
'user. I have set the default Valor to False as an error message can and
'does become irritating after a while. Turn this Valor to true if you want
'to debug your programming code when reading and writing to your system
'registry, as any errors will be displayed in a message box.
Private Const BRM_MSG1 = "Ocorreu o erro número "

Const DisplayErrorMsg = False
Private mvarMsgErro As String
Public Property Get MsgErro() As String
    MsgErro = mvarMsgErro
End Property
Function ListarChaves(NomeDaChave As String) As Variant
    Dim Ret&, Chave&, Cnt&, Tam&
    Dim BUF As String * 100
    Dim Chaves() As String
    ReDim Chaves(0)  'alimenta a variável com um vetor
    AnalisarChave NomeDaChave, ChavePrincipal
    
    Do While Ret = 0
        Ret = RegOpenKey(ChavePrincipal, NomeDaChave, Chave)
        If Ret = 0 Then
            Ret = RegEnumKey(Chave, Cnt, BUF, 100&)
            If Ret = 0 Then
                Tam = InStr(BUF, vbNullChar)
                ReDim Preserve Chaves(UBound(Chaves()) + 1)
                Chaves(UBound(Chaves())) = Left(BUF, Tam - 1)
            End If
            Cnt = Cnt + 1
        End If
    Loop
    ListarChaves = Chaves
    RegCloseKey Chave
End Function
Function WinRegAdicionarDWORD(NomeDaChave As String, _
                        ItemDaChave As String, _
                        Optional Valor As Long = 0) As Boolean

    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" Or Trim(ItemDaChave) = "" Then Err.Raise 449
    Err.Clear
    
    AnalisarChave NomeDaChave, ChavePrincipal
    
    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, NomeDaChave, 0, KEY_WRITE, hKey) 'abre a chave
        If rtn = ERROR_SUCCESS Then
            rtn = RegSetValueExA(hKey, ItemDaChave, 0, REG_DWORD, Valor, 4) 'gravar o Valor
            If Not rtn = ERROR_SUCCESS Then
                Err.Raise 31036 'Error saving to file
            End If
        Else
            Err.Raise 361  'Can't load or unload this object
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If

ErroCls:
    If Err.Number = 0 Then
        WinRegAdicionarDWORD = True
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegAdicionarDWORD - " & Err.Description
        rtn = RegCloseKey(hKey)
    End If
End Function
Function WinRegLerDWORD(NomeDaChave As String, ItemDaChave As String) As String
    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" Or Trim(ItemDaChave) = "" Then Err.Raise 449
    Err.Clear
    
    AnalisarChave NomeDaChave, ChavePrincipal

    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, NomeDaChave, 0, KEY_READ, hKey)
        If rtn = ERROR_SUCCESS Then
            rtn = RegQueryValueExA(hKey, ItemDaChave, 0, REG_DWORD, lBuffer, 4)
            If rtn = ERROR_SUCCESS Then
                rtn = RegCloseKey(hKey)
                WinRegLerDWORD = lBuffer
            Else
                WinRegLerDWORD = "-Erro-"
                Err.Raise 361
            End If
        Else 'otherwise, if the key couldnt be opened
            WinRegLerDWORD = "-Erro-"        'return Error to the user
            Err.Raise 361
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If

ErroCls:
    If Err.Number = 0 Then
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegLerDWORD - " & Err.Description
        rtn = RegCloseKey(hKey)
    End If
End Function
Function WinRegAdicionarBinário(NomeDaChave As String, _
                          ItemDaChave As String, _
                          Optional Valor As String) As Boolean
    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" _
    Or Trim(ItemDaChave) = "" Then Err.Raise 449
    Err.Clear
    
    AnalisarChave NomeDaChave, ChavePrincipal

    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, NomeDaChave, 0, KEY_WRITE, hKey)
        If rtn = ERROR_SUCCESS Then
            lDataSize = Len(Valor)
            ReDim ByteArray(lDataSize)
            For i = 1 To lDataSize
                ByteArray(i) = Asc(Mid$(Valor, i, 1))
            Next
            rtn = RegSetValueExB(hKey, ItemDaChave, 0, REG_BINARY, ByteArray(1), lDataSize)
            If Not rtn = ERROR_SUCCESS Then
                Err.Raise 31036 'Error saving to file
            End If
        Else
            Err.Raise 361  'Can't load or unload this object
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If
    
ErroCls:
    If Err.Number = 0 Then
        WinRegAdicionarBinário = True
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegAdicionarBinário - " & Err.Description
    End If
    rtn = RegCloseKey(hKey)
End Function


Function WinRegLerBinário(NomeDaChave As String, ItemDaChave As String)
    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" _
    Or Trim(ItemDaChave) = "" Then Err.Raise 449
    Err.Clear
    
    AnalisarChave NomeDaChave, ChavePrincipal

    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, NomeDaChave, 0, KEY_READ, hKey)
        If rtn = ERROR_SUCCESS Then
            lBufferSize = 1
            rtn = RegQueryValueEx(hKey, ItemDaChave, 0, REG_BINARY, 0, lBufferSize)
            sBuffer = Space(lBufferSize)
            rtn = RegQueryValueEx(hKey, ItemDaChave, 0, REG_BINARY, sBuffer, lBufferSize)
            If rtn = ERROR_SUCCESS Then
                rtn = RegCloseKey(hKey)
                WinRegLerBinário = sBuffer
            Else
                WinRegLerBinário = "-Erro-"
                Err.Raise 361
            End If
        Else
            WinRegLerBinário = "-Erro-"
            Err.Raise 361
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If
    
ErroCls:
    If Err.Number = 0 Then
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegLerBinário - " & Err.Description
        rtn = RegCloseKey(hKey)
    End If
End Function
Function WinRegExcluirChave(Chave As String) As Boolean
    If Trim(Chave) = "" Then Err.Raise 449
    Err.Clear
    AnalisarChave Chave, ChavePrincipal

    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, Chave, 0, KEY_WRITE, hKey)
        If rtn = ERROR_SUCCESS Then
            rtn = RegDeleteKey(hKey, Chave)
            If Not rtn = ERROR_SUCCESS Then Err.Raise 335 'Could not access system registry
            rtn = RegCloseKey(hKey)
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If
    
ErroCls:
    If Err.Number = 0 Then
        WinRegExcluirChave = True
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegExcluirChave - " & Err.Description
        rtn = RegCloseKey(hKey)
    End If

End Function

Private Function IdentificarChaveRaiz(MainKeyName As String) As Long

    Const HKEY_CLASSES_ROOT = &H80000000
    Const HKEY_CURRENT_USER = &H80000001
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const HKEY_USERS = &H80000003
    Const HKEY_PERFORMANCE_DATA = &H80000004
    Const HKEY_CURRENT_CONFIG = &H80000005
    Const HKEY_DYN_DATA = &H80000006
   
    Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            IdentificarChaveRaiz = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            IdentificarChaveRaiz = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            IdentificarChaveRaiz = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            IdentificarChaveRaiz = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            IdentificarChaveRaiz = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            IdentificarChaveRaiz = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            IdentificarChaveRaiz = HKEY_DYN_DATA
    End Select

End Function

Private Function MensagemDeErro(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

    Select Case lErrorCode
       Case 1009, 1015
            MensagemDeErro = "O arquivo de registro está corrompido!"
       Case 2, 1010
            MensagemDeErro = "Item da Chave Inválido"
       Case 1011
            MensagemDeErro = "Não posso abrir a Chave"
       Case 4, 1012
            MensagemDeErro = "Não posso ler a Chave"
       Case 5
            MensagemDeErro = "Impossível o acesso a Chave"
       Case 1013
            MensagemDeErro = "Não posso gravar a Chave"
       Case 8, 14
            MensagemDeErro = "Memória Insuficiente"
       Case 87
            MensagemDeErro = "Parâmetro Inválido"
       Case 234
            MensagemDeErro = "There is more data than the buffer has been allocated to hold."
       Case Else
            MensagemDeErro = "Código de erro indefinido:  " & Str$(lErrorCode)
    End Select

End Function

Function WinRegLerSequência(NomeDaChave As String, ItemDaChave As String) As String
    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" _
    Or Trim(ItemDaChave) = "" Then Err.Raise 449
    Err.Clear
    
    AnalisarChave NomeDaChave, ChavePrincipal

    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, NomeDaChave, 0, KEY_READ, hKey)
        If rtn = ERROR_SUCCESS Then
            sBuffer = Space(255)
            lBufferSize = Len(sBuffer)
            rtn = RegQueryValueEx(hKey, ItemDaChave, 0, REG_SZ, sBuffer, lBufferSize)
            If rtn = ERROR_SUCCESS Then
                rtn = RegCloseKey(hKey)
                sBuffer = Trim(sBuffer)
                WinRegLerSequência = Left(sBuffer, Len(sBuffer) - 1)
            Else
                WinRegLerSequência = "-Erro-"
                Err.Raise 361
            End If
        Else
            WinRegLerSequência = "-Erro-"
            Err.Raise 361
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If

ErroCls:
    If Err.Number = 0 Then
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegLerSequência - " & Err.Description
    End If
    rtn = RegCloseKey(hKey)

End Function

Private Sub AnalisarChave(Chave As String, Keyhandle As Long)
    
    rtn = InStr(Chave, "\")

    If Left(Chave, 5) <> "HKEY_" Or Right(Chave, 1) = "\" Then
        'MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + Chave
        Exit Sub
    ElseIf rtn = 0 Then
        Keyhandle = IdentificarChaveRaiz(Chave)
        Chave = ""
    Else
        Keyhandle = IdentificarChaveRaiz(Left(Chave, rtn - 1))
        Chave = Right(Chave, Len(Chave) - rtn)
    End If

End Sub
Function WinRegCriarChave(NomeDaChave As String) As Boolean
    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" Then Err.Raise 449
    Err.Clear
    
    AnalisarChave NomeDaChave, ChavePrincipal
    If ChavePrincipal Then
        rtn = RegCreateKey(ChavePrincipal, NomeDaChave, hKey)
        If Not rtn = ERROR_SUCCESS Then
            Err.Raise 31036 'Error saving to file
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If
    
ErroCls:
    If Err.Number = 0 Then
        WinRegCriarChave = True
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegCriarChave - " & Err.Description
    End If
    rtn = RegCloseKey(hKey)
End Function
Function WinRegAdicionarSequência(NomeDaChave As String, _
                         ItemDaChave As String, _
                         Optional Valor As String = " ") As Boolean
    On Error GoTo ErroCls
    If Trim(NomeDaChave) = "" _
    Or Trim(ItemDaChave) = "" Then Err.Raise 449
    Err.Clear
    AnalisarChave NomeDaChave, ChavePrincipal
    If Valor = "" Then Valor = " "
    
    If ChavePrincipal Then
        rtn = RegOpenKeyEx(ChavePrincipal, NomeDaChave, 0, KEY_WRITE, hKey) '
        If rtn = ERROR_SUCCESS Then
            rtn = RegSetValueEx(hKey, ItemDaChave, 0, REG_SZ, ByVal Valor, Len(Valor))
            If Not rtn = ERROR_SUCCESS Then
                Err.Raise 31036 'Error saving to file
            End If
        Else
            Err.Raise 361  'Can't load or unload this object
        End If
    Else
        Err.Raise 327 ' Data value named not found
    End If
    
ErroCls:
    If Err.Number = 0 Then
        WinRegAdicionarSequência = True
        mvarMsgErro = ""
    Else
        mvarMsgErro = BRM_MSG1 & Str(Err.Number) & " - " & "clRegedit.WinRegAdicionarSequência - " & Err.Description
    End If
    rtn = RegCloseKey(hKey)
End Function






