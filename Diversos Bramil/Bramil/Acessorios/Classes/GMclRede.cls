VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclRede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Const NO_ERROR = 0
Const CONNECT_UPDATE_PROFILE = &H1
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_ANY = &H0
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_GLOBALNET = &H2
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
                         (lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
                         ByVal lpUserName As String, ByVal dwFlags As Long) As Long

Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" _
                         (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long


Public Function Mapear(LetraDaUnidade As String, PastaRemota As String) As Boolean
    Dim NetR As NETRESOURCE
    Dim ErrInfo As Long
    LetraDaUnidade = MontaUnidade(LetraDaUnidade)
    If LetraDaUnidade = "" Then Exit Function
            
    'LetraUnidade = Trim(LetraUnidade)
    'If Len(LetraUnidade) > 2 Then Exit Sub
    'If Len(LetraUnidade) = 0 Then Exit Sub
    'If Len(LetraUnidade) = 1 Then LetraUnidade = LetraUnidade & ":"

    NetR.dwScope = RESOURCE_GLOBALNET
    NetR.dwType = RESOURCETYPE_DISK
    NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    NetR.lpLocalName = LetraDaUnidade
    NetR.lpRemoteName = PastaRemota
    ErrInfo = WNetAddConnection2(NetR, Senha, Usuário, CONNECT_UPDATE_PROFILE)

    If ErrInfo = NO_ERROR Then
        Mapear = True
    Else
        
    End If
End Function

Public Function DesMapear(LetraDaUnidade As String) As Boolean
    Dim ret As Long
    LetraDaUnidade = MontaUnidade(LetraDaUnidade)
    If LetraDaUnidade = "" Then Exit Function

    ret = WNetCancelConnection2(LetraDaUnidade, CONNECT_UPDATE_PROFILE, False)
    If ret = NO_ERROR Then
        DesMapear = True
    Else
        'vai entrar string de erro em propriedade de erro
    End If

End Function


Private Function MontaUnidade(strLetraUnidade As String)
    LetraUnidade = Trim(strLetraUnidade)
    If Len(strLetraUnidade) > 2 Then Exit Function
    If Len(strLetraUnidade) = 0 Then Exit Function
    If Len(strLetraUnidade) = 1 Then strLetraUnidade = strLetraUnidade & ":"
    MontaUnidade = strLetraUnidade
End Function

