VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'###################################################################################
'
' Rede
' Classe para efetuar um ping na rede
'
' Por : Alexandre Liondas Neto
'
'###################################################################################

Option Explicit

'Constantes para api Ping
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const PING_TIMEOUT = 200
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

'Variaveis tipadas para api Ping
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type
Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

'Apis para o Ping
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'Variaveis para api Ping
Dim ICMPOPT As ICMP_OPTIONS
Dim ECHO    As ICMP_ECHO_REPLY

'PROPRIEDADES ######################################################################

'Nome do computador
Public Property Get ComputadorNome() As String
   ComputadorNome = NomeComputador
End Property

'Retorna o endereço MAC da placa de rede
Public Property Get RedeMAC(ByVal NumeroAdaptador As Long) As String
   RedeMAC = EthernetAddress(NumeroAdaptador)
End Property

'Retorna o endereço IP da placa de rede
Public Property Get RedeIP() As String
   RedeIP = GetIPAddress
End Property

'Cria a propriedade Address do ping
Property Get ECHO_Address() As String
   ECHO_Address = ECHO.Address
End Property

'Cria a propriedade Status do ping
Property Get ECHO_Status() As String
   ECHO_Status = ECHO.Status
   'GetStatusCode(ECHO.status)
End Property

'Cria a propriedade RoundTripTime ms do ping
Property Get ECHO_RoundTripTime() As String
   ECHO_RoundTripTime = ECHO.RoundTripTime
End Property

'Cria a propriedade DataSize bytes do ping
Property Get ECHO_DataSize() As String
   ECHO_DataSize = ECHO.DataSize
End Property

'Cria a propriedade Data do ping
Property Get ECHO_Data() As String
   Dim Pos As Long
   ECHO_Data = ""
   If Left$(ECHO.Data, 1) <> Chr$(0) Then
      Pos = InStr(ECHO.Data, Chr$(0))
      ECHO_Data = Left$(ECHO.Data, Pos - 1)
   End If
End Property

'Cria a propriedade DataPointer do ping
Property Get ECHO_DataPointer() As String
   ECHO_DataPointer = ECHO.DataPointer
End Property

'METODOS ###########################################################################

'Funcao para retornar a mensagem de status do ping
Public Function GetStatusCode(Status As Long) As String
   Dim msg As String
   Select Case Status
      Case IP_SUCCESS:               msg = "ip success"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   GetStatusCode = CStr(Status) & "   [ " & msg & " ]"
End Function

'Funcao para efetuar o ping
Public Function Ping(szAddress As String) As Long
   Dim hPort       As Long
   Dim dwAddress   As Long
   Dim sDataToSend As String
   Dim iOpt        As Long
   sDataToSend = "Echo This"
   dwAddress = AddressStringToLong(szAddress)
   Call SocketsInitialize
   hPort = IcmpCreateFile()
   If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), PING_TIMEOUT) Then
      Ping = ECHO.RoundTripTime
   Else
      Ping = ECHO.Status * -1
   End If
   Call IcmpCloseHandle(hPort)
   Call SocketsCleanup
End Function

'FUNCOES AUXILIARES ################################################################

Private Function HiByte(ByVal wParam As Integer)
   HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
   LoByte = wParam And &HFF&
End Function

Function AddressStringToLong(ByVal tmp As String) As Long
   Dim i As Integer
   Dim parts(1 To 4) As String
   i = 0
   While InStr(tmp, ".") > 0
      i = i + 1
      parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
      tmp = Mid(tmp, InStr(tmp, ".") + 1)
   Wend
   i = i + 1
   parts(i) = tmp
   If i <> 4 Then
      AddressStringToLong = 0
      Exit Function
   End If
   AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
End Function

Private Function SocketsCleanup() As Boolean
   Dim x As Long
   x = WSACleanup()
   If x <> 0 Then
      MsgBox "Windows Sockets error " & Trim$(Str$(x)) & " occurred in Cleanup.", vbExclamation
      SocketsCleanup = False
   Else
      SocketsCleanup = True
   End If
End Function

Private Function SocketsInitialize() As Boolean
   Dim WSAD As WSADATA
   Dim x As Integer
   Dim szLoByte As String, szHiByte As String, szBuf As String
   x = WSAStartup(WS_VERSION_REQD, WSAD)
   If x <> 0 Then
      MsgBox "Windows Sockets for 32 bit Windows " & "environments is not successfully responding."
      SocketsInitialize = False
      Exit Function
   End If
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      szHiByte = Trim$(Str$(HiByte(WSAD.wVersion)))
      szLoByte = Trim$(Str$(LoByte(WSAD.wVersion)))
      szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
      szBuf = szBuf & " is not supported by Windows " & "Sockets for 32 bit Windows environments."
      MsgBox szBuf, vbExclamation
      SocketsInitialize = False
      Exit Function
   End If
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
      szBuf = "This application requires a minimum of " & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
      MsgBox szBuf, vbExclamation
      SocketsInitialize = False
      Exit Function
   End If
   SocketsInitialize = True
End Function

