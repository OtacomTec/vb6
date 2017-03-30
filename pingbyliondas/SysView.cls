VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###################################################################################
'
' System View
' Classe para retornar informacoes sobre a conexao de rede ethernet de um computador
'
' Por : Alexandre Liondas Neto
'
'###################################################################################

Option Explicit

'API ###############################################################################

Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" ( _
        ByVal lpBuffer As String, _
        nSize As Long) As Long
     
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
        ByVal lpBuffer As String, _
        nSize As Long) As Long

Private Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "KERNEL32" () As Long
Private Declare Function HeapAlloc Lib "KERNEL32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "KERNEL32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Const NCBASTAT = &H33
Private Const NCBNAMSZ = 16
Private Const HEAP_ZERO_MEMORY = &H8
Private Const HEAP_GENERATE_EXCEPTIONS = &H4
Private Const NCBRESET = &H32
Private Type NCB
  ncb_command As Byte
  ncb_retcode As Byte
  ncb_lsn As Byte
  ncb_num As Byte
  ncb_buffer As Long
  ncb_length As Integer
  ncb_callname As String * NCBNAMSZ
  ncb_name As String * NCBNAMSZ
  ncb_rto As Byte
  ncb_sto As Byte
  ncb_post As Long
  ncb_lana_num As Byte
  ncb_cmd_cplt As Byte
  ncb_reserve(9) As Byte ' Reserved, must be 0
  ncb_event As Long
End Type
Private Type ADAPTER_STATUS
  adapter_address(5) As Byte
  rev_major As Byte
  reserved0 As Byte
  adapter_type As Byte
  rev_minor As Byte
  duration As Integer
  frmr_recv As Integer
  frmr_xmit As Integer
  iframe_recv_err As Integer
  xmit_aborts As Integer
  xmit_success As Long
  recv_success As Long
  iframe_xmit_err As Integer
  recv_buff_unavail As Integer
  t1_timeouts As Integer
  ti_timeouts As Integer
  Reserved1 As Long
  free_ncbs As Integer
  max_cfg_ncbs As Integer
  max_ncbs As Integer
  xmit_buf_unavail As Integer
  max_dgram_size As Integer
  pending_sess As Integer
  max_cfg_sess As Integer
  max_sess As Integer
  max_sess_pkt_size As Integer
  name_count As Integer
End Type
Private Type NAME_BUFFER
  name As String * NCBNAMSZ
  name_num As Integer
  name_flags As Integer
End Type
Private Type ASTAT
  adapt As ADAPTER_STATUS
  NameBuff(30) As NAME_BUFFER
End Type

Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" (ByVal szHost As String) As Long
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus   As Long = 128
Private Const ERROR_SUCCESS      As Long = 0
Private Const WS_VERSION_REQD    As Long = &H101
Private Const WS_VERSION_MAJOR   As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR   As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD   As Long = 1
Private Const SOCKET_ERROR       As Long = -1
Private Type HOSTENT
   hName     As Long
   hAliases  As Long
   hAddrType As Integer
   hLen      As Integer
   hAddrList As Long
End Type
Private Type WSADATA
   wVersion     As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)  As Byte
   wMaxSockets  As Integer
   wMaxUDPDG    As Integer
   dwVendorInfo As Long
End Type

'PROPRIEDADES #######################################################################

'Nome do computador
Public Property Get ComputadorNome() As String
   ComputadorNome = NomeComputador
End Property

'Usuário do sistema
Public Property Get ComputadorUsuario() As String
   ComputadorUsuario = WinUsuario
End Property

'Retorna o endereço MAC da placa de rede
Public Property Get RedeMAC(ByVal NumeroAdaptador As Long) As String
   RedeMAC = EthernetAddress(NumeroAdaptador)
End Property

'Retorna o endereço IP da placa de rede
Public Property Get RedeIP() As String
   RedeIP = GetIPAddress
End Property

'FUNCOES AUXILIARES ################################################################

'Nome do computador
Private Function NomeComputador() As String
  Dim sBuffer As String, lBufSize As Long
  lBufSize = 255
  sBuffer = String$(lBufSize, " ")
  NomeComputador = IIf(GetComputerName(sBuffer, lBufSize) = 0, "", Left(sBuffer, lBufSize))
End Function

'Nome do usuário no Windows
Private Function WinUsuario() As String
   Dim Bufstr As String
   Bufstr = Space$(50)
   WinUsuario = IIf(GetUserName(Bufstr, 50) = 0, "", Left(Trim(Bufstr), Len(Trim(Bufstr)) - 1))
End Function

'Localiza o endereço MAC
Private Function EthernetAddress(LanaNumber As Long) As String
  Dim udtNCB       As NCB
  Dim bytResponse  As Byte
  Dim udtASTAT     As ASTAT
  Dim udtTempASTAT As ASTAT
  Dim lngASTAT     As Long
  Dim strOut       As String
  Dim x            As Integer
  udtNCB.ncb_command = NCBRESET
  bytResponse = Netbios(udtNCB)
  udtNCB.ncb_command = NCBASTAT
  udtNCB.ncb_lana_num = LanaNumber
  udtNCB.ncb_callname = "* "
  udtNCB.ncb_length = Len(udtASTAT)
  lngASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, udtNCB.ncb_length)
  strOut = ""
  If lngASTAT Then
     udtNCB.ncb_buffer = lngASTAT
     bytResponse = Netbios(udtNCB)
     CopyMemory udtASTAT, udtNCB.ncb_buffer, Len(udtASTAT)
     With udtASTAT.adapt
        For x = 0 To 5
           strOut = strOut & Right$("00" & Hex$(.adapter_address(x)), 2)
        Next x
     End With
     HeapFree GetProcessHeap(), 0, lngASTAT
  End If
  If Len(strOut) >= 2 Then
     For x = 1 To Len(strOut) Step 2
        EthernetAddress = EthernetAddress & Mid(strOut, x, 2) & "-"
     Next x
     If Right(EthernetAddress, 1) = "-" Then EthernetAddress = Left(EthernetAddress, Len(EthernetAddress) - 1)
   End If
End Function

'Retorna o endereço IP
Private Function GetIPAddress() As String
   Dim sHostName   As String * 256
   Dim lpHost      As Long
   Dim HOST        As HOSTENT
   Dim dwIPAddr    As Long
   Dim tmpIPAddr() As Byte
   Dim i           As Integer
   Dim sIPAddr     As String
   If Not SocketsInitialize() Then Exit Function
   'Ocorre um erro de Socket : " & Str$(WSAGetLastError()) & " , não é possivel obter nome do Host."
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      SocketsCleanup
      Exit Function
   End If
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
   'MsgBox "Windows Sockets não esta respondendo. " & "não é possivel obter nome do Host"
   If lpHost = 0 Then
     SocketsCleanup
     Exit Function
   End If
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   For i = 1 To HOST.hLen
     sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   SocketsCleanup
End Function
Private Function HiByte(ByVal wParam As Integer) As Byte
   HiByte = (wParam And &HFF00&) \ (&H100)
End Function
Private Function LoByte(ByVal wParam As Integer) As Byte
   LoByte = wParam And &HFF&
End Function
Private Function SocketsInitialize() As Boolean
   Dim WSAD As WSADATA
   SocketsInitialize = False
   '32-bit Windows Socket não esta respondendo.
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then Exit Function
   'Esta aplicação requer um minimo de " & CStr(MIN_SOCKETS_REQD) & " sockets suportados.
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then Exit Function
   'A versao Sockets " & CStr(LoByte(WSAD.wVersion)) & "." & CStr(HiByte(WSAD.wVersion)) & " não é suportada por 32-bit Windows Sockets.
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then Exit Function
   SocketsInitialize = True
End Function
Private Sub SocketsCleanup()
   'If WSACleanup() <> ERROR_SUCCESS Then MsgBox " Erro de Socket."
   WSACleanup
End Sub

