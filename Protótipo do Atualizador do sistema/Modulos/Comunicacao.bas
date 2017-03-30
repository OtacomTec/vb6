Attribute VB_Name = "Comunicacao"
Private Declare Function NetMessageBufferSend Lib "NETAPI32.DLL" (yServer As Any, yToName As Byte, yFromName As Any, yMsg As Byte, ByVal lSize As Long) As Long
Private Const NERR_Success As Long = 0&
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
  "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, _
  ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long

Private Declare Function WNetEnumResource Lib "mpr.dll" Alias _
  "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, _
  ByVal lpBuffer As Long, lpBufferSize As Long) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" _
   (ByVal hEnum As Long) As Long

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" _
  (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias _
  "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, _
   ByVal cbCopy As Long)
   
Private Declare Function CopyPointer2String Lib _
  "kernel32" Alias "lstrcpyA" (ByVal NewString As _
  String, ByVal OldString As Long) As Long

Public Function DoNetEnum(list As Object) As Boolean
Dim hEnum As Long, lpBuff As Long, NR As NETRESOURCE
Dim cbBuff As Long, cCount As Long
Dim P As Long, res As Long, i As Long

On Error Resume Next
If Err.Number > 0 Then Exit Function

On Error GoTo ErrorHandler

NR.lpRemoteName = 0

cbBuff = 1024 * 31
cCount = &HFFFFFFFF

res = WNetOpenEnum(RESOURCE_GLOBALNET, _
  RESOURCETYPE_ANY, 0, NR, hEnum)

If res = 0 Then

   lpBuff = GlobalAlloc(GPTR, cbBuff)

   res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
   If res = 0 Then
      P = lpBuff
      For i = 1 To cCount
         CopyMemory NR, ByVal P, LenB(NR)
         list.AddItem PointerToString(NR.lpRemoteName)
         DoNetEnum2 NR, list, "."
         P = P + LenB(NR)
      Next i
      End If
DoNetEnum = True

ErrorHandler:
On Error Resume Next
   If lpBuff <> 0 Then GlobalFree (lpBuff)
   WNetCloseEnum (hEnum)

End If

End Function

Private Function PointerToString(P As Long) As String

   Dim s As String
   s = String(65535, Chr$(0))
   CopyPointer2String s, P
   PointerToString = Left(s, InStr(s, Chr$(0)) - 1)

End Function

Private Sub DoNetEnum2(NR As NETRESOURCE, list As Object, pts As String)

   Dim hEnum As Long, lpBuff As Long
   Dim cbBuff As Long, cCount As Long
   Dim P As Long, res As Long, i As Long

   cbBuff = 1024 * 31
   cCount = &HFFFFFFFF

   res = WNetOpenEnum(RESOURCE_GLOBALNET, _
     RESOURCETYPE_ANY, 0, NR, hEnum)
   If res = 0 Then

      lpBuff = GlobalAlloc(GPTR, cbBuff)
      res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
      If res = 0 Then
         P = lpBuff
         For i = 1 To cCount
            CopyMemory NR, ByVal P, LenB(NR)
            Dim st As String
            Select Case NR.dwDisplayType
                Case &H1
                    st = "Domain"
                Case &H2
                    st = "Server"
                Case &H3
                    st = "Share"
                Case &H4
                    st = "File"
                Case &H5
                    st = "Groups"
                Case &H6
                    st = "Protocol Categories"
            End Select
            list.AddItem pts & Replace(PointerToString(NR.lpRemoteName), "\", "") '& " is  a : " & st
            DoEvents
            If Not NR.dwDisplayType = 2 Then DoNetEnum2 NR, list, pts & "."
            P = P + LenB(NR)
         Next i
      End If

      If lpBuff <> 0 Then GlobalFree (lpBuff)
      WNetCloseEnum (hEnum)

   End If

End Sub

Public Function BroadcastMessage(sToUser As String, sFromUser As String, sMessage As String) As Boolean
  
    Dim yToName() As Byte
    Dim yFromName() As Byte
    Dim yMsg() As Byte
    Dim l As Long
 
    yToName = sToUser & vbNullChar
    yFromName = sFromUser & vbNullChar
    yMsg = sMessage & vbNullChar

    If NetMessageBufferSend(ByVal 0&, yToName(0), ByVal 0&, yMsg(0), UBound(yMsg)) = NERR_Success Then
       BroadcastMessage = True
    End If

End Function
Public Function NameOfPC(MachineName As String) As Long
    
    Dim NameSize As Long
    Dim X As Long

    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    
    'retorna para o text box o nome do computador
    'TextBox.Text = MachineName
    
End Function
