Attribute VB_Name = "modFunctions"
Option Explicit
'========================================================================
'   Class:          VBDeskview          | (C) 2001 Michael A. Schmidt   |
'   Author:         Mike Schmidt        =================================
'   Date:           October 2001
'   E-mail:         mikes@mtdmarketing.com
'========================================================================
'   References:     MSWINSCK.OCX    Microsoft Winsock
'                   SCRRUN.DLL      Microsoft Scripting Runtime (FSO)
'========================================================================
'========================================================================
' How this works:

' [DeskIN]  - .LocalPort = 1005  [DeskIN]   - .Listen
' DeskIN will now listen for an incoming connection on port 1005.
'
' [DeskOUT] - .RemotePort = 1005 [DeskOUT]  - .RemoteIP = 127.0.0.1
' [DeskOUT] - .PacketSize = 2048
' DeskOUT will now connect to localhost 127.0.0.1 on port 1005. Packet
' size is set to 2048. Try larger for LAN, smaller for slow machines,
' errors, or dialup.
'
' Once both have connected, DeskIN will send information regarding the
' size screenshot it wants back. This size is taken from PICDESKTOP. However
' large PICDESKTOP is, that size file will be sent back.
'
' Once DeskIN sends this information, DeskOUT will compose the picture
' and send. DeskOUT will receive and view the picture. All picture transfer
' files are stored within the Windows Temporary Directory.

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Const NewCom As String = "§"


'====================================
'   TempDir
'====================================
' We use this function to grab the temporary
' directory used by Windows. This is where we
' will store files being transfered.
'====================================
Function TempDir$()
    Dim sTmp1$
    sTmp1 = Environ$("temp")

    While Right$(sTmp1, 1) = "\"
        sTmp1 = Left$(sTmp1, Len(sTmp1) - 1)
    Wend

    If sTmp1 <> "" Then
        On Error Resume Next
        MkDir sTmp1
        On Error GoTo 0
    End If
    TempDir = sTmp1
End Function


'====================================
'   GrabFilename
'====================================
Public Function GrabFilename(FullPath As String)

    ' Pulls the filename from a full path and filename.
    ' Returns filename.
    GrabFilename = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))

End Function


'====================================
'   isValidFile
'====================================
Public Function isValidFile(ByVal iFile As String)
Dim FSO As New FileSystemObject

    ' Check to see if file exists.
    ' Return Boolean.
    isValidFile = FSO.FileExists(iFile)

End Function


'====================================
'   GetFileSize
'====================================
Public Function GetFileSize(ByVal iFile As String)
On Error GoTo ErrSub
Dim FSO As New FileSystemObject
Dim FSOfile As File

    ' Get Size of File.
    ' Return File Size.
    Set FSOfile = FSO.GetFile(iFile)
    GetFileSize = FSOfile.Size

Exit Function
ErrSub:
    MsgBox "Fatal Error: Please update SCRRUN.DLL to 5.1.0.5010 or later.", vbCritical, "OCX"
    GetFileSize = GetOpenFileSize(iFile)
End Function


'====================================
'   GetOpenFileSize
'====================================
Public Function GetOpenFileSize(ByVal iFile As String)
Dim FileNumber

    ' - Open files return a different filesize
    ' - than just reading off the disk.
    
    ' - First grab an open filehandle (random = freefile)
    ' - Then open the file, grab the size (LOF) and
    ' - close the file.
    FileNumber = FreeFile
    Open iFile For Binary Access Read As #FileNumber
    GetOpenFileSize = LOF(FileNumber)
    Close #FileNumber


End Function


'====================================
'   FormatBytes
'====================================
Public Function FormatBytes(iBytes As Long) As String

    If iBytes < 1024 Then
        FormatBytes = iBytes & " b"
    ElseIf iBytes < 1048576 Then
        FormatBytes = Format(iBytes / 1024, "0.0") & " kb"
    Else 'If iBits < 1000000000 Then
        FormatBytes = Format(iBytes / 1048576, "0.00") & " mb"
    End If

End Function


'====================================
'   Take Screenshot
'====================================
Public Sub TakeScreenshot(ByVal sWidth As Integer, ByVal sHeight As Integer, FileName As String)
On Error GoTo ErrSub
Dim wScreen As Long
Dim hScreen As Long
Dim w As Long
Dim h As Long
Dim r
Dim hdcScreen
frmDesk.PICSCREEN.Cls

    ' Set Width of picture to grab.
    frmDesk.PICSCREEN.Width = sWidth
    frmDesk.PICSCREEN.Height = sHeight

    ' Calculate Width & Height
    wScreen = Screen.Width \ Screen.TwipsPerPixelX
    hScreen = Screen.Height \ Screen.TwipsPerPixelY

    ' Set Scale Mode
    frmDesk.PICSCREEN.ScaleMode = vbPixels
    w = frmDesk.PICSCREEN.ScaleWidth
    h = frmDesk.PICSCREEN.ScaleHeight

    ' Grab Desktop Screen.
    hdcScreen = GetDC(0)

    ' Stretch it to fit our picture box. (THIS line creates the momentary
    ' jerk in Windows).
    r = StretchBlt(frmDesk.PICSCREEN.hdc, 0, 0, w, h, hdcScreen, 0, 0, wScreen, hScreen, vbSrcCopy)
    
    ' Save to disk. Ready for display or transfer.
    SavePicture frmDesk.PICSCREEN.Image, FileName

Exit Sub
ErrSub:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "TakeScreenshot SUB"
End Sub


'====================================
'   Word Function
'====================================
Public Function Word(ByVal sSource As String, n As Long, SP As String) As String
' This function is used to parse data. Data is send as
' multiple commands in one packet. Each command is seperated
' by a special character. We retrieve specific commands by
' calling 'word' and specifying what seperates each 'word'.
'=================================================
' Word retrieves the nth word from sSource
' Usage:
'    Word("red blue green ", 2)   "blue"
'=================================================
Dim pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim x       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

'sSource = CSpace(sSource)

'find the nth word
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
  
End Function
Public Function SocketState(numState As Integer)
'%%%% This function returns the text-state of
'%%%% the socket, when given the numeric-state.

    Select Case numState
    Case 0: SocketState = "Closed."
    Case 1: SocketState = "Open."
    Case 2: SocketState = "Listening."
    Case 3: SocketState = "Connection Pending."
    Case 4: SocketState = "Resolving Host."
    Case 5: SocketState = "Host Resolved."
    Case 6: SocketState = "Connecting."
    Case 7: SocketState = "Connected."
    Case 8: SocketState = "Peer Closing."
    Case 9: SocketState = "Error."
    End Select
    
End Function
