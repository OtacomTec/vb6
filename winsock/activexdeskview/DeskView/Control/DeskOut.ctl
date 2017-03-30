VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DeskOut 
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ScaleHeight     =   1260
   ScaleWidth      =   4110
   ToolboxBitmap   =   "DeskOut.ctx":0000
   Begin VB.Timer timBPS 
      Interval        =   1000
      Left            =   600
      Top             =   1245
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   180
      Top             =   1245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgConnect 
      Left            =   1035
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DeskOut.ctx":0312
            Key             =   "OFF"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DeskOut.ctx":06CA
            Key             =   "ON"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "DeskOut.ctx":082C
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblTransfered 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / 0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   585
      Width           =   2895
   End
   Begin VB.Label lblBPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / bps"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   885
      Width           =   2895
   End
   Begin VB.Label lblLog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Not Available)"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   135
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Completed:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   585
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   915
      Width           =   510
   End
End
Attribute VB_Name = "DeskOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private Ready As Boolean                    ' Controls if data is sent or not.
Private FileDone As Long                    ' Amount of data sent: 90k of 150k transfered...
Private FileSize As Long                    ' File Size: 90k of 150k transfered...
Private ByteSecond As Long                  ' Counter For BPS
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event DeskError(ErrorDescription As String)
Public Event Disconnected()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarLocalPort As Variant
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarsFile As String
Private mvarsPacketSize As Long


'==============================
'   .Remote IP
'==============================
Public Function RemoteIP()
    RemoteIP = Socket.RemoteHostIP
End Function


'==============================
'   .Listen
'==============================
Public Sub Listen()
    Socket.Close
    Socket.LocalPort = LocalPort
    Socket.Listen
    
    LocalPort = Socket.LocalPort
End Sub


'==============================
'   .Disconnect
'==============================
Public Function Disconnect()
    Socket.Close
End Function


'==============================
'   Resize Control
'==============================
Private Sub UserControl_Resize()
    UserControl.Width = 4110
    UserControl.Height = 1260
End Sub


'==============================
'   Timer (Bytes Second)
'==============================
Private Sub timBPS_Timer()

    ' Calculate Bytes Second
    lblBPS = FormatBytes(ByteSecond) & "s"
    ByteSecond = 0
    lblLog = SocketState(Socket.State)
    
    ' If Socket is disconnected, Raise Event.
    Select Case Socket.State
        Case 8: RaiseEvent Disconnected
        Case 9: RaiseEvent Disconnected
    End Select
    
End Sub


'==============================
'   Connection Request
'==============================
Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
    
    Socket.Close                    '- Must close socket.
    Socket.Accept requestID         '- Accept incoming.
    Socket.SendData "010" & NewCom  '- Data Ready ( Begin's Transfer Sequence )

End Sub


'==============================
'   Data Arrival
'==============================
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrSub
Dim iPacket As String
Dim iCom As String

    'BitsSecond = BitsSecond + bytesTotal

    ' Parse Packet
    Socket.GetData iPacket
    ByteSecond = ByteSecond + Len(iPacket)

    iCom = Word(iPacket, 1, NewCom)
    iPacket = Right(iPacket, Len(iPacket) - Len(iCom) - 1)

    Select Case iCom
    Case "010": Call SendFileStats(Word(iPacket, 1, NewCom), Word(iPacket, 2, NewCom), Word(iPacket, 3, NewCom))
    Case "020": SendFileData
    Case "999": Ready = True
    End Select

Exit Sub
ErrSub:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "DATA ARRIVAL"
End Sub


'==============================
'   .WinsockState
'==============================
Public Function WinsockState()
    WinsockState = Socket.State
End Function


'==============================
'   Send File Data
'==============================
Private Sub SendFileData()
Dim iPacket As String
Dim Buffer As Long
Dim FileNumber
Dim DelLastByte As Boolean
  On Error GoTo ErrSub
    
    FileSize = GetFileSize(sFile)

'    BitsSecond = 0
    FileDone = 0
    ' Buffer size size will be BUFFER - 4bits because every packet
    ' has 4 bits of command line added to it (040XRAWDATA...
    Buffer = PacketSize - 4

    ' - At the end of the file, we usually end up sending
    ' - the EOF character, meaning each sent file gets 1byte
    ' - larger than it should, so now we subtract at the end...
    DelLastByte = False
    ' - Grab Local File Handle (Random = FreeFile)
    ' - Open it. Loop through while not at end of file.
    FileNumber = FreeFile
    Open sFile For Binary Access Read As #FileNumber
    Do While Not EOF(FileNumber)
     
    ' end of file...size buffer appropriately
    If FileSize - Loc(FileNumber) <= Buffer Then
        Buffer = FileSize - Loc(FileNumber) + 1
        DelLastByte = True
    End If

        iPacket = ""
        iPacket = Space$(Buffer)
        Get FileNumber, , iPacket
        
        ' end of file, our packet is 1 bit to large.
        If DelLastByte = True Then iPacket = Left(iPacket, Len(iPacket) - 1)
        
        ' Set bit trackers...
        FileDone = FileDone + Len(iPacket)
        'BitsSecond = BitsSecond + Len(iPacket) + 4
        
        ByteSecond = ByteSecond + Len(iPacket)
        Socket.SendData "030" & NewCom & iPacket
        lblTransfered = CInt(FileDone / FileSize * 100) & ":" & FormatBytes(FileDone) & ":" & FormatBytes(FileSize)

        WaitForRemote

    Loop

    Close #FileNumber
    
    ' Tell Remote Finished!
    Socket.SendData "040" & NewCom
    'RaiseEvent FileComplete
 
 Exit Sub

ErrSub:
Select Case Err.Number
    Case 40006 ' No Connection Detected
        Close #FileNumber

    Case Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Transfer Error"
            Exit Sub 'Resume Next
End Select
End Sub


'==============================
'   Wait For Remote
'==============================
Private Sub WaitForRemote()

    Ready = False
    While Ready = False
    DoEvents
    Wend

End Sub


'==============================
'   Send File Stats
'==============================
Private Sub SendFileStats(iWidth As Long, iHeight As Long, pSize As Long)
On Error GoTo ErrSub
Dim iPacket As String

    ' We are connected.
    TakeScreenshot iWidth, iHeight, sFile
    
    ' Set up packet. Tell remote about file stats.
    iPacket = "020" & NewCom
    iPacket = iPacket & GetFileSize(sFile) & NewCom
    
    ' Send Packet
    Socket.SendData iPacket

Exit Sub
ErrSub:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "SUB SENDFILESTATS"
End Sub


'==============================
'   Error
'==============================
Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent DeskError("Error " & Number & " (" & Description & ")")
End Sub


'==============================
'   Terminate Class
'==============================
Private Sub UserControl_Terminate()
    Socket.Close
End Sub


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   CLASS OBJECT PROPERTY
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Let LocalPort(ByVal vData As Variant)
'Syntax: X.LocalPort = 5
    mvarLocalPort = vData
End Property
Public Property Set LocalPort(ByVal vData As Variant)
'Syntax: Set x.LocalPort = Form1
    Set mvarLocalPort = vData
End Property
Public Property Get LocalPort() As Variant
'Syntax: Debug.Print X.LocalPort
        LocalPort = mvarLocalPort
End Property
Public Property Let sFile(ByVal vData As String)
'Syntax: X.sFile = 5
    mvarsFile = vData
End Property
Public Property Get sFile() As String
'Syntax: Debug.Print X.sFile
    If mvarsFile = "" Then mvarsFile = TempDir & "\SH.MEM"
    sFile = mvarsFile
End Property
Public Property Let PacketSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFile = 5
    mvarsPacketSize = vData
End Property
Public Property Get PacketSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFile
    If mvarsPacketSize = 0 Then mvarsPacketSize = 2048
    PacketSize = mvarsPacketSize
End Property
