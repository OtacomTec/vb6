VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DeskIn 
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   PropertyPages   =   "DeskIn.ctx":0000
   ScaleHeight     =   4770
   ScaleWidth      =   5190
   ToolboxBitmap   =   "DeskIn.ctx":0011
   Begin VB.Timer timState 
      Interval        =   1000
      Left            =   3825
      Top             =   4635
   End
   Begin VB.PictureBox PICDESKTOP 
      Height          =   3915
      Left            =   60
      ScaleHeight     =   3855
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      Begin MSComctlLib.ImageList imgConnect 
         Left            =   4380
         Top             =   3240
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
               Picture         =   "DeskIn.ctx":0323
               Key             =   "OFF"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DeskIn.ctx":06DB
               Key             =   "ON"
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtLog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4230
      Width           =   5115
   End
   Begin VB.Timer timBPS 
      Interval        =   1000
      Left            =   4260
      Top             =   4620
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   4695
      Top             =   4620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4515
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   413
            MinWidth        =   413
            Key             =   "icon"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0.0.0.0"
            TextSave        =   "0.0.0.0"
            Key             =   "ip"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1196
            MinWidth        =   1196
            Text            =   "0"
            TextSave        =   "0"
            Key             =   "port"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4763
            Key             =   "state"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTransfered 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / 0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   3930
      Width           =   1095
   End
   Begin VB.Label lblBPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / bps"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2670
      TabIndex        =   3
      Top             =   3930
      Width           =   1305
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuManual 
         Caption         =   "Manual Mode"
      End
      Begin VB.Menu mnuAutomatic 
         Caption         =   "Automatic Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetScreen 
         Caption         =   "Get Screen"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "DeskIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event SocketError(ErrorDescription As String)
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarRemoteIP As Variant                                     ' Remote IP
Private mvarRemotePort As Variant                                   ' Remote Port
Private mvarsWidth As Long                                          ' Image Width
Private mvarsHeight As Long                                         ' Image Height
Private mvarsFile As String                                         ' File Name
Private mvarsPacketSize As Long                                     ' Packet Size
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private FileNumber As Long                                          ' Holds File Handle
Private ByteSecond As Long                                          ' Counter For BPS
Private FileSize As Long                                            ' File Size
Private FileDone As Long                                            ' File Transfered
Private FileComplete As Boolean                                     ' All done?
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


'==============================
'   Terminate Control
'==============================
Private Sub UserControl_Terminate()
    Socket.Close
End Sub


'==============================
'   Desktop Image Click
'==============================
Private Sub PicDesktop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' On click, popup Menu.
    If Button = 2 Then PopupMenu mnuOptions
End Sub


'==============================
'   Menu Items
'========================================================================
' Three Options:
' (*) Automatic     - Continually transfer image.
' ( ) Manual        - Only grab an image when Grab Image is selected.
' ( ) Grab Image    - Used for Manual Mode.
'========================================================================
Private Sub mnuAutomatic_click()
    mnuAutomatic.Checked = True
    mnuManual.Checked = False
    mnuGetScreen.Enabled = False
    
    If FileComplete Then Call mnuGetScreen_click
End Sub
Private Sub mnuManual_click()
    mnuAutomatic.Checked = False
    mnuManual.Checked = True
    mnuGetScreen.Enabled = True
End Sub
Private Sub mnuGetScreen_click()
    If Socket.State = 7 Then _
    Socket.SendData "010" & NewCom & sWidth & NewCom & sHeight & NewCom & PacketSize & NewCom
    mnuGetScreen.Enabled = False
End Sub


'==============================
'   Control Resize Event
'==============================
Private Sub UserControl_Resize()

    ' Here we simply resize all controls
    ' on the control (or form). Important
    ' stuff below...

    PICDESKTOP.Width = UserControl.Width - 210
    PICDESKTOP.Height = UserControl.Height - 900
    
    ProgressBar.Width = UserControl.Width / 2
    ProgressBar.Top = UserControl.Height - 850
    
    txtLog.Width = UserControl.Width - 190
    txtLog.Top = UserControl.Height - 550
    
    lblBPS.Width = UserControl.Width / 4
    lblBPS.Top = UserControl.Height - 850
    lblBPS.Left = UserControl.Width / 4 * 2
    
    lblTransfered.Width = UserControl.Width / 4 - 190
    lblTransfered.Top = UserControl.Height - 850
    lblTransfered.Left = UserControl.Width / 4 * 3
    
    ' Now. Whatever we have changed our viewing size to, aka our
    ' form size, we should tell the remote to send a file of that size.
    ' No sense in sending something larger than we can display or smaller
    ' than we are looking at.
    sWidth = PICDESKTOP.Width
    sHeight = PICDESKTOP.Height
    ' Set sWidth and sHeight = our current picturebox size!
End Sub


'==============================
'   Winsock State
'==============================
Public Function WinsockState()
    ' Return socket state.
    WinsockState = Socket.State
End Function


'==============================
'   Disconnect Winsock
'==============================
Public Function Disconnect()
    Socket.Close
    txtLog.Text = "Disconnected."
End Function


'==============================
'   Connect
'==============================
Public Sub Connect()

    txtLog.Text = "Connecting to " & RemoteIP & ":" & RemotePort
    Socket.Close
    Socket.Connect RemoteIP, RemotePort

End Sub


'==============================
'   Error Event
'==============================
Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent SocketError("Error " & Number & " (" & Description & ")")
    txtLog.Text = "Error " & Number & " (" & Description & ")"
End Sub


'==============================
'   Data Arrival
'==============================
' Data arrives in a special form. The packet structure: [COM][DATA]
' COM is a three part command we use to figure out what to do with DATA.
' COM will look like 010, 020, 030, etc. This is separated from the rest.

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim iPacket As String
Dim iCom As String
On Error GoTo ErrSub

    ' Parse Packet
    Socket.GetData iPacket                  '- Pull Data from socket.
    ByteSecond = ByteSecond + Len(iPacket)  '- Calculate Rate of Transfer

    ' Now we separate [COM] from [DATA] and give each their own variable.
    iCom = Word(iPacket, 1, NewCom)                         ' - Parse [COM]
    iPacket = Right(iPacket, Len(iPacket) - Len(iCom) - 1)  ' - Parse [DATA]

    ' Decide what to do with [DATA] or "iPacket"
    ' (010) - Send Information
    ' (020) - Ready File (Open, Initialize...)
    ' (030) - Write Data to File
    ' (040) - Close File. View File. Begin Process over (unless MANUAL mode selected).
    Select Case iCom
    Case "010": FileComplete = False: Socket.SendData "010" & NewCom & sWidth & NewCom & sHeight & NewCom & PacketSize & NewCom
    Case "020": ReadyFile (iPacket)
    Case "030": WriteFileData (iPacket)
    Case "040": Close #FileNumber
                PICDESKTOP.Picture = LoadPicture(sFile)
                txtLog.Text = "Transfer Complete " & Time
                FileComplete = True
                If mnuAutomatic.Checked = True Then _
                Socket.SendData "010" & NewCom & sWidth & NewCom & sHeight & NewCom & PacketSize & NewCom _
                Else: mnuGetScreen.Enabled = True
    End Select

Exit Sub
ErrSub:
Select Case Err.Number
    Case 5: ' This error occurs when too large a packet is received, or packets
            ' are being sent too fast. The packet is broken from one whole, and
            ' seperated:
            ' PACKET: "010-XXXXXXXXXXXXXXXXX..." Turns into two packets:
            ' PACKET: "010-XXXXXXXXX..." & "XXXXXXXX..."
            ' We lose the command line, and problems occur. Fortunately, the
            ' image format is very forgiving, and we can just 'POP' these broken
            ' packets into the file, note that we assume the broken packet is part
            ' of a file transfer string because the only large packets we send are these..
            RaiseEvent SocketError("Collision (Appending Media)")
            Put FileNumber, , iPacket
            FileDone = FileDone + Len(iPacket)
            Exit Sub
            ' I would not recommend this method sending regular files, because
            ' after we receive a packet we tell the remote to send more, and if
            ' we are receiving a packet that is broken, we have no way of knowing
            ' it is broken until we receive an error, therefore we always tell the
            ' remote to send more, and if we are telling the remote to send more
            ' while we have strings waiting in the buffer, it floods the buffer...
            ' and things may be written in the wrong order....
    Case Else
            RaiseEvent SocketError("Error " & Err.Number & " (" & Err.Description & ")")
End Select
End Sub


'==============================
'   Write File (Disk)
'==============================
Private Sub WriteFileData(iData As String)

    Put FileNumber, , iData

    FileDone = FileDone + Len(iData)
    lblTransfered = (CInt(FileDone / FileSize * 100) & "% " & FormatBytes(FileDone) & "/" & FormatBytes(FileSize))

    On Error Resume Next
    ProgressBar.Value = CInt(FileDone / FileSize * 100)

    Socket.SendData "999" & NewCom

End Sub


'==============================
'   Ready File (Open, Create)
'==============================
Private Sub ReadyFile(iData As String)

    ' Set Packet Data
    FileSize = Word(iData, 1, NewCom)

    ' Raise Event Transfered
    FileDone = 0
    lblTransfered = 0 & "% " & FormatBytes(FileDone) & "/" & FormatBytes(FileSize)
    ProgressBar.Value = 0

    ' File Operations
    FileNumber = FreeFile
    Open sFile For Binary Access Write As #FileNumber
    
    ' Tell remote send next packet.
    Socket.SendData "020" & NewCom

End Sub


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   CLASS OBJECT PROPERTY FOLLOWS:
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Let RemotePort(ByVal vData As Variant)
'Syntax: X.RemotePort = 5
    mvarRemotePort = vData
End Property
Public Property Set RemotePort(ByVal vData As Variant)
'Syntax: Set x.RemotePort = Form1
    Set mvarRemotePort = vData
End Property
Public Property Get RemotePort() As Variant
Attribute RemotePort.VB_ProcData.VB_Invoke_Property = "SocketPage"
'Syntax: Debug.Print X.RemotePort
    If IsObject(mvarRemotePort) Then
        Set RemotePort = mvarRemotePort
    Else
        RemotePort = mvarRemotePort
    End If
End Property
Public Property Let RemoteIP(ByVal vData As Variant)
'Syntax: X.RemoteIP = 5
    mvarRemoteIP = vData
End Property
Public Property Set RemoteIP(ByVal vData As Variant)
'Syntax: Set x.RemoteIP = Form1
    Set mvarRemoteIP = vData
End Property
Public Property Get RemoteIP() As Variant
Attribute RemoteIP.VB_ProcData.VB_Invoke_Property = "SocketPage"
'Syntax: Debug.Print X.RemoteIP
    If IsObject(mvarRemoteIP) Then
        Set RemoteIP = mvarRemoteIP
    Else
        RemoteIP = mvarRemoteIP
    End If
End Property
Public Property Let sHeight(ByVal vData As Long)
'Syntax: X.sHeight = 5
    mvarsHeight = vData
End Property
Public Property Get sHeight() As Long
Attribute sHeight.VB_ProcData.VB_Invoke_Property = "SocketPage"
'Syntax: Debug.Print X.sHeight
    sHeight = mvarsHeight
End Property
Public Property Let sWidth(ByVal vData As Long)
'Syntax: X.sWidth = 5
    mvarsWidth = vData
End Property
Public Property Get sWidth() As Long
Attribute sWidth.VB_ProcData.VB_Invoke_Property = "SocketPage"
'Syntax: Debug.Print X.sWidth
    sWidth = mvarsWidth
End Property
Public Property Let sFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFile = 5
    mvarsFile = vData
End Property
Public Property Get sFile() As String
Attribute sFile.VB_ProcData.VB_Invoke_Property = "SocketPage"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFile
    If mvarsFile = "" Then mvarsFile = TempDir & "\SC.MEM"
    sFile = mvarsFile
End Property
Private Sub timBPS_Timer()
    lblBPS = FormatBytes(ByteSecond) & "s"
    ByteSecond = 0
End Sub
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
