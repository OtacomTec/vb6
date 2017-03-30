VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ScreenShot Maker"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeRefresh 
      Caption         =   "Change Refresh"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timRefresh 
      Left            =   120
      Top             =   720
   End
   Begin VB.Label lblRefresh 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "number of screenshots sent"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Current Refresh"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTotalConnects 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Total Number of Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblScrSent 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblConnections 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Current Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Refresh rate on pics
Private Const RefreshRate As Integer = 5000 ' default refresh in ms
Private Const ServerQuality As Integer = 90 ' default quality, 0-100

'Statistics
Private currentConnections As Integer
Private totalConnections As Long
Private numShotsSent As Long

Private keepServerAlive As Boolean

Private theDIB As New cDIBSection
Private scrRight As Long
Private scrBottom As Long
Private lhDC As Long
Private finalPic() As Byte

'used to get the desktop
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'used to get the area of the entire screen
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Sub cmdChangeRefresh_Click()
    MsgBox "Not working yet"
End Sub

Private Sub Form_Load()
    'set up our refresh rate in ms
    timRefresh.Interval = RefreshRate
    lblRefresh.Caption = CStr(RefreshRate)
    
    'set up the first server listening connection
    wskServer(0).LocalPort = 1234
    wskServer(0).Listen
    
    generatePic
    timRefresh.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    keepServerAlive = False
End Sub

Private Sub timRefresh_Timer()
    'If keepServerAlive = False Then
    '    timRefresh.Enabled = False
    '    Exit Sub
    'End If
    generatePic
End Sub

Private Sub wskServer_Close(Index As Integer)
    Debug.Print "The client disconnected"
    
    'see if we should stop the server
    'If Index = 0 And wskServer.Count = 1 Then
    '    keepServerAlive = False
    'End If
    
    'remove the component by unloading it
    If Index <> 0 Then
        Unload wskServer(Index)
    Else
        If wskServer(0).State <> sckClosed Then wskServer(0).Close
        If wskServer.Count = 1 Then
            wskServer(0).LocalPort = 1234
            wskServer(0).Listen
        End If
    End If
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'first accept the connection
    Debug.Print "A client has connected"
    If wskServer(Index).State <> sckClosed Then wskServer(Index).Close
    wskServer(Index).Accept requestID
    
    'send the user their packet
    wskServer(Index).SendData finalPic
    'now create a new winsock object
    Load wskServer(Index + 1)
    
    wskServer(Index + 1).LocalPort = 1234
    wskServer(Index + 1).Listen
End Sub

Private Sub generatePic()
    Debug.Print "generated pic"
    'lots of doevents so this hopefully doesn't bottleneck the program too bad.
    lhDC = GetDC(0)
    scrRight = GetSystemMetrics(SM_CXSCREEN)
    scrBottom = GetSystemMetrics(SM_CYSCREEN)
    Debug.Print "The screen is " & scrRight & "x" & scrBottom
    theDIB.Create scrRight, scrBottom
    
    'theDIB.LoadPictureBlt hdc, -157, -175
    theDIB.LoadPictureBlt hdc
    DoEvents
    'theDIB.CopyToClipboard False ' test it by copying it to the clipboard
    'DoEvents
    'ok compress it to jpg for transmission
    Dim b() As Byte
    Dim lBufSize As Long
    Dim lPtr As Long

    ' To save to a byte array, we first need to create
    ' a buffer which will be at least large enough to
    ' hold the image.  Here I create a buffer the same
    ' size as the DIB bits divide by 4 (seems about right)
    ReDim b(0 To theDIB.Height * theDIB.BytesPerScanLine / 4) As Byte
    ' Get a pointer to the buffer:
    lPtr = VarPtr(b(0))
    ' Pass in the buffer size:
    lBufSize = UBound(b) - 1
    DoEvents
    If SaveJPGToPtr(theDIB, lPtr, lBufSize, ServerQuality) Then
        ' If we succeed, then lBufSize will be set to the actual
        ' size of the JPG in bytes, so we can trim the image:
        DoEvents
        ReDim Preserve b(0 To lBufSize - 1) As Byte
        Debug.Print "Picture is " & lBufSize & " bytes"
        ' Just to prove that worked, load the image back in
        ' again from the buffer!
        'lPtr = VarPtr(b(0))
        'If LoadJPGFromPtr(theDIB, lPtr, lBufSize) Then
            '        Debug.Print "Ok!"
        'End If
    End If
    'finally set the new jpeg to the finalPic
    DoEvents
    finalPic = b
End Sub

Private Sub wskServer_SendComplete(Index As Integer)
    If wskServer(Index).State <> sckClosed Then wskServer(Index).Close
End Sub
