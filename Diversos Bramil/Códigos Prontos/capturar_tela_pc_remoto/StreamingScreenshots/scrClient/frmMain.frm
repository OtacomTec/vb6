VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   120
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Text            =   "1234"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picClient 
      AutoRedraw      =   -1  'True
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   480
      Width           =   8775
   End
   Begin VB.Label Label2 
      Caption         =   "Source Port"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Source IP"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private theDib As cDIBSection
Private totalPacket As String

Private Sub cmdConnect_Click()
    Dim theIP As String
    theIP = txtIP.Text
    Dim thePort As String
    thePort = txtPort.Text
    txtIP.Enabled = False
    txtPort.Enabled = False
    wskClient.RemoteHost = theIP
    wskClient.RemotePort = thePort
    wskClient.Connect
End Sub

Private Sub cmdDisconnect_Click()
    txtIP.Enabled = True
    txtPort.Enabled = True
    If wskClient.State <> sckClosed Then wskClient.Close
End Sub

Private Sub Form_Load()
    Set theDib = New cDIBSection
    theDib.Create 1024, 768
End Sub

Private Sub wskClient_Close()
    On Error Resume Next
    If wskClient.State <> sckClosed Then wskClient.Close
    Dim b() As Byte
    b = StrConv(totalPacket, vbFromUnicode)
    Dim lPtr As Long
    Dim lSize As Long
    'and then time to change the picture
    lPtr = VarPtr(b(0))
    lSize = UBound(b)
    Debug.Print "picture is " & lSize & " bytes"
    If Not LoadJPGFromPtr(theDib, lPtr, lSize) Then
        Debug.Print "Did not load the picture"
    Else
        'theDib.Resample picClient.Height, picClient.Width
        picClient.Cls
        theDib.PaintPicture picClient.hdc
        'theDib.CopyToClipboard False
    End If
    totalPacket = vbNullString
    wskClient.Connect
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    If bytesTotal = 0 Then Exit Sub 'nothing there

    Debug.Print "received " & bytesTotal & " bytes"
    Dim thePacket As String
    wskClient.GetData thePacket
    totalPacket = totalPacket & thePacket
End Sub
