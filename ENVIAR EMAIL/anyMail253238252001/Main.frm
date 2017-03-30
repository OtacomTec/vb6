VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "anyMail"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   FillColor       =   &H8000000C&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Send As HTML"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Tag             =   "plain;"
      ToolTipText     =   "Click to Send the E-mail as HTML"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   4680
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6360
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   1838
      Picture         =   "Main.frx":0442
      TabIndex        =   5
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox txtMessage 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "0"
      Text            =   "Main.frx":0544
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "0"
      Text            =   "Type Subject Here . . ."
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtReceiver 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "0"
      Text            =   "reciever@anydomain.com"
      Top             =   600
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "0"
      Text            =   "you@anydomain.com"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Programmed by Saurabh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2415
      MousePointer    =   15  'Size All
      TabIndex        =   14
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Body:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose A SMTP Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reciever's E-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's E-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##################################'
'    anyMail Anonymous Mailer      '
'    Programmed by Saurabh         '
'    http://www.saurabhonline.org  '
'    saurabh_gupta@india.com       '
'##################################'

Option Explicit
Private DataAvailable As Boolean
Dim inData As String
Private timer As Long
Private change As Boolean
Private Const TIME_OUT = 30


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check1.Tag = "html;"        'HTML E-mail
    Else
        Check1.Tag = "plain;"       'Plain text E-mail
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim fAdd As New Add
    fAdd.Show vbModal
    If fAdd.OK Then
        List1.AddItem Trim(fAdd.txtServer.Text) + ":" + fAdd.txtPort.Text
        change = True
    End If
    Unload fAdd
End Sub

Private Sub cmdRemove_Click()
    If Not List1.ListIndex < 0 Then
        List1.RemoveItem List1.ListIndex    'Remove item
        change = True
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim str As String
    DataAvailable = False
    timer = 0
    change = False
    On Error GoTo errhandler
    Open "servers.txt" For Input As #1      'Open SMTP server list file
    While Not EOF(1)
        Line Input #1, str
        List1.AddItem Trim(str)
    Wend
    Close #1
    Exit Sub
errhandler:
    MsgBox "Error opening servers.txt"
    End
End Sub


Private Sub Label6_Click()
    MsgBox "Programmed By Saurabh Gupta" + vbCrLf + "E-mail: saurabh_gupta@india.com" + vbCrLf + "Homepage: http://www.saurabhonline.org", vbOKOnly, "About anyMail"
End Sub

Private Sub txtSender_GotFocus()
    If txtSender.Tag = 0 Then
        txtSender.Tag = 1
        txtSender.Text = ""
    End If
End Sub
Private Sub txtSender_Validate(KeepFocus As Boolean)
    If txtSender.Text = "" Then
        txtSender.Text = "you@anydomain.com"
        KeepFocus = False
        txtSender.Tag = 0
    End If
End Sub
Private Sub txtReceiver_GotFocus()
    If txtReceiver.Tag = 0 Then
        txtReceiver.Tag = 1
        txtReceiver.Text = ""
    End If
End Sub
Private Sub txtReceiver_Validate(KeepFocus As Boolean)
    If txtReceiver.Text = "" Then
        txtReceiver.Text = "receiver@anydomain.com"
        KeepFocus = False
        txtReceiver.Tag = 0
    End If
End Sub

Private Sub txtSubject_GotFocus()
    If txtSubject.Tag = 0 Then
        txtSubject.Tag = 1
        txtSubject.Text = ""
    End If
End Sub
Private Sub txtSubject_Validate(KeepFocus As Boolean)
    If txtSubject.Text = "" Then
        txtSubject.Text = "Type Subject Here . . ."
        KeepFocus = False
        txtSubject.Tag = 0
    End If
End Sub

Private Sub txtMessage_GotFocus()
    If txtMessage.Tag = 0 Then
        txtMessage.Tag = 1
        txtMessage.Text = ""
    End If
End Sub
Private Sub txtMessage_Validate(KeepFocus As Boolean)
    If txtMessage.Text = "" Then
        txtMessage.Text = "Type Message Here . . ."
        KeepFocus = False
        txtMessage.Tag = 0
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not Number = sckSuccess Then
        MsgBox Description          'Display error
        Timer1.Enabled = False
        CloseConn True
    End If
End Sub

Private Sub cmdSend_Click()
    If List1.ListIndex < 0 Then
        MsgBox "Please Choose a Server First"
        Exit Sub
    End If
    If txtSender.Tag = 0 Then
        MsgBox "Please enter Sender's E-mail"
        txtSender.SetFocus
        Exit Sub
    End If
    If txtReceiver.Tag = 0 Then
        MsgBox "Please enter Receiver's E-mail"
        txtReceiver.SetFocus
        Exit Sub
    End If
    If txtSubject.Tag = 0 Then
        MsgBox "Please enter a subject"
        txtSubject.SetFocus
        Exit Sub
    End If
    
    Dim tmp() As String
    tmp = Split(List1.List(List1.ListIndex), ":")
    cmdSend.Enabled = False
    cmdSend.Caption = "Connecting..."
    Winsock1.Connect tmp(0), Val(tmp(1))    'Connect to server
    txtSender.Enabled = False
    txtReceiver.Enabled = False
    txtSubject.Enabled = False
    txtMessage.Enabled = False
    List1.Enabled = False
End Sub

Private Sub Winsock1_DataArrival _
(ByVal bytesTotal As Long)
    Dim data As String
    Winsock1.GetData data, vbString
    'Add data arrived data to the already arrived data
    inData = inData + data
    'Wait till a line is recieved (with CR LF in the end)
    If StrComp(Right$(inData, 2), vbCrLf) = 0 Then DataAvailable = True
End Sub
Private Sub Winsock1_Connect()
    cmdSend.Caption = "Connected"
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    
    Dim reply As String
    Dim tmp() As String
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 220 Then           'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    cmdSend.Caption = "Receiving Welcome Message"
    'Start the process
    Winsock1.SendData "HELO " + Winsock1.LocalHostName + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    'Send MAIL FROM
    Winsock1.SendData "MAIL FROM:<" + txtSender.Text + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn True
        Exit Sub
    End If
    'Send RCPT TO
    Winsock1.SendData "RCPT TO:<" + txtReceiver.Text + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn True
        Exit Sub
    End If
    'Send DATA
    DoEvents
    Winsock1.SendData "DATA" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 354 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    cmdSend.Caption = "Sending Mail . . ."
    'Send the E-Mail
    Winsock1.SendData "From: <" + txtSender.Text + ">" + vbCrLf + _
                      "To: " + txtReceiver.Text + vbCrLf + _
                      "Subject: " + txtSubject.Text + vbCrLf + _
                      "X-Mailer: anyMail v1.1" + vbCrLf + _
                      "Mime-Version: 1.0" + vbCrLf + _
                      "Content-Type: text/" + Check1.Tag + vbTab + "charset=us-ascii" + vbCrLf + vbCrLf + _
                      txtMessage.Text
    Winsock1.SendData vbCrLf + "." + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable             'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then               'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    Winsock1.SendData "QUIT"
    MsgBox "Message Sent"
    CloseConn False
End Sub

Private Sub Timer1_Timer()
    timer = timer + 1
    If timer = TIME_OUT Then
        CloseConn True              'Disconnect if timed out
        MsgBox "Could not connect to host " + List1.List(List1.ListIndex) + vbCrLf + "Operation timed out"
        Timer1.Enabled = False
    End If
End Sub
Private Sub CloseConn(Err As Boolean)           'Close Connection & enable contrls
    Winsock1.Close
    cmdSend.Caption = "Send"
    cmdSend.Enabled = True
    txtSender.Enabled = True
    txtReceiver.Enabled = True
    txtSubject.Enabled = True
    txtMessage.Enabled = True
    List1.Enabled = True
    If Err Then If MsgBox("Do you want to remove this server from the list of servers", vbYesNo) = vbYes Then cmdRemove_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu And change Then
    If MsgBox("Servers list has changed. Do you want to save changes?", vbYesNo) = vbYes Then
        Open "servers.txt" For Output As #1     'Save list before exit
        Dim i As Integer
        For i = 0 To List1.ListCount - 1
            Print #1, List1.List(i)
        Next i
        Close #1
    End If
End If
End Sub
