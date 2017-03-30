VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClient 
   Caption         =   "frmClient"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0484
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":06F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":07C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6060
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   177165
            MinWidth        =   177165
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9600
      Top             =   4920
   End
   Begin VB.CheckBox Check1 
      Caption         =   "List Completed Jobs"
      Height          =   195
      Left            =   6960
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdNew 
         Caption         =   "Create New Job"
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9240
      Top             =   5880
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8705
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   9120
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "FrmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MaxCN As Long
Public Authenticated As Boolean
Public AuthCompleted As Boolean
Public RScount, ConCurrent As Long

Private Sub Check1_Click()
ShowJobs
End Sub

Private Sub CmdConnect_Click()
MaxCN = MaxCN + 1               'creates another winsock instant
Authenticated = False
Load sckClient(MaxCN)
sckClient(MaxCN).Connect Text1.Text, "9456"
End Sub

Private Sub cmdNew_Click()
FrmNewJob.Show 1
ShowJobs
End Sub
Public Sub ShowJobs()
If sckClient(MaxCN).State = sckClosed Then Exit Sub
With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear

    .ColumnHeaders.Add , , "JobID", 600                 'Create Columns once connected
    .ColumnHeaders.Add , , "Date", 1200
    .ColumnHeaders.Add , , "Client", 2000
    .ColumnHeaders.Add , , "Phone", 1200
    .ColumnHeaders.Add , , "Job Description", 2700
    .ColumnHeaders.Add , , "Techinican", 1200
End With

If Check1.Value = 0 Then sckClient(MaxCN).SendData "ShowJobs" & FrmClient.ConCurrent
If Check1.Value = 1 Then sckClient(MaxCN).SendData "ShowCompletedJobs" & FrmClient.ConCurrent

End Sub
Private Sub cmdRefresh_Click()

ShowJobs
ListView1.Refresh

End Sub

Private Sub Form_Load()
Text1.Text = sckClient(MaxCN).LocalIP   'Loopback. Handly if you only have one computer

End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Height = FrmClient.Height - 2000
If FrmClient.Width <> 9270 Then FrmClient.Width = 9270
End Sub

Private Sub Form_Unload(Cancel As Integer)
For i = 0 To MaxCN
    sckClient(i).Close      ' Closes all the winsock connections
Next i

reponse = MsgBox("Would you like to see more of my projects? Then press the YES button." & vbNewLine & _
"If you like this program, feel free to contact me at: Chris@Hatton.com" & _
vbNewLine & "Please take the time in sending me in a vote for this project, as " & _
"i appreciate it", vbInformation + vbYesNo, "Thanks!")

If reponse = vbYes Then

Shell "start http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=Chris+Hatton&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&B1=Quick+Search&optSort=Alphabetical"
Shell "start http://www.chris.hatton.com"

Else

    End
End If

End Sub

Private Sub ListView1_DblClick()
FrmEditJob.Label11.Caption = ListView1.SelectedItem.Text
sckClient(MaxCN).SendData "JobNumber" & ListView1.SelectedItem.Text & "~~" & FrmClient.ConCurrent
FrmEditJob.Show 1
ShowJobs

End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Recieve As String
 sckClient(MaxCN).GetData Recieve
 
If Len(Recieve) = Null Then Exit Sub Else Call ParseRecv(Recieve) 'Analyze data

End Sub

Private Sub Timer1_Timer()
                                                'Get the status of the winsock control
If sckClient(MaxCN).State = sckConnected Then
    Label1.Caption = "Connected:  " & "Port: " & sckClient(MaxCN).LocalPort & " Socket Number: " & MaxCN
    Timer2.Enabled = True
End If

If sckClient(MaxCN).State = sckClosed Then Label1.Caption = "Connection Closed:  " & sckClient(MaxCN).LocalPort & " Socket Number: " & MaxCN
If sckClient(MaxCN).State = sckConnecting Then Label1.Caption = "Connecting:  " & "Port: " & sckClient(MaxCN).LocalPort & " Socket Number: " & MaxCN
If sckClient(MaxCN).State = sckConnectionPending Then Label1.Caption = "Connection Pending:  " & "Port: " & sckClient(MaxCN).LocalPort & " Socket Number: " & MaxCN
If sckClient(MaxCN).State = sckBadState Then Label1.Caption = "Bad State Connection:  " & "Port: " & sckClient(MaxCN).LocalPort & " Socket Number: " & MaxCN

If InStr(Label1.Caption, "Connecting") Then
    CmdConnect = True
    Label1.Caption = "Reconnecting... " & MaxCN
    
    If MaxCN > 20 Then
        Label1.Caption = "Error No Connection Found"
        Timer1.Enabled = False
    End If
End If

If AuthCompleted = True Then
cmdRefresh.Enabled = True
cmdNew.Enabled = True
Check1.Enabled = True
Else
cmdRefresh.Enabled = False
cmdNew.Enabled = False
Check1.Enabled = False

End If

End Sub

Private Sub Timer2_Timer()
If Authenticated = True Then Exit Sub

If sckClient(MaxCN).State = sckConnected Then
    sckClient(MaxCN).SendData "VerifyUser" & ConCurrent 'Send in our Notification.
    StatusBar1.Panels.Item(1).Text = "Requesting Authentication"
    StatusBar1.Panels.Item(1).Picture = ImageList1.ListImages.Item(2).Picture
    Timer2.Enabled = False
    Authenticated = True
End If
End Sub
