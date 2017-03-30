VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmServer 
   Caption         =   "frmServer"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   " User Accounts "
      Height          =   1935
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   4335
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create User"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "UserName"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8880
      Top             =   4680
   End
   Begin VB.Frame Frame1 
      Caption         =   "                    Connected Users                         "
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      Begin MSComctlLib.ListView lvusrs 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8493
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   8880
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   " All Users "
      Height          =   3255
      Left            =   3720
      TabIndex        =   9
      Top             =   2400
      Width           =   4335
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Save Changes"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "UserName"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Last Error:"
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu menu1 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu MnuDisc 
         Caption         =   "&Disconnect User"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FrmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RX As String
Public MaxCN As Long
Public JobNumber, InitMax As Long

Private Sub cmdCreate_Click()
If Text1(0).Text = "" Then
    MsgBox "Can't Create User Account", vbCritical, "Account Creation Error"
        Exit Sub
    End If
    
If Text1(1).Text = "" Then
    MsgBox "Can't Create User Account", vbCritical, "Account Creation Error"
        Exit Sub
    End If
    
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

Rs.Open "Select * from authentication", cn, adOpenKeyset, adLockOptimistic

Rs.AddNew

Rs!UserName = Text1(0).Text
Rs!Password = Text1(1).Text

Rs.Update
Rs.Close
Set Rs = Nothing

Call LtUsrs


End Sub

Private Sub cmdUpdate_Click()
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

SQL = "Select * from authentication where username = " & Chr(34) & List1.Text & Chr(34)

Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

 Rs!UserName = Text1(2).Text
 Rs!Password = Text1(3).Text

Rs.Update
Rs.Close
Set Rs = Nothing

Call LtUsrs

End Sub

Private Sub Form_Load()
lvusrs.ColumnHeaders.Add , , "Connected Users", 3300

MaxCN = 0

DBConnect

If DBConnect = True Then

sckServer(InitMax).LocalPort = "9456"
sckServer(InitMax).Listen

Call LtUsrs
Else

End

End If


End Sub
Private Sub LtUsrs()
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim i As Integer
List1.Clear
Text1(0).Text = "": Text1(1).Text = "": Text1(2).Text = "": Text1(3).Text = ""
Rs.Open "Select * from authentication", cn, adOpenForwardOnly, adLockReadOnly

    For i = 1 To Rs.RecordCount

        List1.AddItem Rs!UserName
        
        Rs.MoveNext
    Next i

Rs.Close
Set Rs = Nothing
End Sub

Private Sub List1_Click()
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

SQL = "Select * from authentication where username = " & Chr(34) & List1.Text & Chr(34)

Rs.Open SQL, cn, adOpenForwardOnly, adLockReadOnly

Text1(2).Text = "" & Rs!UserName
Text1(3).Text = "" & Rs!Password

Rs.Close
Set Rs = Nothing


End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

MaxCN = MaxCN + 1                           'Increases the user count
Load sckServer(MaxCN)                       'Loads up the new winsock control
sckServer(MaxCN).LocalPort = 0              'sets a random port
sckServer(MaxCN).Accept requestID           'Accept connection


End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Recieve As String
    If Len(Recieve) = Null Then
        Exit Sub
    Else
    InitMax = MaxCN
        
        sckServer(InitMax).GetData Recieve
            Call ParseRecv(Recieve) 'Analyze data
            
    End If

End Sub


Private Sub Timer1_Timer()
If sckServer(InitMax).State = sckConnected Then Label1.Caption = "Connected:  " & "Port: " & sckServer(InitMax).LocalPort & " Socket Number: " & MaxCN
If sckServer(InitMax).State = sckClosed Then Label1.Caption = "Connection Closed: " & "Port: " & sckServer(MaxCN).LocalPort & " Socket Number: " & MaxCN
If sckServer(InitMax).State = sckConnecting Then Label1.Caption = "Connecting: " & "Port: " & sckServer(MaxCN).LocalPort & " Socket Number: " & MaxCN
If sckServer(InitMax).State = sckConnectionPending Then Label1.Caption = "Connection Pending: " & "Port: " & sckServer(MaxCN).LocalPort & " Socket Number: " & MaxCN
If sckServer(InitMax).State = sckBadState Then Label1.Caption = "Bad State Connection: " & "Port: " & sckServer(MaxCN).LocalPort & " Socket Number: " & MaxCN


End Sub
