VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "FrmClient"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdADO 
      Caption         =   "Start ADO"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Starts the ADO Service on the server"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   7680
      Top             =   4800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7200
      Top             =   4800
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Services"
      Height          =   2415
      Left            =   4920
      TabIndex        =   13
      Top             =   120
      Width           =   4695
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Server Status:"
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   4455
      End
   End
   Begin MSWinsockLib.Winsock UserSock 
      Left            =   240
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "FrmClient.frx":05E8
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   " Network Connection "
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton CmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Enter Server IP Address"
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   4920
      TabIndex        =   17
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4683
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1000
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Chris Hatton in VB 6.0.
'This is a simple example of using Winsock Client to View parts of a Microsoft Access
'database on a local machine with out having any Database Objects installed.
'Be sure to include in your project references the Microsoft
'ActiveX Data Objects 2.5 Library.
'Feel free to email me with your comments or suggestions. chris@hatton.com
Public ServerStatus As String
Public ADOEnabled As Boolean

Private Sub CmdADO_Click()
SendData "RemoteSTARTUP"    'this enable you to remotely activate the ado service on the server
End Sub

Private Sub CmdConnect_Click()
On Error GoTo SckErr
UserSock.Connect Text1(0).Text, 9456
Timer2.Enabled = True
Timer2.Interval = 1000
Exit Sub

SckErr:
Call WinsockStatus

End Sub

Private Sub CmdDel_Click()
SendData "~" & Text1(1).Text        'sends the server the current user to delete
ClrTxt
Timer2.Enabled = True
End Sub

Private Sub CmdExit_Click()
End
End Sub
Private Sub CmdNew_Click()
If UserSock.State = sckClosed Then
MsgBox "Not Connected"
Exit Sub
End If

If ADOEnabled = False Then          'if the database connection is avialable you cant save changes
MsgBox "ADO Connection is not Running" & vbNewLine & "Records cannot be saved"
Exit Sub
End If

Form2.Show 1
End Sub

Private Sub cmdSave_Click()
                                'sends the server the uptodate changes
SendSaveData Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text


End Sub

Private Sub Form_Load()
ClrTxt                  ' Clear Text Boxes
Label5.Caption = ""
Text1(0).Text = UserSock.LocalIP 'this is optional to have the local ip as the default server ip
ServerStatus = Label4.Caption
End Sub
Sub ClrTxt()

For i = 0 To 4
Text1(i).Text = ""
Next i
TreeView1.Nodes.Clear           ' clears all the text fields.

End Sub

Private Sub Timer1_Timer()
Call WinsockStatus          'current winsock Status.
End Sub

Private Sub Timer2_Timer()

Call GetStuff               'continally polls the server to see if the ADO service is enabled
Timer2.Enabled = False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
SendData TreeView1.SelectedItem.Text    'Sends the Current User to the Server for Database lookup.
End Sub

Private Sub UserSock_DataArrival(ByVal bytesTotal As Long)
Dim getresults As String

UserSock.GetData getresults        'Main provider of retrieving all data.

If getresults = "ADO-True" Then
    ServerStatus = ServerStatus & vbNewLine & "ADO Services Have been Started."
    Label4.Caption = ServerStatus
    ADOEnabled = True               'Ado is avialable to access
Call GetUsers

Else
If getresults = "ADO-False" Then
    ServerStatus = "Server Status:" & vbNewLine & "ADO Services Are Closed."
    Label4.Caption = ServerStatus
    Timer2.Enabled = True
    ADOEnabled = False
Else
    Call ProcessResults(getresults) ' this processes all the incomming data from the server
    End If: End If

End Sub

