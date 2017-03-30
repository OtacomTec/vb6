VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5640
      Top             =   4800
   End
   Begin VB.Frame Frame1 
      Caption         =   " Server Status:"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton CmdDisc 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4080
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton CmdStart 
         Caption         =   "Start ADO Services"
         Height          =   495
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
   Begin MSWinsockLib.Winsock ServerSock 
      Index           =   0
      Left            =   6240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
Dim MaxID As Long
Dim ADO As Boolean

Private Sub CmdDisc_Click()
On Error Resume Next
cn.Close
Call SendAllSorts
ServerSock(MaxID).Close
Set cn = Nothing

Label1.Caption = "Database Status:   Disconnected"
Label2.Caption = "Network Status:   Disconnected"
CmdStart.Enabled = True
CmdDisc.Enabled = False
SendData "ADO-False"       'tells the client that the ADO Service is down
End Sub

Private Sub CmdStart_Click()

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

On Error GoTo OpenErr

cn.Provider = "Microsoft.Jet.OLEDB.4.0"
Label1.Caption = "Loading Database"
cn.Open App.Path & "\Contacts.mdb", admin           'opens the database
Label1.Caption = "Database Status:   Connected"
CmdDisc.Enabled = True
CmdStart.Enabled = False
ADO = True
Exit Sub
OpenErr:

Label1.Caption = "Error Opening Database"           'if theres an error we need to know about it.

MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
ServerSock(0).LocalPort = 9456
ServerSock(0).Listen

Label1.Caption = "Database Status: "
Label2.Caption = "Network Status:  Not Connected"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cn.Close
ServerSock(MaxID).Close
End Sub

Private Sub ServerSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    MaxID = MaxID + 1                   '
    Load ServerSock(MaxID)              '
    ServerSock(MaxID).LocalPort = 0     '
    ServerSock(MaxID).Accept requestID  'for every new connection made, lets load up a new winsock control
  
If ServerSock.Item(MaxID).Protocol = 0 Then ProtocolType = "TCP Connection" Else ProtocolType = "UDP Connection"

If ServerSock.Item(MaxID).State = sckConnected Then
Form1.Label2.Caption = "Network Status:   Connection Established " & vbNewLine & vbNewLine & "Remote IP: " & ServerSock(MaxID).RemoteHostIP & _
 vbNewLine & "Remote Port: " & ServerSock(MaxID).RemotePort & vbNewLine & ProtocolType
 
 
 CmdDisc.Enabled = True
End If

End Sub

Private Sub ServerSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim GetRecieve As String
Dim RecieveQry As String
Dim RecieveSave(4) As String
Dim RecieveDel As String
Dim RecieveAdd(3) As String
Dim Rqry As Integer

ServerSock(MaxID).GetData GetRecieve    'get the incomming data from the client
List1.AddItem "User: " & ServerSock(MaxID).RemoteHostIP & " Getting info: " & GetRecieve 'add it to the server list
If GetRecieve = "$RemoteSTARTUP" Then CmdStart = True 'enables the client to activate the ADO Services remotely

If GetRecieve = "GetUsers" Then ' send the client the list of users
    List1.Clear
    Call ListUsers                 'this does all the hardwork of getting the users to a string.
End If

If GetRecieve = "GetStuff" Then     'Sends the client other stuff such as the Ado Service status.
    List1.Clear
    Call SendAllSorts
    Exit Sub
End If

On Error Resume Next                'heaps of error in this script can't be bothered sorting it out.

Rqry = InStr(RecieveQry, "$")      ' compares the string to see if it is a Query that the user is wanting, and sends the details
                                    ' this keeps the data recieved Unique from the rest
If Rqry >= 0 Then
    RecieveQry = Split(GetRecieve, "$")(1) 'simply splits the "$" from the contact name
    GenerateSQL RecieveQry                  'now that we have the Contact Name in one piece query the database
       
End If


    For i = 1 To 4
           
        RecieveSave(i) = Split(GetRecieve, "%")(i) 'a unique idenifier
                                                   'this saves all text field data to the database
    Next i                                         '
                                                   '
        SaveChanges RecieveSave(1), RecieveSave(2), RecieveSave(3), RecieveSave(4) 'save it all
        
        RecieveDel = Split(GetRecieve, "$~")(1)    'another unique idenifier, when this one recieved
        DelRecord RecieveDel                       'from the client, we know we want to delete this data
        
    For i = 1 To 3

        RecieveAdd(i) = Split(GetRecieve, "~~")(i) 'New User wanting to be added
      Next i
        AddUser RecieveAdd(1), RecieveAdd(2), RecieveAdd(3) 'new record is currently added


End Sub

Sub SendAllSorts()
If ADO = True Then SendData "ADO-True" Else SendData "ADO-False"


End Sub
Private Sub Timer1_Timer()
Call ServerSck
End Sub

Sub ListUsers()
On Error GoTo ListErr
Set rs = New ADODB.Recordset
Dim i As Integer
Dim j As Long
Dim SQLQry, strNames As String                          'this procedure simple gets all the users
                                                        'from the database

SQLQry = "Select Name from users"

rs.Open SQLQry, cn, adOpenKeyset, adLockReadOnly

j = rs.RecordCount

For i = 1 To j


strNames = strNames & rs!Name & "-"

rs.MoveNext
Next i

SendData strNames

rs.Close
Set rs = Nothing


Exit Sub

ListErr:
Exit Sub

End Sub
