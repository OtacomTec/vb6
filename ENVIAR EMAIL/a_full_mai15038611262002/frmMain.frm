VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail Client by Irnchen"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "Send a message"
      Height          =   615
      Left            =   4815
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadMessage 
      Caption         =   "Read Selected Message"
      Height          =   615
      Left            =   2895
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCheckMailbox 
      Caption         =   "Check Mailbox"
      Height          =   615
      Left            =   975
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame frame 
      Caption         =   "Messages"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6975
      Begin VB.ListBox lstMessages 
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
   Begin MSWinsockLib.Winsock pop3 
      Left            =   7200
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   4905
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   2595
      TabIndex        =   7
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "POP3 Host:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################################
'#a full mail client                                             #
'#   by Irnchen                                                  #
'#                                                               #
'#I didn't find something like this on PSC so I coded it myself. #
'#It can recieve and send mails per POP3 or SMTP sever.          #
'#Also it can encode and decode attachments from files and save  #
'#them to disk.                                                  #
'#Nearly every line is commented                                 #
'#so it's easy to understand how the servers are working.        #
'#                       PLEASE VOTE                             #
'#IF SOMEBODY KNOWS HOW TO CREATE 16BIT LITTLE ENDIANS           #
'#PLEASE MAIL ME!                                                #
'#I NEED IT FOR A UPDATE OF MY ISO-BURN-CODE                     #
'#SO IT CAN WRITE FILES ON CD-R!!!!!!!!!                         #
'#                                                               #
'#The encoding module for files isn't mine.                      #
'#But it also can be found on PSC.                               #
'#################################################################
Public ArrayBuffer      As Variant      'save requested data

Private Sub cmdCheckMailbox_Click()
    'Check the mailbox
    pop3.Connect txtHost.Text, 110
End Sub

Private Sub cmdDelete_Click()
    'Delete the selected message
End Sub

Private Sub cmdReadMessage_Click()
    'Create a new read-message form for reading a message
    CreateNewReadMessageForm
End Sub

Private Sub cmdSendMessage_Click()
    'Create a new form for sending a mail
    CreateNewSendMailForm sAdressData
End Sub

Private Sub pop3_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    'Save the received data into strData variable
    pop3.GetData strData
    Debug.Print strData
    If Left$(strData, 1) = "+" Or pop3state = POP3_RETR Then
        'If the first character of the server's response is "+" then
        'the server accept the request
        Select Case pop3state
            Case POP3_Connect
                'Reset the number of messages
                intMessages = 0
                pop3state = POP3_USER
                'Send the username
                pop3.SendData "USER " & txtUsername.Text & vbCrLf
                Debug.Print "USER " & txtUsername.Text
            Case POP3_USER
                'send the password
                pop3state = POP3_PASS
                pop3.SendData "PASS " & txtPass.Text & vbCrLf
                Debug.Print "PASS " & txtPass.Text
            Case POP3_PASS
                pop3state = POP3_STAT
                'Send STAT command to know how many messages in the mailbox
                pop3.SendData "STAT" & vbCrLf
                Debug.Print "STAT"
            Case POP3_STAT
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                intMessages = CInt(Mid$(strData, 5, _
                              InStr(5, strData, " ") - 5))
                If intMessages > 0 Then
                    'mail in the mailbox!
                    pop3state = POP3_RETR
                    'Increment the number of messages by one
                    intCurrentMessage = intCurrentMessage + 1
                    'send the RETR command to the server in
                    'order to retrieve the first message
                    pop3.SendData "RETR 1" & vbCrLf '1 for the first message
                    Debug.Print "RETR 1"
                Else
                    'The mailbox is empty.
                    pop3state = POP3_QUIT
                    pop3.SendData "QUIT" & vbCrLf
                    Debug.Print "QUIT"
                    Debug.Print "You have not mail."
                End If
            Case POP3_RETR
                'this part of the sub retrieves the messages
                strBuffer = strBuffer & strData
                'If in the string is a point, we got the message
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    'Delete the first string-the server's response
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    'Add new message to the list of new messages
                    ArrayBuffer = SplitMessage(strBuffer)
                    MessageBuffer(intCurrentMessage - 1) = ArrayBuffer(3)
                    FromBuffer(intCurrentMessage - 1) = Trim(ArrayBuffer(2))
                    SubjectBuffer(intCurrentMessage - 1) = ArrayBuffer(0)
                    lstMessages.AddItem SubjectBuffer(intCurrentMessage - 1)
                    'Clear buffer for next message
                    strBuffer = ""
                    If intCurrentMessage = intMessages Then
                        'If we got the last message, close the connection
                        pop3state = POP3_QUIT
                        pop3.SendData "QUIT" & vbCrLf
                        Debug.Print "QUIT"
                    Else
                        intCurrentMessage = intCurrentMessage + 1
                        pop3state = POP3_RETR
                        'Send RETR command to download next message
                        pop3.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                        Debug.Print "RETR " & intCurrentMessage
                    End If
                End If
            Case POP3_QUIT
                'close the connection
                pop3.Close
                'handle the messages
                '#######################################################
        End Select
    Else
        'If an error occured...
            pop3.Close
            Debug.Print "POP3 Error: " & strData
    End If
End Sub

Private Sub pop3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'An error occured
    MsgBox "Error: #" & Number & vbCrLf & Description
End Sub
