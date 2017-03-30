VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSendMail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Mail"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dialog 
      Left            =   5040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "open file"
      Filter          =   "All files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send mail"
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtAttachment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   10
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   5535
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtSender 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtRecipient 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock smtp 
      Left            =   0
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Attachment:"
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SMTP Host:"
      Height          =   195
      Left            =   900
      TabIndex        =   7
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Your e-mail address:"
      Height          =   195
      Left            =   345
      TabIndex        =   6
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Recipient e-mail address:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   1185
      TabIndex        =   4
      Top             =   1200
      Width           =   585
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    dialog.ShowOpen
    txtAttachment.Text = dialog.FileName
End Sub

Private Sub cmdClear_Click()
    'Clear all fields
    Me.txtAttachment = ""
    Me.txtHost = ""
    Me.txtMessage = ""
    Me.txtRecipient = ""
    Me.txtSender = ""
    Me.txtSubject = ""
End Sub

Private Sub cmdSend_Click()
    'Connect to the smtp server.
    'Smtp server port is everytime 25
    If Not txtAttachment.Text = "" Then
        EncodedFile = UUEncodeFile(txtAttachment)
    End If
    smtp.Connect txtHost.Text, 25
    'reset the state
    smtpState = MAIL_CONNECT
End Sub

Private Sub smtp_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    'Retrive data from winsock buffer
    smtp.GetData strServerResponse
    Debug.Print strServerResponse
    'Get server response code (first three symbols)
    strResponseCode = Left(strServerResponse, 3)
    'Only these three codes from the server tell us
    'that the command was accepted
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
        Select Case smtpState
            Case MAIL_CONNECT
                smtpState = MAIL_HELO
                'Remove blank spaces
                strDataToSend = Trim$(txtSender.Text)
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                smtp.SendData "HELO " & strDataToSend & vbCrLf
                Debug.Print "HELO " & strDataToSend
            Case MAIL_HELO
                smtpState = MAIL_FROM
                'Send MAIL FROM command to the server
                'so it knows from who the message comes
                smtp.SendData "MAIL FROM:" & Trim$(txtSender.Text) & vbCrLf
                Debug.Print "MAIL FROM:" & Trim$(txtSender.Text)
            Case MAIL_FROM
                smtpState = MAIL_RCPTTO
                'Send RCPT TO command to the server
                'so it knows where to send the message
                smtp.SendData "RCPT TO:" & Trim$(txtRecipient.Text) & vbCrLf
                Debug.Print "RCPT TO:" & Trim$(txtRecipient.Text)
            Case MAIL_RCPTTO
                smtpState = MAIL_DATA
                'Send DATA command to the server
                'so it knows that we want to send the message
                smtp.SendData "DATA" & vbCrLf
                Debug.Print "DATA"
            Case MAIL_DATA
                smtpState = MAIL_DOT
                'Send Subject
                smtp.SendData "Subject:" & txtSubject.Text & vbLf & vbCrLf
                Debug.Print "Subject:" & txtSubject.Text
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                'Add atacchments
                strMessage = txtMessage.Text & vbCrLf & vbCrLf & EncodedFile
                'clear the buffer for the encoded files
                EncodedFiles = ""
                'Parse message to get lines
                varLines = Split(strMessage, vbCrLf)
                'clear message buffer
                strMessage = ""
                'Send each line of the message
                'so no line gets lost
                For Each varLine In varLines
                    smtp.SendData CStr(varLine) & vbLf
                Next
                'Send a dot symbol so the server knows
                'that the end of the message is reached
                smtp.SendData "." & vbCrLf
                Debug.Print "."
            Case MAIL_DOT
                smtpState = MAIL_QUIT
                'Send QUIT command
                smtp.SendData "QUIT" & vbCrLf
                Debug.Print "QUIT"
            Case MAIL_QUIT
                'Close the connection to the smtp server
                smtp.Close
        End Select
    Else
        'Check if an error occured
        smtp.Close
        If Not smtpState = MAIL_QUIT Then
            'If yes then print the error
            MsgBox "Error: " & strServerResponse, vbCritical, "Error"
            Unload Me
        Else
            'if the message sent successfully, print it
            Unload Me
            Debug.Print "Message sent"
        End If
    End If
End Sub

Private Sub smtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Tell the user that an error occured
    MsgBox "Winsock Error number " & Number & vbCrLf & Description, vbExclamation
End Sub

