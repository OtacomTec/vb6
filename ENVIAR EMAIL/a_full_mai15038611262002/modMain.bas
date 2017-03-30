Attribute VB_Name = "modMain"
Public Enum SMTP_State              'The smtp states enum
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Public Enum POP3States             'The POP3 states enum
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_QUIT
End Enum

Public Type AdressData             'buffer for a response
    message         As String
    responseAdress  As String
    subject         As String
End Type

Public sAdressData       As AdressData   'the Adress data buffer
Public pop3state         As POP3States   'the pop3 states
Public smtpState         As SMTP_State   'the smtp states
Public EncodedFile       As String       'the encoded files buffer (attachments)
Dim frm                  As Form         'the variable for new forms
Public MessageBuffer(50) As String  'some Buffer variables
Public FromBuffer(50)    As String
Public SubjectBuffer(50) As String

Sub CreateNewReadMessageForm()
    'Create a new frmReadMessage form
    Set frm = New frmReadMessage
    frm.Show
End Sub

Sub CreateNewSendMailForm(ByRef Adresses As AdressData)
    'Create a new frmSendMail form
    Set frm = New frmSendMail
    frm.Show
    'fill in the string for a response
    frm.txtRecipient = Adresses.responseAdress
    frm.txtMessage = Adresses.message
    frm.txtSubject = "Re: " & Adresses.subject
    'Clear the buffer
    Adresses.message = ""
    Adresses.responseAdress = ""
    Adresses.subject = ""
End Sub

Public Function SplitMessage(message As String) As Variant
    On Error Resume Next
    Dim Pos             As Long
    Dim Pos2            As Long
    Dim arrx(0 To 3)    As String
    Dim br1             As Long
    Dim br2             As Long
    'extract the message body
    Pos = InStr(1, message, vbCrLf & vbCrLf)
    arrx(3) = Right$(message, Len(message) - Pos - 3)
    'Split the message into peaces, so we can handle it better
    Splitter = Split(message, vbCrLf)
    'get every line
    For i = 0 To UBound(Splitter)
        'get every char in the line
        For i2 = 1 To Len(Splitter(i))
            If LCase(Mid(Splitter(i), i2, 8)) = "subject:" Then
                'found the subject
                'fill the array
                arrx(0) = Mid(Splitter(i), i2 + 8)
            ElseIf LCase(Mid(Splitter(i), i2, 7)) = "sender:" Then
                'found the from: adress
                'fill the array
                arrx(1) = Mid(Splitter(i), 8)
            ElseIf LCase(Mid(Splitter(i), i2, 3)) = "to:" Then
                'found the to: adress
                'fill the array
                br1 = InStr(1, Splitter(i), "<")
                br2 = InStr(1, Splitter(i), ">")
                arrx(2) = Mid(Splitter(i), 4)
            End If
        Next i2
    Next i
    'check for brackets in the e-mail adress
    If InStr(1, arrx(1), "<") <> 0 Then
        arrx(1) = Replace(arrx(1), "<", " ")
    End If
    If InStr(1, arrx(1), ">") <> 0 Then
        arrx(1) = Replace(arrx(1), ">", " ")
    End If
    If InStr(1, arrx(2), "<") <> 0 Then
        arrx(2) = Replace(arrx(2), "<", " ")
    End If
    If InStr(1, arrx(2), ">") <> 0 Then
        arrx(2) = Replace(arrx(2), ">", " ")
    End If
    'return the array
    SplitMessage = arrx
End Function
