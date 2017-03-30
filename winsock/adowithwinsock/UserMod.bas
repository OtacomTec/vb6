Attribute VB_Name = "UserMod"

Sub WinsockStatus()
If Form1.UserSock.State = sckConnected Then Form1.Label5.Caption = "Connection Established"
If Form1.UserSock.State = sckClosed Then Form1.Label5.Caption = "Connection is Closed"


If Form1.UserSock.State = sckError Then             'if there is a socket error let us know about it
    Form1.Label5.Caption = "Network Sock Err"
    Form1.UserSock.Close
End If

If Form1.UserSock.State = sckInProgress Then Form1.Label5.Caption = "Connection Open"
If Form1.UserSock.State = sckConnecting Then Form1.Label5.Caption = "Connecting"

If Form1.UserSock.State = 8 Then
    Form1.Label5.Caption = "Connection is Closed"       'Socket is closed
    Form1.ClrTxt
    Form1.UserSock.Close
End If

End Sub

Sub GetUsers()
On Error Resume Next
Form1.UserSock.SendData "GetUsers"      ' Send a request for the contact list
End Sub

Sub GetStuff()
On Error Resume Next
Form1.UserSock.SendData "GetStuff"      ' Get other info such as if the ADO service is active
End Sub

Sub ProcessResults(Digest As String)
On Error Resume Next
Dim Result(2) As String

Result(1) = InStr(Digest, "QRYNAME")        'Compares the incomming data if its a SQL Qry
Result(2) = InStr(Digest, "-")              'Compares the incomming data if it is a Contact List
                
If Result(1) >= 1 Then                      'if the SQL Query result is true then Display
                                            'the details in the text fields. (name,address,etc)
       For i = 1 To 6550
            processed = Split(Digest, "QRYNAME")(i)
            Form1.Text1(i).Text = processed
            If processed = "" Then Exit Sub
    Next i
Else
         
    If Result(2) >= 1 Then                   'if the contact List is the incomming data then
                                             'Split up the string.
            For i = 0 To 6550                'Heres a limited value of the amount of records
            processed = Split(Digest, "-")(i) 'that will be processed if need be. it can be higher.
            If processed = "" Then Exit Sub  'When all the records have been processed end the search.

    With Form1.TreeView1
             .Nodes.Add , tvwChild, , processed, 1  'Add the contact list to the treeview control

    End With

Next i

Else
MsgBox "Theres a Problem with the Results " & Result(1) & " " & Result(2) 'Problem? heres the results

End If: End If


End Sub

Sub SendData(SQL As String)

If Form1.UserSock.State = sckClosed Then    'if theres no connection to the server no point in sending data to it.
MsgBox "Not Connected"
Exit Sub
End If
Form1.UserSock.SendData "$" & SQL           'sends the server the Contact Name to search up in the database

End Sub

Sub SendSaveData(string1 As String, string2 As String, string3 As String, string4 As String)

If Form1.UserSock.State = sckClosed Then
MsgBox "Not Connected"
Exit Sub
End If

Form1.UserSock.SendData "%" & string1 & "%" & string2 & "%" & string3 & "%" & string4 'saves the name,address,location to the current record.

End Sub




