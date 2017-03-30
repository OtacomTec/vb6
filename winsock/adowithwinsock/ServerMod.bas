Attribute VB_Name = "ServerMod"
Sub ServerSck()

If Form1.ServerSock(0).State = sckConnected Then Form1.Label2.Caption = "Network Status:   Connection Established"
If Form1.ServerSock(0).State = sckClosed Then Form1.Label2.Caption = "Network Status:   Connection Closed"

End Sub

Sub SendData(Qry As String)
On Error Resume Next
Dim i As Integer
For i = 0 To 10
DoEvents
Form1.ServerSock(i).SendData Qry

Next i

End Sub
Sub GenerateSQL(Var As String)
Dim Qry As String
Set rs = New ADODB.Recordset

Qry = "Select * from users where name = " & Chr(34) & Var & Chr(34) 'Simply goes to the record that i client requested

rs.Open Qry, cn, adOpenKeyset, adLockReadOnly

SendData "QRYNAME" & rs!Name & "QRYNAME" & rs!Address & "QRYNAME" & rs!Location & "QRYNAME" & rs!Comments & "QRYNAME"
                                                                        'make sure you add "QRYNAME" to the end of the list as it saves errors.
rs.Close

Set rs = Nothing
End Sub


Sub SaveChanges(bit1 As String, bit2 As String, bit3 As String, bit4 As String)
If bit1 = "" Then Exit Sub
Set rs = New ADODB.Recordset
Dim Qry As String

Qry = "Select * from users where name = " & Chr(34) & bit1 & Chr(34)    'this procedure does all the hardwork in saving
rs.Open Qry, cn, adOpenKeyset, adLockOptimistic                         'any updates to the recordset.

If Len(bit1) Then rs!Name = bit1 Else rs!Name = Null
If Len(bit2) Then rs!Address = bit2 Else rs!Address = Null
If Len(bit3) Then rs!Location = bit3 Else rs!Location = Null
If Len(bit3) Then rs!Comments = " " & bit4 Else rs!Comments = ""


rs.Update
rs.Close
Set rs = Nothing


End Sub

Sub DelRecord(Record As String)
If Record = "" Then Exit Sub

Set rs = New ADODB.Recordset
Dim Qry As String


Qry = "Select * from users where name = " & Chr(34) & Record & Chr(34)
rs.Open Qry, cn, adOpenKeyset, adLockOptimistic

                                                               

rs.Delete    'this deletes the record selected by the client machine.
rs.Update
rs.Close
Set rs = Nothing


End Sub


Sub AddUser(String1 As String, String2 As String, String3 As String)
If String1 = "" Then Exit Sub

Set rs = New ADODB.Recordset
Dim Qry As String


Qry = "Select * from users "
rs.Open Qry, cn, adOpenKeyset, adLockOptimistic
rs.AddNew                                           'add the new user to the database

rs!Name = String1
rs!Address = String2
rs!Location = String3


rs.Update
rs.Close
Set rs = Nothing
End Sub


