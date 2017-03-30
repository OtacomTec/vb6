Attribute VB_Name = "Recordsets"
Public Sub RsDel(RecordID As Long)

If RecordID = Null Then Exit Sub

Dim Rs As ADODB.Recordset
Dim SQL As String
 
Set Rs = New ADODB.Recordset

SQL = "select * from Joblist where jobid = " & RecordID

Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic


Rs.Delete adAffectCurrent

Rs.Update
Rs.Close

Set Rs = Nothing

If Len(DetectError) = 0 Then
    FrmServer.sckServer(FrmServer.InitMax).SendData "RecordEditSaved" 'Tells the client to close
                                                                      'there new job form
Else

    FrmServer.sckServer(FrmServer.InitMax).SendData "RecordError" 'Tells the client that theres
FrmServer.Label7.Caption = Err.Description                        'there new job form
End If


End Sub
Public Sub RsEditJob(Name, Addy1, Addy2 As String, Jobdate, DateRequired, _
JobDescription, Bookedby, Tech, Location, Top, Med, Phone, ComDescription, _
Completed As String)


Dim DetectError As String

If Name = "" Then Exit Sub


Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset            'Saves the Edited job to the database

SQL = "Select * from Joblist where jobid = " & FrmServer.JobNumber

Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

If Len(Name) Then Rs!Name = Name
If Len(Addy1) Then Rs!Address1 = Addy1
If Len(Addy2) Then Rs!Address2 = Addy2
If Len(Jobdate) Then Rs!Date = Jobdate
If Len(DateRequired) Then Rs!DateRequired = DateRequired
If Len(JobDescription) Then Rs!JobDescription = JobDescription
If Len(Bookedby) Then Rs!Bookedby = Bookedby
If Len(Tech) Then Rs!Technician = Tech
If Len(Location) Then Rs!Location = Location
If Len(Phone) Then Rs!Phone = Phone
If Len(ComDescription) Then Rs!ComDescription = ComDescription
If Completed = 1 Then Rs!Completedjobs = 1 Else Rs!Completedjobs = 0

If Med = 1 Then
    Rs!pority = "Med"
    Else
If Top = 0 Then
    Rs!pority = Null                    'If top pority = 1 or med pority = 1 then
 End If                                 'save it
    End If
            
If Top = 1 Then
    Rs!pority = "High"
    Else
If Med = 0 Then
    Rs!pority = Null
End If
    End If
    
DetectError = Err.Description

Rs.Update
Rs.Close

Set Rs = Nothing


If Len(DetectError) = 0 Then
    FrmServer.sckServer(FrmServer.InitMax).SendData "RecordEditSaved" 'Tells the client to close
                                                                      'there new job form
Else

    FrmServer.sckServer(FrmServer.InitMax).SendData "RecordError" 'Tells the client that theres
FrmServer.Label7.Caption = Err.Description                        'there new job form
End If



End Sub
Public Sub RsAddNew(Name, Addy1, Addy2 As String, Jobdate, DateRequired, _
JobDescription, Bookedby, Tech, Location, Top, Med, Phone As String)

Dim DetectError As String
On Error Resume Next
If Name = "" Then Exit Sub


Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

Rs.Open "select * from joblist", cn, adOpenKeyset, adLockOptimistic
Rs.AddNew

If Len(Name) Then Rs!Name = Name                     'If the values are not zero length
If Len(Addy1) Then Rs!Address1 = Addy1               'then add it to the database
If Len(Addy2) Then Rs!Address2 = Addy2
If Len(Jobdate) Then Rs!Date = Jobdate
If Len(DateRequired) Then Rs!DateRequired = DateRequired
If Len(JobDescription) Then Rs!JobDescription = JobDescription
If Len(Bookedby) Then Rs!Bookedby = Bookedby
If Len(Tech) Then Rs!Technician = Tech
If Len(Location) Then Rs!Location = Location
If Top = 1 Then Rs!pority = "High"
If Med = 1 Then Rs!pority = "Med"
If Len(Phone) Then Rs!Phone = Phone
DetectError = Err.Description
Rs.Update
Rs.Close

Set Rs = Nothing


If Len(DetectError) = 0 Then
    FrmServer.sckServer(FrmServer.InitMax).SendData "RecordSaved" 'Tells the client to close
                                                                  'there new job form
Else

    FrmServer.sckServer(FrmServer.InitMax).SendData "RecordError" 'Tells the client that theres
FrmServer.Label7.Caption = Err.Description                        'an error.

End If

End Sub
Public Sub RecordCount()
Dim Rs As ADODB.Recordset
Dim j As Integer
Set Rs = New ADODB.Recordset

Rs.Open "Select * from joblist", cn, adOpenForwardOnly, adLockReadOnly

j = Rs.RecordCount
If FrmServer.InitMax = "0" Then
    FrmServer.sckServer(1).SendData "RS~" & j
Else
    FrmServer.sckServer(FrmServer.InitMax).SendData "RS~" & j
End If

Rs.Close
Set Rs = Nothing

FrmServer.Label7.Caption = Err.Description



End Sub
Public Sub SendEditJob()
Dim Rs As ADODB.Recordset
Dim SQL, EditRecord As String
Dim i As Integer

Set Rs = New ADODB.Recordset

SQL = "Select * from Joblist where jobid = " & FrmServer.JobNumber

Rs.Open SQL, cn, adOpenKeyset, adLockReadOnly


    For i = 1 To Rs.RecordCount

        EditRecord = "~&" & "~~" & Rs!Date & "~~" & Rs!Name & "~~" & Rs!Phone & _
        "~~" & Rs!JobDescription & "~~" & Rs!Technician & "~~" & Rs!pority & "~~" & _
        Rs!Completedjobs & "~~" & Rs!ComDescription & "~~" & Rs!DateRequired & "~~" & _
        Rs!Address1 & "~~" & Rs!Address2 & "~~" & Rs!Bookedby & "~~" & Rs!Location

        Rs.MoveNext
           
            FrmServer.sckServer(FrmServer.InitMax).SendData EditRecord
    Next i

FrmServer.Label7.Caption = Err.Description

Rs.Close
Set Rs = Nothing
End Sub

Public Sub SendJobs(Completed As Boolean)

Dim Rs As ADODB.Recordset
Dim SQL, strRecords As String
Dim i As Integer

Set Rs = New ADODB.Recordset

SQL = "Select * from Joblist where CompletedJobs = " & Completed  'Querys the Database for Completed and UnCompleted jobs

Rs.Open SQL, cn, adOpenKeyset, adLockReadOnly


    For i = 1 To Rs.RecordCount

        strRecords = "~!" & Rs!JobID & "~~" & Rs!Date & "~~" & Rs!Name & "~~" & Rs!Phone & _
        "~~" & Rs!JobDescription & "~~" & Rs!Technician & "~~" & Rs!pority & "~~"

        Rs.MoveNext
           
            FrmServer.sckServer(FrmServer.InitMax).SendData strRecords
    Next i


Rs.Close
Set Rs = Nothing
FrmServer.Label7.Caption = Err.Description

End Sub

Public Sub Validate(UserName, Password As String)
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim SQL As String

SQL = "Select * from Authentication where username = " & Chr(34) & UserName & Chr(34)

Rs.Open SQL, cn, adOpenKeyset, adLockReadOnly

With FrmServer
If Rs.EOF = True Then .sckServer(.InitMax).SendData "InValid UserName": Exit Sub
    If UCase(Password) = UCase(Rs!Password) Then
        .sckServer(.InitMax).SendData "Password Validated"
            .lvusrs.ListItems.Add , , UserName & "/" & .lvusrs.ListItems.Count + 1
    Else
        .sckServer(.InitMax).SendData "InValid Password"
    End If

End With

Rs.Close
Set Rs = Nothing
FrmServer.Label7.Caption = Err.Description

End Sub

Public Sub ListedUsers()

Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Dim AllUsers, AddUser As String
Dim i As Integer

Rs.Open "Select UserName from Authentication", cn, adOpenForwardOnly, adLockReadOnly

    For i = 0 To Rs.RecordCount - 1

        AllUsers = Split(Rs!UserName, " ")(0)
    
            AddUser = AddUser & "~$" & AllUsers
    
    Rs.MoveNext
    Next i

            FrmServer.sckServer(FrmServer.InitMax).SendData AddUser
Rs.Close
Set Rs = Nothing
FrmServer.Label7.Caption = Err.Description


End Sub
