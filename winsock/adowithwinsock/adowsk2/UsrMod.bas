Attribute VB_Name = "ServerMod"
Public Sub ParseRecv(RX As Variant)
Dim i As Integer
Dim ProcessUsr, ProcessPass As String
Dim NewJob As String
Dim SaveJob(12) As String
Dim SaveEditJob(14) As String
Dim EditJob As String
Dim GetJobNumber, RealJobNumber, PortNumber As String
Dim DelJob As Long
Debug.Print RX
If Mid(RX, 1, 10) = "VerifyUser" Then
        FrmServer.InitMax = Mid(RX, 11)
        FrmServer.sckServer(FrmServer.MaxCN).SendData "ShowAuthFrm" & FrmServer.MaxCN ' Shows User Authentication Screen
 
    Else
    
If Mid(RX, 1, 8) = "UserName" Then
        ProcessUsr = Split(RX, "~~")(1)
        ProcessPass = Split(RX, "~~")(2)
        PortNumber = Split(RX, "~~")(3)
            FrmServer.InitMax = PortNumber
            Call Validate(ProcessUsr, ProcessPass)  'Validates the user against the database
    End If: End If
    
If Mid(RX, 1, 8) = "ShowJobs" Then
        Call SendJobs(False)                     'Sends all jobs that are not finished.
            FrmServer.InitMax = Mid(RX, 9)      'Tells the server who we are sending the data too.
Else

If Mid(RX, 1, 17) = "ShowCompletedJobs" Then
        Call SendJobs(True)                    'Sends all jobs that are Completed
            FrmServer.InitMax = Mid(RX, 18)
                 
End If: End If

        If Mid(RX, 1, 10) = "GetRsCount" Then
                Call RecordCount                    'Client has requested how many records are in the recordset.
                 FrmServer.InitMax = Mid(RX, 11)
                                
        
        End If

        If Mid(RX, 1, 9) = "ListUsers" Then
                Call ListedUsers
'                FrmServer.InitMax = Mid(RX, 10)
                                   
        End If


If Mid(RX, 1, 2) = "~@" Then
    NewJob = Mid(RX, 3)
         For i = 0 To 12
            SaveJob(i) = Split(NewJob, "~~")(i)
            
               If Len(SaveJob(12)) Then FrmServer.InitMax = SaveJob(12)
                    Next i
                          Call RsAddNew(SaveJob(0), SaveJob(1), SaveJob(2), SaveJob(3), SaveJob(4), SaveJob(5), _
                          SaveJob(6), SaveJob(7), SaveJob(8), SaveJob(9), SaveJob(10), SaveJob(11))

End If

If Mid(RX, 1, 9) = "JobNumber" Then
    GetJobNumber = Mid(RX, 10)
            PortNumber = Split(GetJobNumber, "~~")(1)          'Recieves the jobnumber from the frmclient
            RealJobNumber = Split(GetJobNumber, "~~")(0)       'listview control and passes to the server
                FrmServer.JobNumber = RealJobNumber            'to query and send the details back.
                FrmServer.InitMax = PortNumber
                Call SendEditJob
End If

If Mid(RX, 1, 2) = "~%" Then                'Saves the information from the frmeditjob
    EditJob = Mid(RX, 3)
        For i = 0 To 14
            SaveEditJob(i) = Split(EditJob, "~~")(i)
                                            'Send this to the Database to save the current record
        Next i
            If Len(SaveEditJob(14)) Then FrmServer.InitMax = SaveEditJob(14)
            Call RsEditJob(SaveEditJob(0), SaveEditJob(1), SaveEditJob(2), SaveEditJob(3), SaveEditJob(4), SaveEditJob(5), _
                           SaveEditJob(6), SaveEditJob(7), SaveEditJob(8), SaveEditJob(9), SaveEditJob(10), SaveEditJob(11), _
                           SaveEditJob(12), SaveEditJob(13))

End If


If Mid(RX, 1, 12) = "DeleteRecord" Then             'Deletes the Record
    DelJob = Mid(RX, 13, 13)
          RsDel (DelJob)
        
End If

End Sub




