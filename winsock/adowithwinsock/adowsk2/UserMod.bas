Attribute VB_Name = "UserMod"
'This Module is does all the dirty work, it basically filters the incomming data
'from the winsock control.
'This method of filtering the data, probably isn't the best way but however it
'does the job. if you know of a easier and realible way than this please feel free to
'alter the code and send it back to chris@hatton.com.

Public Sub ParseRecv(RX As Variant)
Dim SortRecords, ListRecords, ListPority, SortPority, UserList, AddUsers As String
Dim EditImport(13) As Variant

If Mid(RX, 1, 11) = "ShowAuthFrm" Then
    FrmClient.ConCurrent = Mid(RX, 12)              'server has requested authorization
    FrmAuthentic.Show 1
   
End If

If Mid(RX, 1, 18) = "Password Validated" Then
    With FrmAuthentic
    .Hide
    FrmClient.StatusBar1.Panels.Item(1).Text = "Password Validated"
    
    .Label4.Visible = True
    .Label4.Caption = "Password Validated"
    FrmClient.AuthCompleted = True              'Tells the client that you have
    FrmClient.ShowJobs                          'verified the username and password.

End With
End If




If Mid(RX, 1, 16) = "InValid UserName" Then
    
    MsgBox "Incorrect UserName", vbCritical, "UserName Error"

With FrmAuthentic
    .Text1(0).Enabled = True                            'Incorrect Username
    .Text1(1).Enabled = True
    .Text1(1).Visible = True
    .Text1(0).Visible = True
    .Label1.Visible = True
    .Label2.Visible = True
    .cmdOK.Enabled = True
    .Label3.Visible = False
    .Label4.Caption = "UserName Error"
End With

End If


If Mid(RX, 1, 16) = "InValid Password" Then
    
    MsgBox "Incorrect Password", vbCritical, "Password Error"

With FrmAuthentic
    .Text1(0).Enabled = True
    .Text1(1).Enabled = True
    .Text1(1).Visible = True                                'incorrect Password
    .Text1(0).Visible = True
    .Label1.Visible = True
    .Label2.Visible = True
    .cmdOK.Enabled = True
    .Label3.Visible = False
    .Label4.Caption = "Password Error"
End With

End If

If Mid(RX, 1, 3) = "RS~" Then FrmClient.RScount = Mid(RX, 4, 10)  'Establish how many records in the database
If Mid(RX, 1, 11) = "RecordSaved" Then Unload FrmNewJob 'Unload the new job form once it has been saved.
If Mid(RX, 1, 15) = "RecordEditSaved" Then Unload FrmEditJob
If Mid(RX, 1, 11) = "RecordError" Then MsgBox "We have a Problem" & vbNewLine _
                                                & "Could not save record on the server" & vbNewLine _
                                                & "Contact your Administrator", vbCritical, "Server Error"


If Mid(RX, 1, 2) = "~!" Then


    For i = 1 To FrmClient.RScount

On Error GoTo SkipRecord


SortRecords = Split(RX, "~!")(i)                    'sorts out the recordset
SortPority = Split(RX, "~!")(i)                     'finds out what the pority is

     
ListRecords = Split(SortRecords, "~~")(0)
 ListPority = Split(SortPority, "~~")(6)
    If ListPority = "High" Then FrmClient.ListView1.ListItems.Add , , ListRecords, , 4
            If ListPority = "Med" Then FrmClient.ListView1.ListItems.Add , , ListRecords, , 5

       
            If ListPority = "" Then FrmClient.ListView1.ListItems.Add , , ListRecords
        
        
FrmClient.StatusBar1.Panels.Item(1).Picture = _
FrmClient.ImageList1.ListImages.Item(1).Picture
FrmClient.StatusBar1.Panels.Item(1).Text = "Downloading " & "((" & i & ") of (" & FrmClient.RScount & ")) Records"
    
        With FrmClient.ListView1
             
             
             ListRecords = Split(SortRecords, "~~")(1)
                    .ListItems(i).ListSubItems.Add , , ListRecords  'Job Number
             ListRecords = Split(SortRecords, "~~")(2)
                    .ListItems(i).ListSubItems.Add , , ListRecords  'Job Date
             ListRecords = Split(SortRecords, "~~")(3)
                    .ListItems(i).ListSubItems.Add , , ListRecords  'Name
             ListRecords = Split(SortRecords, "~~")(4)
                    .ListItems(i).ListSubItems.Add , , ListRecords  'Phone
             ListRecords = Split(SortRecords, "~~")(5)
                    .ListItems(i).ListSubItems.Add , , ListRecords  'Description
             ListRecords = Split(SortRecords, "~~")(6)
                    .ListItems(i).ListSubItems.Add , , ListRecords  'Tech
                    
                FrmClient.StatusBar1.Panels.Item(1).Picture = _
                FrmClient.ImageList1.ListImages.Item(3).Picture
                FrmClient.StatusBar1.Panels.Item(1).Text = _
                "Downloaded " & "((" & i & ") of (" & FrmClient.RScount & ")) Records" 'Statistics
        End With
 
            
 
 Next i

            
SkipRecord:     'if the job is completed then skip it, else if we what uncompleted jobs then skip completed jobs



End If
           

If Mid(RX, 1, 2) = "~$" Then
On Error GoTo StopTransfer
    UserList = Mid(RX, 1)
   
   
        FrmNewJob.Combo1.Clear
        FrmNewJob.Combo2.Clear
        FrmNewJob.Combo3.Clear
        FrmEditJob.Combo1.Clear
        FrmEditJob.Combo2.Clear                 'clears the new job & edit job combo boxes
        FrmEditJob.Combo3.Clear
   
        FrmNewJob.Combo3.AddItem "Workshop"
        FrmNewJob.Combo3.AddItem "Onsite"
        FrmEditJob.Combo3.AddItem "Workshop"
        FrmEditJob.Combo3.AddItem "Onsite"
            
            For i = 1 To 10             'Alter this value if you require more users.
    
    AddUsers = Split(UserList, "~$")(i) 'gets the users from the server and splits it up
                                        'and sends it to the new job/edit job form.
        
        
        FrmNewJob.Combo1.AddItem AddUsers
        FrmNewJob.Combo2.AddItem AddUsers
        FrmEditJob.Combo1.AddItem AddUsers
        FrmEditJob.Combo2.AddItem AddUsers
            
            Next i
        
        
        
StopTransfer:


End If

If Mid(RX, 1, 2) = "~&" Then

    For i = 1 To 13
        EditImport(i) = Split(RX, "~~")(i)
            With FrmEditJob
              If Len(EditImport(1)) Then .Text1(3).Text = EditImport(1)  'date
                If Len(EditImport(2)) Then .Text1(0).Text = EditImport(2) 'Name
                If Len(EditImport(5)) Then .Text1(5).Text = EditImport(3) 'Phone
                If Len(EditImport(4)) Then .RichTextBox1.Text = EditImport(4) 'Job Description
                If EditImport(6) = "High" Then
                    .Check1.Value = 1
                        Else
                    .Check1.Value = 0 'Top Pority
                End If
                If EditImport(6) = "Med" Then
                    .Check2.Value = 1
                        Else
                    .Check2.Value = 0 'Medium Pority
                End If
                
                    If EditImport(7) = True Then .Check3.Value = 1  'completedjob
                If Len(EditImport(8)) Then .RichTextBox2 = EditImport(8) 'Completed Description
                If Len(EditImport(9)) Then .Text1(4).Text = EditImport(9) 'Required date
                If Len(EditImport(10)) Then .Text1(1).Text = EditImport(10) 'Address1
                If Len(EditImport(11)) Then .Text1(2).Text = EditImport(11) 'Address2
                               
            End With
            
    Next i

            With FrmEditJob
            
             If Len(EditImport(12)) Then .Combo1.AddItem EditImport(12): .Combo1.Text = EditImport(12)  'Bookedby
             If Len(EditImport(13)) Then .Combo3.AddItem EditImport(13): .Combo3.Text = EditImport(13)  'Location
             If Len(EditImport(5)) Then .Combo2.AddItem EditImport(5):   .Combo2.Text = EditImport(5) 'Techinican
             

            End With
            'FrmClient.sckClient(FrmClient.MaxCN).SendData "ListUsers" & FrmClient.ConCurrent
End If




End Sub
