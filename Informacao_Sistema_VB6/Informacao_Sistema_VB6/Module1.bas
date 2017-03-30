Attribute VB_Name = "Module1"

Dim FSO As FileSystemObject

Dim clmX            As ColumnHeader
Dim itmAdd          As ListItem


Public Sub SetUp()
On Error Resume Next

Set FSO = New FileSystemObject

With Form1

    .Height = 8905
    
    .ListView1.View = 3
    .ListView2.View = 3
    
    .Picture1.Height = 8735
    
    .Command1.Left = -5
    .Command1.Top = -5
    .Command1.Width = .Picture1.Width / 15 + 10
    .Command1.Height = .Picture1.Height / 15 + 10
    
    .VScroll1.Top = 0
    .VScroll1.Height = .Picture1.ScaleHeight
    .VScroll1.Left = .Picture1.ScaleWidth - .VScroll1.Width
    .VScroll1.Min = -5
    .VScroll1.Max = 500
    .VScroll1.SmallChange = 10
    .VScroll1.LargeChange = 20
    
End With

If FSO.FileExists((App.Path) & "\progress.htm") Then

    Form2.WebBrowser1.Navigate ((App.Path) & "/" & "progress.htm")
    
Else

Set a = FSO.CreateTextFile((App.Path) & "\progress.htm", True)

    a.Write ("<htm><head><title>progresso</title></head><p>&nbsp;&nbsp;<img border=""0"" src=""YJNHXU287296.GIF"" width=""32"" height=""32""></p></htm>")
    
    a.Close

    Form2.WebBrowser1.Navigate ((App.Path) & "/" & "progress.htm")

End If
  
With Form1.ListView1.ColumnHeaders

Set clmAdd = .Add(Text:="Domain")
Set clmAdd = .Add(Text:="Account")
Set clmAdd = .Add(Text:="Description")
Set clmAdd = .Add(Text:="Password Required")
Set clmAdd = .Add(Text:="Status")

End With

With Form1.ListView2.ColumnHeaders

Set clmAdd = .Add(Text:="Partition")
Set clmAdd = .Add(Text:="Size")
Set clmAdd = .Add(Text:="BootPartition")
Set clmAdd = .Add(Text:="NumberOfBlocks")
Set clmAdd = .Add(Text:="PrimaryPartition")
Set clmAdd = .Add(Text:="Description")
Set clmAdd = .Add(Text:="StartingOffSet")

End With

Set a = Nothing
Set FSO = Nothing

End Sub

Public Sub WMISystemDetails()
On Error Resume Next

Dim DomainLength        As Integer
Dim WshNetwork          As WshNetwork

Set WNetwork = New WshNetwork

Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

Set UAcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount", , 48)
Set TZcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_TimeZone", , 48)
Set PRcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor", , 48)
Set SDcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_SystemDriver", , 48)
Set LDcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk", , 48)
Set SBcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS", , 48)
Set BCcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration", , 48)
Set DPcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskPartition", , 48)
Set DDcolItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive", , 48)
Set OScolItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", , 48)

For Each objItem In UAcolItems

Set itmAdd = Form1.ListView1.ListItems.Add(Text:=(objItem.Domain))

    DomainLength = Len(objItem.Domain) + 2

    itmAdd.SubItems(1) = Mid((objItem.Caption), DomainLength, Len(objItem.Caption))

    itmAdd.SubItems(2) = objItem.Description
    
    itmAdd.SubItems(3) = objItem.PasswordRequired
    
    itmAdd.SubItems(4) = objItem.Status
    
With Form1

    .Text1(4) = WNetwork.UserName
    .Text1(5) = objItem.Domain
    
End With

Next objItem

For Each objItem In TZcolItems
   
    Form1.Text1(6) = objItem.Caption
 
Next objItem

For Each objItem In SBcolItems

If IsNull(objItem.BIOSVersion) Then

    Form1.Text1(2) = "Error"
Else

    Form1.Text1(2) = Join(objItem.BIOSVersion, ",")
    
End If

    Form1.Text1(7) = objItem.SerialNumber
    
    Form1.Text1(3) = objItem.Manufacturer
    
    Form1.Text1(12) = objItem.CurrentLanguage
    
    Form1.Text1(13) = Join(objItem.ListOfLanguages, ",")

Next objItem

For Each objItem In BCcolItems
   
    Form1.Text1(8) = objItem.BootDirectory
    
    Form1.Text1(9) = objItem.Caption
    
    Form1.Text1(10) = objItem.ConfigurationPath
    
    Form1.Text1(11) = objItem.ScratchDirectory
    
Next objItem

For Each objItem In DPcolItems

Set itmAdd = Form1.ListView2.ListItems.Add(Text:=(objItem.Caption))

    itmAdd.SubItems(1) = objItem.Size
    itmAdd.SubItems(2) = objItem.BootPartition
    itmAdd.SubItems(3) = objItem.NumberOfBlocks
    itmAdd.SubItems(4) = objItem.PrimaryPartition
    itmAdd.SubItems(5) = objItem.Description
    itmAdd.SubItems(6) = objItem.StartingOffset
    
Next objItem

For Each objItem In LDcolItems

If objItem.FileSystem <> vbNullString Then

    FileSystemData = objItem.FileSystem

Else

    FileSystemData = "NoData"

End If

If objItem.VolumeName <> vbNullString Then

    VolData = objItem.VolumeName

Else

    VolData = "NoData"

End If

If objItem.VolumeSerialNumber <> vbNullString Then

    SerialData = objItem.VolumeSerialNumber

Else

    SerialData = "NoData"

End If
                                                                      
    Form1.List1.AddItem objItem.Caption & " [" & FileSystemData & "] " & " [" & VolData & "] " & " [ SERIAL: " & SerialData & " ] " & " [" & objItem.Description & "]"
    
Next objItem

For Each objItem In DDcolItems
  
    Form1.List2.AddItem objItem.Caption & " [ SERIAL: " & Trim(objItem.SerialNumber) & " ]"
   
Next objItem

For Each objItem In OScolItems

With Form1

    .Text1(14) = objItem.Caption
    .Text1(15) = objItem.BuildType
    .Text1(16) = objItem.BuildNumber
    .Text1(17) = objItem.FreePhysicalMemory
    .Text1(18) = objItem.FreeVirtualMemory
    .Text1(19) = objItem.DataExecutionPrevention_32BitApplications
    .Text1(20) = objItem.OSArchitecture
    .Text1(21) = objItem.SerialNumber
    .Text1(22) = objItem.ServicePackMajorVersion
    
End With

Next objItem

Set FSO = Nothing

Set WNetwork = Nothing

Set objWMIService = Nothing

Set UAcolItems = Nothing
Set TZcolItems = Nothing
Set PRcolItems = Nothing
Set SDcolItems = Nothing
Set LDcolItems = Nothing
Set SBcolItems = Nothing
Set BCcolItems = Nothing
Set DPcolItems = Nothing
Set DDcolItems = Nothing
Set OScolItems = Nothing

End Sub

Public Sub Main()
On Error Resume Next
    Load Form1

    Load Form2
    Form2.Show
    
    Form1.Timer2.Enabled = True
    
End Sub













