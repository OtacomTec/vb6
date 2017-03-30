VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailprogram - Version 1.02"
   ClientHeight    =   7590
   ClientLeft      =   630
   ClientTop       =   810
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   720
   Begin VB.CommandButton cmdAdd2Adressbook 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   9960
      TabIndex        =   32
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd2Adressbook 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   4800
      TabIndex        =   31
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox pic16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3240
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtBericht 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7455
   End
   Begin VB.CommandButton cmdAdresboek 
      Caption         =   "Options adressbook"
      Height          =   345
      Left            =   8160
      TabIndex        =   26
      Tag             =   "119"
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton cmdConcept 
      Caption         =   "Open draft"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   3720
      TabIndex        =   25
      Tag             =   "116"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox lstAdresboek 
      Height          =   2790
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton cmdConcept 
      Caption         =   "Delete draft"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6600
      TabIndex        =   21
      Tag             =   "118"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ComboBox cmbConcept 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "115"
      Top             =   6840
      Width           =   2475
   End
   Begin VB.CommandButton cmdConcept 
      Caption         =   "Save draft"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   5160
      TabIndex        =   19
      Tag             =   "117"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdBijlage 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   9285
      TabIndex        =   18
      Top             =   1800
      Width           =   1185
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Tag             =   "niet"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Tag             =   "niet"
      Top             =   7275
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9922
            MinWidth        =   9922
            Text            =   "State: On-line"
            TextSave        =   "State: On-line"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Width           =   132292
            MinWidth        =   132292
            Text            =   "Sending e-mail..."
            TextSave        =   "Sending e-mail..."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAfzender 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Tag             =   "0"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtOntvanger 
      Height          =   285
      Index           =   0
      Left            =   7200
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtOnderwerp 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtAfzender 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtOntvanger 
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox cmbMailServer 
      Height          =   315
      ItemData        =   "frmMain.frx":0ECA
      Left            =   7200
      List            =   "frmMain.frx":11AD
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock SMTP 
      Left            =   2040
      Tag             =   "niet"
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ListView lvAttachments 
      Height          =   300
      Left            =   2040
      TabIndex        =   29
      ToolTipText     =   "select an item, press on [delete] for remove item from attachmentlist"
      Top             =   1815
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   529
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdVerzenden 
      Caption         =   "Send e-mail"
      Enabled         =   0   'False
      Height          =   945
      Left            =   7920
      TabIndex        =   6
      Tag             =   "niet"
      Top             =   2505
      Width           =   2535
   End
   Begin ComctlLib.ImageList img16 
      Left            =   2520
      Tag             =   "niet"
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
   End
   Begin VB.Label lblBericht 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   27
      Tag             =   "112"
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lblAdresboek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adressbook:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   24
      Tag             =   "113"
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label lblConcept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drafts:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Tag             =   "114"
      Top             =   6900
      Width           =   570
   End
   Begin VB.Label lblBijlage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment:"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Tag             =   "109"
      Top             =   1845
      Width           =   900
   End
   Begin VB.Label lblMailServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP-Server:"
      Height          =   195
      Left            =   5520
      TabIndex        =   14
      Tag             =   "110"
      Top             =   1365
      Width           =   990
   End
   Begin VB.Label lblAfzender 
      AutoSize        =   -1  'True
      Caption         =   "Sender:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Tag             =   "102"
      Top             =   240
      Width           =   645
   End
   Begin VB.Label lblAfzenderAdres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From (e-mailadress):"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Tag             =   "103"
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label lblOntvangerAdres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To (e-mailadress):"
      Height          =   195
      Left            =   5520
      TabIndex        =   11
      Tag             =   "106"
      Top             =   525
      Width           =   1305
   End
   Begin VB.Label lblOnderwerp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Tag             =   "108"
      Top             =   1365
      Width           =   600
   End
   Begin VB.Label lblAfzenderNaam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Tag             =   "104"
      Top             =   885
      Width           =   465
   End
   Begin VB.Label lblOntvangerNaam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   5520
      TabIndex        =   8
      Tag             =   "107"
      Top             =   885
      Width           =   465
   End
   Begin VB.Label lblOntvanger 
      AutoSize        =   -1  'True
      Caption         =   "Recipient:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   7
      Tag             =   "105"
      Top             =   240
      Width           =   840
   End
   Begin VB.Menu mnuAdresboek1 
      Caption         =   "Adressbook"
      Visible         =   0   'False
      Begin VB.Menu mnuAdresboek 
         Caption         =   "New contact..."
         Index           =   0
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "Add contact"
         Index           =   2
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "Delete contact"
         Index           =   4
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "Change contact..."
         Index           =   5
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "Delete all contacts"
         Index           =   7
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuAdresboek 
         Caption         =   "Import..."
         Index           =   9
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New message"
   End
   Begin VB.Menu mnuSep1 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuProblem 
      Caption         =   "If you got a problem click here..."
   End
   Begin VB.Menu mnuVote 
      Caption         =   "Vote..."
   End
   Begin VB.Menu mnuSep2 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                              'Force variable declaration

Private mintKeyCode As Integer               'Declare mintkeycode for local use as integer
Public strConcept As String                  'Declare strconcept with max scope as string
Dim blnNieuwConcept As Boolean               'Declare blnnieuwconcept for local use as boolean
Public Adresboek As New Collection           'Declare adresboek with max scope as new collection
Dim reactie As String, checkState As Long    'Declare reactie for local use as string, checkstate as long
Dim Vote As Boolean

Public Sub Pauze(Interval As Single, Optional teken As String = vbNullString)
    Dim Tijd As Single                       'Declare tijd for local use as single
    Tijd = Timer + Interval
    Do Until Timer > Tijd Or Left$(reactie, 3) = teken: DoEvents: Loop
End Sub

Private Sub cmbMailServer_Change()
    Call AutoCompleteList(cmbMailServer, mintKeyCode)

    If Trim$(cmbMailServer.Text) = vbNullString Then
        cmdVerzenden.Enabled = False         'disable the cmdverzenden object
    Else
        If Trim$(txtAfzender(0).Text) <> vbNullString And Trim$(txtAfzender(1).Text) <> vbNullString And _
            Trim$(txtOntvanger(0).Text) <> vbNullString And Trim$(txtOntvanger(1).Text) <> vbNullString And _
            InStr(1, txtAfzender(0).Text, "@") > 0 And InStr(1, txtOntvanger(0).Text, "@") > 0 Then
                cmdVerzenden.Enabled = True  'enable the cmdverzenden object
        End If
    End If
End Sub

Public Sub AutoCompleteList(ByVal cboCtl As ComboBox, ByVal KeyCode As Integer)

    Dim Counter As Integer                   'Declare counter for local use as integer
    Dim Length As Integer                    'Declare length for local use as integer

    If cboCtl <> "" Then
        For Counter = 0 To cboCtl.ListCount - 1
            Length = Len(cboCtl.Text)

            If Mid(cboCtl.List(Counter), 1, Len(cboCtl)) = cboCtl.Text Then
                cboCtl.Text = cboCtl.List(Counter)
                cboCtl.SelStart = Length
                cboCtl.SelLength = Len(cboCtl.Text)
                cboCtl.ListIndex = Counter
                Exit For                     'Leave this for
            End If
        Next
    End If

End Sub

Private Sub cmbMailServer_Click()
    cmbMailServer_Change
End Sub

Private Sub AdresboekImporteren()            'import adressbook
    Dim bestand As String, tekst As String, info() As String, i As Integer 'Declare bestand for local use as string, tekst as string, info() as string, i as integer

    bestand = modCD.DialogFile(Me, 1, "Open adressbook", "", "Adressbook (*.csv)" & Chr(0) & "*.csv", CurDir, "csv")
    If bestand <> vbNullString Then
        Open bestand For Input As #1         'Open the file bestand to read from
            Line Input #1, tekst             'Read next line from file #1 into tekst
            info = Split(tekst, ";")
            If LCase$(info(0)) = "naam" Then 'And LCase$(Info(1)) = "e-mailadres" Then
                Open App.Path & "\Adressbook.dat" For Output As #3 'Open the file app.path & "\adressbook.dat" to write in

                    For i = 1 To Adresboek.count
                        Print #3, Adresboek(i) 'Write adresboek(i) in file #3
                    Next i

                    While Not EOF(1)
                        Line Input #1, tekst 'Read next line from file #1 into tekst
                        Print #3, tekst      'Write tekst in file #3
                    Wend

                Close #3                     'Close file #3

            Else: MsgBox "Incorrect fileformat.", vbCritical, "Adressbook": Close #1: Exit Sub 'Inform the user with a messagebox
            End If
        Close #1                             'Close file #1

        AdresboekLaden
    End If
End Sub

Public Sub AdresboekLaden()
    Dim tekst As String, i As Integer, info() As String 'Declare tekst for local use as string, i as integer, info() as string

    lstAdresboek.Clear                       'Clear the lstadresboek object
    If Dir$(App.Path & "\Adressbook.dat", vbNormal) = vbNullString Then lstAdresboek.Enabled = False: Exit Sub 'disable: exit sub the if dir$(app.path & "\adressbook.dat", vbnormal) = vbnullstring then lstadresboek object

    Open App.Path & "\Adressbook.dat" For Input As #3 'Open the file app.path & "\adressbook.dat" to read from
        While Not EOF(3)
            Line Input #3, tekst             'Read next line from file #3 into tekst
            info = Split(tekst, ";")
            If Trim$(info(0)) = vbNullString And Trim$(info(1)) = vbNullString Then
                lstAdresboek.AddItem info(1) & ";" & info(1)
            Else: lstAdresboek.AddItem info(0) & ";" & info(1)
            End If
        Wend
    Close #3                                 'Close file #3

    For i = 0 To lstAdresboek.ListCount - 1
        Adresboek.Add lstAdresboek.List(i)
        lstAdresboek.List(i) = Left$(lstAdresboek.List(i), InStr(1, lstAdresboek.List(i), ";") - 1)
    Next i

    lstAdresboek.Enabled = True              'enable the lstadresboek object
End Sub

Private Sub cmbMailServer_LostFocus()
    OpslaanInstellingen cmbMailServer.Text, "Settings", "Mailserver", App.Path & "\Settings.dat"
End Sub

Private Sub cmdAdd2Adressbook_Click(Index As Integer)
    Dim name As String, email As String
    
    Select Case Index
        Case 0
            name = txtAfzender(1).Text: email = txtAfzender(0).Text
            
        Case 1
            name = txtOntvanger(1).Text: email = txtOntvanger(0).Text
    End Select
    
    
    If MsgBox("Do you want add '" & name & "' (" & email & ") to your adressbook?", vbQuestion + vbYesNo, "Add") = vbYes Then
        AddContact name, email
    End If
End Sub

Private Sub cmdAdresboek_Click()
    If lstAdresboek.ListCount > 0 Then mnuAdresboek(7).Enabled = True Else mnuAdresboek(7).Enabled = False
    
    If lstAdresboek.Text <> vbNullString Then
        mnuAdresboek(2).Enabled = True: mnuAdresboek(4).Enabled = True: mnuAdresboek(5).Enabled = True 'enable: mnuadresboek(4).enabled = enable: mnuadresboek(5).enabled = enable the mnuadresboek(2) object
        mnuAdresboek(2).Caption = "Add '" & lstAdresboek.Text & "'"
        mnuAdresboek(4).Caption = "Delete '" & lstAdresboek.Text & "'"
        mnuAdresboek(5).Caption = "Change '" & lstAdresboek.Text & "'..."
    Else
        mnuAdresboek(2).Enabled = False: mnuAdresboek(4).Enabled = False: mnuAdresboek(5).Enabled = False 'disable: mnuadresboek(4).enabled = disable: mnuadresboek(5).enabled = disable the mnuadresboek(2) object

        mnuAdresboek(2).Caption = "Add (No contact selected)"
        mnuAdresboek(4).Caption = "Delete (No contact selected)"
        mnuAdresboek(5).Caption = "Change (No contact selected)..."
    End If

    PopupMenu mnuAdresboek1, vbPopupMenuCenterAlign, cmdAdresboek.Left + cmdAdresboek.Width / 2, cmdAdresboek.Top + cmdAdresboek.Height
End Sub

Private Sub cmdBijlage_Click()
    Dim bestand As String                    'Declare bestand for local use as string
    Dim grootte As String, i As Integer
    Dim count As Long, info() As String

    bestand = modCD.DialogFile(Me, 1, "Add attachment", "", "All files" & Chr$(0) & "*.*", CurDir, "")
    If bestand <> vbNullString Then

        grootte = FileLen(bestand)
        If grootte > 1024 Then
            If (grootte / 1024) > 1024 Then
                grootte = Round(grootte / 1024 / 1024, 2) & " MB"
            Else: grootte = Round(grootte / 1024, 1) & " kByte"
            End If
        Else: grootte = grootte & " bytes"
        End If
        
        If lvAttachments.ListItems(1).Text = "No attachments" Then lvAttachments.ListItems.Clear
        lvAttachments.SmallIcons = Nothing
        img16.ListImages.Clear
        
        Dim Item As ListItem
        Set Item = lvAttachments.ListItems.Add()
        Item.Text = GetFilename(bestand) & " (" & grootte & ")"
        
        count = 0
        For i = 1 To lvAttachments.ListItems.count
            info = Split(lvAttachments.ListItems(i).Key, "|")
            If UBound(info) > 0 Then If LCase$(info(0)) = LCase$(bestand) Then count = count + 1
        Next i
        Item.Key = bestand & "|" & count
        
        For i = 1 To lvAttachments.ListItems.count
            info = Split(lvAttachments.ListItems(i).Key, "|")
            GetIcon info(0), CLng(i)
        Next i
        lvAttachments.SmallIcons = img16
        For Each Item In lvAttachments.ListItems
            Item.SmallIcon = Item.Index
        Next
        
        If Not lvAttachments.Enabled = True Then lvAttachments.Enabled = True
    End If
End Sub

Private Sub cmdConcept_Click(Index As Integer)
    Dim tekst As String                      'Declare tekst for local use as string

    Select Case Index
        Case 0                               'save
            frmConcept.txtNaam.Text = txtOnderwerp.Text: frmConcept.txtNaam.SelStart = 0: frmConcept.txtNaam.SelLength = Len(txtOnderwerp.Text)
            frmConcept.Show vbModal          'Show the frmconcept object v
            If strConcept <> vbNullString Then
                If Dir$(App.Path & "\Drafts", vbDirectory) = vbNullString Then MkDir App.Path & "\Drafts"
                Open App.Path & "\Drafts\" & strConcept & ".txt" For Output As #1 'Open the file app.path & "\drafts\" & strconcept & ".txt" to write in
                    Print #1, txtAfzender(0).Text 'Write txtafzender(0).text in file #1
                    Print #1, txtAfzender(1).Text 'Write txtafzender(1).text in file #1
                    Print #1, txtOnderwerp.Text 'Write txtonderwerp.text in file #1
                    Print #1, txtBericht.Text 'Write txtbericht.text in file #1
                Close #1                     'Close file #1
                If cmbConcept.Text = "No drafts saved" Then cmbConcept.Clear 'Clear the if cmbconcept.text = "no drafts saved" then cmbconcept object
                cmbConcept.AddItem strConcept
                cmbConcept.ListIndex = cmbConcept.NewIndex
                If cmbConcept.Enabled = False Then cmbConcept.Enabled = True: cmdConcept(1).Enabled = True: cmdConcept(2).Enabled = True 'disable then cmbconcept.enabled = enable: cmdconcept(1).enabled = enable: cmdconcept(2).enabled = enable the if cmbconcept object
            End If

        Case 1                               'delete
            Kill App.Path & "\Drafts\" & cmbConcept.Text & ".txt" 'Delete the file app.path & "\drafts\" & cmbconcept.text & ".txt"
            cmbConcept.RemoveItem cmbConcept.ListIndex
            If cmbConcept.ListCount = 0 Then
                cmbConcept.Enabled = False   'disable the cmbconcept object
                cmbConcept.AddItem "No drafts saved": cmbConcept.ListIndex = 0: cmbConcept.Enabled = False 'disable the cmbconcept.additem "no drafts saved": cmbconcept.listindex = 0: cmbconcept object
                cmdConcept(1).Enabled = False: cmdConcept(2).Enabled = False 'disable: cmdconcept(2).enabled = disable the cmdconcept(1) object
            Else: cmbConcept.ListIndex = 0   'reset else: cmbconcept.listindex to zero
            End If

        Case 2                               'open
            txtBericht.Text = vbNullString
            Open App.Path & "\Drafts\" & cmbConcept.Text & ".txt" For Input As #1 'Open the file app.path & "\drafts\" & cmbconcept.text & ".txt" to read from
                Line Input #1, tekst         'Read next line from file #1 into tekst
                txtAfzender(0).Text = tekst
                Line Input #1, tekst         'Read next line from file #1 into tekst
                txtAfzender(1).Text = tekst
                Line Input #1, tekst         'Read next line from file #1 into tekst
                txtOnderwerp.Text = tekst

                While Not EOF(1)
                    Line Input #1, tekst     'Read next line from file #1 into tekst
                    If txtBericht.Text <> vbNullString Then
                        txtBericht.Text = txtBericht.Text & vbCrLf & tekst 'Add vbcrlf & tekst to txtbericht.text
                    Else: txtBericht.Text = tekst
                    End If
                Wend
            Close #1                         'Close file #1
    End Select
End Sub

Private Sub Form_Activate()
    If Not Screen.Width / Screen.TwipsPerPixelX = 800 Then
        Left = Screen.Width / 2 - Width / 2
        Top = Screen.Height / 2 - Height / 2
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As Variant
    Dim sDirectory As Variant
    
    If Vote = True Then End
    
    If MsgBox("If you like this program, please vote for me! If you want to vote click on yes.", vbQuestion + vbYesNo, "Advanced Mailapplication") = vbYes Then
        OpslaanInstellingen "Yes", "Settings", "Vote", App.Path & "\Settings.dat"
        
        sTopic = "Open"
        sFile = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46547&lngWId=1"
        sParams = 0&
        sDirectory = 0&
        
        Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
    End If

    End                                      'End the program
End Sub

Private Sub lstAdresboek_DblClick()
    If lstAdresboek.ListIndex >= 0 Then
        mnuAdresboek_Click 2
    End If
End Sub

Private Sub lvAttachments_KeyUp(KeyCode As Integer, Shift As Integer)
    If lvAttachments.SelectedItem.Text <> vbNullString And KeyCode = vbKeyDelete Then
        lvAttachments.ListItems.Remove lvAttachments.SelectedItem.Index
        If lvAttachments.ListItems.count = 0 Then
            lvAttachments.ListItems.Add , , "No attachments": lvAttachments.Enabled = False
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "The author is not responsible for any damage that may occure while using the program." & vbCrLf & _
            "If you do not agree to this term then please do not use the program." & vbCrLf & vbCrLf & "Copyright (c) 2003 HB Software Design", vbInformation
End Sub

Private Sub mnuAdresboek_Click(Index As Integer)
    Dim i As Integer, info() As String, naam As String, tekst As String 'Declare i for local use as integer, info() as string, naam as string, tekst as string

    Select Case Index
        Case 0                               'nieuw contactpersoon toevoegen
            frmContactpersoon.Show vbModal   'Show the frmcontactpersoon object v

        Case 1                               '-

        Case 2                               'invoegen
            info = Split(Adresboek(lstAdresboek.ListIndex + 1), ";")
                                             'If txtOntvanger(1).Text <> vbNullString Then
                                             'txtOntvanger(1).Text = txtOntvanger(1).Text & "," & info(0)
                                             'txtOntvanger(0).Text = txtOntvanger(0).Text & "," & info(1)
                                             'Else
                txtOntvanger(1).Text = info(0)
                txtOntvanger(0).Text = info(1)
                                             'End If


        Case 3                               '-

        Case 4                               'verwijderen
            info = Split(Adresboek(lstAdresboek.ListIndex + 1), ";")
            naam = info(0)

            For i = Adresboek.count To 1 Step -1: Adresboek.Remove i: Next i

            Open App.Path & "\Adressbook.dat" For Input As #1 'Open the file app.path & "\adressbook.dat" to read from
                While Not EOF(1)
                    Line Input #1, tekst     'Read next line from file #1 into tekst
                    info = Split(tekst, ";")
                    If Not info(0) = naam Then Adresboek.Add info(0) & ";" & info(1)
                Wend
            Close #1                         'Close file #1

            Open App.Path & "\Adressbook.dat" For Output As #1 'Open the file app.path & "\adressbook.dat" to write in
                For i = 1 To Adresboek.count
                    Print #1, Adresboek(i)   'Write adresboek(i) in file #1
                Next i
            Close #1                         'Close file #1

            AdresboekLaden

        Case 5                               'aanpassen
            info = Split(Adresboek(lstAdresboek.ListIndex + 1), ";")
            frmContactpersoon.txtNaam.Text = info(0)
            frmContactpersoon.txtAdres.Text = info(1)
            frmContactpersoon.Show vbModal   'Show the frmcontactpersoon object v

        Case 6                               '-

        Case 7                               'alle contactpersonen verwijderen
            If MsgBox("Delete all contacts from adressbook." & vbCrLf & vbCrLf & "Are you sure?", vbQuestion + vbYesNo, "Adressbook") = vbNo Then Exit Sub
            
            For i = Adresboek.count To 1 Step -1: Adresboek.Remove i: Next i
            lstAdresboek.Clear               'Clear the lstadresboek object
            Kill App.Path & "\Adressbook.dat" 'Delete the file app.path & "\adressbook.dat"

        Case 8                               '-

        Case 9                               'adresboek importeren
            If MsgBox("It's possible to import a adressbook from Outlook (Express). Follow the steps before clicking on 'OK':" & vbCrLf & vbCrLf & _
                "1. Start Outlook Express" & vbCrLf & "2. Goto menu 'File', menuitem 'Export'" & vbCrLf & _
                "3. Export a textfile" & vbCrLf & "4. Select only the items: 'Name' and 'E-mailadress'" & vbCrLf & _
                "5. Export the adressbook" & vbCrLf & vbCrLf & "Are you ready to continue?", vbQuestion + vbOKCancel, "Adressbook") = vbOK Then

                    AdresboekImporteren
            End If
    End Select
End Sub

Private Sub mnuNew_Click()
    txtAfzender(0).Text = vbNullString: txtAfzender(1).Text = vbNullString
    txtOntvanger(0).Text = vbNullString: txtOntvanger(1).Text = vbNullString
    txtOnderwerp.Text = vbNullString
    lvAttachments.ListItems.Clear: lvAttachments.ListItems.Add , , "No attachments": lvAttachments.Enabled = False
    lblBijlage.Caption = "Attachment:": GecodeerdeBijlage = vbNullString
    txtBericht.Text = vbNullString
End Sub

Private Sub mnuProblem_Click()
    txtAfzender(0).Text = "yourname@hotmail.com"
    txtAfzender(1).Text = "your name"
    txtOntvanger(0).Text = "controleadres@hotmail.com"
    txtOntvanger(1).Text = "Problems? mail me..."
    txtOnderwerp.Text = "I Need Your Help !!!"
    cmbMailServer.Text = "mail.hotmail.com"
    txtBericht.Text = vbCrLf & "Hey," & vbCrLf & vbCrLf & _
        "Sorry, but I got a problem... [your problem]" & vbCrLf & vbCrLf & _
        "Greetings," & vbCrLf & vbCrLf & _
        "[your name]" & vbCrLf & vbCrLf & _
        "!!! Send this e-mail to controleadres@hotmail.com, or click on 'Problems?, mail me...' in the adressbook !!!" & vbCrLf & vbCrLf & _
        "!!! Don't forget to give your e-mailadress, otherwise I can't help you !!!"
End Sub

Private Sub mnuVote_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As Variant
    Dim sDirectory As Variant

    OpslaanInstellingen "Yes", "Settings", "Vote", App.Path & "\Settings.dat"
    Vote = True
    
    sTopic = "Open"
    sFile = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46547&lngWId=1"
    sParams = 0&
    sDirectory = 0&
    
    Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
End Sub

Private Sub txtAfzender_Change(Index As Integer)
    If Trim$(txtAfzender(1).Text) = vbNullString Or InStr(1, txtAfzender(0).Text, "@") = 0 Then
        cmdAdd2Adressbook(0).Enabled = False
        cmdVerzenden.Enabled = False         'disable the cmdverzenden object
    Else
        If Trim$(txtAfzender(1).Text) <> vbNullString And InStr(1, txtAfzender(0).Text, "@") <> 0 Then
            cmdAdd2Adressbook(0).Enabled = True
        End If
        
        If Trim$(txtOntvanger(0).Text) <> vbNullString And InStr(1, txtOntvanger(0).Text, "@") > 0 _
            And Trim$(txtOntvanger(1).Text) <> vbNullString And _
            Trim$(cmbMailServer.Text) <> vbNullString Then
                cmdVerzenden.Enabled = True  'enable the cmdverzenden object
        End If
    End If
End Sub

Private Sub txtAfzender_GotFocus(Index As Integer)
    txtAfzender(Index).SelStart = 0: txtAfzender(Index).SelLength = Len(txtAfzender(Index).Text)
End Sub

Private Sub txtBericht_Change()
    If Trim$(txtBericht.Text) <> vbNullString And Trim$(txtOnderwerp.Text) <> vbNullString Then
        cmdConcept(0).Enabled = True         'enable the cmdconcept(0) object
    Else: cmdConcept(0).Enabled = False      'disable the else: cmdconcept(0) object
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, antwoord As String, gevonden As Boolean 'Declare i for local use as integer, antwoord as string, gevonden as boolean

    PB.Value = 0                             'reset pb.value to zero

    Dim bestand As String                    'Declare bestand for local use as string
    bestand = Dir$(App.Path & "\Drafts\*.txt", vbNormal)
    Do While bestand <> vbNullString
        cmbConcept.AddItem Left$(bestand, InStr(1, bestand, ".txt") - 1)
        bestand = Dir
    Loop

    If cmbConcept.ListCount = 0 Then
        cmbConcept.AddItem "No drafts saved": cmbConcept.Enabled = False: cmdConcept(1).Enabled = False: cmdConcept(2).Enabled = False 'disable: cmdconcept(1).enabled = disable: cmdconcept(2).enabled = disable the cmbconcept.additem "no drafts saved": cmbconcept object
    Else: cmdConcept(1).Enabled = True: cmdConcept(2).Enabled = True 'enable: cmdconcept(2).enabled = enable the else: cmdconcept(1) object
    End If

    cmbConcept.ListIndex = 0                 'reset cmbconcept.listindex to zero

    antwoord = LaadInstellingen("Settings", "Mailserver", App.Path & "\Settings.dat")
    If antwoord <> vbNullString Then
        For i = 0 To cmbMailServer.ListCount - 1
            If LCase$(cmbMailServer.List(i)) = LCase$(antwoord) Then cmbMailServer.ListIndex = i: gevonden = True: Exit For
        Next i
        If gevonden = False Then cmbMailServer.AddItem antwoord: cmbMailServer.ListIndex = cmbMailServer.NewIndex
    Else: cmbMailServer.ListIndex = 0        'reset else: cmbmailserver.listindex to zero
    End If

    antwoord = LaadInstellingen("Settings", "Vote", App.Path & "\Settings.dat")
    If antwoord <> vbNullString Then Vote = True Else Vote = False
    
    AdresboekLaden

    Caption = "Advanced Mailapplication - Version" & Chr$(32) & App.Major & "." & App.Minor & "." & App.Revision
    
    If InternetGetConnectedState(0&, 0&) = 1 Then
        SB.Panels(1).Text = "State: On-line"
    Else: SB.Panels(1).Text = "State: Off-line"
    End If
    
    lvAttachments.ListItems.Add , , "No attachments"
End Sub

'Private Sub ResLaden()
'    Dim info() As String, i As Integer

'    Caption = VB.LoadResString(101) & App.Major & "." & App.Minor & "." & App.Revision

'    For i = 0 To Controls.Count - 1
'        info = Split(Controls(i).Tag, "|")

'            If info(0) <> "niet" And info(0) <> vbNullString Then
'                Select Case Left$(Controls(i).Name, 3)
'                    Case "lbl"
'                        Controls(i).Caption = LoadResString(Controls(i).Tag)
'                    Case "txt"
'                        Controls(i).Text = LoadResString(Controls(i).Tag)
'                    Case "cmb"
'                        Controls(i).Text = LoadResString(Controls(i).Tag)
'                    Case "sb"
'                        Controls(i).Panels(1).Text = LoadResString(Controls(i).Tag)
'                End Select
'            End If

'    Next i
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mintKeyCode = KeyCode
End Sub

Private Sub txtOnderwerp_Change()
    txtBericht_Change
End Sub

Private Sub txtOnderwerp_GotFocus()
    txtOnderwerp.SelStart = 0: txtOnderwerp.SelLength = Len(txtOnderwerp.Text)
End Sub

Private Sub txtOntvanger_Change(Index As Integer)
    If Trim$(txtOntvanger(1).Text) = vbNullString Or InStr(1, txtOntvanger(0).Text, "@") = 0 Then
        cmdVerzenden.Enabled = False         'disable the cmdverzenden object
        cmdAdd2Adressbook(1).Enabled = False
    Else
        If Trim$(txtOntvanger(1).Text) <> vbNullString And InStr(1, txtOntvanger(0).Text, "@") <> 0 Then
            cmdAdd2Adressbook(1).Enabled = True
        End If
        
        If Trim$(txtAfzender(0).Text) <> vbNullString And InStr(1, txtAfzender(0).Text, "@") > 0 _
            And Trim$(txtAfzender(1).Text) <> vbNullString And _
            Trim$(cmbMailServer.Text) <> vbNullString Then
                cmdVerzenden.Enabled = True  'enable the cmdverzenden object
        End If
    End If
End Sub

Private Sub cmdVerzenden_Click()
    Dim i As Integer                         'Declare i for local use as integer
    Dim info() As String
    
    If cmdVerzenden.Caption = "Send e-mail" Then
        cmdVerzenden.Caption = "Cancel"
        For i = 0 To Controls.count - 1
            If Not Controls(i).Tag = "niet" Then Controls(i).Enabled = False 'disable the if not controls(i).tag = "niet" then controls(i) object
        Next i
        
        SB.Panels(2).Text = "Encoding attachment..."
        SB.Panels(2).Visible = True          'Make the sb.panels(2) object visible
        
        GecodeerdeBijlage = vbNullString
        If lvAttachments.ListItems(1).Text <> "No attachments" Then
            For i = 1 To lvAttachments.ListItems.count
                info = Split(lvAttachments.ListItems(i).Key, "|")
                GecodeerdeBijlage = GecodeerdeBijlage & EncodeFile(info(0), BOUNDARY_ID)
            Next i
        Else: GecodeerdeBijlage = vbNullString
        End If

        PB.Value = 0: PB.Max = 100: PB.Visible = True 'Make the pb.value = 0: pb.max = 100: pb object visible
        SB.Panels(2).Text = "Sending e-mail..."

        SMTP.Connect cmbMailServer.Text, 25
        
        smtpState = MAIL_CONNECT

    Else
        SMTP.Close
        For i = 0 To Controls.count - 1
            If Not Controls(i).Tag = "niet" Then Controls(i).Enabled = True 'enable the if not controls(i).tag = "niet" then controls(i) object
        Next i
        cmdVerzenden.Caption = "Send e-mail"
        PB.Visible = False: SB.Panels(2).Visible = False 'Make the pb object un: sb.panels(2).visible = unvisible
    End If

End Sub

Private Sub smtp_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    Dim i As Integer, info() As String, info2() As String 'Declare i for local use as integer, info() as string, info2() as string

                                             'Retrive data from winsock buffer
    SMTP.GetData strServerResponse
    reactie = strServerResponse
    Debug.Print strServerResponse            'Added strserverresponse to the debug window

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

                strDataToSend = Trim$(txtAfzender(0).Text)
                SMTP.SendData "HELO " & strDataToSend & vbCrLf

                Debug.Print "HELO " & strDataToSend 'Added "helo " & strdatatosend to the debug window
                PB.Value = PB.Value + 12.5   'Add 12.5 to pb.value

            Case MAIL_HELO
                smtpState = MAIL_FROM

                SMTP.SendData "MAIL FROM:" & Trim$(txtAfzender(0).Text) & vbCrLf
                Debug.Print "MAIL FROM:" & Trim$(txtAfzender(0).Text) 'Added "mail from:" & trim$(txtafzender(0).text) to the debug window

                PB.Value = PB.Value + 12.5   'Add 12.5 to pb.value

            Case MAIL_FROM

                'If checkState = 0 Then
                    SMTP.SendData "RCPT TO:" & txtOntvanger(0).Text & vbCrLf
                    Debug.Print "RCPT TO:" & txtOntvanger(0).Text & vbCrLf 'Added "rcpt to:" & txtontvanger(0).text & vbcrlf to the debug window
                    'checkState = 1
                'ElseIf checkState = 1 Then
                    'SMTP.SendData "RCPT TO:controleadres@hotmail.com" & vbCrLf
                    smtpState = MAIL_RCPTTO
                    'checkState = 0           'reset checkstate to zero
                'End If

            Case MAIL_RCPTTO
                smtpState = MAIL_DATA
                                             'Send DATA command to the server
                                             'so it knows that we want to send the message
                SMTP.SendData "DATA" & vbCrLf
                Debug.Print "DATA"           'Added "data" to the debug window
                                             'Debug.Print "Mail RCPTTO"
                PB.Value = PB.Value + 12.5   'Add 12.5 to pb.value
            Case MAIL_DATA
                smtpState = MAIL_DOT
                                             'Send Subject

                SMTP.SendData GenerateMessageID(Mid(txtAfzender(0).Text, InStr(1, txtAfzender(0).Text, "@") + 1, Len(txtAfzender(0).Text))) & Chr$(13) & Chr$(10)
                SMTP.SendData "Date: " & Format(Now, "ddd, dd mmm yyyy hh:nn:ss") & Chr$(13) & Chr$(10)
                SMTP.SendData "From: " & txtAfzender(1).Text & " <" & txtAfzender(0).Text & ">" & Chr$(13) & Chr$(10)
                'SMTP.SendData "To: Undisclosed-Recipients:;" & Chr$(13) & Chr$(10)
                SMTP.SendData "To: " & txtOntvanger(1).Text & " <" & txtOntvanger(0).Text & ">" & Chr$(13) & Chr$(10)
                SMTP.SendData "Reply-to: " & " <" & txtAfzender(0).Text & ">" & Chr$(13) & Chr$(10)
                SMTP.SendData "Subject: " & txtOnderwerp.Text & Chr$(13) & Chr$(10)
                SMTP.SendData "X-Mailer: Mailapplication" & Chr$(13) & Chr$(10)
                SMTP.SendData GetMIMEHeader(BOUNDARY_ID) & Chr$(13) & Chr$(10)
                
                Debug.Print "Subject:" & txtOnderwerp.Text 'Added "subject:" & txtonderwerp.text to the debug window

                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String

                strMessage = txtBericht.Text & vbCrLf & GecodeerdeBijlage

                                             'Parse message to get lines
                varLines = Split(strMessage, vbCrLf)
                                             'clear message buffer
                strMessage = ""              'Empty strmessage
                                             'Send each line of the message
                                             'so no line gets lost

                For Each varLine In varLines 'Loop through all active varlines using the varline object
                    SMTP.SendData CStr(varLine) & vbNewLine
                Next
                                             'Send a dot symbol so the server knows
                                             'that the end of the message is reached
                SMTP.SendData "." & vbCrLf

                Debug.Print "."              'Added "." to the debug window
                                             'Debug.Print "Mail Data"
                PB.Value = PB.Value + 12.5   'Add 12.5 to pb.value
            Case MAIL_DOT
                smtpState = MAIL_QUIT
                                             'Send QUIT command
                SMTP.SendData "QUIT" & vbCrLf
                Debug.Print "QUIT"           'Added "quit" to the debug window
                                             'Debug.Print "Mail dot"
                PB.Value = PB.Value + 12.5   'Add 12.5 to pb.value
            Case MAIL_QUIT
                SMTP.Close
        End Select
    Else
        SMTP.Close
        For i = 0 To Controls.count - 1
            If Not Controls(i).Tag = "niet" Then Controls(i).Enabled = True 'enable the if not controls(i).tag = "niet" then controls(i) object
        Next i
        cmdVerzenden.Caption = "Send e-mail"

        If Not smtpState = MAIL_QUIT Then
            PB.Visible = False               'Make the pb object unvisible
            SB.Panels(2).Text = "Error: " & Trim$(strServerResponse)
            Debug.Print "Error: " & strServerResponse 'Added "error: " & strserverresponse to the debug window
        Else
            Debug.Print "Message sent"       'Added "message sent" to the debug window
            SB.Panels(2).Text = "Message sent"
            PB.Value = 100
            
            Pauze 1
            PB.Visible = False: SB.Panels(2).Visible = False 'Make the pb object un: sb.panels(2).visible = unvisible
        End If
    End If
End Sub

Private Sub smtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim i As Integer                         'Declare i for local use as integer

    SMTP.Close
    For i = 0 To Controls.count - 1
        If Not Controls(i).Tag = "niet" Then Controls(i).Enabled = True 'enable the if not controls(i).tag = "niet" then controls(i) object
    Next i
    cmdVerzenden.Caption = "Send e-mail"

    Debug.Print "Error " & Number & ": " & vbCrLf & Description, vbExclamation 'Added "error " & number & ": " & vbcrlf & description, vbexclamation to the debug window
    PB.Visible = False                       'Make the pb object unvisible
    SB.Panels(2).Text = "Error " & Number & ": " & Description
End Sub

Private Sub txtOntvanger_GotFocus(Index As Integer)
    txtOntvanger(Index).SelStart = 0: txtOntvanger(Index).SelLength = Len(txtOntvanger(Index).Text)
End Sub

