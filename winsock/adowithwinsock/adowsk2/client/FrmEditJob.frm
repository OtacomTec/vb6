VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmEditJob 
   Caption         =   "Edit Job"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Record"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   " Technician Details "
      Height          =   1575
      Left            =   0
      TabIndex        =   26
      Top             =   3600
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Booked By:"
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Allocated To:"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Job Location"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   1005
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Job Description "
      Height          =   2775
      Left            =   4560
      TabIndex        =   25
      Top             =   120
      Width           =   4455
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2415
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4260
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FrmEditJob.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Client Details "
      Height          =   1695
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Client Name"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Street Address "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "City Address"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Telephone Number"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Job Details "
      Height          =   1695
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   4455
      Begin VB.CheckBox Check3 
         Caption         =   "Job Completed"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Top Pority"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Medium Pority"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Top             =   1280
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Job Number"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Job Date:"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Required Date:"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   645
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1935
      Left            =   4560
      TabIndex        =   13
      Top             =   3240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmEditJob.frx":00E1
   End
   Begin VB.Label Label10 
      Caption         =   "Completed Job Description"
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   3000
      Width           =   2415
   End
End
Attribute VB_Name = "FrmEditJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
With FrmClient.sckClient(FrmClient.MaxCN)

If .State = sckConnected Then
    .SendData "~%" & Text1(0).Text & "~~" & Text1(1).Text & "~~" _
    & Text1(2).Text & "~~" & Text1(3).Text & "~~" & Text1(4).Text & "~~" _
    & RichTextBox1.Text & "~~" & Combo1.Text & "~~" & Combo2.Text & "~~" _
    & Combo3.Text & "~~" & Check1.Value & "~~" & Check2.Value & "~~" _
    & Text1(5).Text & "~~" & RichTextBox2.Text & "~~" & Check3.Value & "~~" _
    & FrmClient.ConCurrent & "~~" 'Send the edit form details to the server
                                  'to be save to the database
                                            
    Else

MsgBox "Problem with Connection: Probably not connected", vbCritical, "Error! With Connection"
Exit Sub
End If

End With

End Sub
Private Sub Command1_Click()
Dim Comfirm As String

Comfirm = MsgBox("Delete this Record? " & vbNewLine & "Job Number " & Label11.Caption _
, vbCritical + vbYesNo, "Delete This Record?")

If Comfirm = vbYes Then
    FrmClient.sckClient(FrmClient.MaxCN).SendData "DeleteRecord" & Label11.Caption
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Text1(0).Text = "": Text1(1).Text = "": Text1(2).Text = "": Text1(4).Text = ""
RichTextBox2.Text = "": RichTextBox1.Text = "": Text1(3).Text = "": Text1(5).Text = ""
End Sub
Private Sub Check1_Click()

If Check1.Value = 1 Then Check2.Value = 0

End Sub

Private Sub Check2_Click()
 
If Check2.Value = 1 Then Check1.Value = 0
 
End Sub

