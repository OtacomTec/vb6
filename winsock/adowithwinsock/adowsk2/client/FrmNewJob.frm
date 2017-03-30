VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmNewJob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Job"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   " Job Details "
      Height          =   1695
      Left            =   4560
      TabIndex        =   19
      Top             =   0
      Width           =   3495
      Begin VB.CheckBox Check2 
         Caption         =   "Medium Pority"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Top Pority"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Required Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   640
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Todays Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   260
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Client Details "
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Telephone Number"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "City Address"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Street Address "
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Client Name"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Job Description "
      Height          =   2775
      Left            =   0
      TabIndex        =   18
      Top             =   1680
      Width           =   4455
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4260
         _Version        =   393217
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"FrmNewJob.frx":0000
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Technician Details "
      Height          =   1575
      Left            =   4560
      TabIndex        =   22
      Top             =   1680
      Width           =   3495
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Job Location"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1000
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Allocated To:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Booked By:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   280
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmNewJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

If Check1.Value = 1 Then Check2.Value = 0

End Sub

Private Sub Check2_Click()
 
If Check2.Value = 1 Then Check1.Value = 0
 
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()

With FrmClient.sckClient(FrmClient.MaxCN)

If .State = sckConnected Then
    .SendData "~@" & Text1(0).Text & "~~" & Text1(1).Text & "~~" _
    & Text1(2).Text & "~~" & Text1(3).Text & "~~" & Text1(4).Text & "~~" _
    & RichTextBox1.Text & "~~" & Combo1.Text & "~~" & Combo2.Text & "~~" _
    & Combo3.Text & "~~" & Check1.Value & "~~" & Check2.Value & "~~" _
    & Text1(5).Text & "~~" & FrmClient.ConCurrent & "~~" 'Looks pretty ugly aye?
                                                         'This sends all the client details to the server
    Else

MsgBox "Problem with Connection: Probably not connected", vbCritical, "Error! With Connection"
Exit Sub
End If

End With


End Sub

Private Sub Form_Load()

Text1(0).Text = "": Text1(1).Text = "": Text1(2).Text = "": Text1(4).Text = ""
RichTextBox1.Text = "": Text1(3).Text = Date: Text1(5).Text = ""

                                            'gets the all users into the combos
FrmClient.sckClient(FrmClient.MaxCN).SendData "ListUsers" & FrmClient.ConCurrent


End Sub
