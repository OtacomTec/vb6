VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   2550
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "New User"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1100
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   400
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Form1.UserSock.SendData "~~" & Text1(0).Text & "~~" & Text1(1).Text & "~~" & Text1(2).Text & "~~"
Form1.ClrTxt
Form1.Timer2 = True
Unload Me
                                    ' sends the server the the new user details to add
End Sub

Private Sub Form_Load()
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
End Sub
