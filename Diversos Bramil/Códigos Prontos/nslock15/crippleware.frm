VERSION 5.00
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "nslock15vb5.ocx"
Begin VB.Form Form1 
   Caption         =   "ActiveLock sample"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin nslock15vb5.ActiveLock ActiveLock1 
      Left            =   3720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "PA$$WORD..."
      SoftwareName    =   "Crippleware Sample"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnRegister 
      Caption         =   "Unregister"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame frameRegister 
      Caption         =   "Please, register !"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register !"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Liberation key:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Software code:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "You've been using this software for 0 days."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Label Label4 
      Caption         =   $"crippleware.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRegister_Click()
  ActiveLock1.LiberationKey = Text2
  If ActiveLock1.RegisteredUser Then
    frameRegister.Visible = False
    MsgBox "Thanks for registering!"
    Command1.Enabled = True
    Command2.Enabled = True
    cmdUnRegister.Enabled = True
  Else
    MsgBox "Wrong key, try again!"
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
    Text2.SetFocus
  End If
End Sub

Private Sub cmdUnRegister_Click()
  Dim R As VbMsgBoxResult
  R = MsgBox("Are you sure that you want to unregister this software?", vbYesNo)
  If R = vbYes Then
    ActiveLock1.LiberationKey = "0"
    Unload Me
  End If
End Sub

Private Sub Command1_Click()
  MsgBox "Command1 code goes here."
End Sub

Private Sub Command2_Click()
  MsgBox "Command2 code goes here."
End Sub

Private Sub Form_Load()
  ' If the user hasn't registered yet,
  ' shows the registration frame
  If ActiveLock1.RegisteredUser Then
    frameRegister.Visible = False
  Else
    Text1 = ActiveLock1.SoftwareCode
    Text2 = ""
    Label1 = "You've been using this software for " _
          & ActiveLock1.UsedDays & " day(s)."
    Command1.Enabled = False
    Command2.Enabled = False
    cmdUnRegister.Enabled = False
  End If
End Sub
