VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sample Application 1.0"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLiberationKey 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtCounter 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtUsedDays 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtRegisteredUser 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtSoftwareCode 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label LabelHelp 
      Caption         =   "Compile and place ActiveLock on this form!"
      Height          =   300
      Left            =   720
      TabIndex        =   12
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label LabelLiberationKey 
      Caption         =   "ActiveLock1.LiberationKey ="
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label LabelCounter 
      Caption         =   "ActiveLock1.Counter ="
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label LabelUsedDays 
      Caption         =   "ActiveLock1.UsedDays ="
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label LabelRegisteredUser 
      Caption         =   "ActiveLock1.RegisteredUser ="
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label LabelSoftwareCode 
      Caption         =   "ActiveLock1.SoftwareCode ="
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label LabelTitle 
      Caption         =   "Sample Application 1.0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    UpdateControls
End Sub

Private Sub cmdSet_Click()
    ' Set the liberation key
    ' (This method will raise the ActiveLock1_Registration event)
    ActiveLock1.Register (txtLiberationKey)
End Sub

Private Sub ActiveLock1_Registration(WasSuccessful As Boolean)
    If WasSuccessful Then
        MsgBox "Thank you!", vbInformation, "ActiveLock1_Registration"
    Else
        MsgBox "You didn't type the right key, sorry.", vbExclamation, "ActiveLock1_Registration"
    End If

    UpdateControls
End Sub

Private Sub UpdateControls()
    txtSoftwareCode = ActiveLock1.SoftwareCode
    txtRegisteredUser = ActiveLock1.RegisteredUser
    txtUsedDays = ActiveLock1.UsedDays
    txtCounter = ActiveLock1.Counter
End Sub
