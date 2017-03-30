VERSION 5.00
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "nslock15vb5.ocx"
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please, register !"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin nslock15vb5.ActiveLock ActiveLock1 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "PASSWORD"
      SoftwareName    =   "Trialware Sample"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Liberation key:"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Software code:"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Users should register in order to use the software for a longer period."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "The program is fully functional, but it will stop working after 21 days. (Try to change the time settings in your computer)"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "This sample shows how to create ""trialware"" using ActiveLock."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Calculator.Caption = "Calculator evaluation"
  Unload Me
End Sub

Private Sub cmdRegister_Click()
  ' Set the LiberationKey:
  ActiveLock1.LiberationKey = Text2
  
  ' Check if it was correct:
  If Not (ActiveLock1.RegisteredUser) Then
    MsgBox "Invalid liberation key!"
  Else
    MsgBox "Thank you for registering!"
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Text1 = ActiveLock1.SoftwareCode
  Text2 = ""
End Sub
