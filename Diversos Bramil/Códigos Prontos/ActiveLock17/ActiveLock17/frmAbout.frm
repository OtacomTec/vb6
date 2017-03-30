VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ActiveLock"
   ClientHeight    =   3795
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4710
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2619.375
   ScaleMode       =   0  'User
   ScaleWidth      =   4422.935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":030A
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   112.686
      X2              =   4282.077
      Y1              =   2070.652
      Y2              =   2070.652
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "April / 2002"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblURL 
      Caption         =   "http://www.activelock.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Get the lastest version of ActiveLock, free of charge, at:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4282.077
      Y1              =   942.147
      Y2              =   942.147
   End
   Begin VB.Label lblTitle 
      Caption         =   "lblTitle"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   4282.077
      Y1              =   952.5
      Y2              =   952.5
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   112.686
      X2              =   4282.077
      Y1              =   2070.652
      Y2              =   2070.652
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Author: Nelson Ferraz
' Date  : 1998-2002

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "About " & App.Title
  lblTitle.Caption = App.Title & " version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
