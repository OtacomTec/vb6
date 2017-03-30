VERSION 5.00
Begin VB.Form dlgAbout 
   Caption         =   "About VertMenu ActiveX Control"
   ClientHeight    =   2775
   ClientLeft      =   2535
   ClientTop       =   4575
   ClientWidth     =   5385
   Icon            =   "dlgAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5385
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright 1997 WinResources Computing, Inc., Solana Beach, CA 90275"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   225
      TabIndex        =   2
      Top             =   1575
      Width           =   4875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Vertical Menu ActiveX Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   585
      TabIndex        =   0
      Top             =   180
      Width           =   4245
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

