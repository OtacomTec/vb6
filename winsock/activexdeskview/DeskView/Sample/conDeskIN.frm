VERSION 5.00
Begin VB.Form conDeskIN 
   Caption         =   "Viewer"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin Project1.DeskIn DeskIn1 
      Height          =   6465
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   11404
   End
End
Attribute VB_Name = "conDeskIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
DeskIn1.Width = Me.Width - 100
DeskIn1.Height = Me.Height - 200
End Sub

