VERSION 5.00
Begin VB.Form Add 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add SMTP Server"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2708
      TabIndex        =   5
      Top             =   675
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   668
      TabIndex        =   4
      Top             =   675
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "25"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean

Private Sub cmdAdd_Click()
    If Len(Trim(txtServer.Text)) = 0 Then
        MsgBox "Please enter a server"
        txtServer.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtPort.Text)) = 0 Or Not IsNumeric(txtPort.Text) Then
        MsgBox "Please enter a valid port"
        txtPort.SetFocus
        Exit Sub
    End If
    OK = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub
