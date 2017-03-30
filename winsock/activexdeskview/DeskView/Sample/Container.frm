VERSION 5.00
Begin VB.Form Container 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "Container.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGo 
      Caption         =   "Continue!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5025
      TabIndex        =   2
      Top             =   810
      Width           =   1725
   End
   Begin VB.CommandButton ShowNotes 
      Caption         =   "Read Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5025
      TabIndex        =   3
      Top             =   120
      Width           =   1725
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4020
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Container.frx":030A
      Top             =   3480
      Width           =   6720
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2190
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Container.frx":0641
      Top             =   1230
      Width           =   6720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "mikes@mtdmarketing.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3030
      TabIndex        =   6
      Top             =   615
      Width           =   1860
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Send any bugs or comments to "
      Height          =   240
      Left            =   795
      TabIndex        =   5
      Top             =   615
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use right-click on the picture window to switch between Automatic and Manual modes."
      Height          =   360
      Left            =   795
      TabIndex        =   4
      Top             =   165
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "Container.frx":08C8
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1125
      Left            =   90
      Top             =   60
      Width           =   6735
   End
End
Attribute VB_Name = "Container"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()

    With conDeskOUT
        conDeskOUT.DeskOut1.LocalPort = 10078
        conDeskOUT.DeskOut1.Listen
        conDeskOUT.DeskOut1.PacketSize = "2048"
        .Show
    End With
    
    With conDeskIN
        .DeskIn1.RemotePort = 10078
        .DeskIn1.RemoteIP = "127.0.0.1"
        .DeskIn1.Connect
        .Show
    End With
    
    Unload Me

End Sub

Private Sub ShowNotes_Click()
    Me.Height = 7950
    ShowNotes.Enabled = False
End Sub
