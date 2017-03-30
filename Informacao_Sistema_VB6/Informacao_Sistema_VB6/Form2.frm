VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   1740
   ClientLeft      =   11925
   ClientTop       =   3345
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1740
   ScaleWidth      =   5265
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      ScaleHeight     =   615
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   600
      Width           =   975
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1215
         Left            =   -120
         TabIndex        =   2
         Top             =   -120
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   2143
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Relevante sobre o sistema"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Recolhendo Informação"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
On Error Resume Next

Unload Me
End
End Sub

Private Sub Form_Load()
On Error Resume Next

With Me

    .Height = 1860
    .Width = 5385
    .Left = Screen.Width / 2 - Me.Width / 2
    .Top = Screen.Height / 1.5 - Me.Height / 2

End With

With Picture1

    .Left = Me.Width / 6
    .Top = Me.Height / 1.7 - .Height

End With

With WebBrowser1

    .Left = -100
    .Top = -100
    .Width = Picture1.Width + 400
    .Height = Picture1.Height + 200

End With

End Sub

Private Sub Form_Resize()
On Error Resume Next

With Me

    .Height = 1860
    .Width = 5385
    .Left = Screen.Width / 2 - Me.Width / 2
    .Top = Screen.Height / 1.5 - Me.Height / 2

End With

End Sub
