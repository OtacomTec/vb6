VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   3495
   ClientTop       =   2760
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Mapear"
      Height          =   525
      Left            =   2880
      TabIndex        =   4
      Top             =   1290
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Desmapear"
      Height          =   525
      Left            =   150
      TabIndex        =   3
      Top             =   930
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2850
      TabIndex        =   2
      Text            =   "\\nt1\pasta-i"
      Top             =   810
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Text            =   "t:"
      Top             =   420
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Text            =   "t"
      Top             =   450
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DriveRede As GMclRede

Private Sub Command1_Click()
    DriveRede.DesMapear Text1.Text
End Sub

Private Sub Command2_Click()
    DriveRede.Mapear Text2.Text, Text3.Text
End Sub

Private Sub Form_Load()
    Set DriveRede = New GMclRede
    
End Sub
