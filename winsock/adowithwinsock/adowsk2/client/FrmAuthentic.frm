VERSION 5.00
Begin VB.Form FrmAuthentic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Authentication Required"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmAuthentic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
FrmClient.AuthCompleted = False
End Sub
Public Function SendAuthentication(Username, Password As String) As Boolean

With FrmClient
    
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    cmdOK.Enabled = False
    Label3.Visible = True
    Label3.Caption = "Verifing Username and Password"
    .StatusBar1.Panels.Item(1).Text = "Verifing Username and Password"
    .sckClient(.MaxCN).SendData "UserName" & "~~" & Username & "~~" & Password & "~~" & FrmClient.ConCurrent & "~~"
    Text1(1).Visible = False
    Text1(0).Visible = False
    Label1.Visible = False
    Label2.Visible = False
End With

End Function
Private Sub cmdOK_Click()

 j = SendAuthentication(Text1(0).Text, Text1(1).Text) = True

End Sub

Private Sub Form_Load()
'Text1(1).Text = "": Text1(0).Text = "" 'Clears the TextBoxes.
Label3.Visible = False
Label4.Visible = False
Text1(0).Text = "Chris Hatton"      'Just saves me from having to enter in my details.
Text1(1).Text = "Password"          'Im to lazy to keep inputting my password.
FrmClient.sckClient(FrmClient.MaxCN).SendData "GetRsCount" & FrmClient.ConCurrent
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(0).SelLength = Len(Text1(0).Text)
Text1(1).SelLength = Len(Text1(1).Text)
End Sub











