VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   2010
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   6585
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   990
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   390
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private s As VetorDeMensagens.ServidorDeMensagens
Private Sub Command1_Click()
    Set s = New VetorDeMensagens.ServidorDeMensagens
    s.EnviarMensagem Me.hWnd, Text1.Text, "JanelaDestino"
    's.EnviarMensagem Me.hWnd, Text1.Text, "Cliente"
End Sub
