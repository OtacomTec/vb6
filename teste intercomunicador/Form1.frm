VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   585
      Left            =   2790
      TabIndex        =   4
      Top             =   330
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   1230
      Width           =   4545
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sobrenome"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private servidor As VetorDeMensagens.ServidorDeMensagens
Dim msg As String

Private Sub Command1_Click()
    Shell (App.Path & "\BANCOS.exe")
    Set servidor = New VetorDeMensagens.ServidorDeMensagens
    msg = Me.Text1.Text + " " + Me.Text2.Text
    servidor.EnviarMensagem Me.hWnd, msg, "Banco"
End Sub

