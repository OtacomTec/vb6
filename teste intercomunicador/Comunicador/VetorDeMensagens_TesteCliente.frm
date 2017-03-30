VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JanelaDestino"
   ClientHeight    =   3435
   ClientLeft      =   2010
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   1650
      TabIndex        =   0
      Top             =   780
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private c As VetorDeMensagens.ClienteDeMensagens

Private Sub Command1_Click()
    MsgBox c.MensagemRecebida
End Sub

Private Sub Form_Load()
    Set c = New VetorDeMensagens.ClienteDeMensagens
    c.ID_Aplicativo = Me.hWnd
    c.Interceptar
    
End Sub
