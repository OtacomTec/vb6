VERSION 5.00
Begin VB.Form frmClientes 
   Caption         =   "Banco"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Receiver"
      Height          =   465
      Left            =   3180
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Cliente As VetorDeMensagens.ClienteDeMensagens
Private Sub Command1_Click()
    MsgBox Cliente.MensagemRecebida
End Sub

Private Sub Form_Load()
    Set Cliente = New VetorDeMensagens.ClienteDeMensagens
    Cliente.ID_Aplicativo = Me.hWnd
    Cliente.Interceptar
End Sub
