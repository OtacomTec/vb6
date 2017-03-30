VERSION 5.00
Begin VB.Form AppForm 
   Caption         =   "(^^!)"
   ClientHeight    =   3195
   ClientLeft      =   -255
   ClientTop       =   405
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botaoMensagem 
      Caption         =   "Mensagem..."
      Height          =   825
      Left            =   210
      TabIndex        =   2
      Top             =   2115
      Width           =   2430
   End
   Begin VB.CommandButton botaoCriarXExe 
      Caption         =   "Lançar ActiveX Exe"
      Height          =   810
      Left            =   210
      TabIndex        =   1
      Top             =   195
      Width           =   2430
   End
   Begin VB.CommandButton botaoTravar 
      Caption         =   "Travar este executavel"
      Height          =   810
      Left            =   210
      TabIndex        =   0
      Top             =   1155
      Width           =   2430
   End
End
Attribute VB_Name = "AppForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botaoCriarXExe_Click()
    Dim xExe As xExe.XClass
    Set xExe = New xExe.XClass
    xExe.AbrirInterface
End Sub

Private Sub botaoTravar_Click()
    Dim d As Double
    
    botaoMensagem.Caption = "Ih Travou! Tente clicar aqui..."
    botaoTravar.Enabled = False
    botaoCriarXExe.Enabled = False
    DoEvents
    
    For d = -90000000# To 900000000#
        '(...)
    Next
    
    botaoMensagem.Caption = "Ok! Liberado!"
    botaoTravar.Enabled = True
    botaoCriarXExe.Enabled = True
    
End Sub

Private Sub botaoMensagem_Click()
    MsgBox "Formulario liberado...", vbInformation
End Sub
