VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   16110
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As New clsMail.SmtpEnvio

Private Sub Form_Load()

    ABC = a.Envia("smtp.onlytechsolutions.com.br", "587", "marcos@onlytechsolutions.com.br", "TESTE", "marcos_onlytech@hotmail.com", "MARCOS tão LINDO", "marcos", "marcos@onlytechsolutions.com.br", "mar123***", False)
        
End Sub
