VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim dataLink As New MSDASC.DataLinks
Dim connString As String
Dim cn As New ADODB.Connection

' abaixo fazemos um vínculo da janela oledb com o formulário atual
dataLink.hWnd = Me.hWnd

' Exibimos o diálogo no promptnew
On Error Resume Next
connString = dataLink.PromptNew
If Err = 0 Then
' Utilizamos a connection string obtida em um objeto de conexão
cn.ConnectionString = connString
Else
' Usuário cancelou a operação
End If


'No exemplo abaixo, outra forma de abrir a janela, utiliza-se o método
' promptEdit, ao invés de promptnew, para editar a string de conexão de um
' objeto de conexão existente
If dataLink.PromptEdit(cn) Then
MsgBox cn.ConnectionString
Else
MsgBox "Cancelou"
End If
End Sub
