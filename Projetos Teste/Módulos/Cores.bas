Attribute VB_Name = "Cores"
'*******************************************************************************************
'M�dulo............................: Nenhum
'Formul�rio........................: Cores
'Objetivo do formul�rio............: Tratar cores na execu��o dos eventos relacionadas no textbox
'Programa��o.......................: Marcos Bai�o
'Data..............................: 14/11/2002
'*******************************************************************************************
Option Explicit

Public Sub Cor_text_box(textbox As textbox, Evento As String)
'Evento:Sempre ser� E --> Entrada  ou S --> Sa�da
  If Evento = "E" Then
     textbox.BackColor = &H80FFFF
  ElseIf Evento = "S" Then
     textbox.BackColor = &H80000018
  End If
  
End Sub
