Attribute VB_Name = "Cores"
'*******************************************************************************************
'Módulo............................: Nenhum
'Formulário........................: Cores
'Objetivo do formulário............: Tratar cores na execução dos eventos relacionadas no textbox
'Programação.......................: Marcos Baião
'Data..............................: 14/11/2002
'*******************************************************************************************
Option Explicit

Public Sub Cor_text_box(textbox As textbox, Evento As String)
'Evento:Sempre será E --> Entrada  ou S --> Saída
  If Evento = "E" Then
     textbox.BackColor = &H80FFFF
  ElseIf Evento = "S" Then
     textbox.BackColor = &H80000018
  End If
  
End Sub
