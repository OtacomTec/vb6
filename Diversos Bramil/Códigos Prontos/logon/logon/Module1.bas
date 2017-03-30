Attribute VB_Name = "Module1"
'Para esconder a Área de Trabalho faça:
'AreaTrabalho False
'
'Para apresentar a Área de Trabalho faça:
'AreaTrabalho True
'
'Para esconder a Barra de Tarefas faça:
'BarraTarefas False
'
'Para apresentar a Barra de Tarefas faça:
'BarraTarefas True
'
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Public Sub BarraTarefas(Visible As Boolean)
Dim hWnd As Long
hWnd = FindWindow("Shell_TrayWnd", "")
If Visible Then
ShowWindow hWnd, SW_SHOW
Else
ShowWindow hWnd, SW_HIDE
End If
EnableWindow hWnd, Visible
End Sub
Public Sub AreaTrabalho(Visible As Boolean)
Dim hWnd As Long
hWnd = FindWindow("Progman", "Program Manager")
If Visible Then
ShowWindow hWnd, SW_SHOW
Else
ShowWindow hWnd, SW_HIDE
End If
EnableWindow hWnd, Visible
End Sub

