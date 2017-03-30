VERSION 5.00
Begin VB.Form frmAplicativo 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   0
   ScaleWidth      =   0
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmAplicativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Const GW_HWNDNEXT = 2
Dim mWnd As Long
Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long, test_pid As Long, test_thread_id As Long
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        If GetParent(test_hwnd) = 0 Then
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
            If test_pid = target_pid Then
                InstanceToWnd = test_hwnd
                Exit Do
            End If
        End If
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function
Private Sub Form_Load()
    Dim Pid As Long
    LockWindowUpdate GetDesktopWindow
    'Pid = Shell("c:\arquivos de programas\office\excel.exe", vbNormalFocus) 'Substitua a string pela localização do seu aplicativo (isto funciona com qualquer programa que rode em janelas).
    Pid = Shell("C:\WINDOWS\NOTEPAD.EXE", vbNormalFocus) 'Substitua a string pela localização do seu aplicativo (isto funciona com qualquer programa que rode em janelas).
    If Pid = 0 Then MsgBox "Erro iniciando programa."
    mWnd = InstanceToWnd(Pid)
    SetParent mWnd, Me.hwnd
    Putfocus mWnd
    LockWindowUpdate False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DestroyWindow mWnd
    TerminateProcess GetCurrentProcess, 0
End Sub

