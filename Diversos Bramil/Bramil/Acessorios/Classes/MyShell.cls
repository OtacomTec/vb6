VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFF

Public Function SyncShell(ByVal CaminhoDoArquivo As String, Optional ExibirJanela As Integer = 0) As Boolean
    'Aciona a execução de um processo e deixa a aplicação que chamou em modo
    'de espera até o processo ser encerrado.
    'CaminhoDoArquivo: Caminho do programa a ser chamado, podem ser acrescentados os parâmetros de chamada do programa
    'ExibirJanela: modo de execução (0 - modo invisível, usuário não vê que outro prorgama foi chamado)
    On Error GoTo SyncShell_Error
    Dim processID As Long
    Dim ProcessHandle As Long
    processID = Shell(CaminhoDoArquivo, ExibirJanela)
    ProcessHandle = OpenProcess(SYNCHRONIZE, True, processID)
    WaitForSingleObject ProcessHandle, INFINITE
    SyncShell = True
    Exit Function
SyncShell_Error:
    On Error GoTo 0
    SyncShell = False
    Exit Function
End Function
'Chamada da função
'If SyncShell("C:\util\pkzip.exe -a - b", 0) = True Then
       'Implemente seu código aqui
'End If

