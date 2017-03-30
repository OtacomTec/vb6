VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AreaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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


Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Const PROCESS_TERMINATE As Long = &H1
Private Const SYNCHRONIZE = &H100000

Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_NORMAL = 1
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

Public Enum EstadoDaJanela
    EDJ_Normal = 1
    EDJ_Minimizada = 6
    EDJ_Maximizada = 3
    EDJ_Restaurada = 9
End Enum
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Servidor_mensagem_admin As VetorDeMensagens.ServidorDeMensagens

Private Const GW_HWNDNEXT = 2
'local variable(s) to hold property value(s)


Function PegarIDdoAplicativo(ByVal PID_AplicativoAlvo As Long) As Long
    Dim ID_Teste As Long, ID_AplicativoTeste As Long, IDParcial_Aplicativo As Long
    ID_Teste = FindWindow(ByVal 0&, ByVal 0&)
    
    Do While ID_Teste <> 0
        If GetParent(ID_Teste) = 0 Then
            IDParcial_Aplicativo = GetWindowThreadProcessId(ID_Teste, ID_AplicativoTeste)
            If ID_AplicativoTeste = PID_AplicativoAlvo Then
                PegarIDdoAplicativo = ID_Teste
                Exit Do
            End If
        End If
        ID_Teste = GetWindow(ID_Teste, GW_HWNDNEXT)
    Loop
    'Debug.Print PegarIDdoAplicativo
End Function

Public Function AdicionarAplicativo(CaminhoDoAplicativo As String, IDAreaDeTrabalho As Long, PID As Long, HANDLER_FORM As Long, Caption_form_chamada As String) As Long
    Dim ID_Aplicativo As Long
    Dim ID_AreaDeTrabalho As Long
    Dim PID_Aplicativo As Long
    Dim strMensagem As String
    
    On Error GoTo Erro
    
    'Setando a DLLIntercomunicadora
    Set Servidor_mensagem_admin = New VetorDeMensagens.ServidorDeMensagens
    
    PID_Aplicativo = Shell(CaminhoDoAplicativo, vbMaximizedFocus)
    
    'Monta a mensagem a ser passada para o Exe
    strMensagem = strEstação & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).NomeReduzido & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).Senha & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).Nome & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).Codigo & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).PrivilégioAcessar & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).PrivilégioAlterar & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).PrivilégioConsultar & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).PrivilégioExcluir & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).PrivilégioIncluir & "¤" & _
                  frmAdminMDI.AplicativoUsuário(0).Empresa
          
    Servidor_mensagem_admin.EnviarMensagem HANDLER_FORM, strMensagem, Caption_form_chamada
    
    If PID_Aplicativo = 0 Then MsgBox "Erro iniciando programa."
    
    ID_JanelaDoAplicativo = PegarIDdoAplicativo(PID_Aplicativo) 'Receber o ID(handle) da janela do aplicativo no Windows
    SetParent ID_JanelaDoAplicativo, IDAreaDeTrabalho 'Define que a Janela do Aplicativo será uma Filha da Área de Trabalho
        
    PID = PID_Aplicativo
    'Putfocus= Atribui o foco a um determinado ID(handle) do windows
    Putfocus IDAreaDeTrabalho
    Putfocus ID_JanelaDoAplicativo
    LockWindowUpdate False 'não sei só sei que tem que colocar
    AdicionarAplicativo = ID_JanelaDoAplicativo
    
    Exit Function
    
Erro:

    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Logicx"
    
End Function

Public Function CriarÁreaDeTrabalho(frm As Form) As Long
    LockWindowUpdate GetDesktopWindow
    Putfocus frm.hwnd ' mWnd
    LockWindowUpdate False
    CriarÁreaDeTrabalho = frm.hwnd
End Function

Public Function AlternarPara(ID As Long, Optional EstadoDaJanelaX As EstadoDaJanela = EDJ_Normal)
    Dim X As Long
    'Putfocus ID
    Call ShowWindow(ID, EstadoDaJanelaX)
    If IsIconic(ID) Then
         Call ShowWindow(ID, EstadoDaJanelaX)
    End If
    Call SetForegroundWindow(ID)
    X = (GetForegroundWindow = ID)
    Call ShowWindow(ID, EstadoDaJanelaX)
    Putfocus ID
End Function

Public Function AlternarPara2(ID As Long)
    Dim X As Long
    'Putfocus ID
    Call ShowWindow(ID, SW_RESTORE)
    If IsIconic(ID) Then
         Call ShowWindow(ID, SW_RESTORE)
    End If
    Call SetForegroundWindow(ID)
    X = (GetForegroundWindow = ID)
    Call ShowWindow(ID, SW_RESTORE)
    Putfocus ID
End Function


Public Function FinalizarTarefa(ByVal ProcessID As Long) As Boolean
   Dim hProc As Long
   Const fdwAccess As Long = SYNCHRONIZE Or PROCESS_TERMINATE
   
   ' Need to open process with terminate rights, or
   ' just give up immediately.
   hProc = OpenProcess(fdwAccess, 0&, ProcessID)
   If hProc Then
      ' Kill it.
      If TerminateProcess(hProc, 0&) Then
         FinalizarTarefa = True
      'Else
      '   ApiErrorDump Err.LastDllError, "TerminateProcess"
      End If
      ' Clean up.
      'Call CloseHandle(hProc)
   End If
End Function

