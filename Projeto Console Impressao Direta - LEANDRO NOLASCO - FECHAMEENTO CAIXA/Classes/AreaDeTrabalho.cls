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
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'------------
Const WS_BORDER = &H800000
Const WS_DLGFRAME = &H400000
Const WS_THICKFRAME = &H40000
Const WS_CAPTION = &HC00000      '  WS_BORDER Or WS_DLGFRAME
Const WS_EX_CLIENTEDGE = &H200
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
'---------------
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

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
Dim ID_JanelaDoAplicativo As Long
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

Public Function AdicionarAplicativo(CaminhoDoAplicativo As String, IDAreaDeTrabalho As Long, PID As Long, HANDLER_FORM As Long, Caption_form_chamada As String, Optional Nao_Redimensona_EXE As Boolean) As Long

    Dim ID_Aplicativo As Long
    Dim ID_AreaDeTrabalho As Long
    Dim PID_Aplicativo As Long
    Dim strMensagem As String
    
    On Error GoTo Erro
    
    '--------------------------------------------------------------------------------------------------
    'Setando a DLLIntercomunicadora
    Set Servidor_mensagem_admin = New VetorDeMensagens.ServidorDeMensagens
    
    '''PID_Aplicativo = Shell(CaminhoDoAplicativo, vbMaximizedFocus)
    PID_Aplicativo = Shell(CaminhoDoAplicativo, vbNormalFocus)
    
    If PID_Aplicativo = 0 Then MsgBox "Erro iniciando programa."
    
    ID_JanelaDoAplicativo = PegarIDdoAplicativo(PID_Aplicativo) 'Receber o ID(handle) da janela do aplicativo no Windows
    
    '--------------------------------------------------------------------------------------------------
    'Dimensionando o Módulo dentro do espaço pré-definido no admin
    
    Dim R As RECT, hDWP As Long
    Dim X As String
    Dim Y As String
    
    'Achando a atual resolução para adequar os módulos no admin
    Call Funcoes_Gerais.Acha_Resolucao(X, Y)
    
    If X = 1024 And Y = 768 Then
       R.Left = 5
       R.Top = 22
       R.Bottom = 615
       R.Right = 905
    ElseIf X = 800 And Y = 600 Then
       R.Left = 5
       R.Top = 22
       R.Bottom = 445
       R.Right = 685
    Else
       R.Left = 5
       R.Top = 22
       R.Bottom = 325
       R.Right = 520
    End If
    
    AdjustWindowRectEx R, WS_CAPTION, False, WS_EX_CLIENTEDGE
    
    If Nao_Redimensona_EXE = False Then: MoveWindow ID_JanelaDoAplicativo, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, False
    
    '--------------------------------------------------------------------------------------------------
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
    
    SetParent ID_JanelaDoAplicativo, IDAreaDeTrabalho 'Define que a Janela do Aplicativo será uma Filha da Área de Trabalho
        
    PID = PID_Aplicativo
    
    'Putfocus= Atribui o foco a um determinado ID(handle) do windows
    'Putfocus IDAreaDeTrabalho
    Putfocus ID_JanelaDoAplicativo
   
    LockWindowUpdate False 'não sei só sei que tem que colocar
    AdicionarAplicativo = ID_JanelaDoAplicativo
    
    Exit Function
    
Erro:

    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Only Tech"
    
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
       'Kill it
      If TerminateProcess(hProc, 0&) Then
         FinalizarTarefa = True
      'Else
      '   ApiErrorDump Err.LastDllError, "TerminateProcess"
      End If
      ' Clean up.
      'Call CloseHandle(hProc)
   End If
   
End Function

