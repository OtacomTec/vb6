Attribute VB_Name = "embrap"
'Option Explicit
'x-x-x-x-x-x-x-x-x-x-x-x-x-x Defini��es para proc �coneNaBandeja
Private Type �cone_Na_Bandeja
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    ImgTray As Long
    TextTray As String * 64
End Type

Private Const BAND_ADIC = &H0
Private Const BAND_MODIF = &H1
Private Const BAND_DEL = &H2
Private Const MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Public Const DUPLO_CLICK_ESQ = &H203
Public Const BOTAO_ESQ_DOWN = &H201 '513
Public Const BOTAO_ESQ_UP = &H202
Public Const DUPLO_CLICK_DIR = &H206
Public Const BOTAO_DIR_DOWN = &H204 '516 Direito em Baixo
Public Const BOTAO_DIR_UP = &H205 ' 517 Direito Em Cima
Public Const BOTAO_MOVENDO = &H200 '512

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const WM_SYSCOMMAND = &H112

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As �cone_Na_Bandeja) As Boolean
Public t As �cone_Na_Bandeja
'x-x-x-x-x-x-x-x-x-x-x-x-x-x T�rmino defini��es para �coneNaBandeja


'*********************** Declara��es para sub ExecutarArq
Declare Function GetActiveWindow Lib "user32" () As Long
'Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'*********************** Declara��es para Fun��o DesligaWindows
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4

'*********************** Declara��o API para chamar URL
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'*************** Declara��es para a fun��o GravaINI e LeINI
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Ret As String



Declare Function GetMenuItemCount Lib "user32" _
        (ByVal hMenu As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" _
        Alias "GetMenuItemInfoA" (ByVal hMenu As _
        Long, ByVal un As Long, ByVal b As Boolean, _
        lpMenuItemInfo As MENUITEMINFO) As Boolean

Public Const MIIM_ID As Long = &H2
Public Const MIIM_TYPE As Long = &H10
Public Const MFT_STRING As Long = &H0&

'********* Fun��o PastaSistema
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const MAX_PATH = 260
   
Public PastaWin As String
Public PastaWinSys As String

'''Declara��es para as fun��es ExecutaWav e TemPlacaDeSomWav
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

'Para poder esperar um per�odo de tempo
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Type Email
    Nome As String
    Servidor As String
    Usu�rio As String
    Senha As String
    SenhaHex As String
End Type
Public ContaEmail As Boolean
Public Xmail() As Email
Public Const MYLOGIN = "VAMOS A LA PLAIA 963171167 B27FLS187V"

Public Sub ExecutaWav(ArquivoDeSom As String)
    Dim rtn As Long
    Const SND_NODEFAULT = &H2
    Const SND_FILENAME = &H20000
    Const SND_ASYNC = &H1
    rtn = PlaySound(ArquivoDeSom + Chr$(0), 0&, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC)
    
    
End Sub
Public Function TemPlacaWav() As Boolean
    Dim rtn As Long
    rtn = waveOutGetNumDevs() 'Verifica se tem Placa de Som
    TemPlacaWav = rtn > 0
End Function


Public Function PastaSistema(Pasta As String) As String
    Pasta = UCase(Pasta)
    Dim Buffer As String
    Select Case Pasta
        Case "WIN"
            
            Buffer = Space(MAX_PATH)
            rtn = GetWindowsDirectory(Buffer, Len(Buffer))
            WinPath = Left(Buffer, rtn)
            PastaSistema = WinPath
            
        Case "WINSYS"
            
            Buffer = Space(MAX_PATH)
            rtn = GetSystemDirectory(Buffer, Len(Buffer))
            WinSysPath = Left(Buffer, rtn)
            PastaSistema = WinSysPath
    End Select
    
End Function


Public Sub GravaINI(NomeDoArquivo As String, Se��o As String, Chave As String, Text As String)
    WritePrivateProfileString Se��o, Chave, Text, NomeDoArquivo
End Sub

Public Function L�INI(NomeDoArquivo As String, Se��o As String, Chave As String)
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Se��o, Chave, "", Ret, Len(Ret), NomeDoArquivo)
    Ret = Left$(Ret, RetLen)
    L�INI = Ret
End Function

Public Sub ExecutarArq(ByVal Arquivo, Caminho)
    Dim temp
    Dim X
    temp = GetActiveWindow()
    X = ShellExecute(temp, "Open", Arquivo, vbNullString, Caminho, 1)
    GravaINI App.Path & "\Config.ini", "Config", "Primeira Vez", "1"
    If X < 32 Then
        MsgBox "Ocorreu um erro na cria��o do Acesso a Rede Dial-Up"
    End If
End Sub
Public Sub DesligaWindows(ByVal Op��oDeDesligar As Byte)
    '0 = Logoff
    '1 = Desligar
    '2 = Reiniciar
    '3 = Force????
  Dim Successo As Boolean
  Successo = ExitWindowsEx(Op��oDeDesligar, 0&)
End Sub

Public Sub �coneNaBandeja(A��o As String, _
                          ctrlInicial As Control, _
                          ctrl As Control, _
                          frm As Form, _
                          Optional msg As String)
    
    A��o = UCase(A��o)
    Select Case A��o
        Case "ADICIONA"
            t.cbSize = Len(t)
            t.hWnd = frm.hWnd 'ctrlInicial.hwnd
            t.uID = 1&
            t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            t.uCallbackMessage = MOUSEMOVE
            t.ImgTray = ctrlInicial.Picture
            t.TextTray = msg & Chr$(0)
            Shell_NotifyIcon BAND_ADIC, t
            
        Case "DELETA"
            t.cbSize = Len(t)
            't.hwnd = ctrlInicial.hwnd
            t.hWnd = frm.hWnd
            t.uID = 1&
            Shell_NotifyIcon BAND_DEL, t
            
        Case "ALTERA"
            t.cbSize = Len(t)
            't.hwnd = ctrlInicial.hwnd
            t.hWnd = frm.hWnd
            t.uID = 1&
            t.uFlags = NIF_ICON
            t.ImgTray = ctrl.Picture
            Shell_NotifyIcon BAND_MODIF, t
            DoEvents
            Shell_NotifyIcon BAND_MODIF, t
    End Select
    
End Sub

Sub ConfiguraIE()
    'Define pagina Inicial
    AdNovaSequ�ncia "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Start Page", (P�ginaInicial)
    'Define Barra de T�tulo
    'AdNovaSequ�ncia "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Window Title", strBarraT�tuloIE

    'N�o permite que seja feita autodiscagem
    AdNovoValBin�rio "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "EnableAutodial", Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0)
    
    'Trava P�gina Inicial
    If TravaP�gina = "0" Then
        CriaChave "HKEY_USERS\.Default\Software\Policies\Microsoft\Internet Explorer\Control Panel"
        AdNovaSequ�ncia "HKEY_USERS\.Default\Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage", ""
        'DelChave "HKEY_USERS\.Default\Software\Policies\Microsoft\Internet Explorer\Control Panel\HomePage\"
    Else
        CriaChave "HKEY_USERS\.Default\Software\Policies\Microsoft\Internet Explorer\Control Panel"
        AdNovaSequ�ncia "HKEY_USERS\.Default\Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage", (TravaP�gina)
    End If
End Sub

