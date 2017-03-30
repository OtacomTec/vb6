Attribute VB_Name = "Robozinho"
Option Explicit

#Const WIN32_IE = &H501 'Shell 5.01 declarations

'Subclassing declarations
Dim hOldProc As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC = (-4)

'Messages that the shell sends to us
Public Const WM_USER = &H400
Public Const VBT_SHELLICON = WM_USER + 101 'Custom notify message
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Shell_NotifyIcon API declarations
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
#If WIN32_IE < &H500 Then
    szTip As String * 64
#Else
    szTip As String * 128
#End If
#If WIN32_IE >= &H500 Then
    dwState As Long '//Version 5.0
    dwStateMask As Long '//Version 5.0
    szInfo As String * 256 'szInfo[256]; '//Version 5.0
    uTimeout As Long '(also uVersion) ' //Version 5.0
    szInfoTitle As String * 64 'szInfoTitle[64]; '//Version 5.0
    dwInfoFlags As Long '//Version 5.0
#End If
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

#If WIN32_IE >= &H500 Then
Public Const NIN_SELECT = (WM_USER + 0)
Public Const NINF_KEY = &H1
Public Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
#End If

#If WIN32_IE >= &H501 Then
Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
#End If

'Parameters for dwMessage
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
#If WIN32_IE >= &H500 Then
Public Const NIM_SETFOCUS = &H3 '//Version 5.0
Public Const NIM_SETVERSION = &H4 '//Version 5.0
Public Const NOTIFYICON_VERSION = 3 '//Version 5.0
#End If

'Flags for the uFlags member of NOTIFYICONDATA
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
#If WIN32_IE >= &H500 Then
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
#End If
#If WIN32_IE >= &H600 Then
Public Const NIF_GUID = &H20
#End If

#If WIN32_IE >= &H500 Then
Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2

'// says this is the source of a shared icon

'// Notify Icon Infotip flags
Public Const NIIF_NONE = &H0
'// icon flags are mutually exclusive
'// and take only the lowest 2 bits
Public Const NIIF_INFO = &H1
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_ICON_MASK = &HF
#If WIN32_IE >= &H501 Then
Public Const NIIF_NOSOUND = &H10
#End If
#End If

Public Sub ShellIconInitialize(Form As Object)

    Dim ni As NOTIFYICONDATA
    
    If hOldProc <> 0 Then
        'ShellIconInitialize was already called
        Exit Sub
    End If
    
    'Subclass the form
    hOldProc = GetWindowLong(Form.hWnd, GWL_WNDPROC)
    SetWindowLong Form.hWnd, GWL_WNDPROC, AddressOf ShellIconCallback
    
    'Add our icon to the tray
    ni.cbSize = Len(ni)
    ni.hWnd = Form.hWnd
    ni.uID = VBT_SHELLICON
    #If WIN32_IE < &H500 Then 'Lower than IE5 - use normal tooltip
    ni.uFlags = NIF_ICON + NIF_MESSAGE + NIF_TIP
    #Else 'IE 5.0 or higher - use balloon tip
    ni.uFlags = NIF_ICON + NIF_MESSAGE + NIF_TIP + NIF_INFO
    #End If
    ni.uCallbackMessage = VBT_SHELLICON
    'ni.hIcon = Form.imgIcon(0).Picture.Handle
    'ni.szTip = "Only Tech - Robô Integrador" & Chr$(0)
    #If WIN32_IE >= &H500 Then 'IE 5.0 or higher - use balloon tip members
    ni.szTip = "Only Tech - Autenticador do Concentrador" & Chr$(0)
    ni.szInfo = "Inicializado o Autenticador do Concentrador!." & Chr$(0)
    ni.uTimeout = 10000 '10000 (10 seconds) is the minimum value anyway
    ni.szInfoTitle = "Only Tech - Autenticador do Concentrador"
    ni.dwInfoFlags = NIIF_INFO
    #End If
    
'    If Shell_NotifyIcon(NIM_ADD, ni) = 0 Then
'        'Failure!
'        MsgBox "Falha ao inicializar o shell icon.", vbExclamation, "Only Tech"
'    End If

End Sub

Public Function ShellIconCallback(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Form As Object) As Long

    If wMsg = VBT_SHELLICON Then
        'Don't check the wParam for this demo - it will always
        'be VBT_SHELLICON (the ID of the icon we've added)
        If (lParam = WM_RBUTTONUP) Or (lParam = WM_CONTEXTMENU) Then
           'Handle shell tray notify
           Form.PopupMenu Form.mnuRobo
        End If
    Else
        'Normal message, just pass it on
        ShellIconCallback = CallWindowProc(hOldProc, hWnd, wMsg, wParam, lParam)
    End If

End Function

Public Sub ShellIconTerminate(Form As Object)
    
    Dim ni As NOTIFYICONDATA
    
    If hOldProc = 0 Then
        'ShellIconInitialize was never called
        Exit Sub
    End If
    
    'Remove the tray icon
    ni.cbSize = Len(ni)
    ni.hWnd = Form.hWnd
    ni.uID = VBT_SHELLICON
    If Shell_NotifyIcon(NIM_DELETE, ni) = 0 Then
        'Failure!
        MsgBox "Falha ao remover shell icon.", vbExclamation, "Only Tech"
    End If
    
    'Unsubclass the shelltray form
    SetWindowLong Form.hWnd, GWL_WNDPROC, hOldProc
    hOldProc = 0

End Sub
