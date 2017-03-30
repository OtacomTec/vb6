VERSION 5.00
Begin VB.UserControl NotifyIcon 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "NotifyIcon.ctx":0000
   PropertyPages   =   "NotifyIcon.ctx":0762
   ScaleHeight     =   375
   ScaleWidth      =   480
   ToolboxBitmap   =   "NotifyIcon.ctx":0778
End
Attribute VB_Name = "NotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias _
    "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As _
    NOTIFYICONDATA) As Boolean
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private NotifyIconStruktur As NOTIFYICONDATA
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MBUTTONUP = &H208
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204

'Default Property Values:
Const m_def_TipText = "NotifyIcon TipText"
'Property Variables:
Dim m_Icon As Picture
Dim m_TipText As String
'Event Declarations:
Event MausClickLinks()
Event MausClickMitte()
Event MausClickRechts()
Event MausDblclkLinks()
Event MausDownLinks()
Event MausDblclkMitte()
Event MausDownMitte()
Event MausDdlclkRechts()
Event MausDownRechts()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event TipTextGeändert(NeuerText As String)
Event SymbolGeändert(NeuesSymbol As Picture)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lngMsg As Long
    Static blnFlag As Boolean
    lngMsg = X / Screen.TwipsPerPixelX
    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
            Case WM_LBUTTONUP
                RaiseEvent MausClickLinks
            Case WM_MBUTTONUP
                RaiseEvent MausClickMitte
            Case WM_RBUTTONUP
                RaiseEvent MausClickRechts
            Case WM_LBUTTONDBLCLK
                RaiseEvent MausDblclkLinks
            Case WM_LBUTTONDOWN
                RaiseEvent MausDownLinks
            Case WM_MBUTTONDBLCLK
                RaiseEvent MausDblclkMitte
            Case WM_MBUTTONDOWN
                RaiseEvent MausDownMitte
            Case WM_RBUTTONDBLCLK
                RaiseEvent MausClickRechts
            Case WM_RBUTTONDOWN
                RaiseEvent MausDownRechts
        End Select
        blnFlag = False
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Function NotifyAdd()
Attribute NotifyAdd.VB_Description = "Plaziert das Symbol, das mit der Eigenschaft 'Icon' festgelgt wurde, im SysTray"
On Error GoTo Fehler
    With NotifyIconStruktur
        .cbSize = Len(NotifyIconStruktur)
        .hwnd = UserControl.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Icon
        .szTip = TipText & vbNullChar
    End With
    Call Shell_NotifyIcon(NIM_ADD, NotifyIconStruktur)
Fehler:
End Function

Public Function NotifyRemove()
Attribute NotifyRemove.VB_Description = "Entfernt das Symbol aus dem SysTray"
    Call Shell_NotifyIcon(NIM_DELETE, NotifyIconStruktur)
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Icon = LoadPicture("")
    m_TipText = m_def_TipText
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_TipText = PropBag.ReadProperty("TipText", m_def_TipText)
End Sub


Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 375
End Sub

Private Sub UserControl_Terminate()
    NotifyRemove
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("TipText", m_TipText, m_def_TipText)
End Sub

Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Zur Laufzeit muß SET verwendet werden. Bsp: Set NotifyIcon1.Icon = Me.Icon"
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    PropertyChanged "Icon"
    NotifyIconStruktur.hIcon = New_Icon
    Call Shell_NotifyIcon(NIM_MODIFY, NotifyIconStruktur)
    RaiseEvent SymbolGeändert(New_Icon)
End Property

Public Property Get TipText() As String
Attribute TipText.VB_Description = "Legt einen Text fest der als QuickInfo angezeigt wird, oder gibt ihn zurück."
Attribute TipText.VB_ProcData.VB_Invoke_Property = "Allgemein"
    TipText = m_TipText
End Property

Public Property Let TipText(ByVal New_TipText As String)
    m_TipText = New_TipText
    PropertyChanged "TipText"
    NotifyIconStruktur.szTip = New_TipText & vbNullChar
    Call Shell_NotifyIcon(NIM_MODIFY, NotifyIconStruktur)
    RaiseEvent TipTextGeändert(New_TipText)
End Property

