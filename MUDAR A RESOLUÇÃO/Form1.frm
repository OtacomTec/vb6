VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const BITSPIXEL = 12
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Dim OldX As Long, OldY As Long, nDC As Long
Sub ChangeRes(X As Long, Y As Long, Bits As Long)
    Dim DevM As DEVMODE
    'Get the info into DevM
    erg& = EnumDisplaySettings(0&, 0&, DevM)
    'This is what we're going to change
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = X 'ScreenWidth
    DevM.dmPelsHeight = Y 'ScreenHeight
    DevM.dmBitsPerPel = Bits '(can be 8, 16, 24, 32 or even 4)
    'Now change the display and check if possible
    erg& = ChangeDisplaySettings(DevM, CDS_TEST)
    'Check if succesfull
    Select Case erg&
        Case DISP_CHANGE_RESTART
            an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
            If an = vbYes Then
                erg& = ExitWindowsEx(EWX_REBOOT, 0&)
            End If
        Case DISP_CHANGE_SUCCESSFUL
            erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
            MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
        Case Else
            MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
    End Select
End Sub
Private Sub Form_Load()
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim nDC As Long
    'retrieve the screen's resolution
    OldX = Screen.Width / Screen.TwipsPerPixelX
    OldY = Screen.Height / Screen.TwipsPerPixelY
    'Create a device context, compatible with the screen
    nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    'Change the screen's resolution
    ChangeRes 1024, 768, GetDeviceCaps(nDC, BITSPIXEL)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'restore the screen resolution
    ChangeRes OldX, OldY, GetDeviceCaps(nDC, BITSPIXEL)
    'delete our device context
    DeleteDC nDC
End Sub


