Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32" () As Long 'windows xp support

Private Const NIIF_NONE = &H0                'Declare niif_none with value &h0 as constant for local use
Private Const NIIF_INFO = &H1                'Declare niif_info with value &h1 as constant for local use
Private Const NIIF_WARNING = &H2             'Declare niif_warning with value &h2 as constant for local use
Private Const NIIF_ERROR = &H3               'Declare niif_error with value &h3 as constant for local use
Private Const NIIF_GUID = &H5                'Declare niif_guid with value &h5 as constant for local use
Private Const NIIF_ICON_MASK = &HF           'Declare niif_icon_mask with value &hf as constant for local use
Private Const NIIF_NOSOUND = &H10            'Declare niif_nosound with value &h10 as constant for local use

Public Enum BalloonStyle
    bsIconExclamation = NIIF_WARNING
    bsIconCritical = NIIF_ERROR
    bsIconInformation = NIIF_INFO
    bsGuid = NIIF_GUID
    bsIconMask = NIIF_ICON_MASK
    bsNoSound = NIIF_NOSOUND
End Enum

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type EDITBALLOONTIP
   cbStruct As Long
   pszTitle As String
   pszText As String
   ttiIcon As Long
End Type

Public Const ECM_FIRST As Long = &H1500      'Declare ecm_first as long with value &h1500 as constant with max scope
Public Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3) 'Declare em_showballoontip as long with value (ecm_first + 3) as constant with max scope
Public Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4) 'Declare em_hideballoontip as long with value (ecm_first + 4) as constant with max scope

Private m_BalloonData As EDITBALLOONTIP      'Declare m_balloondata for local use as editballoontip

Public pop3state         As POP3States
Public smtpState         As SMTP_State
Public GecodeerdeBijlage As String           'encoded attachment

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long 'Declare the writeprivateprofilestring API with max scope
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long 'Declare the getprivateprofilestring API with max scope

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
   hWndDesk = GetDesktopWindow()

   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
   
  If success = SE_ERR_NOASSOC Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
   
End Sub

Public Function LaadInstellingen(kopdata As String, keydata As String, bestand As String) As String 'save settings in ini-file
    Dim antwoord As String * 256             'Declare antwoord for local use as string * 256

    GetPrivateProfileString kopdata, keydata, vbNullString, antwoord, Len(antwoord), bestand
    LaadInstellingen = Left$(antwoord, InStr(1, antwoord, Chr$(0)) - 1)

End Function

Public Sub OpslaanInstellingen(Data As String, kopdata As String, keydata As String, bestand As String) 'load settings from ini-file
    WritePrivateProfileString kopdata, keydata, Data, bestand
End Sub

Public Sub PopupBallon(hTextBoxWnd As Long, sMessage As String, sTitle As String, bsStyle As BalloonStyle)
    With m_BalloonData
        .cbStruct = Len(m_BalloonData)
        .pszTitle = StrConv(sTitle, vbUnicode)
        .pszText = StrConv(sMessage, vbUnicode)
        .ttiIcon = bsStyle
    End With
    SendMessage hTextBoxWnd, EM_SHOWBALLOONTIP, 0&, m_BalloonData
End Sub

Public Function AddContact(name As String, email As String) As Boolean
    Dim info() As String, i As Integer       'Declare info() for local use as string, i as integer
    Dim gevondenNaam As Boolean, gevondenAdres As Boolean, strNaam As String, strAdres As String, strIndex 'Declare gevondennaam for local use as boolean, gevondenadres as boolean, strnaam as string, stradres as string, strindex

    For i = 1 To frmMain.Adresboek.count
        info = Split(frmMain.Adresboek(i), ";")
        If LCase$(info(0)) = LCase$(name) Then
            strAdres = info(1)
            strNaam = info(0)
            strIndex = i
            gevondenNaam = True: Exit For
        ElseIf LCase$(info(1)) = LCase$(email) Then
            strAdres = info(1)
            strNaam = info(0)
            strIndex = i
            gevondenAdres = True: Exit For
        End If
    Next i

    If gevondenNaam = True Then
        If MsgBox("There is already a contact with the same name, namely " & strAdres & ". Do you want to change this e-mailadress?", vbQuestion + vbYesNo, "Contact") = vbNo Then 'Inform the user with a messagebox
            Exit Function                    'Leave this sub
        Else
            frmMain.Adresboek.Remove strIndex
            frmMain.Adresboek.Add name & ";" & email
        End If
    ElseIf gevondenAdres = True Then
        If MsgBox("There is already a contact with the same e-mailadress, namely " & strNaam & ". Do you want to change this name?", vbQuestion + vbYesNo, "Contact") = vbNo Then 'Inform the user with a messagebox
            Exit Function                    'Leave this sub
        Else
            frmMain.Adresboek.Remove strIndex
            frmMain.Adresboek.Add name & ";" & email
        End If
    Else: frmMain.Adresboek.Add name & ";" & email
    End If

    Open App.Path & "\Adressbook.dat" For Output As #3 'Open the file app.path & "\adressbook.dat" to write in
        For i = frmMain.Adresboek.count To 1 Step -1
            Print #3, frmMain.Adresboek(i)   'Write frmmain.adresboek(i) in file #3
            frmMain.Adresboek.Remove i
        Next i
    Close #3                                 'Close file #3

    frmMain.AdresboekLaden
End Function
