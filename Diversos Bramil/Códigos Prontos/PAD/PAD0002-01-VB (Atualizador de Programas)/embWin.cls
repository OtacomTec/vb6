VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Attribute VB_Ext_KEY = "Member1" ,"eWinDesligar"
'Shutdown do windows
Public Enum embSaidaDoWindows
    [_First] = 0
    SDW_LOGOFF = 0&
    SDW_SHUTDOWN = 1&
    SDW_REBOOT = 2&
    SDW_FORCE = 4&
    SDW_POWEROFF = 8&
    SDW_FORCEIFHUNG = 10& ' NT5 only
    SDW_RESET = SDW_LOGOFF Or SDW_FORCE Or SDW_REBOOT
    [_Last] = &H20& - 1
End Enum


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As eWinVersao_OSVERSIONINFO) As Long
Const VER_PLATFORM_WIN32_NT As Long = 2
Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Const VER_PLATFORM_WIN32s As Long = 0
Private Type eWinVersao_OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Enum eWinVersaoVER_SO_en
    emb_VerSONãoIdentificado = 0
    emb_VerSOWindows95 = 10
    emb_VerSOWindows95SE = 11
    emb_VerSOWindows98 = 15
    emb_VerSOWindows98SE = 16
    emb_VerSOWindowsNT3 = 20
    emb_VerSOWindowsNT4 = 25
    emb_VerSOWindows2K = 40
    emb_VerSOWindowsXP = 50
    emb_VerSOWindows2K3 = 60
End Enum

Private mvarSistemaOperacional As String 'local copy
Private mvarVersãoMaior As String 'local copy
Private mvarVersãoMenor As String 'local copy
Private mvarDetalhe As String 'local copy
Private mvarID_SO As Byte 'local copy
Private mvarCompilação As String 'local copy
Private mvarPlataforma As Integer 'local copy
Private mvarNomeDaPlataforma As String 'local copy

'<<<<<<<< Verificar NomeDoComputador
Private Declare Function w32_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function w32_SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private mvarNomeDoComputador As String
'>>>>>>>>

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private mvarPastaWindows As String 'local copy
Private mvarPastaWinSys As String 'local copy
Private mvarDesligar As eWinDesligar


'Public Property Get Desligar() As eWinDesligar
'    If mvarDesligar Is Nothing Then
'        Set mvarDesligar = New eWinDesligar
'    End If
'    Set Desligar = mvarDesligar
'End Property
'
'
'Public Property Set Desligar(vData As eWinDesligar)
'    Set mvarDesligar = vData
'End Property
Private Sub Class_Terminate()
  Set mvarClass1 = Nothing
    Set mvarDesligar = Nothing
End Sub



Public Property Get PastaWinSys() As String
    PastaWinSys = mvarPastaWinSys
End Property

Public Property Get PastaWindows() As String
    PastaWindows = mvarPastaWindows
End Property
Public Property Get NomeDoComputador() As String
    NomeDoComputador = mvarNomeDoComputador
End Property

Public Function AlterarNomeDoComputador(strNovoNomeDoComputador As String) As Boolean
  Dim lngRes As Long
  Dim booRes As Boolean
  lngRes = w32_SetComputerName(strNovoNomeDoComputador)
  If lngRes <> 0 Then
    mvarNomeDoComputador = strNovoNomeDoComputador
    booRes = True
  Else
    booRes = False
  End If
  AlterarNomeDoComputador = booRes
End Function

Public Function VerificarNomeDoComputador() As String ' Boolean
  Dim strNomeDoComputador As String
  Dim lngTamanhoDoRetorno As Long
  Dim lngRes As Long
  
  lngTamanhoDoRetorno = 256
  strNomeDoComputador = Space(lngTamanhoDoRetorno)
  lngRes = w32_GetComputerName(strNomeDoComputador, lngTamanhoDoRetorno)
  If lngRes <> 0 Then
    mvarNomeDoComputador = Left$(strNomeDoComputador, lngTamanhoDoRetorno)
  Else
    mvarNomeDoComputador = Empty
  End If
  VerificarNomeDoComputador = mvarNomeDoComputador
End Function

Public Property Get NomeDaPlataforma() As String
    NomeDaPlataforma = mvarNomeDaPlataforma
End Property

Public Property Get Plataforma() As Integer
    Plataforma = mvarPlataforma
End Property

Public Property Get Compilação() As String
    Compilação = mvarCompilação
End Property

Public Property Get ID_SO() As eWinVersaoVER_SO_en
    ID_SO = mvarID_SO
End Property

Public Property Get Detalhe() As String
    Detalhe = mvarDetalhe
End Property

Public Property Get VersãoMenor() As String
    VersãoMenor = mvarVersãoMenor
End Property

Public Property Get VersãoMaior() As String
    VersãoMaior = mvarVersãoMaior
End Property

Public Property Get SistemaOperacional() As String
    SistemaOperacional = mvarSistemaOperacional
End Property
Public Function VerificarNomeSistOperacional() As String
    Dim osvSistema As eWinVersao_OSVERSIONINFO
    Dim sRetorno As String
    On Error GoTo ERRO
           
    osvSistema.dwOSVersionInfoSize = Len(osvSistema)
    
    If GetVersionEx(osvSistema) = 0 Then
        sRetorno = "Não Identificado"
        mvarSistemaOperacional = "Não Identificado"
        mvarSO = emb_VerSONãoIdentificado
        mvarVersãoMaior = 0
        mvarVersãoMenor = 0
        mvarDetalhe = Empty
        
    Else
        Select Case osvSistema.dwPlatformId
            Case VER_PLATFORM_WIN32_NT
                Select Case osvSistema.dwMajorVersion
                    Case 3
                        If osvSistema.dwMinorVersion = 0 Then
                            sRetorno = "Windows NT3 "
                            mvarSistemaOperacional = "Windows NT3"
                            mvarID_SO = emb_VerSOWindowsNT3
                        End If
                        
                    Case 4
                        If osvSistema.dwMinorVersion = 0 Then
                            sRetorno = "Windows NT4 "
                            mvarSistemaOperacional = "Windows NT4"
                            mvarID_SO = emb_VerSOWindowsNT4
                        End If
                        
                    Case 5
                        Select Case osvSistema.dwMinorVersion
                            Case 0 ' Windows 2000
                                sRetorno = "Windows 2000 "
                                mvarSistemaOperacional = "Windows 2000"
                                mvarID_SO = emb_VerSOWindows2K
                            
                            Case 1 'Windows XP
                                sRetorno = "Windows XP "
                                mvarSistemaOperacional = "Windows XP"
                                mvarID_SO = emb_VerSOWindowsXP
                            Case 2 'Windows 2003
                                sRetorno = "Windows 2003 Server "
                                mvarSistemaOperacional = "Windows 2003 Server"
                                mvarID_SO = emb_VerSOWindows2K3
                                
                            Case Else
                                sRetorno = "Windows Plataforma NT "
                                mvarSistemaOperacional = "Windows Plataforma NT"
                                mvarID_SO = emb_VerSOWindowsXP
                        End Select
                            
                    Case Else
                        sRetorno = "Windows Indefinido "
                        mvarSistemaOperacional = "Windows Indefinido"
                        mvarID_SO = emb_VerSONãoIdentificado
                        
                End Select
                                
                mvarVersãoMaior = osvSistema.dwMajorVersion
                mvarVersãoMenor = osvSistema.dwMinorVersion
                mvarCompilação = osvSistema.dwBuildNumber
                If Left(osvSistema.szCSDVersion, 1) = Chr(0) Then
                    mvarDetalhe = Empty
                Else
                    mvarDetalhe = Mid(osvSistema.szCSDVersion, 1, InStr(osvSistema.szCSDVersion, Chr(0)) - 1)
                End If
                sRetorno = sRetorno & mvarDetalhe
                
            Case VER_PLATFORM_WIN32_WINDOWS
                Select Case osvSistema.dwMajorVersion
                    Case 4
                        If osvSistema.dwMinorVersion = 10 Then
                            sRetorno = "Windows 98 "
                            mvarSistemaOperacional = "Windows 98SE"
                            mvarVersãoMaior = osvSistema.dwMajorVersion
                            mvarVersãoMenor = osvSistema.dwMinorVersion
                            mvarDetalhe = osvSistema.szCSDVersion
                            mvarID_SO = emb_VerSOWindows98SE
                        End If
                        
                    Case Else
                        sRetorno = "Windows 32 "
                        mvarSistemaOperacional = "Windows 32"
                        
                End Select
                
                mvarVersãoMaior = osvSistema.dwMajorVersion
                mvarVersãoMenor = osvSistema.dwMinorVersion
                mvarCompilação = osvSistema.dwBuildNumber
                If Left(osvSistema.szCSDVersion, 1) = Chr(0) Then
                    mvarDetalhe = Empty
                Else
                    mvarDetalhe = Mid(osvSistema.szCSDVersion, 1, InStr(osvSistema.szCSDVersion, Chr(0)) - 1)
                End If
                sRetorno = sRetorno & mvarDetalhe
                
            Case VER_PLATFORM_WIN32s
                sRetorno = "Windows 32s"
                mvarSistemaOperacional = "Windows 32s"
                mvarID_SO = VER_PLATFORM_WIN32s
                mvarVersãoMaior = osvSistema.dwMajorVersion
                mvarVersãoMenor = osvSistema.dwMinorVersion
                mvarCompilação = osvSistema.dwBuildNumber
                If Left(osvSistema.szCSDVersion, 1) = Chr(0) Then
                    mvarDetalhe = Empty
                Else
                    mvarDetalhe = Mid(osvSistema.szCSDVersion, 1, InStr(osvSistema.szCSDVersion, Chr(0)) - 1)
                End If
                
        End Select
    End If
    
    mvarPlataforma = osvSistema.dwPlatformId
    Select Case mvarPlataforma
        Case VER_PLATFORM_WIN32s: mvarNomeDaPlataforma = "Win 32s"
        Case VER_PLATFORM_WIN32_WINDOWS: mvarNomeDaPlataforma = "Win 32"
        Case VER_PLATFORM_WIN32_NT: mvarNomeDaPlataforma = "Win NT"
        Case Else: mvarNomeDaPlataforma = Empty
    End Select
        
    VerificarNomeSistOperacional = sRetorno
    Exit Function
ERRO:
    VerificarNomeSistOperacional = "Erro ao Identificar SO"
End Function

Public Function SairDoWindows(lngSTipoDeSaída As embSaidaDoWindows) As Long
    mvarDesligar.DesligarWindows lngSTipoDeSaída
End Function

'set the shut down privilege for the current application
Function VerificarPastaWin() As String
    Dim strNomeDaPasta As String * 256
    Dim lngTamDaString As Long
    Dim lngTam As Long
    lngTam = 256
    lngTamDaString = GetWindowsDirectory(strNomeDaPasta, lngTam)
    mvarPastaWindows = Left$(strNomeDaPasta, lngTamDaString)
    VerificarPastaWin = mvarPastaWindows
End Function
Function VerificarPastaWinSys() As String
    Dim strNomeDaPasta As String * 256
    Dim lngTamDaString As Long
    Dim lngTam As Long
    lngTam = 256
    lngTamDaString = GetSystemDirectory(strNomeDaPasta, lngTam)
    mvarPastaWinSys = Left$(strNomeDaPasta, lngTamDaString)
    VerificarPastaWinSys = mvarPastaWinSys
End Function
Private Sub Class_Initialize()
    VerificarNomeSistOperacional
    VerificarNomeDoComputador
    Set mvarDesligar = New eWinDesligar
  
End Sub
