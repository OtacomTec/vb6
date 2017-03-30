Attribute VB_Name = "OtherStuff"
Option Private Module
Option Explicit

' Author: Nelson Ferraz
' Date  : 1998-2002

Global gstrRegistryPath As String

Private Declare Function GetVolumeInformation& Lib "kernel32" Alias _
        "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
        pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
        lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
        lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
        ByVal nFileSystemNameSize As Long)
        
Private Declare Function GetComputerName& Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)

' Registry stuff

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Private Const MAX_FILENAME_LEN = 256
Private Const MAX_COMPUTERNAME_LENGTH = 15

Public Function ComputerName() As String
  Dim s$, sz&, dl&
  sz& = MAX_COMPUTERNAME_LENGTH + 1
  s$ = String$(sz&, 0)
  dl& = GetComputerName(s$, sz&)
  ComputerName = s$
End Function

Public Function DriveSerial(ByVal sDrv As String) As Long
  Dim RetVal As Long
  Dim str As String * MAX_FILENAME_LEN
  Dim str2 As String * MAX_FILENAME_LEN
  Dim a As Long
  Dim b As Long
  Call GetVolumeInformation(sDrv & ":\", str, MAX_FILENAME_LEN, RetVal, a, b, str2, MAX_FILENAME_LEN)
  DriveSerial = RetVal
End Function

Public Function WindowsProductKey() As String
    Dim strKey As String
  
    strKey = QueryValue(HKEY_LOCAL_MACHINE, _
                        "SOFTWARE\Microsoft\Windows\CurrentVersion", _
                        "ProductKey")
                        
    If strKey = "" Then
        strKey = QueryValue(HKEY_LOCAL_MACHINE, _
                            "SOFTWARE\Microsoft\Windows NT\CurrentVersion", _
                            "ProductKey")
    End If

    WindowsProductKey = strKey
End Function

Public Sub SoftwareNameError()
    Dim Msg As String
    Msg = "You haven't assigned the software name yet." & vbCrLf _
        & "Solution: Assign the software name property first." _
        & "Example: ActiveLock1.SoftwareName=""MyApp"""
    Err.Raise vbObjectError + 2, "ActiveLock", "ActiveLock Error"
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant

    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)

    QueryValue = vValue
    RegCloseKey (hKey)
End Function
