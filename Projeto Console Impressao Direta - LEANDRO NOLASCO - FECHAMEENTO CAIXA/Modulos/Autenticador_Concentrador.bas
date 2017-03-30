Attribute VB_Name = "Autenticador_Concentrador"
Public Const DATA_DELIMITER = vbCrLf & "====" & vbCrLf

Public Const MAX_USERS = 99 '0 - 99 = 100 Max users allowed to connect at one time
Public User(0 To MAX_USERS) As typUser 'Declare our user type

Type typUser 'Used for the 'User' variable
FreeSocket As Boolean
EncryptionString As String
HasAuthenticated As Boolean
End Type
Public Function GetRandomString(intLength As Integer) As String
    Dim intCharpos As Integer
    Dim intStrLen As Integer
    Dim strRandString As String
    Dim strChars As String
    strChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    intStrLen = 0
    Randomize Timer

    Do Until intStrLen = intLength
        intCharpos = Int((Len(strChars) * Rnd) + 1)
        strRandString = strRandString & Mid(strChars, intCharpos, 1)
        intStrLen = intStrLen + 1
    Loop
    GetRandomString = strRandString
End Function
Public Sub SendData(strData As String, intIndex As Integer, Form As Object)
   Form.sckServer(intIndex).SendData strData & DATA_DELIMITER 'Send the data
   ' frmConexoes_concentrador.sckServer(intIndex).SendData strData & DATA_DELIMITER 'Send the data
   DoEvents 'Dont leave this sub until the data is sent
End Sub
Public Sub DisconnectUser(intIndex As Integer)
    frmConexoes_concentrador.sckServer(intIndex).Close
    User(intIndex).FreeSocket = True
    User(intIndex).HasAuthenticated = False
    User(intIndex).EncryptionString = ""
    Log "Socket ID: " & intIndex & " was dropped."
End Sub
Public Function GenerateAuthString(intIndex As Integer) As String
    Dim strRandomString As String
    'Generates an authentication string, returns the auth string
    strRandomString = GetRandomString(100) 'Generate a 100 character random string
    User(intIndex).EncryptionString = strRandomString 'Set our version to plain text
    GenerateAuthString = TEncrypt(strRandomString) 'Encrypt the clients version
End Function
Public Function CheckAuthentication(strAuthString As String, intIndex As Integer) As Boolean
    'Returns TRUE if this user has sent a valid auth string, FALSE if not.
    If User(intIndex).EncryptionString = strAuthString Then 'If what the user has sent back
    'is what we sent (decrypted version) then this user is authentic!
    CheckAuthentication = True
    Else
    CheckAuthentication = False 'If they do not match, this user is not authentic
    End If
End Function
Public Sub Log(strLog As String)
    'Simply add an event to the log file, has no relevance to the authentication itself
    'but simply displays information to the user
    With frmConexoes_concentrador.rtbLog
        .SelColor = vbBlue
        .SelText = "[" & time & "] "
        .SelColor = vbBlack
        .SelText = strLog & vbCrLf
    End With
End Sub
Public Sub LogRAW(strLog As String)
    'Simply add an event to the log file, has no relevance to the authentication itself
    'but simply displays information to the user
    
    'Logs raw data received by winsock in red :)
    
    With frmMain.rtbLog
        .SelColor = vbRed
        .SelText = "[" & time & "] "
        .SelColor = vbBlack
        .SelText = "Socket Data: " & strLog & vbCrLf
    End With
End Sub
'Function created by Jeffrey C. Talum
'Avaliable at: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=8527&lngWId=1
Function TEncrypt(iString)
    On Error GoTo uhoh
    Q = ""
    a = GetRandomNumber(9) + 32
    b = GetRandomNumber(9) + 32
    C = GetRandomNumber(9) + 32
    d = GetRandomNumber(9) + 32
    Q = Chr(a) & Chr(C) & Chr(b)
    e = 1


    For X = 1 To Len(iString)
        f = Mid(iString, X, 1)
        If e = 1 Then Q = Q & Chr(Asc(f) + a)
        If e = 2 Then Q = Q & Chr(Asc(f) + C)
        If e = 3 Then Q = Q & Chr(Asc(f) + b)
        If e = 4 Then Q = Q & Chr(Asc(f) + d)
        e = e + 1
        If e > 4 Then e = 1
    Next X
    Q = Q & Chr(d)
    TEncrypt = Q
    Exit Function
uhoh:
    TEncrypt = "Error: Invalid text To Encrypt"
    Exit Function
End Function

'Function created by Jeffrey C. Talum
'Avaliable at: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=8527&lngWId=1
Function TDecrypt(iString)
    On Error GoTo uhohs
    Q = ""
    zz = Left(iString, 3)
    a = Left(zz, 1)
    b = Mid(zz, 2, 1)
    C = Mid(zz, 3, 1)
    d = Right(iString, 1)
    a = Int(Asc(a)) 'key 1
    b = Int(Asc(b)) 'key 2
    C = Int(Asc(C)) 'key 3
    d = Int(Asc(d)) 'key 4
    txt = Left(iString, Len(iString) - 1)
    txt2 = Mid(txt, 4, Len(txt)) 'encrypted text
    e = 1


    For X = 1 To Len(txt2)
        f = Mid(txt2, X, 1)
        If e = 1 Then Q = Q & Chr(Asc(f) - a)
        If e = 2 Then Q = Q & Chr(Asc(f) - b)
        If e = 3 Then Q = Q & Chr(Asc(f) - C)
        If e = 4 Then Q = Q & Chr(Asc(f) - d)
        e = e + 1
        If e > 4 Then e = 1
    Next X
    TDecrypt = Q
    Exit Function
uhohs:
    TDecrypt = "Error: Invalid text To Decrypt"
    Exit Function
End Function


Function GetRandomNumber(intSeed As Integer)
    Randomize
    GetRandomNumber = Int((Val(intSeed) * Rnd) + 1)
End Function

