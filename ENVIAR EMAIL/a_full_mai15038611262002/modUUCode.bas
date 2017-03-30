Attribute VB_Name = "modUUCode"
Public Function UUDecodeToFile(strUUCodeData As String, strFilePath As String)
    On Error Resume Next
    Dim vDataLine   As Variant      'some variables needed for decoding
    Dim vDataLines  As Variant
    Dim strDataLine As String
    Dim intSymbols  As Integer
    Dim intFile     As Integer
    Dim strTemp     As String
    If Left$(strUUCodeData, 6) = "begin " Then  'check if it is a encoded file
        strUUCodeData = Mid$(strUUCodeData, InStr(1, strUUCodeData, vbLf) + 1)
    End If
    If Right$(strUUCodeData, 4) = "end" + vbLf Then 'check if "end" is available
        strUUCodeData = Left$(strUUCodeData, Len(strUUCodeData) - 7)
    End If
    intFile = FreeFile
    Open strFilePath For Binary As intFile  'open output file
        vDataLines = Split(strUUCodeData, vbLf)
        For Each vDataLine In vDataLines    'get every line
                strDataLine = CStr(vDataLine)
                intSymbols = Asc(Left$(strDataLine, 1)) 'get number of chars in
                                                        'one line. This is important
                                                        'for decoding
                strDataLine = Mid$(strDataLine, 2, intSymbols)
                For i = 1 To Len(strDataLine) Step 4
                    'now some decoding
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i, 1)) - 32) * 4 + _
                              (Asc(Mid(strDataLine, i + 1, 1)) - 32) \ 16)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 1, 1)) Mod 16) * 16 + _
                              (Asc(Mid(strDataLine, i + 2, 1)) - 32) \ 4)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 2, 1)) Mod 4) * 64 + _
                              Asc(Mid(strDataLine, i + 3, 1)) - 32)
                Next i
                'put the decoded data in the file
                Put intFile, , strTemp
                strTemp = ""
        Next
    'close the file
    Close intFile
End Function

Public Function UUEncodeFile(strFilePath As String) As String
    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    Dim strResult       As String
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    'This important: "begin 664"
    strResult = "begin 664 " + strFileName + vbLf
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize / 45 + 1
    'you need to encode every 45 bytes
    strFileData = Space(45)
    intFile = FreeFile
    'open the output file
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            If i = lEncodedLines Then
                strFileData = Space(lFileSize Mod 45)
            End If
            'get data
            Get intFile, , strFileData
            'the first byte in a line is a char, which number describes
            'how many bytes are in the line
            strTempLine = Chr(Len(strFileData) + 32)
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            'now some encoding
            For j = 1 To Len(strFileData) Step 3
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            strResult = strResult + strTempLine + vbLf
            strTempLine = ""
            'get next line
        Next i
        'close the file
    Close intFile
    'add the "end" string
    strResult = strResult & "'" & vbLf + "end" + vbLf
    'return the encoded string
    UUEncodeFile = strResult
End Function

