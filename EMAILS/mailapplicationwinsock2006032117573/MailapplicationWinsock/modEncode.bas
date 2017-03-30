Attribute VB_Name = "modEncode"
Public Enum SMTP_State                       'The smtp states enum
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Public Enum POP3States                       'The POP3 states enum
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_QUIT
End Enum

Dim strContentType(999) As String

Public Const BOUNDARY_ID As String = "NextMimePart"

Public Function GenerateMessageID(ByVal sHost As String) As String
    Dim idnum As Double                      'Declare idnum for local use as double
    Dim sMessageID As String                 'Declare smessageid for local use as string
    sMessageID = "Message-ID: "
    Randomize Int(CDbl((Now))) + Timer
    idnum = GetRandom(9999999999999#, 99999999999999#)
    sMessageID = sMessageID & CStr(idnum)    'Add cstr(idnum) to smessageid
    idnum = GetRandom(9999, 99999)
    sMessageID = sMessageID & "." & CStr(idnum) & ".qmail@" & sHost 'Add "." & cstr(idnum) & ".qmail@" & shost to smessageid
    GenerateMessageID = sMessageID
End Function
Private Function GetRandom(ByVal dFrom As Double, ByVal dTo As Double) As Double
    Dim x As Double                          'Declare x for local use as double
    Randomize
    x = dTo - dFrom
    GetRandom = Int((x * Rnd) + 1) + dFrom
End Function

Public Function GetMIMEHeader(ByVal vsBoundaryID As String) As String
    'related
    GetMIMEHeader = "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: multipart/mixed; boundary=" & _
        Chr(34) & vsBoundaryID & Chr(34) & "; type=" & Chr(34) & _
        "text/plain" & Chr(34) & vbCrLf & _
        "Text displayed only to non-MIME-compliant mailers" & vbCrLf & _
        "--" & vsBoundaryID & vbCrLf & _
        "Content-Type: text/plain; charset=iso-8859-1" & Chr$(13) & Chr$(10) & _
        "Content-Transfer-Encoding: 7bit" & Chr$(13) & Chr$(10)
        
        '"Content-Type: text/html; charset=iso-8859-1" & Chr$(13) & Chr$(10) & _
        '"Content-Transfer-Encoding: quoted printable" & Chr$(13) & Chr$(10)
        
End Function

Public Function EncodeFile(ByVal vsFullPathname As String, ByVal vsBoundaryID As String) As String
    Dim sResult As String
    Dim sFileName As String
    
    sFileName = GetFilename(vsFullPathname)
    
    'Preparing the Mime Header
    sResult = vbCrLf & "--" & vsBoundaryID & vbNewLine
    sResult = sResult & "Content-Type: " & SelectExt(vsFullPathname) & ";" & vbCrLf & vbTab & "name=" & Chr(34) & sFileName & Chr(34) & vbNewLine
    sResult = sResult & "Content-Transfer-Encoding: base64" & vbNewLine
    sResult = sResult & "Content-Disposition: attachment;" & vbCrLf & vbTab & "filename=" & Chr(34) & sFileName & Chr(34) & vbNewLine
    
    sResult = sResult & EncodeBase64(vsFullPathname)

    EncodeFile = sResult
    
End Function

Public Function GetFilename(ByVal vsFullPathname As String, Optional ByVal vbOmitExtension As Boolean = False) As String
    Dim iBackslashPos As Integer
    Dim iExtensionPos As Integer
    Dim i As Integer
    
    For i = Len(vsFullPathname) To 1 Step -1
        iBackslashPos = InStr(i, vsFullPathname, "\")
        If iBackslashPos > 0 Then Exit For
    Next
    
    If Not vbOmitExtension Then
        GetFilename = Mid(vsFullPathname, iBackslashPos + 1)
    Else
    
        For i = Len(vsFullPathname) To 1 Step -1
            iExtensionPos = InStr(i, vsFullPathname, ".")
            If iExtensionPos > 0 Then Exit For
        Next
        
        GetFilename = Mid(vsFullPathname, iBackslashPos + 1, iExtensionPos - iBackslashPos - 1)
    
    End If
    
End Function

Public Function EncodeBase64(ByVal vsFullPathname As String) As String
    'For Encoding BASE64
    Dim b           As Integer
    Dim Base64Tab   As Variant
    Dim bin(3)      As Byte
    Dim s           As String
    Dim l           As Long
    Dim i           As Long
    Dim FileIn      As Long
    Dim sResult     As String
    Dim n           As Long
    
    'Base64Tab=>tabla de tabulación
    Base64Tab = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
    
    Erase bin
    l = 0: i = 0: FileIn = 0: b = 0:
    s = ""
    
    'Gets the next free filenumber
    FileIn = FreeFile
    
    'Open Base64 Input File
    Open vsFullPathname For Binary As FileIn
    
    sResult = s & vbCrLf
    s = ""
    
    l = LOF(FileIn) - (LOF(FileIn) Mod 3)
    
    For i = 1 To l Step 3

        'Read three bytes
        Get FileIn, , bin(0)
        Get FileIn, , bin(1)
        Get FileIn, , bin(2)
        
        'Always wait until there're more then 64 characters
        If Len(s) > 64 Then

            s = s & vbCrLf
            sResult = sResult & s
            s = ""

        End If

        'Calc Base64-encoded char
        b = (bin(n) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
        s = s & Base64Tab(b) 'the character s holds the encoded chars
        
        b = ((bin(n) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
        s = s & Base64Tab(b)
        
        b = ((bin(n + 1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
        s = s & Base64Tab(b)
        
        b = bin(n + 2) And &H3F
        s = s & Base64Tab(b)
        
    Next i

    'Now, you need to check if there is something left
    If Not (LOF(FileIn) Mod 3 = 0) Then

        'Reads the number of bytes left
        For i = 1 To (LOF(FileIn) Mod 3)
            Get FileIn, , bin(i - 1)
        Next i
    
        'If there are only 2 chars left
        If (LOF(FileIn) Mod 3) = 2 Then
            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            
            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            
            b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
            s = s & Base64Tab(b)
            
            s = s & "="
        
        Else 'If there is only one char left
            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            
            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            
            s = s & "=="
        End If
    End If

    'Send the characters left
    If s <> "" Then
        s = s & vbCrLf
        sResult = sResult & s
    End If
    
    'Send the last part of the MIME Body
    s = ""
    
    Close FileIn
    EncodeBase64 = sResult
    
End Function

Private Function SelectExt(ByVal vsFullPathname As String) As String
    
    Dim extension As String
    extension = StrReverse(LCase$(Left(StrReverse(vsFullPathname), InStr(1, StrReverse(vsFullPathname), "."))))
    If Len(extension) = 0 Then extension = vsFullPathname
    
    Select Case extension
    
        Case ".323"
            SelectExt = "text/h323"
    
        Case ".aab"
            SelectExt = "application/x-authorware-bin"
    
        Case ".aam"
            SelectExt = "application/x-authorware-map"
    
        Case ".ace"
            SelectExt = "application/x-compressed"
    
        Case ".acp"
            SelectExt = "audio/x-mei-aac"
    
        Case ".ai"
            SelectExt = "application/postscript"
    
        Case ".aif"
            SelectExt = "audio/aiff"
    
        Case ".aifc"
            SelectExt = "audio/aiff"
    
        Case ".aiff"
            SelectExt = "audio/aiff"
    
        Case ".aip"
            SelectExt = "text/x-audiosoft-intra"
    
        Case ".ARJ"
            SelectExt = "application/x-compressed"
    
        Case ".art"
            SelectExt = "image/x-jg"
    
        Case ".asa"
            SelectExt = "text/asa"
    
        Case ".asf"
            SelectExt = "video/x-ms-asf"
    
        Case ".asp"
            SelectExt = "text/asp"
    
        Case ".asx"
            SelectExt = "video/x-ms-asf"
    
        Case ".asx"
            SelectExt = "video/x-ms-asx"
    
        Case ".au"
            SelectExt = "audio/basic"
    
        Case ".aut"
            SelectExt = "application/pbautomation"
    
        Case ".avi"
            SelectExt = "video/avi"
    
        Case ".avi"
            SelectExt = "video/x-msvideo"
    
        Case ".bmo"
            SelectExt = "audio/blue-matter-offer"
    
        Case ".bmp"
            SelectExt = "image/bmp"
    
        Case ".bmp"
            SelectExt = "image/x-bmp"
    
        Case ".bmr"
            SelectExt = "text/blue-matter-content-ref"
    
        Case ".bmt"
            SelectExt = "audio/blue-matter-song"
    
        Case ".bub"
            SelectExt = "application/photobubble"
    
        Case ".cat"
            SelectExt = "application/vnd.ms-pki.seccat"
    
        Case ".cdf"
            SelectExt = "application/x-cdf"
    
        Case ".cel"
            SelectExt = "video/flc"
    
        Case ".cer"
            SelectExt = "application/pkix-cert"
    
        Case ".cer"
            SelectExt = "application/x-x509-ca-cert"
    
        Case ".class"
            SelectExt = "java/*"
    
        Case ".crl"
            SelectExt = "application/pkix-crl"
    
        Case ".crt"
            SelectExt = "application/pkix-cert"
    
        Case ".crt"
            SelectExt = "application/x-x509-ca-cert"
    
        Case ".css"
            SelectExt = "text/css"
    
        Case ".dcr"
            SelectExt = "application/x-director"
    
        Case ".der"
            SelectExt = "application/pkix-cert"
    
        Case ".der"
            SelectExt = "application/x-x509-ca-cert"
    
        Case ".dib"
            SelectExt = "image/bmp"
    
        Case ".dib"
            SelectExt = "image/x-bmp"
    
        Case ".dif"
            SelectExt = "video/x-dv"
    
        Case ".dir"
            SelectExt = "application/x-director"
    
        Case ".dll"
            SelectExt = "application/x-msdownload"
    
        Case ".doc"
            SelectExt = "application/msword"
    
        Case ".dot"
            SelectExt = "application/msword"
    
        Case ".dpg"
            SelectExt = "application/vnd.dpgraph"
    
        Case ".dpgraph"
            SelectExt = "application/vnd.dpgraph"
    
        Case ".dv"
            SelectExt = "video/x-dv"
    
        Case ".dxr"
            SelectExt = "application/x-director"
    
        Case ".eml"
            SelectExt = "message/rfc822"
    
        Case ".emm"
            SelectExt = "application/x-emms-content"
    
        Case ".eps"
            SelectExt = "application/postscript"
    
        Case ".exe"
            SelectExt = "application/x-msdownload"
    
        Case ".fdf"
            SelectExt = "application/vnd.fdf"
    
        Case ".fif"
            SelectExt = "application/fractals"
    
        Case ".flc"
            SelectExt = "video/flc"
    
        Case ".fli"
            SelectExt = "video/flc"
    
        Case ".fml"
            SelectExt = "application/file-mirror-list"
    
        Case ".fpx"
            SelectExt = "image/x-xbitmap"
    
        Case ".gif"
            SelectExt = "image/gif"
    
        Case ".grv"
            SelectExt = "application/vnd.groove-injector"
    
        Case ".gz"
            SelectExt = "application/x-compressed"
    
        Case ".gz"
            SelectExt = "application/x-gzip"
    
        Case ".hpf"
            SelectExt = "application/x-icq-hpf"
    
        Case ".hqx"
            SelectExt = "application/mac-binhex40"
    
        Case ".hta"
            SelectExt = "application/hta"
    
        Case ".htc"
            SelectExt = "text/x-component"
    
        Case ".htm"
            SelectExt = "text/html"
    
        Case ".html"
            SelectExt = "text/html"
    
        Case ".htt"
            SelectExt = "text/webviewhtml"
    
        Case ".htx"
            SelectExt = "text/html"
    
        Case ".ico"
            SelectExt = "image/x-icon"
    
        Case ".iii"
            SelectExt = "application/x-iphone"
    
        Case ".ins"
            SelectExt = "application/x-internet-signup"
    
        Case ".ips"
            SelectExt = "application/x-ipscript"
    
        Case ".ipx"
            SelectExt = "application/x-ipix"
    
        Case ".isp"
            SelectExt = "application/x-internet-signup"
    
        Case ".IVF"
            SelectExt = "video/x-ivf"
    
        Case ".ivr"
            SelectExt = "i-world/i-vrml"
    
        Case ".java"
            SelectExt = "java/*"
    
        Case ".java"
            SelectExt = "text/java"
    
        Case ".jfif"
            SelectExt = "image/pjpeg"
    
        Case ".jpe"
            SelectExt = "image/jpeg"
    
        Case ".jpeg"
            SelectExt = "image/jpeg"
    
        Case ".jpg"
            SelectExt = "image/jpeg"
    
        Case ".JS"
            SelectExt = "application/x-javascript"
    
        Case ".la1"
            SelectExt = "audio/x-liquid-file"
    
        Case ".lar"
            SelectExt = "application/x-laplayer-reg"
    
        Case ".latex"
            SelectExt = "application/x-latex"
    
        Case ".lav"
            SelectExt = "audio/x-liquid"
    
        Case ".lavs"
            SelectExt = "audio/x-liquid-secure"
    
        Case ".lha"
            SelectExt = "application/x-compressed"
    
        Case ".lks"
            SelectExt = "application/x-lk-rlestream"
    
        Case ".lmsff"
            SelectExt = "audio/x-la-lms"
    
        Case ".lqt"
            SelectExt = "audio/x-liquid-file"
    
        Case ".ls"
            SelectExt = "application/x-javascript"
    
        Case ".lsf"
            SelectExt = "video/x-la-asf"
    
        Case ".lsx"
            SelectExt = "video/x-la-asf"
    
        Case ".LZH"
            SelectExt = "application/x-compressed"
    
        Case ".m1v"
            SelectExt = "video/mpeg"
    
        Case ".m3u"
            SelectExt = "audio/mpegurl"
    
        Case ".m3u"
            SelectExt = "audio/x-mpegurl"
    
        Case ".mac"
            SelectExt = "image/x-macpaint"
    
        Case ".man"
            SelectExt = "application/x-troff-man"
    
        Case ".mbc"
            SelectExt = "application/x-pn-virtualink"
    
        Case ".mbo"
            SelectExt = "application/x-previewsystems-vbox-music"
    
        Case ".mbox"
            SelectExt = "application/x-previewsystems-vbox-music"
    
        Case ".mdb"
            SelectExt = "application/msaccess"
    
        Case ".med"
            SelectExt = "application/x-att-a2bmusic-purchase"
    
        Case ".mes"
            SelectExt = "application/x-att-a2bmusic"
    
        Case ".mht"
            SelectExt = "message/rfc822"
    
        Case ".mhtml"
            SelectExt = "message/rfc822"
    
        Case ".mid"
            SelectExt = "audio/mid"
    
        Case ".midi"
            SelectExt = "audio/mid"
    
        Case ".mix"
            SelectExt = "image/x-xbitmap"
    
        Case ".mjf"
            SelectExt = "audio/x-vnd.AudioExplosion.MjuiceMediaFile"
    
        Case ".mjv"
            SelectExt = "audio/audio/mjuice_voucher"
    
        Case ".mmjb_mime"
            SelectExt = "application/x-musicmatch-mmjb5.20detect"
    
        Case ".mmz"
            SelectExt = "application/x-mmjb-mmz"
    
        Case ".mocha"
            SelectExt = "application/x-javascript"
    
        Case ".mov"
            SelectExt = "video/quicktime"
    
        Case ".movie"
            SelectExt = "video/x-sgi-movie"
    
        Case ".mp1"
            SelectExt = "audio/mpeg"
    
        Case ".mp2"
            SelectExt = "video/mpeg"
    
        Case ".mp2v"
            SelectExt = "video/mpeg"
    
        Case ".mp3"
            SelectExt = "audio/mpeg"
    
        Case ".mpa"
            SelectExt = "video/mpeg"
    
        Case ".mpe"
            SelectExt = "video/mpeg"
    
        Case ".mpeg"
            SelectExt = "video/mpeg"
    
        Case ".mpg"
            SelectExt = "video/mpeg"
    
        Case ".mpga"
            SelectExt = "audio/mpeg"
    
        Case ".mpv"
            SelectExt = "video/mpg"
    
        Case ".mpv2"
            SelectExt = "video/mpeg"
    
        Case ".mwc"
            SelectExt = "application/vnd.dpgraph"
    
        Case ".mxp"
            SelectExt = "application/x-mmxp"
    
        Case ".npi"
            SelectExt = "application/x-pn-npistream"
    
        Case ".nws"
            SelectExt = "message/rfc822"
    
        Case ".p10"
            SelectExt = "application/pkcs10"
    
        Case ".p12"
            SelectExt = "application/x-pkcs12"
    
        Case ".p7b"
            SelectExt = "application/x-pkcs7-certificates"
    
        Case ".p7c"
            SelectExt = "application/pkcs7-mime"
    
        Case ".p7m"
            SelectExt = "application/pkcs7-mime"
    
        Case ".p7r"
            SelectExt = "application/x-pkcs7-certreqresp"
    
        Case ".p7s"
            SelectExt = "application/pkcs7-signature"
    
        Case ".pct"
            SelectExt = "image/pict"
    
        Case ".pdf"
            SelectExt = "application/pdf"
    
        Case ".pfx"
            SelectExt = "application/x-pkcs12"
    
        Case ".pic"
            SelectExt = "image/pict"
    
        Case ".pict"
            SelectExt = "image/pict"
    
        Case ".pko"
            SelectExt = "application/vnd.ms-pki.pko"
    
        Case ".pl"
            SelectExt = "application/x-perl"
    
        Case ".plg"
            SelectExt = "text/html"
    
        Case ".pls"
            SelectExt = "audio/scpls"
    
        Case ".pls"
            SelectExt = "audio/x-scpls"
    
        Case ".png"
            SelectExt = "image/png"
    
        Case ".pnq"
            SelectExt = "application/x-icq-pnq"
    
        Case ".pntg"
            SelectExt = "image/x-macpaint"
    
        Case ".POT"
            SelectExt = "application/vnd.ms-powerpoint"
    
        Case ".ppa"
            SelectExt = "application/vnd.ms-powerpoint"
    
        Case ".pps"
            SelectExt = "application/vnd.ms-powerpoint"
    
        Case ".ppt"
            SelectExt = "application/x-mspowerpoint"
    
        Case ".prf"
            SelectExt = "application/pics-rules"
    
        Case ".ps"
            SelectExt = "application/postscript"
    
        Case ".pwz"
            SelectExt = "application/vnd.ms-powerpoint"
    
        Case ".py"
            SelectExt = "text/plain"
    
        Case ".pyw"
            SelectExt = "text/plain"
    
        Case ".qht"
            SelectExt = "text/x-html-insertion"
    
        Case ".qhtm"
            SelectExt = "text/x-html-insertion"
    
        Case ".qt"
            SelectExt = "video/quicktime"
    
        Case ".qti"
            SelectExt = "image/x-quicktime"
    
        Case ".qtif"
            SelectExt = "image/x-quicktime"
    
        Case ".qtl"
            SelectExt = "application/x-quicktimeplayer"
    
        Case ".r00"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r01"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r02"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r03"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r04"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r05"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r06"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r07"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r08"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r09"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r10"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r11"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r12"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r13"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r14"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r15"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r16"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r17"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r18"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r19"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r20"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r21"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r22"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r23"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r24"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r25"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r26"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r27"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r28"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r29"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r30"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r31"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r32"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r33"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r34"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r35"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r36"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r37"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r38"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r39"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r3t"
            SelectExt = "text/vnd.rn-realtext3d"
    
        Case ".r40"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r41"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r42"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r43"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r44"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r45"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r46"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r47"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r48"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r49"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r50"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r51"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r52"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r53"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r54"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r55"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r56"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r57"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r58"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r59"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r60"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r61"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r62"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r63"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r64"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r65"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r66"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r67"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r68"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r69"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r70"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r71"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r72"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r73"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r74"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r75"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r76"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r77"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r78"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r79"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r80"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r81"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r82"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r83"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r84"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r85"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r86"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r87"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r88"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r89"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r90"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r91"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r92"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r93"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r94"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r95"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r96"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r97"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r98"
            SelectExt = "application/x-rar-compressed"
    
        Case ".r99"
            SelectExt = "application/x-rar-compressed"
    
        Case ".ra"
            SelectExt = "audio/vnd.rn-realaudio"
    
        Case ".ram"
            SelectExt = "audio/x-pn-realaudio"
    
        Case ".rar"
            SelectExt = "application/x-rar-compressed"
    
        Case ".rat"
            SelectExt = "application/rat-file"
    
        Case ".rf"
            SelectExt = "image/vnd.rn-realflash"
    
        Case ".rjs"
            SelectExt = "application/vnd.rn-realsystem-rjs"
    
        Case ".rjt"
            SelectExt = "application/vnd.rn-realsystem-rjt"
    
        Case ".rm"
            SelectExt = "application/vnd.rn-realmedia"
    
        Case ".rmi"
            SelectExt = "audio/mid"
    
        Case ".rmj"
            SelectExt = "application/vnd.rn-realsystem-rmj"
    
        Case ".rmm"
            SelectExt = "audio/x-pn-realaudio"
    
        Case ".rmp"
            SelectExt = "application/vnd.rn-rn_music_package"
    
        Case ".rmx"
            SelectExt = "application/vnd.rn-realsystem-rmx"
    
        Case ".rnx"
            SelectExt = "application/vnd.rn-realplayer"
    
        Case ".rob"
            SelectExt = "application/vnd.rn-objects"
    
        Case ".rp"
            SelectExt = "image/vnd.rn-realpix"
    
        Case ".rpm"
            SelectExt = "audio/x-pn-realaudio-plugin"
    
        Case ".rsml"
            SelectExt = "application/vnd.rn-rsml"
    
        Case ".rt"
            SelectExt = "text/vnd.rn-realtext"
    
        Case ".rtf"
            SelectExt = "application/msword"
    
        Case ".rtsp"
            SelectExt = "application/x-rtsp"
    
        Case ".rv"
            SelectExt = "video/vnd.rn-realvideo"
    
        Case ".sc"
            SelectExt = "application/vnd.optx-screenwatch"
    
        Case ".scm"
            SelectExt = "application/x-icq-scm"
    
        Case ".sct"
            SelectExt = "text/scriptlet"
    
        Case ".sd2"
            SelectExt = "audio/x-sd2"
    
        Case ".sdf"
            SelectExt = "application/x-server-launch"
    
        Case ".sdp"
            SelectExt = "application/sdp"
    
        Case ".sgi"
            SelectExt = "image/x-sgi"
    
        Case ".sit"
            SelectExt = "application/x-stuffit"
    
        Case ".sma"
            SelectExt = "application/x-smb-directive"
    
        Case ".smi"
            SelectExt = "application/smil"
    
        Case ".smil"
            SelectExt = "application/smil"
    
        Case ".snd"
            SelectExt = "audio/basic"
    
        Case ".spc"
            SelectExt = "application/x-pkcs7-certificates"
    
        Case ".spl"
            SelectExt = "application/futuresplash"
    
        Case ".spn"
            SelectExt = "application/vnd.spinnerplus"
    
        Case ".ssm"
            SelectExt = "application/streamingmedia"
    
        Case ".sst"
            SelectExt = "application/vnd.ms-pki.certstore"
    
        Case ".stl"
            SelectExt = "application/vnd.ms-pki.stl"
    
        Case ".stm"
            SelectExt = "text/html"
    
        Case ".svg"
            SelectExt = "image/svg+xml"
    
        Case ".svg"
            SelectExt = "image/svg-xml"
    
        Case ".svgz"
            SelectExt = "image/svg+xml"
    
        Case ".svgz"
            SelectExt = "image/svg-xml"
    
        Case ".swf"
            SelectExt = "application/x-shockwave-flash"
    
        Case ".tar"
            SelectExt = "application/x-compressed"
    
        Case ".tar"
            SelectExt = "application/x-tar"
    
        Case ".tga"
            SelectExt = "image/x-targa"
    
        Case ".tgz"
            SelectExt = "application/x-compressed"
    
        Case ".tif"
            SelectExt = "image/tiff"
    
        Case ".tiff"
            SelectExt = "image/tiff"
    
        Case ".txt"
            SelectExt = "text/plain"
    
        Case ".uin"
            SelectExt = "application/x-icq"
    
        Case ".uls"
            SelectExt = "text/iuls"
    
        Case ".ultact"
            SelectExt = "application/x-UltimateAction"
    
        Case ".ulw"
            SelectExt = "audio/basic"
    
        Case ".urls"
            SelectExt = "application/x-url-list"
    
        Case ".UU"
            SelectExt = "application/x-compressed"
    
        Case ".UUE"
            SelectExt = "application/x-compressed"
    
        Case ".vcf"
            SelectExt = "text/x-vcard"
    
        Case ".vcg"
            SelectExt = "application/vnd.groove-vcard"
    
        Case ".vcl"
            SelectExt = "text/html"
    
        Case ".vfw"
            SelectExt = "video/x-msvideo"
    
        Case ".vpg"
            SelectExt = "application/x-vpeg"
    
        Case ".vsl"
            SelectExt = "application/x-cnet-vsl"
    
        Case ".wav"
            SelectExt = "audio/wav"
    
        Case ".wax"
            SelectExt = "audio/x-ms-wax"
    
        Case ".wiz"
            SelectExt = "application/msword"
    
        Case ".wm"
            SelectExt = "video/x-ms-wm"
    
        Case ".wma"
            SelectExt = "audio/x-ms-wma"
    
        Case ".wmd"
            SelectExt = "application/x-ms-wmd"
    
        Case ".wme"
            SelectExt = "text/xml"
    
        Case ".wmp"
            SelectExt = "video/x-ms-wmp"
    
        Case ".wms"
            SelectExt = "application/x-ms-wms"
    
        Case ".wmv"
            SelectExt = "video/x-ms-wmv"
    
        Case ".wmx"
            SelectExt = "video/x-ms-wmx"
    
        Case ".wmz"
            SelectExt = "application/x-ms-wmz"
    
        Case ".wsc"
            SelectExt = "text/scriptlet"
    
        Case ".wvx"
            SelectExt = "video/x-ms-wvx"
    
        Case ".xbm"
            SelectExt = "image/x-xbitmap"
    
        Case ".xls"
            SelectExt = "application/vnd.ms-excel"
    
        Case ".xls"
            SelectExt = "application/x-msexcel"
    
        Case ".xml"
            SelectExt = "text/xml"
    
        Case ".xpl"
            SelectExt = "audio/mpegurl"
    
        Case ".xsl"
            SelectExt = "text/xml"
    
        Case ".XXE"
            SelectExt = "application/x-compressed"
    
        Case ".ymg"
            SelectExt = "application/ymsgr"
    
        Case ".yps"
            SelectExt = "application/ymsgr"
    
        Case ".z"
            SelectExt = "application/x-compress"
    
        Case ".zip"
            SelectExt = "application/x-zip-compressed"
    
        Case "ratfile"
            SelectExt = "application/rat-file"
    
        Case "smafile"
            SelectExt = "application/x-smb-directive"
        
        Case Else
            SelectExt = "application/octet-stream"
    End Select

End Function
