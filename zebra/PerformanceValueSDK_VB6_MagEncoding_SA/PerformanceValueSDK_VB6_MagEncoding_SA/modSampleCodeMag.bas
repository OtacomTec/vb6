Attribute VB_Name = "modSampleCodeMag"
'**********************************************
'* CONFIDENTIAL AND PROPRIETARY
'*
'* The source code and other information contained herein is the confidential and the exclusive property of
'* ZIH Corp. and is subject to the terms and conditions in your end user license agreement.
'* This source code, and any other information contained herein, shall not be copied, reproduced, published,
'* displayed or distributed, in whole or in part, in any medium, by any means, for any purpose except as
'* expressly permitted under such license agreement.
'*
'* Copyright ZIH Corp. 2010
'*
'* ALL RIGHTS RESERVED
'***********************************************
'File: modSampleCodeMag.bas
'Description: Example code showing how to apply magnetic encoding.
'$Revision: 1 $
'$Date: 2010/12/13 $
'*******************************************************************************/

Option Explicit

' Local Functions -------------------------------------------------------------------------------------------

' Byte Array to String --------------------------------------------------------------------------------------

Private Function ByteToString(buf() As Byte) As String

    On Error GoTo ByteToString_Error
    
    Dim i As Integer
    Dim s As String
    s = ""
    i = 0
    
    'Do
    '    If buf(i) <= 13 Then Exit Do
    '    s = s & Chr$(buf(i))
    '    i = i + 1
    'Loop
    
    While buf(i) > 13
        s = s & Chr$(buf(i))
        i = i + 1
    Wend
            
ByteToString_Exit:
    ByteToString = s
    
    On Error GoTo 0
    Exit Function
    
ByteToString_Error:
    s = ""
    MsgBox "Error in ByteToString: " & Err.Description
    GoTo ByteToString_Exit
End Function

' String to Byte Array --------------------------------------------------------------------------------------

Private Sub StringToByte(buf() As Byte, ByVal s As String)
    
    On Error GoTo StringToByte_Error
    
    Dim i As Integer
    For i = 1 To Len(s)
        buf(i - 1) = Asc(Mid(s, i, 1))
    Next
    
StringToByte_Exit:
    On Error GoTo 0
    Exit Sub
    
StringToByte_Error:
    MsgBox "Error in StringToByte: " & Err.Description
    GoTo StringToByte_Exit
End Sub

' Gets the ZBRPrinter.dll version ---------------------------------------------------------------------------

Public Sub GetPrinterDllVersion(ByRef version As String)

    On Error GoTo GetPrinterDllVersion_Error
    
    Dim engLevel    As Long
    Dim major       As Long
    Dim minor       As Long
    
    engLevel = 0
    major = 0
    minor = 0
    
    version = ""
    
    ZBRPRNGetSDKVer major, minor, engLevel
    If (major + minor + engLevel) = 0 Then
        version = "None"
    Else
        version = CStr(major) & "." & CStr(minor) + "." + CStr(engLevel)
    End If
        
GetPrinterDllVersion_Exit:
    On Error GoTo 0
    Exit Sub

GetPrinterDllVersion_Error:
    MsgBox "Error in GetPrinterDllVersion: " & Err.Description
    GoTo GetPrinterDllVersion_Exit
End Sub

' Magnetic Encoding Example Code ----------------------------------------------------------------------------

Public Sub MagCode(ByVal prnDriver As String, ByVal track1 As String, ByVal track2 As String, _
                    ByVal track3 As String, ByVal eject As Boolean, ByRef msg As String)
                    
    On Error GoTo MagCode_Error
    
    ' Opens a connection to a printer driver
    
    Dim handle      As Long
    Dim errValue    As Long
    Dim prnType     As Long
    
    If ZBRGetHandle(handle, prnDriver, prnType, errValue) = 0 Then
        msg = "Mag Encoder Error [" & CStr(errValue) & "]: Opening Printer Driver"
        GoTo MagCode_Exit
    End If
    
    ' Declares and clears the track buffers
    
    Dim inTrack1(20)    As Byte
    Dim inTrack2(20)    As Byte
    Dim inTrack3(20)    As Byte
    
    Dim i As Integer
    
    For i = 0 To 19
        inTrack1(i) = 0
        inTrack2(i) = 0
        inTrack3(i) = 0
    Next i

    ' Initializes the track buffers
    
    StringToByte inTrack1(), track1
    StringToByte inTrack2(), track2
    StringToByte inTrack3(), track3

    ' Encodes tracks 1,2,3
    '     if track data is "", that track is not encoded
    
    ' Note that we can only encode 6 characters or less to track 2
    ' when using printer firmware version lower than 2.00.03
    
    If ZBRPRNWriteMag(handle, prnType, 7, inTrack1(0), inTrack2(0), inTrack3(0), errValue) = 0 Then
        msg = "Mag Encoder Error [" & CStr(errValue) & "]: Writing Magnetic Tracks"
        GoTo MagCode_Exit
    End If
    
    ' Reads all 3 tracks
    
    Dim outTrack1(20)   As Byte
    Dim outTrack2(20)   As Byte
    Dim outTrack3(20)   As Byte
    Dim sz1             As Long
    Dim sz2             As Long
    Dim sz3             As Long
    
    If ZBRPRNReadMag(handle, prnType, 7, outTrack1(0), sz1, outTrack2(0), sz2, _
                                outTrack3(0), sz3, errValue) = 0 Then
        msg = "Mag Encoder Error [" & CStr(errValue) & "]: Reading Magnetic Tracks"
        GoTo MagCode_Exit
    End If
    
    ' Verifies that encoded equals read
    
    For i = 0 To 7
        If inTrack1(i) <> outTrack1(i) Or inTrack2(i) <> outTrack2(i) Or inTrack3(i) <> outTrack3(i) Then
            msg = "Mag Error : Verification failed"
            GoTo MagCode_Exit
        End If
    Next
    
MagCode_Exit:
    If eject = True Then ZBRPRNEjectCard handle, prnType, errValue
    
    On Error GoTo 0
    Exit Sub

MagCode_Error:
    MsgBox "Error in MagCode: " & Err.Description
    GoTo MagCode_Exit
End Sub



