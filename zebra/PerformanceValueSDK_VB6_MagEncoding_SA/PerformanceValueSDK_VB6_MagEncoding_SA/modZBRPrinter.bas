Attribute VB_Name = "modZBRPrinter"
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
'File: modZBRPrinter.bas
'Description: A wrapper class for the Value/Performance Class SDK's Printer/Magnetic Encoding functionality.
'$Revision: 1 $
'$Date: 2010/12/13 $
'*******************************************************************************/

Option Explicit

' DLL Version -----------------------------------------------------------------------------------------------

Public Declare Sub ZBRPRNGetSDKVer Lib "ZBRPrinter.dll" ( _
    ByRef major As Long, _
    ByRef minor As Long, _
    ByRef engLevel As Long)

' Handle ----------------------------------------------------------------------------------------------------

Public Declare Function ZBRGetHandle Lib "ZBRPrinter.dll" ( _
    ByRef prnHandle As Long, _
    ByVal drvName As String, _
    ByRef prnType As Long, _
    ByRef errValue As Long) As Long
                                                          
Public Declare Function ZBRCloseHandle Lib "ZBRPrinter.dll" ( _
    ByVal prnHandle As Long, _
    ByRef errValue As Long) As Long
                                                            
' Card Movement ---------------------------------------------------------------------------------------------

Public Declare Function ZBRPRNEjectCard Lib "ZBRPrinter.dll" ( _
    ByVal prnHandle As Long, _
    ByVal prnType As Long, _
    ByRef errValue As Long) As Long
    
' Magnetic Encoding -----------------------------------------------------------------------------------------

Public Declare Function ZBRPRNReadMag Lib "ZBRPrinter.dll" ( _
    ByVal prnHandle As Long, _
    ByVal prnType As Long, _
    ByVal trksToRead As Long, _
    ByRef trkBuf1 As Byte, _
    ByRef sz1 As Long, _
    ByRef trkBuf2 As Byte, _
    ByRef sz2 As Long, _
    ByRef trkBuf3 As Byte, _
    ByRef sz3 As Long, _
    ByRef errValue As Long) As Long
    
Public Declare Function ZBRPRNWriteMag Lib "ZBRPrinter.dll" ( _
    ByVal prnHandle As Long, _
    ByVal prnType As Long, _
    ByVal trksToWrite As Long, _
    ByRef trkBuf1 As Byte, _
    ByRef trkBuf2 As Byte, _
    ByRef trkBuf3 As Byte, _
    ByRef errValue As Long) As Long
        


