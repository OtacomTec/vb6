Attribute VB_Name = "modZBRGraphics"
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
'File: modZBRGraphics.bas
'Description: A wrapper class for the Value/Performance Class SDK's Graphics functionality.
'$Revision: 1 $
'$Date: 2010/12/13 $
'*******************************************************************************/

Option Explicit

' Check Print Spooler ---------------------------------------------------------------------------------------

Public Declare Function ZBRGDIIsPrinterReady Lib "ZBRGRAPHICS.DLL" ( _
    ByVal devName As String, _
    ByRef errValue As Long) As Long

' DLL Version -----------------------------------------------------------------------------------------------

Public Declare Sub ZBRGDIGetSDKVer Lib "ZBRGRAPHICS.DLL" ( _
    ByRef major As Long, _
    ByRef minor As Long, _
    ByRef engLevel As Long)
    
' Initialization --------------------------------------------------------------------------------------------

Public Declare Function ZBRGDIInitGraphics Lib "ZBRGRAPHICS.DLL" ( _
    ByVal devName As String, _
    ByRef hDC As Long, _
    ByRef errValue As Long) As Long

Public Declare Function ZBRGDICloseGraphics Lib "ZBRGRAPHICS.DLL" ( _
    ByVal hDC As Long, _
    ByRef errValue As Long) As Long

Public Declare Function ZBRGDIClearGraphics Lib "ZBRGRAPHICS.DLL" ( _
    ByRef errValue As Long) As Long
    
' Print -----------------------------------------------------------------------------------------------------

Public Declare Function ZBRGDIPrintGraphics Lib "ZBRGRAPHICS.DLL" ( _
    ByVal hDC As Long, _
    ByRef errValue As Long) As Long

' Draw ------------------------------------------------------------------------------------------------------

Public Declare Function ZBRGDIDrawText Lib "ZBRGRAPHICS.DLL" ( _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal data As String, _
    ByVal font As String, _
    ByVal fontSize As Long, _
    ByVal fontType As Long, _
    ByVal fontColor As Long, _
    ByRef errValue As Long) As Long
    
Public Declare Function ZBRGDIDrawLine Lib "ZBRGRAPHICS.DLL" ( _
    ByVal x1 As Long, _
    ByVal y1 As Long, _
    ByVal x1 As Long, _
    ByVal y1 As Long, _
    ByVal color As Long, _
    ByVal thickness As Single, _
    ByRef errValue As Long) As Long
   
Public Declare Function ZBRGDIDrawImageRect Lib "ZBRGRAPHICS.DLL" ( _
    ByVal filename As String, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal sizeX As Long, _
    ByVal sizeY As Long, _
    ByRef errValue As Long) As Long

Public Declare Function ZBRGDIDrawBarCode Lib "ZBRGRAPHICS.DLL" ( _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal rotation As Long, _
    ByVal barcodeType As Long, _
    ByVal barWidthRatio As Long, _
    ByVal barcodeMultiplier As Long, _
    ByVal barCodeHeight As Long, _
    ByVal textUnder As Long, _
    ByVal barcodeData As String, _
    ByRef errValue As Long) As Long


