Attribute VB_Name = "modSampleCodeGraphics"
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
'File: modSampleCodeGraphics.bas
'Description: Example code showing how to print graphics and text to a card.
'$Revision: 1 $
'$Date: 2010/12/13 $
'*******************************************************************************/

Option Explicit

' Local Constants -------------------------------------------------------------------------------------------

Private Const FontBold = 1
Private Const FontItalic = 2
Private Const FontUnderline = 4
Private Const FontStrikeThru = 8

Private Const Red = &HFF0000

' Gets SDK Dll Version --------------------------------------------------------------------------------------

Public Sub GetGraphicsDllVersion(ByRef version As String)

    Dim engLevel    As Long
    Dim major       As Long
    Dim minor       As Long
    
    On Error GoTo GetGraphicsDllVersion_Error
    
    engLevel = 0
    major = 0
    minor = 0
    
    version = ""
    
    ZBRGDIGetSDKVer major, minor, engLevel
    If (major + minor + engLevel) <> 0 Then
        version = CStr(major) & "." & CStr(minor) + "." + CStr(engLevel)
    End If
    
GetGraphicsDllVersion_Exit:
    On Error GoTo 0
    Exit Sub
    
GetGraphicsDllVersion_Error:
    MsgBox "Error in GetGraphicsDllVersion: " & Err.Description
    GoTo GetGraphicsDllVersion_Exit
End Sub

' Printing on both sides ------------------------------------------------------------------------------------

Public Sub PrintBothSides(ByVal prnDriver As String, ByVal frontText As String, ByVal imgPath As String, _
                                ByVal backText, ByRef msg As String)

    On Error GoTo PrintBothSides_Error

    ' Gets a Device Context Handle and initializes a Graphics Buffer
    
    Dim hDC As Long
    Dim errValue As Long
    
    If ZBRGDIInitGraphics(prnDriver, hDC, errValue) = 0 Then
        msg = "Printing : InitGraphics : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_Exit
    End If
    
   On Error GoTo PrintBothSides_CloseGraphicsDevice
   
    ' Draws Text into the graphics buffer for the front side
    
    Dim fontStyle As Long
    fontStyle = FontBold Or FontItalic Or FontUnderline Or FontStrikeThru
    
    If ZBRGDIDrawText(35, 575, frontText, "Arial", 12, fontStyle, &HFF0000, errValue) = 0 Then
        msg = "Printing : DrawText : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    'Draws a Line into the graphics buffer
    
    If ZBRGDIDrawLine(35, 300, 300, 300, &HFF0000, 5#, errValue) = 0 Then
        msg = "Printing : DrawLine : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    ' Places an image into the graphics buffer
    
    If ZBRGDIDrawImageRect(imgPath, 30, 30, 200, 150, errValue) = 0 Then
        msg = "Printing : DrawImage : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    ' Sends barcode data to the monochrome buffer
    
    If ZBRGDIDrawBarCode(35, 500, 0, 0, 1, 3, 30, 1, "123456789", errValue) = 0 Then
        msg = "Printing : DrawBarCode : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    ' Prints the graphics buffer (front side)
    
    If ZBRGDIPrintGraphics(hDC, errValue) = 0 Then
        msg = "Printing : PrintGraphics : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If

    ' Clears the graphics buffer
    
    If ZBRGDIClearGraphics(errValue) = 0 Then
        msg = "Printing : ClearGraphics : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    ' Draws Text into the graphics buffer for the back side
    
    If ZBRGDIDrawText(30, 575, backText, "Arial", 12, fontStyle, 0, errValue) = 0 Then
        msg = "Printing : DrawText : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    ' Prints the graphics buffer (back side)
    
    If ZBRGDIPrintGraphics(hDC, errValue) = 0 Then
        msg = "Printing : PrintGraphics : Error[" & CStr(errValue) & "]"
        GoTo PrintBothSides_CloseGraphicsDevice
    End If
    
    ' Starts the printing process and releases the graphics buffer
    If ZBRGDICloseGraphics(hDC, errValue) = 0 Then
        msg = "Printing : CloseGraphics : Error[" & CStr(errValue) & "]"
    End If
    
PrintBothSides_Exit:
    On Error GoTo 0
    Exit Sub
    
PrintBothSides_CloseGraphicsDevice:
    On Error GoTo PrintBothSides_Error
    If ZBRGDICloseGraphics(hDC, errValue) = 0 Then
        msg = "Printing : CloseGraphics : Error[" & CStr(errValue) & "]"
    End If
    
PrintBothSides_Error:
    MsgBox "Error in PrintBothSides: " & Err.Description
    GoTo PrintBothSides_Exit
End Sub

' Printing on front side only -------------------------------------------------------------------------------

Public Sub PrintFrontSideOnly(ByVal prnDriver As String, ByVal text As String, ByVal imgPath As String, _
                                ByRef msg As String)

    On Error GoTo PrintFrontSideOnly_Error

    ' Gets a Device Context Handle and initializes a Graphics Buffer
    
    Dim hDC As Long
    Dim errValue As Long
    
    If ZBRGDIInitGraphics(prnDriver, hDC, errValue) = 0 Then
        msg = "Printing : InitGraphics : Error[" & CStr(errValue) & "]"
        GoTo PrintFrontSideOnly_Exit
    End If
    
    On Error GoTo PrintFrontSideOnly_CloseGraphicsDevice
    
    ' Draws Text into the graphics buffer for the front side
    
    Dim fontStyle As Long
    fontStyle = FontBold Or FontItalic Or FontUnderline Or FontStrikeThru
    
    If ZBRGDIDrawText(35, 575, text, "Arial", 12, fontStyle, &HFF0000, errValue) = 0 Then
        msg = "Printing : DrawText : Error[" & CStr(errValue) & "]"
        GoTo PrintFrontSideOnly_CloseGraphicsDevice
    End If
    
    'Draws a Line into the graphics buffer
    
    If ZBRGDIDrawLine(35, 300, 300, 300, &HFF0000, 5#, errValue) = 0 Then
        msg = "Printing : DrawLine : Error[" & CStr(errValue) & "]"
        GoTo PrintFrontSideOnly_CloseGraphicsDevice
    End If
    
    ' Places an image into the graphics buffer
    
    If ZBRGDIDrawImageRect(imgPath, 30, 30, 200, 150, errValue) = 0 Then
        msg = "Printing : DrawImage : Error[" & CStr(errValue) & "]"
        GoTo PrintFrontSideOnly_CloseGraphicsDevice
    End If
    
    ' Sends barcode data to the monochrome buffer
    
    If ZBRGDIDrawBarCode(35, 500, 0, 0, 1, 3, 30, 1, "123456789", errValue) = 0 Then
        msg = "Printing : DrawBarCode : Error[" & CStr(errValue) & "]"
        GoTo PrintFrontSideOnly_CloseGraphicsDevice
    End If
    
    ' Prints the graphics buffer (front side)
    
    If ZBRGDIPrintGraphics(hDC, errValue) = 0 Then
        msg = "Printing : PrintGraphics : Error[" & CStr(errValue) & "]"
        GoTo PrintFrontSideOnly_CloseGraphicsDevice
    End If

    ' Starts the printing process and releases the graphics buffer
    
    If ZBRGDICloseGraphics(hDC, errValue) = 0 Then
        msg = "Printing : CloseGraphics : Error[" & CStr(errValue) & "]"
    End If
    
PrintFrontSideOnly_Exit:
    On Error GoTo 0
    Exit Sub
    
PrintFrontSideOnly_CloseGraphicsDevice:
    On Error GoTo PrintFrontSideOnly_Error
    If ZBRGDICloseGraphics(hDC, errValue) = 0 Then
        msg = "Printing : CloseGraphics : Error[" & CStr(errValue) & "]"
    End If
    
PrintFrontSideOnly_Error:
    MsgBox "Error in PrintFrontSideOnly: " & Err.Description
    GoTo PrintFrontSideOnly_Exit
End Sub



