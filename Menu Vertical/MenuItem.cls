VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private picMenu As PictureBox
Private picCache As PictureBox
Private msCaption As String      ' caption of MenuItem
Private mlIndex As Long          ' index of icon on Menu (1 based)
Private picButton As Picture    ' icon picture
Private msCaptionX As Long
Private msCaptionY As Long
Private mlButtonHeight As Long
Private mbButtonDownOnMe As Boolean
Private msPictureURL As String
Private msKey As String
Private msTag As String

Private Type BUTTON_STRUCT
    RECT     As RECT
    State    As Long
    OnScreen As Boolean
End Type
Private mButtonStruct As BUTTON_STRUCT

Private mHitStruct As RECT
Private m3DStruct As RECT

#If USE_WING Then
    Private Declare Function WinGBitBlt Lib "wing32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long) As Long
#Else
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
#End If
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal SavedDC As Long) As Long

Const SRCCOPY = &HCC0020

Const ICON_SIZE = 32
Const MOUSE_UP = 1
Const MOUSE_DOWN = -1
Const MOUSE_MOVE = 0
Const RAISED = 1
Const SUNKEN = -1
Const NONE = 0
Const HITTEXT_EXTRA_PIXELS = 4
Const CLIPPING_NO = True
Const CLIPPING_YES = False

Public Property Get Caption() As String
    On Error Resume Next
    Caption = msCaption
End Property

Public Property Let Caption(ByVal sNewValue As String)
    On Error Resume Next
    msCaption = sNewValue
End Property

Public Property Get Index() As Long
    On Error Resume Next
    Index = mlIndex
End Property

Public Property Let Index(ByVal lNewValue As Long)
    On Error Resume Next
    mlIndex = lNewValue
End Property

Public Property Get Button() As Object
    On Error Resume Next
    Set Button = picButton
End Property

Public Property Set Button(ByVal vNewValue As Object)
    On Error Resume Next
    Set picButton = vNewValue
End Property

Public Property Get Left() As Long
    On Error Resume Next
    Left = mButtonStruct.RECT.Left
End Property

Public Property Let Left(ByVal lNewValue As Long)
    On Error Resume Next
    With mButtonStruct.RECT
        .Left = lNewValue
        .Right = lNewValue + ICON_SIZE
    End With
End Property

Public Property Get Top() As Long
    On Error Resume Next
    Top = mButtonStruct.RECT.Top
End Property

Public Property Let Top(ByVal lNewValue As Long)
    On Error Resume Next
    With mButtonStruct.RECT
        .Top = lNewValue
        .Bottom = lNewValue + ICON_SIZE
    End With
End Property

Public Property Get Right() As Long
    On Error Resume Next
    Right = mButtonStruct.RECT.Right
End Property

Public Property Get Bottom() As Long
    On Error Resume Next
    Bottom = mButtonStruct.RECT.Bottom
End Property

Public Property Get State() As Long
    On Error Resume Next
    State = mButtonStruct.State
End Property

Public Property Let State(ByVal lNewValue As Long)
    On Error Resume Next
    mButtonStruct.State = lNewValue
End Property

Public Property Get CaptionX() As Long
    On Error Resume Next
    CaptionX = msCaptionX
End Property

Public Property Let CaptionX(ByVal lNewValue As Long)
    On Error Resume Next
    msCaptionX = lNewValue
End Property

Public Property Get CaptionY() As Long
    On Error Resume Next
    CaptionY = msCaptionY
End Property

Public Property Let CaptionY(ByVal lNewValue As Long)
    On Error Resume Next
    msCaptionY = lNewValue
End Property

' paint the icon (32x32) and its caption
Public Function PaintButton(lTopMenuItemDisplayed, lIconStart As Long, lMenuCur As Long, lClipY As Long) As Boolean
    Dim lCenter As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim lRight As Long
    Dim lBottom As Long
    Dim lResult As Long
    Dim lHeight As Long
    Dim bClipping As Boolean
    Dim lPositionFromTop As Long
    Dim RgnRect As RECT
    Dim hRgn As Long
    Dim lRetCod As Long
    Dim hSavedDC As Long
    
    On Error Resume Next
    
    If mlIndex < lTopMenuItemDisplayed Then
        mButtonStruct.OnScreen = False
        PaintButton = CLIPPING_NO
        Exit Function
    End If
    
    ' position the image
    lPositionFromTop = mlIndex - lTopMenuItemDisplayed + 1
    Top = (lPositionFromTop * 2 * ICON_SIZE) - ICON_SIZE + ((lPositionFromTop + 1 = 1) * 4) + (lMenuCur - 1) * mlButtonHeight
    With picMenu
        .ScaleMode = vbPixels
        lCenter = .ScaleWidth \ 2
        Left = lCenter - (ICON_SIZE \ 2)
    End With
    
    With mButtonStruct
        lLeft = .RECT.Left
        lTop = .RECT.Top
        lRight = .RECT.Right
        lBottom = .RECT.Bottom
    
        ' see if it will fit in the control's viewing area
        If lTop > lClipY Then
            .OnScreen = False
            PaintButton = CLIPPING_YES
            Exit Function
        End If
        If lBottom > lClipY Then
            bClipping = True
            lBottom = lClipY
        End If
        .OnScreen = True
    End With
    ' position the menu caption
    CaptionX = lCenter - (CLng(picMenu.TextWidth(Caption())) \ 2)
    CaptionY = lTop + ICON_SIZE + 4

    ' calculate the hittest structure
    With mHitStruct
        .Left = lLeft - HITTEXT_EXTRA_PIXELS - 2
        .Top = lTop - HITTEXT_EXTRA_PIXELS - 2
        .Right = lRight + HITTEXT_EXTRA_PIXELS + 2
        ' hittest includes the caption below the icon
        .Bottom = lBottom + picMenu.TextHeight(Caption()) + 5
        If bClipping Then
            .Bottom = lBottom
        End If
    End With
        
    ' calculate the 3d structure
    With m3DStruct
        .Left = lLeft - 2
        .Top = lTop - 2
        .Right = lRight + 2
        ' hittest includes the caption below the icon
        If Not bClipping Then
            .Bottom = lBottom + 2
        Else
            .Bottom = lBottom
        End If
    End With
    
    With mButtonStruct.RECT
        If Not bClipping Then
            lHeight = ICON_SIZE
        Else
            lHeight = lBottom - lTop
        End If
#If USE_WING Then
        lResult = WinGBitBlt(picMenu.hdc, .Left, _
            .Top, _
            ICON_SIZE, lHeight, _
            picCache.hdc, 0, mlButtonHeight * 2 + (lIconStart + lPositionFromTop) * ICON_SIZE)
#Else
        lResult = BitBlt(picMenu.hdc, .Left, _
            .Top, _
            ICON_SIZE, lHeight, _
            picCache.hdc, 0, mlButtonHeight * 2 + (lIconStart + lPositionFromTop) * ICON_SIZE, SRCCOPY)
#End If
    End With
    
    ' bClipping is set just for the icon
    ' if we are already clipping, set a clipping region
    ' so we can display part of the caption.
    
    ' position the caption
    If Not bClipping Then
        With picMenu
            .CurrentX = msCaptionX
            .CurrentY = msCaptionY
            .ForeColor = vbWhite
            If .CurrentY + .TextHeight(msCaption) < lClipY Then
                picMenu.Print msCaption
                PaintButton = CLIPPING_NO
            Else
                ' set the region
                With picMenu
                    RgnRect.Left = 0
                    RgnRect.Top = msCaptionY
                    RgnRect.Right = .Width
                    RgnRect.Bottom = lClipY
                    ' save the original DC
                    hSavedDC = SaveDC(.hdc)
                    ' create a region for the text
                    hRgn = CreateRectRgnIndirect(RgnRect)
                    ' set clipping
                    lRetCod = SelectClipRgn(.hdc, hRgn)
                    ' print the caption
                    picMenu.Print msCaption
                    ' delete the object
                    hRgn = DeleteObject(hRgn)
                    ' restore the original DC
                    lRetCod = RestoreDC(.hdc, hSavedDC)
                    
                    PaintButton = CLIPPING_YES
                End With
            End If
        End With
    Else
        PaintButton = CLIPPING_YES
    End If
End Function

Public Property Set Parent(ByVal picNewValue As Control)
    On Error Resume Next
    Set picMenu = picNewValue
End Property

Public Function HitTest(ByVal iMousePosition As Integer, ByVal X As Long, ByVal Y As Long) As Boolean
    
    ' don't bother if it is not on screen
    If Not mButtonStruct.OnScreen Then
        Exit Function
    End If
    
    If PtInRect(mHitStruct, X, Y) Then
        HitTest = True
        
        Select Case iMousePosition
            Case MOUSE_UP
                Select Case mButtonStruct.State
                    Case SUNKEN, NONE
                        DrawBorder RAISED
                    Case Else
                        ' nothing to do
                End Select
                mbButtonDownOnMe = False
            Case MOUSE_DOWN
                Select Case mButtonStruct.State
                    Case SUNKEN
                        ' nothing to do - it's already drawn
                    Case Else
                        DrawBorder SUNKEN
                        mbButtonDownOnMe = True
                End Select
            Case MOUSE_MOVE
                Select Case mButtonStruct.State
                    Case RAISED
                        ' nothing to do - it's already drawn
                    Case NONE
                        ' if the mouse went down on me, moved off me
                        ' and now returns and no mouse up yet, draw
                        ' me as down
                        If Not mbButtonDownOnMe Then
                            DrawBorder RAISED
                        Else
                            DrawBorder SUNKEN
                        End If
                    Case SUNKEN
                        ' leave it that way
                End Select
        End Select
    Else
        ' there is no hit
        HitTest = False
        
        If iMousePosition <> MOUSE_MOVE Then
            mbButtonDownOnMe = False
        End If
        
        ' if any border is currently drawn, remove it
        If mButtonStruct.State <> NONE Then
            DrawBorder NONE
        End If
    End If
End Function

Public Sub DrawBorder(iDirection As Integer)
    On Error Resume Next
    picMenu.ScaleMode = vbPixels
       
    If Not mButtonStruct.OnScreen Then
        Exit Sub
    End If
    
    ' save the state of the button
    State = iDirection
    
    ' icon not clipped
    If m3DStruct.Bottom - m3DStruct.Top = ICON_SIZE + 4 Then
        Select Case iDirection
            Case RAISED
                DrawEdge picMenu.hdc, m3DStruct, BDR_RAISEDOUTER, BF_RECT
            Case SUNKEN
                DrawEdge picMenu.hdc, m3DStruct, BDR_SUNKENINNER, BF_RECT
            Case NONE
                With m3DStruct
                    picMenu.Line (.Left, .Top)-(.Right - 1, .Bottom - 1), BACKGROUND_COLOR, B
                End With
        End Select
    ' icon clipped
    Else
        Select Case iDirection
            Case RAISED
                DrawEdge picMenu.hdc, m3DStruct, BDR_RAISEDOUTER, BF_LEFT Or BF_TOP Or BF_RIGHT
            Case SUNKEN
                DrawEdge picMenu.hdc, m3DStruct, BDR_SUNKENINNER, BF_LEFT Or BF_TOP Or BF_RIGHT
            Case NONE
                With m3DStruct
                    picMenu.Line (.Left, .Top)-(.Right - 1, .Top), BACKGROUND_COLOR
                    picMenu.Line (.Left, .Top)-(.Left, .Bottom), BACKGROUND_COLOR
                    picMenu.Line (.Right - 1, .Top)-(.Right - 1, .Bottom), BACKGROUND_COLOR
                End With
        End Select
    End If
End Sub

Public Property Set Cache(ByVal oNewValue As Object)
    On Error Resume Next
    Set picCache = oNewValue
End Property

Public Property Let ButtonHeight(ByVal lNewValue As Long)
    On Error Resume Next
    mlButtonHeight = lNewValue
End Property

Public Property Get PictureURL() As String
    On Error Resume Next
    PictureURL = msPictureURL
End Property

Public Property Let PictureURL(ByVal sNewValue As String)
    On Error Resume Next
    msPictureURL = PictureURL
End Property

Public Property Get Key() As String
    On Error Resume Next
    Key = msKey
End Property

Public Property Let Key(ByVal sNewValue As String)
    On Error Resume Next
    msKey = sNewValue
End Property

Public Property Get Tag() As String
    On Error Resume Next
    Tag = msTag
End Property

Public Property Let Tag(ByVal sNewValue As String)
    On Error Resume Next
    msTag = sNewValue
End Property
