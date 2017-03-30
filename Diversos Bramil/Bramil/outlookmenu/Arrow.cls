VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mpicUp As Picture
Private mParent As PictureBox
Private mIsDisplayed As Boolean
Private mImgRect As RECT
Private mArrowType As Integer
Private mlButtonHeight As Long
Private mlState As Long
Private mbLastButtonDown As Boolean

Const PIXELS_FROM_TOP = 6
Const PIXELS_FROM_RIGHT = 6
Const PIXELS_FROM_BOTTOM = 6
Const PIXEL_WIDTH = 16
Const ARROW_UP = 1
Const ARROW_DOWN = -1
Const MOUSE_UP = 1
Const MOUSE_DOWN = -1
Const MOUSE_MOVE = 0
Const RAISED = 1
Const SUNKEN = -1

Public Property Get Bitmap() As Object
    On Error Resume Next
    Set Bitmap = mpicUp
End Property

Public Property Set Bitmap(ByVal oNewValue As Object)
    On Error Resume Next
    Set mpicUp = oNewValue
End Property

Public Property Set Parent(ByVal picNewValue As Control)
    On Error Resume Next
    Set mParent = picNewValue
End Property

Public Sub Show(iDirection As Integer, Optional MenusAtTop As Long, Optional MenusAtBottom, Optional TotalMenus As Long)
    On Error Resume Next
    If Not mParent Is Nothing Then
        ' this procedure is called during initializing in
        ' Internet Explorer.  so don't run it unless the
        ' parent is visible
        If Not mParent.Visible Then
            Exit Sub
        End If
        mParent.ScaleMode = vbPixels
        With mImgRect
            .Left = mParent.ScaleWidth - PIXELS_FROM_RIGHT - PIXEL_WIDTH
            If mArrowType = ARROW_UP Then
                If MenusAtTop = 0 Then
                    .Left = 0
                    Exit Sub
                End If
                .Top = PIXELS_FROM_TOP + MenusAtTop * mlButtonHeight
            Else
                .Top = mParent.ScaleHeight - PIXELS_FROM_BOTTOM - MenusAtBottom * mlButtonHeight
            End If
            .Right = .Left + PIXEL_WIDTH
            .Bottom = .Top + PIXEL_WIDTH
        
            If .Left <> 0 Then
                mIsDisplayed = True
                DrawBorder RAISED
            End If
        End With
    End If
End Sub

Public Sub Hide()
    On Error Resume Next
    If mIsDisplayed Then
        With mImgRect
            mParent.Line (.Left, .Top)-(.Right, .Bottom), BACKGROUND_COLOR, BF
        End With
        mIsDisplayed = False
    End If
    'miLastPosition = 0
End Sub

Public Property Get IsDisplayed() As Boolean
    On Error Resume Next
    IsDisplayed = mIsDisplayed
End Property

Public Property Let ArrowType(ByVal iNewValue As Integer)
    On Error Resume Next
    mArrowType = iNewValue
End Property

Public Property Let ButtonHeight(ByVal lNewValue As Long)
    On Error Resume Next
    mlButtonHeight = lNewValue
End Property

Public Function HitTest(ByVal iMousePosition As Integer, ByVal x As Long, ByVal y As Long) As Boolean
    On Error Resume Next
    If mIsDisplayed Then
        If PtInRect(mImgRect, x, y) <> 0 Then
            HitTest = True
        
            Select Case iMousePosition
                Case MOUSE_UP
                    Select Case mlState
                        Case SUNKEN
                            DrawBorder RAISED
                        Case Else
                            ' nothing to do
                    End Select
                    mbLastButtonDown = False
                Case MOUSE_DOWN
                    Select Case mlState
                        Case SUNKEN
                            ' nothing to do
                        Case Else
                            DrawBorder SUNKEN
                    End Select
                    mbLastButtonDown = True
                Case MOUSE_MOVE
                    If mbLastButtonDown And mlState = RAISED Then
                        DrawBorder SUNKEN
                    End If
            End Select
        Else
            Select Case iMousePosition
                Case MOUSE_UP
                    mbLastButtonDown = False
                Case MOUSE_MOVE
                    If mlState = SUNKEN Then
                        DrawBorder RAISED
                    End If
                Case MOUSE_DOWN
                    If mlState = SUNKEN Then
                        DrawBorder RAISED
                    End If
                    mbLastButtonDown = False
            End Select
        End If
    End If
End Function

Public Sub DrawBorder(iDirection As Integer)
    On Error Resume Next
    If mIsDisplayed Then
        Select Case iDirection
            Case RAISED
                With mImgRect
                    mParent.PaintPicture mpicUp, .Left, .Top
                End With
                DrawEdge mParent.hdc, mImgRect, BDR_RAISED, BF_RECT
                mlState = RAISED
            Case SUNKEN
                With mImgRect
                    mParent.PaintPicture mpicUp, .Left + 1, .Top + 1, .Right - .Left - 1, .Bottom - .Top - 1
                End With
                DrawEdge mParent.hdc, mImgRect, BDR_SUNKEN, BF_RECT
                mlState = SUNKEN
        End Select
    End If
End Sub

Public Sub Reset()
    On Error Resume Next
    mbLastButtonDown = False
    mIsDisplayed = False
End Sub
