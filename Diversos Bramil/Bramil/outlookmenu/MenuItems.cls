VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colMenuItems As New Collection
Const MOUSE_UP = 1
Const MOUSE_DOWN = -1
Const MOUSE_MOVE = 0

' add a new MenuItem to the collection
' Parameters:   sCaption        Caption of the MenuItem
'               lIndex          Location of the MenuItem in MenuItems collection
'               picIcon         Icon
Public Function Add(ByVal sCaption As String, lIndex As Long, lButtonHeight As Long, picIcon As Object) As MenuItem
    Dim newMenuItem As New MenuItem
    
    On Error Resume Next
    With newMenuItem
        .Caption = sCaption
        .Index = lIndex
        .ButtonHeight = lButtonHeight
        Set .Button = picIcon
            
        ' add the item to the collection specified by lIndex
        ' note, if there is nothing in the collection, just add it
        ' if there is nothing in the collection or we are adding it at then end, just add it
        ' elseif we are inserting in the first position, add it BEFORE
        ' else add it AFTER the previous item
        If colMenuItems.Count = 0 Or lIndex = colMenuItems.Count + 1 Then
            colMenuItems.Add newMenuItem
        ElseIf lIndex = 1 Then
            colMenuItems.Add newMenuItem, , 1
        Else
            colMenuItems.Add newMenuItem, , , lIndex - 1
        End If
    End With
    
    Set Add = newMenuItem
End Function

' delete the MenuItem from the collection
' Parameters:       lIndex  Index of the collection member
Public Sub Delete(lIndex As Long)
    On Error Resume Next
    colMenuItems.Remove lIndex
End Sub

' return the object of the MenuItem in the collection
' Parameters:       lIndex  Index of the collection member
Public Function Item(lIndex As Long) As MenuItem
    On Error Resume Next
    Set Item = colMenuItems.Item(lIndex)
End Function

' return the number of MenuItems in the collection
Public Function Count() As Long
    On Error Resume Next
    Count = colMenuItems.Count
End Function

' paint all MenuItems (icon & caption) in this collection
' Parameters:       bRecalc
'                       True    Forces a recalc of the icon's position
'                       False   Uses the current icon position
Public Function Paint(lTopMenuItemDisplayed As Long, lIconStart As Long, lMenuCur As Long, lClipY As Long) As Boolean
    Dim MenuItem As MenuItem
    
    For Each MenuItem In colMenuItems
        With MenuItem
            'If .Index >= lTopMenuItemDisplayed Then
            Paint = .PaintButton(lTopMenuItemDisplayed, lIconStart, lMenuCur, lClipY)
            'End If
        End With
    Next
End Function

' process mouse events for all MenuItems in the collection
Public Function MouseProcess(ByVal iMousePosition, ByVal x As Long, ByVal y As Long) As Long
    Dim MenuItem As MenuItem
    Dim bResult As Boolean
    Dim lIndex As Long
    Static lLastDown As Long
    
    On Error Resume Next
    For Each MenuItem In colMenuItems
        With MenuItem
            bResult = .HitTest(iMousePosition, x, y)
            
            ' the mouse can only be over one object at a time (they don't overlap)
            ' if we get a hit, set MouseProcess to return to the calling routine
            ' we need to remember where the mouse went down because if the user
            ' moves the mouse and raises on another item, we don't want to fire the event
            lIndex = .Index
            If bResult Then
                Select Case iMousePosition
                    Case MOUSE_UP
                        If lLastDown = lIndex Then
                            MouseProcess = lIndex
                        End If
                    Case Else
                        MouseProcess = lIndex
                End Select
                If iMousePosition = MOUSE_DOWN Then
                    lLastDown = lIndex
                End If
            End If
        End With
    Next
End Function

