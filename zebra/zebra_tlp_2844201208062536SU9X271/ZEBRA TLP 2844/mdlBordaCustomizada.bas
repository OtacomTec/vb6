Attribute VB_Name = "mdlBordaCustomizada"

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Dim J       As Integer
Private CB  As cBorders

Public Sub AplicaBorda(ByVal Frm As Form, Optional Cor As OLE_COLOR)

    Dim borderType As BorderStyleOptions
    
    Dim Colors(0 To 3) As Long
    
    If Cor = Empty Then
        Colors(0) = &HEA7A37    '&HFF8080 (tipo um roxo) 'esta é a cor padrão passada
    Else
        Colors(0) = Cor
    End If
    
    Colors(1) = bsAutoShade: Colors(2) = bsAutoShade: Colors(3) = bsAutoShade

    borderType = bsFlat1Color
    
    If CB Is Nothing Then Set CB = New cBorders
    
    For Each CTL In Frm.Controls
    
        If TypeOf CTL Is DTPicker Then
            CB.SetBorder CTL.hWnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
    
        'If TypeOf CTL Is MSFlexGrid Then
        '    CB.SetBorder CTL.hwnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
        'End If
    
        If TypeOf CTL Is TextBox Then
            CB.SetBorder CTL.hWnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is MaskEdBox Then
            CB.SetBorder CTL.hWnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is ListView Then
            CB.SetBorder CTL.hWnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
        End If
    
        If TypeOf CTL Is ListBox Then
            CB.SetBorder CTL.hWnd, borderType, ctListBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is DriveListBox Then
            CB.SetBorder CTL.hWnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is ComboBox Then
            CB.SetBorder CTL.hWnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
    
        If TypeOf CTL Is ImageCombo Then
            CB.SetBorder CTL.hWnd, borderType, ctImageCombo, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is DataCombo Then
            CB.SetBorder CTL.hWnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is DBCombo Then
            CB.SetBorder CTL.hWnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is DTPicker Then
            CB.SetBorder CTL.hWnd, borderType, ctImageCombo, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
    
        If TypeOf CTL Is FileListBox Then
            CB.SetBorder CTL.hWnd, borderType, ctListBox, Colors(0), Colors(1), Colors(2), Colors(3)
        End If
    
        'If TypeOf CTL Is Frame Then
        '    CB.SetBorder CTL.hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
        'End If
        
        If TypeOf CTL Is TreeView Then
            CB.SetBorder CTL.hWnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is ProgressBar Then
            CB.SetBorder CTL.hWnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
        End If
        
        If TypeOf CTL Is PictureBox Then
            CB.SetBorder CTL.hWnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
        End If
    Next
    
    
End Sub
