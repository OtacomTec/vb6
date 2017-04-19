Attribute VB_Name = "modZebra"
Option Explicit

Dim h As Long
Dim NumOfCopies As String
Dim Arr() As String
Dim LastIndex As Long
Dim Port As String

Public Enum eTipoEtiqueta
'    cPEQUENA
    cGRANDE
    cMEDIANA
End Enum

Public Enum eOrientation
    c0
    c90
    c180
    c270
End Enum
Dim mOrientation  As String


Dim SendToPrinter As Boolean
Const DEBUGGING = True

Sub Main()
    If ImprimirEtiquetas(cGRANDE, "CAB-0110", "La descripcion", "BLANCO", "01/01/2006-1", 2, "57 MTS", "COM1:", c90) Then
        MsgBox "Impresión exitosa"
    Else
        MsgBox "Se encontró un error al intentar imprimir"
    End If
End Sub
Private Sub SetOrientation(o As eOrientation)
Dim strO As String

    mOrientation = "N"
    Select Case o
    Case c180
        mOrientation = "I"
    Case c270
        mOrientation = "B"
    Case c90
        mOrientation = "R"
    End Select
End Sub
Private Sub PrintOrientation()
    If mOrientation <> "N" Then
        Call Add("^FW" & mOrientation)
    End If
End Sub
Private Sub BeginLabel(Optional blnSendToPrinter As Boolean = True)
    If DEBUGGING Then
        blnSendToPrinter = False
    End If
    SendToPrinter = blnSendToPrinter
    Call Add("^XA", True)
End Sub

Private Sub Add(Valor As String, Optional ClearArray As Boolean = False)
    If Valor <> "" Then
        If SendToPrinter Then
            Print #h, Valor
            LastIndex = -1
        Else
            If ClearArray Then
                ReDim Arr(0) As String
                LastIndex = 0
            Else
                On Local Error Resume Next
                
                LastIndex = UBound(Arr)
                If Err.Number <> 0 Then
                    LastIndex = -1
                End If
                
                LastIndex = LastIndex + 1
                On Local Error GoTo 0
            End If
            
            ReDim Preserve Arr(LastIndex) As String
            Arr(LastIndex) = Valor
        End If
    End If
End Sub

Public Sub CalibrarImpresora(strPuerto As String)
    Call SetPort(strPuerto)
    Call OpenPrinter
            Call BeginLabel
                                
            Call confCalibrate
            
            Call EndLabel
    Call ClosePrinter
End Sub

Private Sub ClosePrinter()
Dim i As Long
    If LastIndex >= 0 Then
        For i = 1 To 20
            Debug.Print
        Next
        For i = LBound(Arr) To UBound(Arr)
            'Print #h, Arr(i)
            Debug.Print Arr(i)
        Next
    End If
    Close #h
End Sub

Private Sub EndLabel()
    Call Add(NumOfCopies)
    Call Add("^XZ")
End Sub

Private Function FileExist(FileName As String) As Boolean
Dim X As Long
    On Local Error Resume Next
    
    X = FreeFile
    Open FileName For Binary As X
    If Err.Number = 0 Then
        FileExist = True
        Close #X
    Else
        FileExist = False
    End If
    
    On Local Error GoTo 0
End Function

Public Function ImprimirEtiquetas(Tipo As eTipoEtiqueta, Codigo As String, Descripcion As String, Color As String, Lote As String, Optional NumCopias As Integer = 1, Optional Cantidad As String = "100 METROS", Optional strPuerto As String = "COM1:", Optional lOrientation As eOrientation = c0) As Boolean

    On Local Error GoTo ErrDrv
    
    Call SetPort(strPuerto)
    Call SetOrientation(lOrientation)
    
    Select Case Tipo
    Case cGRANDE
        Call ImprimirEtiquetaGde(Codigo, Descripcion, Color, Lote, NumCopias, Cantidad)
'    Case cPEQUENA
'        Call ImprimirEtiquetaPeq(Codigo, Descripcion, Color, Lote, NumCopias, Cantidad)
    Case cMEDIANA
        Call ImprimirEtiquetaMed(Codigo, Descripcion, Color, Lote, NumCopias, Cantidad)
    End Select
    
    ImprimirEtiquetas = True
    Exit Function
    
ErrDrv:
    ImprimirEtiquetas = False
End Function

Private Sub LoadImage(Path As String, FileName As String)
Dim X As Long
Dim CurByte As Byte, i As Long, J As Long
Dim TotalBytes As Long ', BytesPerRow As Long
Dim Line As String, BytesPerRow As Byte

    If FileExist(Path & "\" & FileName & ".PCX") Then
        X = FreeFile
        Open Path & "\" & FileName & ".PCX" For Binary As X
        
        TotalBytes = LOF(X) - 128       'Cabezera de los PCX: 128 bytes
        Get #X, 67, BytesPerRow         'Byte 67: # Bytes por Linea gráfica
        Call Add("~DGR:" & FileName & ".GRF," & TotalBytes & "," & BytesPerRow & ",")
        Get #X, 128, CurByte
        
        Do
        
            Line = ""
            For J = 1 To BytesPerRow
                CurByte = 0
                Get #X, , CurByte
                If Not EOF(X) Then
                    If Len(Hex(CurByte)) = 1 Then
                        Line = Line & "0" & Hex(CurByte)
                    Else
                        Line = Line & Hex(CurByte)
                    End If
                Else
                    Exit For
                End If
            Next
            Call Add(Line)
        
        Loop While Not EOF(X)
        
        Close #X
        
    End If
End Sub

Private Sub PrintData(Data As String)
    Call PrintOrientation
    Call Add("^FD" & Data & "^FS")
End Sub

Private Sub PrintImage(FileName As String, Optional XFactor As Integer = 1, Optional YFactor As Integer = 1)
Dim Extra As String
    Call PrintOrientation
    Call Add("^XGR:" & FileName & "," & XFactor & "," & YFactor & "^FS")
End Sub

Private Sub SetBlock(Optional Width As String = "0", Optional MaxLines As String = "1", Optional AddOrDeleteSpace As String = "0", Optional Justify As String = "C", Optional InnerMargin As String = "0")
    Call PrintOrientation
    Call Add("^FB" & Width & "," & MaxLines & "," & AddOrDeleteSpace & "," & Justify & "," & InnerMargin)

    '^FBa,b,c,d,e: Bloque de texto. Precede al ^FD
    '               a=Ancho del texto (0..9999)
    '               b=Número máximo de líneas (1..9999)
    '               c=Agregar o eliminar espacio entre líneas (-9999..9999)
    '               d=justificacion (L,C,R,J)
    '               e=sangría de la segunda línea y sucesivas.  (0..9999)
End Sub

Private Sub OpenPrinter()
    LastIndex = -1

    Call ClosePrinter
    h = FreeFile
    Open Port For Output As h
End Sub

Private Sub PrintBarCode(BarCode As String, Optional Orientation As eOrientation, Optional Height As String = "100", Optional PrintCode As String = "N", Optional CodeOnTop As String = "N", Optional PrintCheckDigit As String = "N")
Dim strO As String
    strO = "N"
    Select Case Orientation
    Case c180
        strO = "R"
    Case c270
        strO = "I"
    Case c90
        strO = "B"
    End Select
    
    Call Add("^BA" & strO & "," & Height & "," & PrintCode & "," & CodeOnTop & "," & PrintCheckDigit & "^FD" & BarCode & "^FS")
    
    'o = Orientación(N, R, i, B)
    'h=Altura (1..32000)
    'f=Imprimir Línea de Interpretación (Y,N)
    'g=Imprimir Línea de Interpretación sobre el código (Y,N)
    'e=Imprimir dígito de chequeo
End Sub

Private Sub PrintBox(Optional Width As Single = 0, Optional Height As Single = 0, Optional BorderThickness As Single = 2)
    Select Case mOrientation
    Case "N", "B"
        Call Add("^GB" & Width & "," & Height & "," & BorderThickness & "^FS")
    Case "R", "I"
        Call Add("^GB" & Height & "," & Width & "," & BorderThickness & "^FS")
    End Select
End Sub


Private Sub SetFont(Optional FontName As String = "A", Optional Orientation As String = "N", Optional Height As Single = 14, Optional Width As Single = 12)
    Call Add("^A" & FontName & Orientation & "," & Height & "," & Width)
End Sub

Private Sub ImprimirEtiquetaMed(Codigo As String, Descripcion As String, Color As String, Lote As String, Optional NumCopias As Integer = 1, Optional Cantidad As String = "100 METROS")
Static BytesPerRow As Long
Dim TipoLetraDescripcion As String
Dim AltoDescripcion As Single
Dim AnchoDescripcion As Single
Dim XDescripcion As Single
Dim YDescripcion As Single
Dim TipoLetraColor As String
Dim AltoColor As Single
Dim AnchoColor As Single
Dim XColor As Single
Dim YColor As Single
Dim TipoLetraCantidad As String
Dim AltoCantidad As Single
Dim AnchoCantidad As Single
Dim XCantidad As Single
Dim YCantidad As Single
Dim TipoLetraCRESMAR As String
Dim AltoCRESMAR As Single
Dim AnchoCRESMAR As Single
Dim XCRESMAR As Single
Dim YCRESMAR As Single
Dim TipoLetraLOTE As String
Dim AltoLOTE As Single
Dim AnchoLOTE As Single
Dim XLOTE As Single
Dim YLOTE As Single

    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 20
    XLOTE = 580
    
    TipoLetraCRESMAR = "G"
    AltoCRESMAR = 70
    AnchoCRESMAR = 18
    YCRESMAR = 20
    XCRESMAR = 30
    
    TipoLetraColor = "O"
    AltoColor = 70
    AnchoColor = 65
    YColor = 200
    XColor = 30
    
    TipoLetraCantidad = "D"
    AltoCantidad = 70
    AnchoCantidad = 30
    YCantidad = 200
    XCantidad = 150
    
    TipoLetraDescripcion = "O"
    AltoDescripcion = 80
    AnchoDescripcion = 70
    YDescripcion = 120
    XDescripcion = 30
    
    Call OpenPrinter
            Call BeginLabel
                Call SetNumOfCopies(NumCopias)
                
                Call SetPos(40, 40)
                    
                    Call SetPos(YCRESMAR, XCRESMAR)
                    Call SetFont(TipoLetraCRESMAR, , AltoCRESMAR, AnchoCRESMAR)
                    Call SetBlock(600, , , "L")
                    Call PrintData("CRESMAR")
                    
                    Call SetPos(YLOTE, XLOTE)
                    Call SetFont(TipoLetraLOTE)
                    Call SetBlock(200, 1)
                        Call PrintData("LOTE/FECHA")
                   Call SetPos(YLOTE + 40, XLOTE - 50)
                    Call SetFont(TipoLetraLOTE)
                    Call SetBlock(300, 1)
                        Call PrintData(Lote)
                        
                        
                    Call SetPos(YDescripcion, XDescripcion)     ' 190,30 antes
                    Call SetFont(TipoLetraDescripcion, , AltoDescripcion, AnchoDescripcion)     '"F", , 34, 34)
                    Call SetBlock(770, 2, , "L")
                        Call PrintData(Descripcion)
                        
                    Call SetPos(YColor, XColor)  '190 + 90, 30)
                    Call SetFont(TipoLetraColor, , AltoColor, AnchoColor)   '"F", ,34,32 antes
                    Call SetBlock(770, 1, , "L")
                        Call PrintData(Color)
                        
                    Call SetPos(YCantidad, XCantidad) '280,250 antes
                    Call SetFont(TipoLetraCantidad, , AltoCantidad, AnchoCantidad)    ',"10" antes
                    Call SetBlock(770, 1)
                        Call PrintData(Cantidad)
                                        

                Call SetPos(300, 20)
                Call PrintBarCode(Codigo, , 80, "N", "N")
                Call SetPos(300 - 20, 20 + 40)  'Relativo al Bar Code
                    Call SetFont("B")
                    Call PrintData(Codigo)
                    
                    Call SetFont
                    Call SetPos(390, 20)
                    Call SetBlock(770, , , "R")
                    Call PrintData("HECHO EN VENEZUELA POR blabla")
                    
                    Call SetFont
                    Call SetPos(390, 20)
                    Call SetBlock(770, , , "L")
                    If Cantidad = "100 MTS" Then
                        Call PrintData("UPO-FR005")
                    Else
                        Call PrintData("UPO-FR006")
                    End If


'                'Imprimiendo Código de Barra y Logo
'                Call SetPos(660, 30)
'                    Call SetFont("F", , 120, 20)
'                    Call SetBlock(308)
'                    Call PrintData("CRESMAR")
'                    'Call LoadImage(App.Path, "CRESMAR")
'                    'Call PrintImage("CRESMAR", 2, 2)
            
            Call EndLabel
    Call ClosePrinter
End Sub

Private Sub ImprimirEtiquetaGde(Codigo As String, Descripcion As String, Color As String, Lote As String, Optional NumCopias As Integer = 1, Optional Cantidad As String = "100 METROS")
Static BytesPerRow As Long
Dim TipoLetraDescripcion As String
Dim AltoDescripcion As Single
Dim AnchoDescripcion As Single
Dim XDescripcion As Single
Dim YDescripcion As Single

Dim TipoLetraColor As String
Dim AltoColor As Single
Dim AnchoColor As Single
Dim XColor As Single
Dim YColor As Single

Dim TipoLetraLOTE As String
Dim AltoLOTE As Single
Dim AnchoLOTE As Single
Dim XLOTE As Single
Dim YLOTE As Single

Dim TipoLetraCRESMAR As String
Dim AltoCRESMAR As Single
Dim AnchoCRESMAR As Single
Dim XCRESMAR As Single
Dim YCRESMAR As Single
    
    TipoLetraDescripcion = "O"
    AltoDescripcion = 130
    AnchoDescripcion = 90
    YDescripcion = 290
    XDescripcion = 80
    
    TipoLetraCRESMAR = "G"
    AltoCRESMAR = 120
    AnchoCRESMAR = 30
    YCRESMAR = 590
    XCRESMAR = 80
    
    TipoLetraColor = "O"
    AltoColor = 120
    AnchoColor = 85
    YColor = 280
    XColor = 80
    
    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 700
    XLOTE = 725
    
    Call OpenPrinter
            Call BeginLabel
                Call SetNumOfCopies(NumCopias)
                
                Call SetPos(40, 40)
                    Call PrintBox(975, 720, 4)     'Marco Principal
                    
                    Call SetPos(YCRESMAR, XCRESMAR)
                    Call SetFont(TipoLetraCRESMAR, , AltoCRESMAR, AnchoCRESMAR)
                    Call SetBlock(600, , , "L")
                    Call PrintData("CRESMAR")
                    
                    Call SetPos(YLOTE, XLOTE)
                    Call SetFont(TipoLetraLOTE, , AltoLOTE, AnchoLOTE)
                    Call SetBlock(300, 1)
                        Call PrintData("LOTE/FECHA")
                        
                    Call SetPos(YLOTE - 50, XLOTE)
                    Call SetFont(TipoLetraLOTE, , AltoLOTE, AnchoLOTE)
                    Call SetBlock(300, 1)
                        Call PrintData(Lote)
                                        
                    Call SetPos(YDescripcion, XDescripcion)
                    Call SetFont(TipoLetraDescripcion, , AltoDescripcion, AnchoDescripcion)
                    Call SetBlock(900, 2, , "L")
                        Call PrintData(Descripcion)
                        
                    Call SetPos(YColor, XColor)
                    Call SetFont(TipoLetraColor, , AltoColor, AnchoColor)
                    Call SetBlock(900, 1, , "L")
                        Call PrintData(Color & "  " & Cantidad)

                Call SetPos(120, 90)
                Call PrintBarCode(Codigo, c90, 80, "N", "N")
                Call SetPos(120 + 90, 90 + 40)  'Relativo al Bar Code
                    Call SetFont("B")
                    Call PrintData(Codigo)
                    
                    Call SetFont
                    Call SetPos(50, 20)
                    Call SetBlock(975, , , "R")
                    Call PrintData("HECHO EN VENEZUELA POR URAPLAST, C.A.")
                    
                    Call SetFont
                    Call SetPos(50, 50)
                    Call SetBlock(975, , , "L")
                    Call PrintData("UPO-FR004")


'                'Imprimiendo Código de Barra y Logo
'                Call SetPos(660, 30)
'                    Call SetFont("F", , 120, 20)
'                    Call SetBlock(308)
'                    Call PrintData("CRESMAR")
'                    'Call LoadImage(App.Path, "CRESMAR")
'                    'Call PrintImage("CRESMAR", 2, 2)
            
            Call EndLabel
    Call ClosePrinter
    
End Sub


Private Sub SetNumOfCopies(Optional Number As Integer = 1)
    If Number > 1 Then
        NumOfCopies = "^PQ" & Str(Number)
    Else
        NumOfCopies = ""
    End If
End Sub

Private Sub SetPort(Optional strPort As String = "COM1:")
    Port = strPort
End Sub


Private Sub SetPos(Y As Single, X As Single)
    Select Case mOrientation
    Case "N", "B"
        Call Add("^FO" & X & "," & Y)
    Case "R", "I"
        Call Add("^FO" & Y & "," & X)
    End Select
    
End Sub

Private Sub confCalibrate()
    Call Add("~JC")
End Sub

Private Sub confCalibrateShowGraphic()
    Call Add("~JG")
End Sub

Private Sub confSetLabelLength()
    Call Add("~JL")
End Sub

