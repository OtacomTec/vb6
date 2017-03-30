VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclConversao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Codigo Programa: GMclConversao
'Descr da Classe: Converte Instrução SQL UPDATE (simples) em INSERT
'Analista.......: Geraldo Coimbra
'Programador....: Luis Henrique Borges
'Data Criação...: 27 Set 2001
'Data Alteração.:
'-------------------------------------------------------------------------------------------
Public Function CInsert(strSQL As String) As String

    Dim lstrCampo()   As String
    Dim lstrValor()   As String
    Dim lstrInstrução As String
    Dim lstrTabela    As String
    Dim lstrSpace     As String
    Dim p1            As Integer
    Dim p2            As Integer
    Dim p3            As Integer
    
    ReDim lstrCampo(0)  'Armazenará os nomes dos campos
    ReDim lstrValor(0)  'Armazenará os valores dos campos de lstrCampo()
    
    lstrInstrução = Trim(Mid(strSQL, 1, InStr(strSQL, " ")))
    lstrSpace = " "
    
    If lstrInstrução = "UPDATE" Then
        p1 = InStr(strSQL, " ")
        p2 = InStr(p1 + 1, strSQL, " ")
        lstrTabela = Trim(Mid(strSQL, p1 + 1, p2 - 1 - p1 + 1))
        
        p1 = p2
        p2 = InStr(p1, UCase(strSQL), "SET ") + 4
        p1 = p2
        lstrFinal = Trim(Mid(strSQL, p1, Len(strSQL) - p1 + 1))
               
        If InStr(UCase(lstrFinal), " WHERE ") <> 0 Then
            lstrCampos = lstrFinal
            p2 = InStr(lstrCampos, "=")
            ReDim Preserve lstrCampo(UBound(lstrCampo) + 1)
            ReDim Preserve lstrValor(UBound(lstrValor) + 1)
            
            lstrCampo(UBound(lstrCampo)) = Trim(Mid(lstrCampos, 1, p2 - 1))
            
            'Verifica se o próximo valor é uma string
            lstrResto = Trim(Mid(lstrCampos, p2 + 1, Len(lstrCampos) - (p2 - 1)))
            If Mid(lstrResto, 1, 1) = "'" Then
                p1 = InStr(lstrResto, "'")
                p2 = InStr(p1 + 1, lstrResto, "'")
                                
                lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, p1, p2 - p1 + 1))
                
            Else
                p2 = InStr(lstrResto, ",") - 1
                lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, 1, p2))
                
            End If
            p3 = InStr(lstrResto, ",") + 1
            'Extrair apenas o resto da string que não foi tratada
            'isto é, descartar a string do campo e valor
            
            p1 = InStr(UCase(lstrResto), " WHERE ")
            lstrWHERE = Trim(Mid(lstrResto, p1 + 7, Len(lstrResto) - p1))
            lstrResto = Mid(lstrResto, p3, p1 - p3) 'Len(lstrResto) - p2 + 2)
            
            
            If InStr(lstrResto, "=") <> 0 Then
                Achou = True
                Do While Achou = True
                    'Verifica o próximo sinal de "="
                    p2 = InStr(lstrResto, "=")
                    If p2 = 0 Then Exit Do
                    
                    ReDim Preserve lstrCampo(UBound(lstrCampo) + 1)
                    ReDim Preserve lstrValor(UBound(lstrValor) + 1)
                    'If UBound(lstrValor) = 30 Then
                    '    MsgBox " "
                    'End If
                    
                    lstrCampo(UBound(lstrCampo)) = Trim(Mid(lstrResto, 1, p2 - 1))
                    
                    'Verifica se o próximo valor é uma string
                    lstrResto = Trim(Mid(lstrResto, p2 + 1, Len(lstrResto) - p2 + 1))
                    If Mid(lstrResto, 1, 1) = "'" Then
                        p1 = InStr(lstrResto, "'")
                        p2 = InStr(p1 + 1, lstrResto, "'")
                
                        lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, p1, p2 - p1 + 1))
                        lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                    Else
                        p2 = InStr(lstrResto, ",") - 1
                        If p2 = -1 Then
                            lstrValor(UBound(lstrValor)) = Trim(lstrResto)
                            Achou = False
                        Else
                            
                            lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, 1, p2))
                            lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                        End If
                    End If
                Loop
                
                
                'Vefica despois do Where
                Achou = True
                lstrResto = lstrWHERE
                Do While Achou = True
                    'Verifica o próximo sinal de "="
                    p2 = InStr(lstrResto, "=")
                    If p2 = 0 Then Exit Do
                    CampoExistente = False
                    For i = 1 To UBound(lstrCampo)
                        If Trim(Mid(lstrResto, 1, p2 - 1)) = lstrCampo(i) Then CampoExistente = True: Exit For
                    Next i
                    If CampoExistente = False Then
                        ReDim Preserve lstrCampo(UBound(lstrCampo) + 1)
                        ReDim Preserve lstrValor(UBound(lstrValor) + 1)
                        lstrCampo(UBound(lstrCampo)) = Trim(Mid(lstrResto, 1, p2 - 1))
                    End If
                    'Verifica se o próximo valor é uma string
                    lstrResto = Trim(Mid(lstrResto, p2 + 1, Len(lstrResto) - p2 + 1))
                    If Mid(lstrResto, 1, 1) = "'" Then
                        p1 = InStr(lstrResto, "'")
                        p2 = InStr(p1 + 1, lstrResto, "'")
                        
                        If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, p1, p2 - p1 + 1))
                        
                        lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                    Else
                        p2 = InStr(UCase(lstrResto), " AND ") - 1
                        If p2 = -1 Then
                            If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(lstrResto)
                            Achou = False
                        Else
                            If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, 1, p2))
                            lstrResto = Trim(Mid(lstrResto, p2 + 6, Len(lstrResto) - p2 + 6))
                        End If
                    End If
                    CampoExistente = False
                Loop
            Else
                'Vefica despois do Where
                Achou = True
                lstrResto = lstrWHERE
                Do While Achou = True
                    'Verifica o próximo sinal de "="
                    p2 = InStr(lstrResto, "=")
                    If p2 = 0 Then Exit Do
                    
                    For i = 1 To UBound(lstrCampo)
                        If Trim(Mid(lstrResto, 1, p2 - 1)) = lstrCampo(i) Then CampoExistente = True
                    Next i
                    If CampoExistente = False Then
                        ReDim Preserve lstrCampo(UBound(lstrCampo) + 1)
                        ReDim Preserve lstrValor(UBound(lstrValor) + 1)
                        lstrCampo(UBound(lstrCampo)) = Trim(Mid(lstrResto, 1, p2 - 1))
                    End If
                    'Verifica se o próximo valor é uma string
                    lstrResto = Trim(Mid(lstrResto, p2 + 1, Len(lstrResto) - p2 + 1))
                    If Mid(lstrResto, 1, 1) = "'" Then
                        p1 = InStr(lstrResto, "'")
                        p2 = InStr(p1 + 1, lstrResto, "'")
                        
                        If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, p1, p2 - p1 + 1))
                        
                        lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                    Else
                        p2 = InStr(UCase(lstrResto), " AND ") - 1
                        If p2 = -1 Then
                            If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(lstrResto)
                            Achou = False
                        Else
                            If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, 1, p2))
                            lstrResto = Trim(Mid(lstrResto, p2 + 6, Len(lstrResto) - p2 + 6))
                        End If
                    End If
                    CampoExistente = False
                Loop
            End If
        End If
    Else
        CInsert = Empty
    End If
    
    If UBound(lstrCampo) > 0 Then
        lstrCInsert = "INSERT INTO" & lstrSpace & lstrTabela & lstrSpace & "("
        For i = 1 To UBound(lstrCampo)
            If i = UBound(lstrCampo) Then
                lstrCInsert = lstrCInsert & lstrCampo(i) & ") "
                Debug.Print lstrCInsert
            Else
                lstrCInsert = lstrCInsert & lstrCampo(i) & ", "
                Debug.Print lstrCInsert
            End If
        Next i
    
        lstrCInsert = lstrCInsert & "VALUES" & lstrSpace & "("
        Debug.Print lstrCInsert
        For i = 1 To UBound(lstrValor)
            If i = UBound(lstrCampo) Then
                lstrCInsert = lstrCInsert & lstrValor(i) & ")"
                Debug.Print lstrCInsert
            Else
                lstrCInsert = lstrCInsert & lstrValor(i) & ", "
            End If
            Debug.Print lstrCInsert
        Next i
    
        lstrCInsert = Replace(Replace(lstrCInsert, Chr(13), ""), Chr(10), "")
        CInsert = Trim(lstrCInsert)
    End If
End Function

