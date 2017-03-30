Attribute VB_Name = "Module1"
Function CInsert(strSQL As String) As String
'--------------------------------------------------------------------------------------------'
'Nome da Fun��o : CInsert
'Descr.Programa.: Converte instru��o SQL UPDATE simples em Instru��o INSERT
'Analista.......: Geraldo Coimbra                                                            '
'Programador....: Luis Henrique Borges                                                       '
'Data Cria��o...: 27/09/2001                                                                 '
'Data Altera��o.:                                                                            '
'--------------------------------------------------------------------------------------------'
    Dim lstrCampo()   As String
    Dim lstrValor()   As String
    Dim lstrInstru��o As String
    Dim lstrTabela    As String
    Dim lstrSpace     As String
    Dim p1            As Integer
    Dim p2            As Integer
    Dim p3            As Integer
    
    ReDim lstrCampo(0)  'Armazenar� os nomes dos campos
    ReDim lstrValor(0)  'Armazenar� os valores dos campos de lstrCampo()
    
    lstrInstru��o = Trim(Mid(strSQL, 1, InStr(strSQL, " ")))
    lstrSpace = " "
    
    If lstrInstru��o = "UPDATE" Then
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
            
            lstrCampo(UBound(lstrCampo)) = Mid(lstrCampos, 1, p2 - 1)
            
            'Verifica se o pr�ximo valor � uma string
            lstrResto = Trim(Mid(lstrCampos, p2 + 1, Len(lstrCampos) - (p2 - 1)))
            If Mid(lstrResto, 1, 1) = "'" Then
                p1 = InStr(lstrResto, "'")
                p2 = InStr(p1 + 1, lstrResto, "'")
                lstrValor(UBound(lstrValor)) = Mid(lstrResto, p1, p2 - p1 + 1)
            Else
                p2 = InStr(lstrResto, ",") - 1
                lstrValor(UBound(lstrValor)) = Trim(Mid(lstrResto, 1, p2))
                
            End If
            p3 = InStr(lstrResto, ",") + 1
            'Extrair apenas o resto da string que n�o foi tratada
            'isto �, descartar a string do campo e valor
            
            p1 = InStr(UCase(lstrResto), " WHERE ")
            lstrWHERE = Trim(Mid(lstrResto, p1 + 7, Len(lstrResto) - p1))
            lstrResto = Mid(lstrResto, p3, p1 - p3) 'Len(lstrResto) - p2 + 2)
            
            
            If InStr(lstrResto, "=") <> 0 Then
                Achou = True
                Do While Achou = True
                    'Verifica o pr�ximo sinal de "="
                    p2 = InStr(lstrResto, "=")
                    If p2 = 0 Then Exit Do
                    
                    ReDim Preserve lstrCampo(UBound(lstrCampo) + 1)
                    ReDim Preserve lstrValor(UBound(lstrValor) + 1)
                    lstrCampo(UBound(lstrCampo)) = Trim(Mid(lstrResto, 1, p2 - 1))
                    
                    'Verifica se o pr�ximo valor � uma string
                    lstrResto = Trim(Mid(lstrResto, p2 + 1, Len(lstrResto) - p2 + 1))
                    If Mid(lstrResto, 1, 1) = "'" Then
                        p1 = InStr(lstrResto, "'")
                        p2 = InStr(p1 + 1, lstrResto, "'")
                        lstrValor(UBound(lstrValor)) = Mid(lstrResto, p1, p2 - p1 + 1)
                        lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                    Else
                        p2 = InStr(lstrResto, ",") - 1
                        If p2 = -1 Then
                            lstrValor(UBound(lstrValor)) = Trim(lstrResto)
                            Achou = False
                        Else
                            lstrValor(UBound(lstrValor)) = Mid(lstrResto, 1, p2)
                            lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                        End If
                    End If
                Loop
                
                
                'Vefica despois do Where
                Achou = True
                lstrResto = lstrWHERE
                Do While Achou = True
                    'Verifica o pr�ximo sinal de "="
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
                    'Verifica se o pr�ximo valor � uma string
                    lstrResto = Trim(Mid(lstrResto, p2 + 1, Len(lstrResto) - p2 + 1))
                    If Mid(lstrResto, 1, 1) = "'" Then
                        p1 = InStr(lstrResto, "'")
                        p2 = InStr(p1 + 1, lstrResto, "'")
                        
                        If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Mid(lstrResto, p1, p2 - p1 + 1)
                        lstrResto = Trim(Mid(lstrResto, p2 + 2, Len(lstrResto) - p2 + 2))
                    Else
                        p2 = InStr(UCase(lstrResto), " AND ") - 1
                        If p2 = -1 Then
                            If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Trim(lstrResto)
                            Achou = False
                        Else
                            If CampoExistente = False Then lstrValor(UBound(lstrValor)) = Mid(lstrResto, 1, p2)
                            lstrResto = Trim(Mid(lstrResto, p2 + 6, Len(lstrResto) - p2 + 6))
                        End If
                    End If
                    CampoExistente = False
                Loop
            End If
        End If
    End If
    
    lstrCInsert = "INSERT INTO" & lstrSpace & lstrTabela & lstrSpace & "("
    For i = 1 To UBound(lstrCampo)
        If i = UBound(lstrCampo) Then
            lstrCInsert = lstrCInsert & lstrCampo(i) & ") "
        Else
            lstrCInsert = lstrCInsert & lstrCampo(i) & ", "
        End If
    Next i
    
    lstrCInsert = lstrCInsert & "VALUES" & lstrSpace & "("
    
    For i = 1 To UBound(lstrValor)
        If i = UBound(lstrCampo) Then
            lstrCInsert = lstrCInsert & lstrValor(i) & ")"
        Else
            lstrCInsert = lstrCInsert & lstrValor(i) & ", "
        End If
    Next i
    
    lstrCInsert = Replace(Replace(lstrCInsert, Chr(13), ""), Chr(10), "")
    CInsert = Trim(lstrCInsert)
    
        
End Function
