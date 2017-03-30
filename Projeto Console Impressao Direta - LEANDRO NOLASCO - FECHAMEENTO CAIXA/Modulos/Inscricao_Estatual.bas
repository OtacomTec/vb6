Attribute VB_Name = "Inscricao_Estadual"
Option Explicit

Public Function ChecaInscrE(pUF As String, pInscr As String, TextBox As TextBox) As Boolean

        Dim strBase As String
        Dim strBase2 As String
        Dim strOrigem As String
        Dim strDigito1 As String
        Dim strDigito2 As String
        Dim intPos As Integer
        Dim intValor As Integer
        Dim intSoma As Integer
        Dim intResto As Integer
        Dim intNumero As Integer
        Dim intPeso As Integer
        Dim intDig As Integer

        ChecaInscrE = False
        
        strBase = Empty
        strBase2 = Empty
        strOrigem = Empty

        If Trim(pInscr) = "ISENTO" Then
           ChecaInscrE = True
           Exit Function
        End If
        
        For intPos = 1 To Len(Trim(pInscr))
            If InStr(1, "0123456789P", Mid$(pInscr, intPos, 1), vbTextCompare) > 0 Then
               strOrigem = strOrigem & Mid$(pInscr, intPos, 1)
            End If
        Next

        Select Case pUF
               Case "AC"         '  Acre
                    strBase = Left(Trim(strOrigem) & "000000000", 9)
                    If Left(strBase, 2) = "01" And Mid$(strBase, 3, 2) <> "00" Then
                       intSoma = 0
                       For intPos = 1 To 8
                           intValor = Val(Mid$(strBase, intPos, 1))
                           intValor = intValor * (10 - intPos)
                           intSoma = intSoma + intValor
                       Next
                       intResto = intSoma Mod 11
                       strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                       strBase2 = Left(strBase, 8) & strDigito1
                       If strBase2 = strOrigem Then
                          ChecaInscrE = True
                          pInscr = Format(pInscr, "00,00,0000-0")
                          TextBox.Text = pInscr
                       End If
                    End If
               Case "AL"         '  Alagoas
                    strBase = Left(Trim(strOrigem) & "000000000", 9)
                    If Left(strBase, 2) = "24" Then
                       intSoma = 0
                       For intPos = 1 To 8
                           intValor = Val(Mid$(strBase, intPos, 1))
                           intValor = intValor * (10 - intPos)
                           intSoma = intSoma + intValor
                       Next
                       intSoma = intSoma * 10
                       intResto = intSoma Mod 11
                       strDigito1 = Right(IIf(intResto = 10, "0", Str(intResto)), 1)
                       strBase2 = Left(strBase, 8) & strDigito1
                       If strBase2 = strOrigem Then
                          ChecaInscrE = True
                          pInscr = Format(pInscr, "000,00000-0")
                          TextBox.Text = pInscr
                       End If
                    End If
            Case "AM"         '  Amazonas
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 If intSoma < 11 Then
                    strDigito1 = Right(Str(11 - intSoma), 1)
                 Else
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 End If
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,000,000-0")
                    TextBox.Text = pInscr
                 End If
            Case "AP"         '  Amapa
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intPeso = 0
                 intDig = 0
                 If Left(strBase, 2) = "03" Then
                    intNumero = Val(Left(strBase, 8))
                    If intNumero >= 3000001 And _
                       intNumero <= 3017000 Then
                       intPeso = 5
                       intDig = 0
                    Else
                      If intNumero >= 3017001 And _
                         intNumero <= 3019022 Then
                         intPeso = 9
                         intDig = 1
                      Else
                        If intNumero >= 3019023 Then
                           intPeso = 0
                           intDig = 0
                        End If
                        intSoma = intPeso
                        For intPos = 1 To 8
                            intValor = Val(Mid$(strBase, intPos, 1))
                            intValor = intValor * (10 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 11
                        intValor = 11 - intResto
                        If intValor = 10 Then
                           intValor = 0
                        Else
                          If intValor = 11 Then
                             intValor = intDig
                          End If
                          strDigito1 = Right(Str(intValor), 1)
                          strBase2 = Left(strBase, 8) & strDigito1
                          If strBase2 = strOrigem Then
                             ChecaInscrE = True
                             pInscr = Format(pInscr, "00,00,0000-0")
                             TextBox.Text = pInscr
                          End If
                        End If
                      End If
                    End If
                 End If
            Case "BA"         '  Bahia
                 strBase = Left(Trim(strOrigem) & "00000000", 8)
                 If InStr(1, "0123458", Left(strBase, 1), vbTextCompare) > 0 Then
                    intSoma = 0
                    For intPos = 1 To 6
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (8 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 10
                    strDigito2 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
                    strBase2 = Left(strBase, 6) & strDigito2
                    intSoma = 0
                    For intPos = 1 To 7
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (9 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 10
                    strDigito1 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
                 Else
                    intSoma = 0
                    For intPos = 1 To 6
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (8 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 6) & strDigito2
                    intSoma = 0
                    For intPos = 1 To 7
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (9 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 End If
                 strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,000-00")
                    TextBox.Text = pInscr
                 End If
            Case "CE"         '  Ceara
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = 0
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "0000,0000-0")
                    TextBox.Text = pInscr
                 End If
            Case "DF"         '  Distrito  Federal
                 strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                 If Left(strBase, 2) = "07" Then
                    intSoma = 0
                    intPeso = 2
                    For intPos = 11 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 11) & strDigito1
                    intSoma = 0
                    intPeso = 2
                    For intPos = 12 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 12) & strDigito2
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "000\.00000\.000\-00")
                       TextBox.Text = pInscr
                    End If
                 End If
            Case "ES"         '  Espirito  Santo
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,000,000")
                    TextBox.Text = pInscr
                 End If
            Case "GO"         '  Goias
                 intPos = Len(strOrigem)
                 If intPos <> 9 Then
                    intPos = 9 - intPos
                    Do While intPos <> 0
                       strOrigem = 0 & strOrigem
                       intPos = intPos - 1
                    Loop
                 End If
                 strBase = Left(Trim(strOrigem) & "000000000", 8)
                 If InStr(1, "10,11,15", Left(strBase, 2), vbTextCompare) > 0 Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    If intResto = 0 And Mid$(strOrigem, 9, 1) = 0 Then
                        ChecaInscrE = True
                        pInscr = Format(pInscr, "00,000,000-0")
                        TextBox.Text = pInscr
'                       strDigito1 = "0"
                    Else
                      If intResto = 1 Then
                         intNumero = Val(Left(strBase, 8))
                         strDigito1 = Right(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
                      Else
                         strDigito1 = Right(Str(11 - intResto), 1)
                      End If
                      strBase2 = Left(strBase, 8) & strDigito1
                      If strBase2 = strOrigem Then
                         ChecaInscrE = True
                         pInscr = Format(pInscr, "00,000,000-0")
                         TextBox.Text = pInscr
                      End If
                    End If
                 End If
            Case "MA"         '  Maranhão
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "12" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "00,00,0000-0")
                       TextBox.Text = pInscr
                    End If
                  End If
            Case "MT"         '  Mato  Grosso
                 intPos = Len(strOrigem)
                 If intPos <> 11 Then
                    intPos = 11 - intPos
                    Do While intPos <> 0
                       strOrigem = 0 & strOrigem
                       intPos = intPos - 1
                    Loop
                 End If
                 strBase = Left(Trim(strOrigem) & "0000000000", 10)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 10 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 9 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 10) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,000,0000-0")
                    TextBox.Text = pInscr
                 End If
            Case "MS"         '  Mato  Grosso  do  Sul
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "28" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "00,00,0000-0")
                       TextBox.Text = pInscr
                    End If
                 End If
            Case "MG"         '  Minas  Gerais
                 strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                 strBase2 = Left(strBase, 3) & "0" & Mid$(strBase, 4, 8)
                 intNumero = 2
                 For intPos = 1 To 12
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intNumero = IIf(intNumero = 2, 1, 2)
                     intValor = intValor * intNumero
                     If intValor > 9 Then
                        strDigito1 = Format(intValor, "00")
                        intValor = Val(Left(strDigito1, 1)) + _
                        Val(Right(strDigito1, 1))
                     End If
                     intSoma = intSoma + intValor
                 Next
                 intValor = intSoma
                 While Right(Format(intValor, "000"), 1) <> "0"
                       intValor = intValor + 1
                 Wend
                 strDigito1 = Right(Format(intValor - intSoma, "00"), 1)
                 strBase2 = Left(strBase, 11) & strDigito1
                 intSoma = 0
                 intPeso = 2
                 For intPos = 12 To 1 Step -1
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 11 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = strBase2 & strDigito2
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,000,000-0000")
                    TextBox.Text = pInscr
                 End If
            Case "PA"         '  Para
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "15" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "00,000,000-0")
                       TextBox.Text = pInscr
                    End If
                 End If
            Case "PB"         '  Paraiba
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = 0
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,000,000-0")
                    TextBox.Text = pInscr
                 End If
            Case "PE"         '  Pernambuco
                 strBase = Left(Trim(strOrigem) & "00000000000000", 14)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 13 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 9 Then
                        intPeso = 1
                     End If
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = intValor - 10
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 13) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,0,000,000000-0")
                    TextBox.Text = pInscr
                 End If
            Case "PI"         '  Piaui
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,000,000-0")
                    TextBox.Text = pInscr
                 End If
            Case "PR"         '  Parana
                 strBase = Left(Trim(strOrigem) & "0000000000", 10)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 8 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 7 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 intSoma = 0
                 intPeso = 2
                 For intPos = 9 To 1 Step -1
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 7 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = strBase2 & strDigito2
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,00000-00")
                    TextBox.Text = pInscr
                 End If
            Case "RJ"         '  Rio  de  Janeiro
                 strBase = Left(Trim(strOrigem) & "00000000", 8)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 7 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 7 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 7) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,000,00-0")
                    TextBox.Text = pInscr
                 End If
            Case "RN"         '  Rio  Grande  do  Norte
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "20" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intSoma = intSoma * 10
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto > 9, "0", Str(intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "00,000,000-0")
                       TextBox.Text = pInscr
                    End If
                 End If
            Case "RO"         '  Rondonia
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 strBase2 = Mid$(strBase, 4, 5)
                 intSoma = 0
                 For intPos = 1 To 5
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intValor = intValor * (7 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = intValor - 10
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,00000-0")
                    TextBox.Text = pInscr
                 End If
            Case "RR"         '  Roraima
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "24" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 9
                    strDigito1 = Right(Str(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "00,000000-0")
                       TextBox.Text = pInscr
                    End If
                 End If
            Case "RS"         '  Rio  Grande  do  Sul
                 strBase = Left(Trim(strOrigem) & "0000000000", 10)
                 intNumero = Val(Left(strBase, 3))
                 If intNumero > 0 And intNumero < 468 Then
                    intSoma = 0
                    intPeso = 2
                    For intPos = 9 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    intValor = 11 - intResto
                    If intValor > 9 Then
                       intValor = 0
                    End If
                    strDigito1 = Right(Str(intValor), 1)
                    strBase2 = Left(strBase, 9) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "000/00000-0")
                       TextBox.Text = pInscr
                    End If
                 End If
            Case "SC"         '  Santa  Catarina
                 intPos = Len(strOrigem)
                 If intPos <> 9 Then
                    intPos = 9 - intPos
                    Do While intPos <> 0
                       strOrigem = 0 & strOrigem
                       intPos = intPos - 1
                    Loop
                 End If
                 strBase = Left(Trim(strOrigem) & "000000000", 8)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,000,000")
                    TextBox.Text = pInscr
                 End If
            Case "SE"         '  Sergipe
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = 0
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "00,000,000-0")
                    TextBox.Text = pInscr
                 End If
            Case "SP"         '  São  Paulo
                 If Left(strOrigem, 1) = "P" Then
                    strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                    strBase2 = Mid$(strBase, 2, 8)
                    intSoma = 0
                    intPeso = 1
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso = 2 Then
                           intPeso = 3
                        End If
                        If intPeso = 9 Then
                           intPeso = 10
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(Str(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 11, 3)
                 Else
                    strBase = Left(Trim(strOrigem) & "000000000000", 12)
                    intSoma = 0
                    intPeso = 1
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso = 2 Then
                           intPeso = 3
                        End If
                        If intPeso = 9 Then
                           intPeso = 10
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(Str(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 10, 2)
                    intSoma = 0
                    intPeso = 2
                    For intPos = 11 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 10 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(Str(intResto), 1)
                    strBase2 = strBase2 & strDigito2
                 End If
                 If strBase2 = strOrigem Then
                    ChecaInscrE = True
                    pInscr = Format(pInscr, "000,000,000,000")
                    TextBox.Text = pInscr
                 End If
            Case "TO"         '  Tocantins
                 strBase = Left(Trim(strOrigem) & "00000000000", 11)
                 If InStr(1, "01,02,03,99", Mid$(strBase, 3, 2), vbTextCompare) > 0 Then
                    strBase2 = Left(strBase, 2) & Mid$(strBase, 5, 6)
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 10) & strDigito1
                    If strBase2 = strOrigem Then
                       ChecaInscrE = True
                       pInscr = Format(pInscr, "00,00000000-0")
                       TextBox.Text = pInscr
                    End If
                 End If
        End Select
        
End Function

Public Function ValidarInscrE(pUF As String, pInscr As String) As String

        Dim strBase As String
        Dim strBase2 As String
        Dim strOrigem As String
        Dim strDigito1 As String
        Dim strDigito2 As String
        Dim intPos As Integer
        Dim intValor As Integer
        Dim intSoma As Integer
        Dim intResto As Integer
        Dim intNumero As Integer
        Dim intPeso As Integer
        Dim intDig As Integer

        ValidarInscrE = Empty
        
        strBase = Empty
        strBase2 = Empty
        strOrigem = Empty

        If Trim(pInscr) = "ISENTO" Then
           ValidarInscrE = "ISENTO"
           Exit Function
        End If
        
        For intPos = 1 To Len(Trim(pInscr))
            If InStr(1, "0123456789P", Mid$(pInscr, intPos, 1), vbTextCompare) > 0 Then
               strOrigem = strOrigem & Mid$(pInscr, intPos, 1)
            End If
        Next

        Select Case pUF
               Case "AC"         '  Acre
                    strBase = Left(Trim(strOrigem) & "000000000", 9)
                    If Left(strBase, 2) = "01" And Mid$(strBase, 3, 2) <> "00" Then
                       intSoma = 0
                       For intPos = 1 To 8
                           intValor = Val(Mid$(strBase, intPos, 1))
                           intValor = intValor * (10 - intPos)
                           intSoma = intSoma + intValor
                       Next
                       intResto = intSoma Mod 11
                       strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                       strBase2 = Left(strBase, 8) & strDigito1
                       If strBase2 = strOrigem Then
                          pInscr = Format(pInscr, "00,00,0000-0")
                          ValidarInscrE = pInscr
                       End If
                    End If
               Case "AL"         '  Alagoas
                    strBase = Left(Trim(strOrigem) & "000000000", 9)
                    If Left(strBase, 2) = "24" Then
                       intSoma = 0
                       For intPos = 1 To 8
                           intValor = Val(Mid$(strBase, intPos, 1))
                           intValor = intValor * (10 - intPos)
                           intSoma = intSoma + intValor
                       Next
                       intSoma = intSoma * 10
                       intResto = intSoma Mod 11
                       strDigito1 = Right(IIf(intResto = 10, "0", Str(intResto)), 1)
                       strBase2 = Left(strBase, 8) & strDigito1
                       If strBase2 = strOrigem Then
                          pInscr = Format(pInscr, "000,00000-0")
                          ValidarInscrE = pInscr
                       End If
                    End If
            Case "AM"         '  Amazonas
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 If intSoma < 11 Then
                    strDigito1 = Right(Str(11 - intSoma), 1)
                 Else
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 End If
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,000,000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "AP"         '  Amapa
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intPeso = 0
                 intDig = 0
                 If Left(strBase, 2) = "03" Then
                    intNumero = Val(Left(strBase, 8))
                    If intNumero >= 3000001 And _
                       intNumero <= 3017000 Then
                       intPeso = 5
                       intDig = 0
                    Else
                      If intNumero >= 3017001 And _
                         intNumero <= 3019022 Then
                         intPeso = 9
                         intDig = 1
                      Else
                        If intNumero >= 3019023 Then
                           intPeso = 0
                           intDig = 0
                        End If
                        intSoma = intPeso
                        For intPos = 1 To 8
                            intValor = Val(Mid$(strBase, intPos, 1))
                            intValor = intValor * (10 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 11
                        intValor = 11 - intResto
                        If intValor = 10 Then
                           intValor = 0
                        Else
                          If intValor = 11 Then
                             intValor = intDig
                          End If
                          strDigito1 = Right(Str(intValor), 1)
                          strBase2 = Left(strBase, 8) & strDigito1
                          If strBase2 = strOrigem Then
                             pInscr = Format(pInscr, "00,00,0000-0")
                             ValidarInscrE = pInscr
                          End If
                        End If
                      End If
                    End If
                 End If
            Case "BA"         '  Bahia
                 strBase = Left(Trim(strOrigem) & "00000000", 8)
                 If InStr(1, "0123458", Left(strBase, 1), vbTextCompare) > 0 Then
                    intSoma = 0
                    For intPos = 1 To 6
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (8 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 10
                    strDigito2 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
                    strBase2 = Left(strBase, 6) & strDigito2
                    intSoma = 0
                    For intPos = 1 To 7
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (9 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 10
                    strDigito1 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
                 Else
                    intSoma = 0
                    For intPos = 1 To 6
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (8 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 6) & strDigito2
                    intSoma = 0
                    For intPos = 1 To 7
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (9 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 End If
                 strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,000-00")
                    ValidarInscrE = pInscr
                 End If
            Case "CE"         '  Ceara
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = 0
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "0000,0000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "DF"         '  Distrito  Federal
                 strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                 If Left(strBase, 2) = "07" Then
                    intSoma = 0
                    intPeso = 2
                    For intPos = 11 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 11) & strDigito1
                    intSoma = 0
                    intPeso = 2
                    For intPos = 12 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 12) & strDigito2
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "000\.00000\.000\-00")
                       ValidarInscrE = pInscr
                    End If
                 End If
            Case "ES"         '  Espirito  Santo
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,000,000")
                    ValidarInscrE = pInscr
                 End If
            Case "GO"         '  Goias
                 intPos = Len(strOrigem)
                 If intPos <> 9 Then
                    intPos = 9 - intPos
                    Do While intPos <> 0
                       strOrigem = 0 & strOrigem
                       intPos = intPos - 1
                    Loop
                 End If
                 strBase = Left(Trim(strOrigem) & "000000000", 8)
                 If InStr(1, "10,11,15", Left(strBase, 2), vbTextCompare) > 0 Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    If intResto = 0 And Mid$(strOrigem, 9, 1) = 0 Then
                        pInscr = Format(pInscr, "00,000,000-0")
                        ValidarInscrE = pInscr
'                       strDigito1 = "0"
                    Else
                      If intResto = 1 Then
                         intNumero = Val(Left(strBase, 8))
                         strDigito1 = Right(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
                      Else
                         strDigito1 = Right(Str(11 - intResto), 1)
                      End If
                      strBase2 = Left(strBase, 8) & strDigito1
                      If strBase2 = strOrigem Then
                         pInscr = Format(pInscr, "00,000,000-0")
                         ValidarInscrE = pInscr
                      End If
                    End If
                 End If
            Case "MA"         '  Maranhão
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "12" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "00,00,0000-0")
                       ValidarInscrE = pInscr
                    End If
                  End If
            Case "MT"         '  Mato  Grosso
                 intPos = Len(strOrigem)
                 If intPos <> 11 Then
                    intPos = 11 - intPos
                    Do While intPos <> 0
                       strOrigem = 0 & strOrigem
                       intPos = intPos - 1
                    Loop
                 End If
                 strBase = Left(Trim(strOrigem) & "0000000000", 10)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 10 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 9 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 10) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,000,0000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "MS"         '  Mato  Grosso  do  Sul
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "28" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "00,00,0000-0")
                       ValidarInscrE = pInscr
                    End If
                 End If
            Case "MG"         '  Minas  Gerais
                 strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                 strBase2 = Left(strBase, 3) & "0" & Mid$(strBase, 4, 8)
                 intNumero = 2
                 For intPos = 1 To 12
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intNumero = IIf(intNumero = 2, 1, 2)
                     intValor = intValor * intNumero
                     If intValor > 9 Then
                        strDigito1 = Format(intValor, "00")
                        intValor = Val(Left(strDigito1, 1)) + _
                        Val(Right(strDigito1, 1))
                     End If
                     intSoma = intSoma + intValor
                 Next
                 intValor = intSoma
                 While Right(Format(intValor, "000"), 1) <> "0"
                       intValor = intValor + 1
                 Wend
                 strDigito1 = Right(Format(intValor - intSoma, "00"), 1)
                 strBase2 = Left(strBase, 11) & strDigito1
                 intSoma = 0
                 intPeso = 2
                 For intPos = 12 To 1 Step -1
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 11 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = strBase2 & strDigito2
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,000,000-0000")
                    ValidarInscrE = pInscr
                 End If
            Case "PA"         '  Para
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "15" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "00,000,000-0")
                       ValidarInscrE = pInscr
                    End If
                 End If
            Case "PB"         '  Paraiba
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = 0
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,000,000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "PE"         '  Pernambuco
                 strBase = Left(Trim(strOrigem) & "00000000000000", 14)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 13 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 9 Then
                        intPeso = 1
                     End If
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = intValor - 10
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 13) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,0,000,000000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "PI"         '  Piaui
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,000,000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "PR"         '  Parana
                 strBase = Left(Trim(strOrigem) & "0000000000", 10)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 8 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 7 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 intSoma = 0
                 intPeso = 2
                 For intPos = 9 To 1 Step -1
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 7 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = strBase2 & strDigito2
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,00000-00")
                    ValidarInscrE = pInscr
                 End If
            Case "RJ"         '  Rio  de  Janeiro
                 strBase = Left(Trim(strOrigem) & "00000000", 8)
                 intSoma = 0
                 intPeso = 2
                 For intPos = 7 To 1 Step -1
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * intPeso
                     intSoma = intSoma + intValor
                     intPeso = intPeso + 1
                     If intPeso > 7 Then
                        intPeso = 2
                     End If
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 7) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,000,00-0")
                    ValidarInscrE = pInscr
                 End If
            Case "RN"         '  Rio  Grande  do  Norte
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "20" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intSoma = intSoma * 10
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto > 9, "0", Str(intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "00,000,000-0")
                       ValidarInscrE = pInscr
                    End If
                 End If
            Case "RO"         '  Rondonia
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 strBase2 = Mid$(strBase, 4, 5)
                 intSoma = 0
                 For intPos = 1 To 5
                     intValor = Val(Mid$(strBase2, intPos, 1))
                     intValor = intValor * (7 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = intValor - 10
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,00000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "RR"         '  Roraima
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 If Left(strBase, 2) = "24" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 9
                    strDigito1 = Right(Str(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "00,000000-0")
                       ValidarInscrE = pInscr
                    End If
                 End If
            Case "RS"         '  Rio  Grande  do  Sul
                 strBase = Left(Trim(strOrigem) & "0000000000", 10)
                 intNumero = Val(Left(strBase, 3))
                 If intNumero > 0 And intNumero < 468 Then
                    intSoma = 0
                    intPeso = 2
                    For intPos = 9 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    intValor = 11 - intResto
                    If intValor > 9 Then
                       intValor = 0
                    End If
                    strDigito1 = Right(Str(intValor), 1)
                    strBase2 = Left(strBase, 9) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "000/00000-0")
                       ValidarInscrE = pInscr
                    End If
                 End If
            Case "SC"         '  Santa  Catarina
                 intPos = Len(strOrigem)
                 If intPos <> 9 Then
                    intPos = 9 - intPos
                    Do While intPos <> 0
                       strOrigem = 0 & strOrigem
                       intPos = intPos - 1
                    Loop
                 End If
                 strBase = Left(Trim(strOrigem) & "000000000", 8)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,000,000")
                    ValidarInscrE = pInscr
                 End If
            Case "SE"         '  Sergipe
                 strBase = Left(Trim(strOrigem) & "000000000", 9)
                 intSoma = 0
                 For intPos = 1 To 8
                     intValor = Val(Mid$(strBase, intPos, 1))
                     intValor = intValor * (10 - intPos)
                     intSoma = intSoma + intValor
                 Next
                 intResto = intSoma Mod 11
                 intValor = 11 - intResto
                 If intValor > 9 Then
                    intValor = 0
                 End If
                 strDigito1 = Right(Str(intValor), 1)
                 strBase2 = Left(strBase, 8) & strDigito1
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "00,000,000-0")
                    ValidarInscrE = pInscr
                 End If
            Case "SP"         '  São  Paulo
                 If Left(strOrigem, 1) = "P" Then
                    strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                    strBase2 = Mid$(strBase, 2, 8)
                    intSoma = 0
                    intPeso = 1
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso = 2 Then
                           intPeso = 3
                        End If
                        If intPeso = 9 Then
                           intPeso = 10
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(Str(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 11, 3)
                 Else
                    strBase = Left(Trim(strOrigem) & "000000000000", 12)
                    intSoma = 0
                    intPeso = 1
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso = 2 Then
                           intPeso = 3
                        End If
                        If intPeso = 9 Then
                           intPeso = 10
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(Str(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 10, 2)
                    intSoma = 0
                    intPeso = 2
                    For intPos = 11 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 10 Then
                           intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(Str(intResto), 1)
                    strBase2 = strBase2 & strDigito2
                 End If
                 If strBase2 = strOrigem Then
                    pInscr = Format(pInscr, "000,000,000,000")
                    ValidarInscrE = pInscr
                 End If
            Case "TO"         '  Tocantins
                 strBase = Left(Trim(strOrigem) & "00000000000", 11)
                 If InStr(1, "01,02,03,99", Mid$(strBase, 3, 2), vbTextCompare) > 0 Then
                    strBase2 = Left(strBase, 2) & Mid$(strBase, 5, 6)
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 10) & strDigito1
                    If strBase2 = strOrigem Then
                       pInscr = Format(pInscr, "00,00000000-0")
                       ValidarInscrE = pInscr
                    End If
                 End If
        End Select
        
End Function





