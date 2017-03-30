Attribute VB_Name = "Module2"
Public Sub ppBarraPerc(lstrNomeObjeto As SSPanel, ByVal liPerc As Integer)
    'recebe perc, um número de 0 a 100
    '100 = "desliga" a barra
    'Acum = boolean, acumula o anterior ou não (true/false)

    Static liVA_Vez
    Static liVA_SaveCor As Long
    
    'If liAcum Then liPerc = liPerc + lstrForm.lstrNomeObjeto.FloodPercent
    
    If liPerc > 100 Or liPerc < -1 Then
        MsgBox "Percentual deve estar entre -1 e 100", 16, "Erro de parâmetro em ppBarraPerc"
        Exit Sub
    End If
    
    If IsEmpty(liVA_Vez) Or liVA_Vez = 1 Then
        'liga barra - altera o painel
        lstrNomeObjeto.Visible = True
        lstrNomeObjeto.Caption = ""
        lstrNomeObjeto.FloodShowPct = True
        liVA_SaveCor = lstrNomeObjeto.ForeColor
        lstrNomeObjeto.ForeColor = RGB(0, 0, 0) 'preto
        lstrNomeObjeto.BevelOuter = 1
        lstrNomeObjeto.BevelWidth = 1
        lstrNomeObjeto.BorderWidth = 3
        lstrNomeObjeto.FloodType = 1 'left to right
        lstrNomeObjeto.FontSize = 8.25
    End If
    
    If liPerc < 100 Then
        If liPerc > 48 Then lstrNomeObjeto.ForeColor = RGB(255, 255, 255) 'branco

        'mostra perc
        lstrNomeObjeto.FloodPercent = liPerc
        liVA_Vez = 2 'ou mais
    Else
        'desliga barra - reestrutura painel
        lstrNomeObjeto.BevelOuter = 1
        lstrNomeObjeto.BevelWidth = 1
        lstrNomeObjeto.BorderWidth = 3
        lstrNomeObjeto.FloodType = 0
        lstrNomeObjeto.FontSize = 8.25
        lstrNomeObjeto.ForeColor = liVA_SaveCor
        lstrNomeObjeto.FloodShowPct = False
        lstrNomeObjeto.Visible = False
        liVA_Vez = 1
    End If
End Sub

Public Function pfstrValidaString(pstrTexto As String) As String
    pfstrValidaString = pstrTexto
    If InStr(1, pstrTexto, "'", 1) = 0 And InStr(1, pstrTexto, """", 1) Then Exit Function
    Dim liContador As Integer
    For liContador = 1 To Len(pstrTexto)
        If InStr(liContador, pstrTexto, "'") = 0 And InStr(liContador, pstrTexto, """") = 0 Then Exit For
        Mid(pstrTexto, InStr(liContador, pstrTexto, "'", 1)) = " "
        If InStr(liContador, pstrTexto, "'") = 0 And InStr(liContador, pstrTexto, """") = 0 Then Exit For
        Mid(pstrTexto, InStr(liContador, pstrTexto, """", 1)) = " "
    Next
    pfstrValidaString = pstrTexto
End Function

Function ppVerifica_CgcCpf(lstrCgcCpf As String) As Boolean
    ' Esta Rotina Devolverá True  Se o Cgc/Cpf Informado For valido
    '                    ou False Se o Cgc/Cpf Não For Correto
    ' Para Chamar esta Rotina de Consistência
    ' 1 ) Atribuir o valor do CgcCpf a uma Variavel String
    ' 2 ) Chamar a Rotina com : ppVerifica_CgcCpf (Variavel)
    ' Uma Forma Simples de fazer a Consistencia
      ' é Copiando as linhas abaixo (exemplo)
    ' para dentro do Programa
    ' Dim lstrCgcCpf As String
    ' lstrCgcCpf = Me.CgcCpf.Text
    ' If ppVerifica_CgcCpf(lstrCgcCpf) = False then
    '  MsgBox "( Cgc/Cpf Informado Não é um Cgc/Cpf Correto )"
    '  Me.CgcCpf.SetFocus
    '  Exit Sub
    ' End if
    
    ppVerifica_CgcCpf = False
    Dim lstrCgcCpf2         As String
    Dim lstrDigito          As String
    Static liNumero(15)     As Integer
    Dim liResto             As Integer
    Dim liResultado         As Integer
    Dim liSomaDigito10      As Integer
    Dim liResto1            As Integer

    lstrCgcCpf2 = Format(lstrCgcCpf, "@@@@@@@@@@@@@@")
    lstrDigito = Mid(lstrCgcCpf2, 13, 2)
    
    liNumero(1) = Val(Mid(lstrCgcCpf2, 1, 1))
    liNumero(2) = Val(Mid(lstrCgcCpf2, 2, 1))
    liNumero(3) = Val(Mid(lstrCgcCpf2, 3, 1))
    liNumero(4) = Val(Mid(lstrCgcCpf2, 4, 1))
    liNumero(5) = Val(Mid(lstrCgcCpf2, 5, 1))
    liNumero(6) = Val(Mid(lstrCgcCpf2, 6, 1))
    liNumero(7) = Val(Mid(lstrCgcCpf2, 7, 1))
    liNumero(8) = Val(Mid(lstrCgcCpf2, 8, 1))
    liNumero(9) = Val(Mid(lstrCgcCpf2, 9, 1))
    liNumero(10) = Val(Mid(lstrCgcCpf2, 10, 1))
    liNumero(11) = Val(Mid(lstrCgcCpf2, 11, 1))
    liNumero(12) = Val(Mid(lstrCgcCpf2, 12, 1))
    liNumero(13) = Val(Mid(lstrCgcCpf2, 13, 1))
    liNumero(14) = Val(Mid(lstrCgcCpf2, 14, 1))
    
    If Len(Trim(lstrCgcCpf)) > 11 Then  ' Cgc
        liResultado = liNumero(1) * 2
        If liResultado > 9 Then
            liSomaDigito10 = liResultado + 1
        Else
            liSomaDigito10 = liResultado
        End If
        liResultado = liNumero(3) * 2
        If liResultado > 9 Then
            liSomaDigito10 = liSomaDigito10 + liResultado + 1
        Else
            liSomaDigito10 = liSomaDigito10 + liResultado
        End If
        liResultado = liNumero(5) * 2
        If liResultado > 9 Then
            liSomaDigito10 = liSomaDigito10 + liResultado + 1
        Else
            liSomaDigito10 = liSomaDigito10 + liResultado
        End If
        liResultado = liNumero(7) * 2
        If liResultado > 9 Then
            liSomaDigito10 = liSomaDigito10 + liResultado + 1
        Else
            liSomaDigito10 = liSomaDigito10 + liResultado
        End If
        liSomaDigito10 = liSomaDigito10 + liNumero(2) + liNumero(4) + liNumero(6)
        If Mid(Str(liSomaDigito10), Len(Str(liSomaDigito10)), 1) = "0" Then
            liResto = 0
        Else
            liResto = 10 - Val(Mid(Str(liSomaDigito10), _
            Len(Str(liSomaDigito10)), 1))
        End If
        If liResto <> liNumero(8) Then
            Exit Function
        End If
        liResultado = (liNumero(1) * 5) + (liNumero(2) * 4) _
            + (liNumero(3) * 3) + (liNumero(4) * 2) _
            + (liNumero(5) * 9) + (liNumero(6) * 8) + _
            (liNumero(7) * 7) + (liNumero(8) * 6) + _
            (liNumero(9) * 5) + (liNumero(10) * 4) + _
            (liNumero(11) * 3) + (liNumero(12) * 2)
        ' Atribui para resto o resto da divisão
            ' de liResultado dividido por 11
            liResto = liResultado Mod 11
            If liResto < 2 Then
            liResto1 = 0
        Else
            liResto1 = 11 - liResto
        End If
        If liResto1 <> liNumero(13) Then
            Exit Function
        End If
        liResultado = (liNumero(1) * 6) + _
            (liNumero(2) * 5) + (liNumero(3) * 4) + _
            (liNumero(4) * 3) + (liNumero(5) * 2) + _
            (liNumero(6) * 9) + (liNumero(7) * 8) + _
            (liNumero(8) * 7) + (liNumero(9) * 6) + _
            (liNumero(10) * 5) + (liNumero(11) * 4) + _
            (liNumero(12) * 3) + (liNumero(13) * 2)
        ' Atribui para resto o resto da divisão
            ' de liResultado dividido por 11
            liResto = liResultado Mod 11
        If liResto < 2 Then
            liResto1 = 0
        Else
            liResto1 = 11 - liResto
        End If
        If liResto1 <> liNumero(14) Then
            Exit Function
        End If
    Else  ' Cpf
        liResultado = (liNumero(4) * 1) + _
            (liNumero(5) * 2) + (liNumero(6) * 3) _
            + (liNumero(7) * 4) + (liNumero(8) * 5) _
            + (liNumero(9) * 6) + (liNumero(10) * 7) _
            + (liNumero(11) * 8) + (liNumero(12) * 9)
        liResto = liResultado Mod 11
        If liResto > 9 Then
            liResto1 = liResto - 10
        Else
            liResto1 = liResto
        End If
        If liResto1 <> liNumero(13) Then
            Exit Function
        End If
        
        liResultado = (liNumero(5) * 1) _
            + (liNumero(6) * 2) + (liNumero(7) * 3) _
            + (liNumero(8) * 4) + (liNumero(9) * 5) + _
            (liNumero(10) * 6) + (liNumero(11) * 7) + _
            (liNumero(12) * 8) + (liResto * 9)
        liResto = liResultado Mod 11
        If liResto > 9 Then
            liResto1 = liResto - 10
        Else
            liResto1 = liResto
        End If
        If liResto1 <> liNumero(14) Then
            Exit Function
        End If

    End If
    ppVerifica_CgcCpf = True
End Function

Function pfCriptaSenha(lstrSenha As String) As String
    If lstrSenha = "" Or Trim(lstrSenha) = "" Then Exit Function

    Dim lstrSenhaCripta As String
    Dim liVarreSenha As Integer

    liVarreSenha = 1

    While liVarreSenha <= Len(Trim(lstrSenha))
        lstrSenhaCripta = lstrSenhaCripta + "*"
        liVarreSenha = liVarreSenha + 1
    Wend
    
    pfCriptaSenha = lstrSenhaCripta
End Function

Function pfStrRight(lstrVerificar As String, lstrAchar As String) As Integer
    If lstrVerificar = "" Or lstrAchar = "" Then
        pfStrRight = 0
        Exit Function
    End If

    Dim liPosicao As Integer
    Dim lstrCaracter As String
    Dim liAchou As Integer

    liPosicao = Len(lstrVerificar)
    liAchou = 0

    While liPosicao > 0 And liAchou = 0
        If Mid(lstrVerificar, liPosicao, 1) = Trim(lstrAchar) Then liAchou = liPosicao
        liPosicao = liPosicao - 1
    Wend

    pfStrRight = IIf(liAchou = 0, 1, liAchou)
End Function


