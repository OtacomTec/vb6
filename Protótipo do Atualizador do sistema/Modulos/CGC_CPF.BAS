Attribute VB_Name = "CGC_CPF"

''*****************************************************************************'
''                                                                             '
''An�lise...........................: Marcos Bai�o                            '
''Programa��o.......................: Marcos Bai�o                             '
''Data..............................: 23/08/1999                               '
''Data da �ltima manuten��o.........:                                          '
''Manuten��o executada por..........:                                          '
''                                                                             '
''Este m�dulo foi desenvolvido com a fun��o de falicitar o desenvolvedor para a'
''valida��o de CGC e CPF. O desenvolvedor dever� tamb�m ter inserido os objetos'
''no formul�rio e nome�-los conforme as instru��es da fun��o.                  '
''                                                                             '
''*****************************************************************************'
'
'Public Function CGC_CPF(Nome_Objeto As Object, Optional strCNPJ_CPF As String) As Boolean
'    On Error GoTo Erro
'
'    Dim intRes As Integer
'    'Vari�vel que vai receber o valor do n�mero de bytes existentes no campo
'    intRes = Len(Nome_Objeto.Text)
'    'Carregando a vari�vel com o n�mero de bytes existentes no campo
'
'
'    'colocado por vitor
'    If strCNPJ_CPF = "CPF" Then
'       GoTo Verificacao
'    ElseIf strCNPJ_CPF = "CNPJ" Then
'       GoTo Verificacao_CGC
'    Else
'       GoTo Verificacao_CGC
'   End If
''******VALIDA��O DO CPF******************
'Verificacao:
'    'Rotina de verifica��o de CGC_CPF
'    'Caso existam 14 d�gitos - CPF cotando com os carecteres
'    If intRes = 14 Then
'        Dim TempCPF_1 As String, TempCPF_2 As String, TempCPF_3 As String, TempCPF_4 As String, TempCPF_5 As String
'        Dim TempCPF_6 As String, TempCPF_7 As String, TempCPF_8, TempCPF_9, TempCPF_0 As String
'        'Se o campo tiver 14 d�gitos, declara��o das vari�veis que receber�o os valores dos
'        'CPF que n�o devem ser validados
'            TempCPF_1 = "111.111.111-11"
'            TempCPF_2 = "222.222.222-22"
'            TempCPF_3 = "333.333.333-33"
'            TempCPF_4 = "444.444.444-44"
'            TempCPF_5 = "555.555.555-55"
'            TempCPF_6 = "666.666.666-66"
'            TempCPF_7 = "777.777.777-77"
'            TempCPF_8 = "888.888.888-88"
'            TempCPF_9 = "999.999.999-99"
'            TempCPF_0 = "000.000.000-00"
'            'Carrega as vari�veis com os valores dos CPF's inv�lidos
'
'        If Nome_Objeto.Text = TempCPF_1 Or Nome_Objeto.Text = TempCPF_2 Or Nome_Objeto.Text = TempCPF_3 Or Nome_Objeto.Text = TempCPF_4 Or Nome_Objeto.Text = TempCPF_5 Or '            Nome_Objeto.Text = TempCPF_6 Or Nome_Objeto.Text = TempCPF_7 Or Nome_Objeto.Text = TempCPF_8 Or Nome_Objeto.Text = TempCPF_9 Or Nome_Objeto.Text = TempCPF_0 Then
'            MsgBox "O n�mero do CPF n�o � v�lido", vbCritical, "Director"
'            Nome_Objeto.Mask = Empty
'            Nome_Objeto.Text = Empty
'            Nome_Objeto.Mask = "###.###.###-##"
'            CGC_CPF = False
'            Exit Function
'            'Verifica se o CPF digitado � um dos CPF's da lista acima:
'        Else
'            Dim CPF_1 As String, CPF_2 As String, CPF_3 As String, CPF_4 As String, CPF_5 As String
'            Dim CPF_6 As String, CPF_7 As String, CPF_8 As String, CPF_9 As String, CPF_10 As String
'            Dim CPF_11 As String
'            Dim CPFDV_1 As Integer, CPFDV_2 As Integer
'            Dim CPFDV_1_1 As Long, CPFDV_2_2 As Long
'            'Declara��o das vari�veis que v�o receber o valor de cada d�gito
'
'            CPF_1 = Val(Mid(Nome_Objeto.Text, 1, 1))
'            CPF_2 = Val(Mid(Nome_Objeto.Text, 2, 1))
'            CPF_3 = Val(Mid(Nome_Objeto.Text, 3, 1))
'            CPF_4 = Val(Mid(Nome_Objeto.Text, 5, 1))
'            CPF_5 = Val(Mid(Nome_Objeto.Text, 6, 1))
'            CPF_6 = Val(Mid(Nome_Objeto.Text, 7, 1))
'            CPF_7 = Val(Mid(Nome_Objeto.Text, 9, 1))
'            CPF_8 = Val(Mid(Nome_Objeto.Text, 10, 1))
'            CPF_9 = Val(Mid(Nome_Objeto.Text, 11, 1))
'            CPF_10 = Val(Mid(Nome_Objeto.Text, 13, 1))
'            CPF_11 = Val(Mid(Nome_Objeto.Text, 14, 1))
'            'Carrega as vari�veis com o valor de cada d�gito do CPF digitado
'
'            CPFDV_1 = (CPF_1 * 10) + (CPF_2 * 9) + (CPF_3 * 8) + (CPF_4 * 7) + (CPF_5 * 6) + (CPF_6 * 5) + (CPF_7 * 4) + (CPF_8 * 3) + (CPF_9 * 2)
'            'Multiplica cada d�gito por um n�mero e soma os valores
'
'            CPFDV_1_1 = CPFDV_1 Mod 11
'            'Divide o resultado por 11 e armazena o resto
'
'            If CPFDV_1_1 = 0 Or CPFDV_1_1 = 1 Then
'                CPFDV_1_1 = 0
'            Else
'                CPFDV_1_1 = CPFDV_1_1 - 11
'                'Verifica o resultado, se for 1 ou 0 , o resultado passa a ser 0,
'                'se o resultado n�o for 1 ou 0 , o resultado � ele mesmo - 11
'
'                If CPFDV_1_1 < 0 Then
'                    CPFDV_1_1 = CPFDV_1_1 * -1
'                End If
'                'Verifica se o resultado � negativo, se for, o trasnforma em positivo
'            End If
'
'            If CPFDV_1_1 <> CPF_10 Then
'                MsgBox "O n�mero do CPF n�o � v�lido", vbCritical, "Director"
'                Nome_Objeto.Mask = Empty
'                Nome_Objeto.Text = Empty
'                Nome_Objeto.Mask = "###.###.###-##"
'                Exit Function
'                'Verifica se o primeiro d�gito � verdadeiro, se for falso, sai da fun��o
'            Else
'                CPFDV_2 = (CPF_1 * 11) + (CPF_2 * 10) + (CPF_3 * 9) + (CPF_4 * 8) + (CPF_5 * 7) + (CPF_6 * 6) + (CPF_7 * 5) + (CPF_8 * 4) + (CPF_9 * 3) + (CPF_10 * 2)
'                CPFDV_2_2 = CPFDV_2 Mod 11
'                'Se for verdadeiro, carrega as vari�veis do segundo d�gito verificador,
'                'multiplica cada vari�vel por um n�mero, divide por 11 e garda o resto.
'
'                If CPFDV_2_2 = 0 Or CPFDV_2_2 = 1 Then
'                    CPFDV_2_2 = 0
'                    'Se o resultado for 1 ou 0, o resultado � 0
'                Else
'                    CPFDV_2_2 = CPFDV_2_2 - 11
'                    'Se n�o, o resultado = resultado - 11
'
'                    If CPFDV_2_2 < 0 Then
'                        CPFDV_2_2 = CPFDV_2_2 * -1
'                        'Se for um n�mero negativo, transforma-o em positivo
'                    End If
'                End If
'
'                If CPFDV_2_2 <> CPF_11 Then
'                    MsgBox "O n�mero de CPF n�o � v�lido", vbCritical, "Director"
'                    Nome_Objeto.Mask = Empty
'                    Nome_Objeto.Text = Empty
'                    Nome_Objeto.Mask = "###.###.###-##"
'                    Exit Function
'
'                    If CPFDV_1_1 & CPFDV_2_2 = (Mid(Nome_Objeto.Text, 12, 2)) Then
'                    Exit Function
'                    End If
'                    'Compara o segundo d�gito, se for falso, emite uma mensagem de erro,
'                    'limpa o campo e prende o foco e sai da fun��o
'                Else
'                    'Nome_Objeto.Text = Format(Nome_Objeto.Text, "&&&.&&&.&&&-&&")
'                    CGC_CPF = True
'                    Exit Function
'                    'Se verdadeiro, formata o campo e sai da fun��o
'                End If
'            End If
'        End If
'
'    Else
'        Nome_Objeto.Mask = Empty
'        Nome_Objeto.Text = Empty
'        Nome_Objeto.Mask = "###.###.###-##"
'        Exit Function         'Caso existam 18 d�gitos - CGC
'    End If
'
'    Exit Function      'Sair
'
''******VALIDA��O DO CNPJ******************
'Verificacao_CGC:
'    Dim TempCGC_1 As String, TempCGC_2 As String, TempCGC_3 As String
'    'Declara��o das vari�veis que ir�o receber valores de CGC's
'    TempCGC_1 = "62.580.000/0000-84"
'    TempCGC_2 = "78.408.606/0001-51"
'    TempCGC_3 = "92.875.673/0001-63"
'        'Carrega as vari�veis com CGC's v�lidos que n�o obedecem � regra
'    If Nome_Objeto.Text = TempCGC_1 Or Nome_Objeto.Text = TempCGC_2 Or Nome_Objeto.Text = TempCGC_3 Then
'        CGC_CPF = False
'        Exit Function
'        'Se o campo conter um destes CGC's, formata o campo e sai da fun��o
'    Else
'        Dim CGC_1 As String, CGC_2 As String, CGC_3 As String, CGC_4 As String, CGC_5 As String
'        Dim CGC_6 As String, CGC_7 As String, CGC_8 As String, CGC_9 As String, CGC_10 As String
'        Dim CGC_11 As String, CGC_12 As String, CGC_13 As String, CGC_14 As String
'        Dim CGCDV_1 As Integer, CGCDV_2 As Integer
'        Dim CGCDV_1_1 As Long, CGCDV_2_2 As Long
'        'Declara��o das vari�veis que ir�o receber o valor de cada d�gito
'        If Len(Nome_Objeto) = 18 Then
'            CGC_1 = Val(Mid(Nome_Objeto.Text, 1, 1))
'            CGC_2 = Val(Mid(Nome_Objeto.Text, 2, 1))
'            CGC_3 = Val(Mid(Nome_Objeto.Text, 4, 1))
'            CGC_4 = Val(Mid(Nome_Objeto.Text, 5, 1))
'            CGC_5 = Val(Mid(Nome_Objeto.Text, 6, 1))
'            CGC_6 = Val(Mid(Nome_Objeto.Text, 8, 1))
'            CGC_7 = Val(Mid(Nome_Objeto.Text, 9, 1))
'            CGC_8 = Val(Mid(Nome_Objeto.Text, 10, 1))
'            CGC_9 = Val(Mid(Nome_Objeto.Text, 12, 1))
'            CGC_10 = Val(Mid(Nome_Objeto.Text, 13, 1))
'            CGC_11 = Val(Mid(Nome_Objeto.Text, 14, 1))
'            CGC_12 = Val(Mid(Nome_Objeto.Text, 15, 1))
'            CGC_13 = Val(Mid(Nome_Objeto.Text, 17, 1))
'            CGC_14 = Val(Mid(Nome_Objeto.Text, 18, 1))
'            'Carrega as vari�veis com o valor de cada d�gito do CGC digitado
'
'            CGCDV_1 = (CGC_1 * 5) + (CGC_2 * 4) + (CGC_3 * 3) + (CGC_4 * 2) + (CGC_5 * 9) + (CGC_6 * 8) + (CGC_7 * 7) + (CGC_8 * 6) + (CGC_9 * 5) + (CGC_10 * 4) + (CGC_11 * 3) + (CGC_12 * 2)
'            CGCDV_1_1 = CGCDV_1 Mod 11
'            'Multiplica cada vari�vel por um n�mero, e as soma,
'            'Divide o resultado por 11 e armazena o resto
'
'            If CGCDV_1_1 = 0 Or CGCDV_1_1 = 1 Then
'                CGCDV_1_1 = 0
'                'Se o resultado for 0 ou 1 , o resultado = 0
'            Else
'                CGCDV_1_1 = CGCDV_1_1 - 11
'                'Se n�o, o resultado = resultado - 11
'
'                If CGCDV_1_1 < 0 Then
'                    CGCDV_1_1 = CGCDV_1_1 * -1
'                    'Se o resultado for um n�mero negativo, transforma-o em positivo
'                End If
'            End If
'
'            If CGCDV_1_1 <> CGC_13 Then
'                MsgBox "O n�mero de CNPJ n�o � v�lido", vbCritical, "Director"
'                Nome_Objeto.Mask = Empty
'                Nome_Objeto.Text = Empty
'                Nome_Objeto.Mask = "##.###.###/####-##"
'                Exit Function
'                'Compara o primeiro d�gito, se for falso, emite mensagem d erro, limpa
'                'o campo, prende o fico e sai da fun��o
'            Else
'                CGCDV_2 = (CGC_1 * 6) + (CGC_2 * 5) + (CGC_3 * 4) + (CGC_4 * 3) + (CGC_5 * 2) + (CGC_6 * 9) + (CGC_7 * 8) + (CGC_8 * 7) + (CGC_9 * 6) + (CGC_10 * 5) + (CGC_11 * 4) + (CGC_12 * 3) + (CGC_13 * 2)
'                CGCDV_2_2 = CGCDV_2 Mod 11
'                'Se for verdadeiro, verifica o segundo d�gito, multiplica cada vari�vel por
'                'um n�mero, divide o resultado por 11 e armazena o resto
'
'                If CGCDV_2_2 = 0 Or CGCDV_2_2 = 1 Then
'                    CGCDV_2_2 = 0
'                    'Se o resultado for 0 ou 1, o resultado = 0
'                Else
'                    CGCDV_2_2 = CGCDV_2_2 - 11
'                    'Se n�o, o resultado = resultado - 11
'
'                    If CGCDV_2_2 < 0 Then
'                        CGCDV_2_2 = CGCDV_2_2 * -1
'                        'Se o resultado for negativo, transforma-o em positivo
'                    End If
'                End If
'
'                If CGCDV_2_2 <> CGC_14 Then
'                    MsgBox "O n�mero de CNPJ n�o � v�lido", vbCritical, "Director"
'                    Nome_Objeto.Mask = Empty
'                    Nome_Objeto.Text = Empty
'                    Nome_Objeto.Mask = "##.###.###/####-##"
'                    Exit Function
'
'                    'Verifica o segundo d�gito,se for falso, emite mensagem de erro,
'                    'limpa o campo, prende o foco e sai da fun��o
'                Else
'                    'Nome_Objeto.Text = Format(Nome_Objeto.Text, "&&.&&&.&&&/&&&&.&&")
'                    CGC_CPF = True
'                    'Nome_Objeto.Text = Format(Nome_Objeto.Text, "&&.&&&.&&&/&&&&.&&")
'                    'Exit Function
'                    'Se verdadeiro, formata o campo e sai da fun��o
'                End If
'
'                If CGCDV_1_1 & CGCDV_2_2 = (Mid(Nome_Objeto.Text, 17, 2)) Then
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
'
'    Exit Function
'Erro:
'    Erro.Erro ("CGC_CPF")
'    Resume Next
'End Function
''----------------------------Instru��es---------------------------------------'
''                                                                             '
''Esta � a fun��o de valida��o de CGC e CPF, objetos necess�rios e nomenclatura'
''   TextBox de acordo com o padr�o processa.                                  '
''                                                                             '
''Como Cham�-la?                                                               '
'' Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.'
''      EX: CALL CGC_CPF.CGC_CPF(                                              '
''                                                                             '
''Como Preench�-la?                                                            '
'' Dentro do par�nteses digite o nome do TextBox. Em seguida feche o par�nteses'
''                                                                             '
''------------------------------------------------------------------------------'
Function FormatarCNPJ(CNPJ As String, TextBox As TextBox) As String
    Dim Conta As Integer
    Dim Soma As Long
    Dim Passo As Integer
    Dim Digito1 As Integer
    Dim Digito2 As Integer
    Dim flag As Integer
    Dim AcrescentarD�gito As Boolean
    
    If Len(CNPJ) <> 14 Then
        If Len(CNPJ) = 12 Then
            AcrescentarD�gito = True
        Else
            MsgBox "N�mero de digitos invalido!Redigite", vbCritical, "Only Tech"
            FormatarCNPJ = "INV�LIDO"
            TextBox.Text = Empty
            TextBox.SetFocus
            Exit Function
        End If
    End If
    
    For Passo = 5 To 6
        Soma = 0
        flag = Passo
        For Conta = 1 To Passo + 7
            Soma = Soma + (Val(Mid(CNPJ, Conta, 1)) * flag)
            flag = IIf(flag > 2, flag - 1, 9)
        Next
        Soma = Soma Mod 11
        If Passo = 5 Then
            Digito1 = IIf(Soma > 1, 11 - Soma, 0)
            
            If AcrescentarD�gito = True Then CNPJ = CNPJ & Digito1
        End If
        If Passo = 6 Then Digito2 = IIf(Soma > 1, 11 - Soma, 0)
    Next
    
    If AcrescentarD�gito = True Then
        FormatarCNPJ = Format(CNPJ & Digito2, "00\.000\.000\/0000\-00")
    Else
        If Digito1 <> Val(Mid(CNPJ, 13, 1)) Then
          MsgBox "CNPJ Inv�lido!Redigite um V�lido!", vbCritical, "Only Tech"
          FormatarCNPJ = Format(CNPJ, "00\.000\.000\/0000\-00")
          FormatarCNPJ = "INV�LIDO"
          TextBox.Text = Empty
          TextBox.SetFocus
        End If
        If Digito2 <> Val(Mid(CNPJ, 14, 1)) Then
          MsgBox "CNPJ Inv�lido!Redigite um V�lido!", vbCritical, "Only Tech"
          FormatarCNPJ = Format(CNPJ, "00\.000\.000\/0000\-00")
          FormatarCNPJ = "INV�LIDO"
          TextBox.Text = Empty
          TextBox.SetFocus
        End If
        TextBox.Text = Format(CNPJ, "00\.000\.000\/0000\-00")
    End If

End Function
Function FormatarCPF(CPF As String, TextBox As TextBox) As String
    Dim Conta As Integer
    Dim Soma As Integer
    Dim Resto As Integer
    Dim Passo As Integer
    Dim D1 As Byte
    Dim D2 As Byte
    Dim AcrescentarD�gito As Boolean
    
    If Len(CPF) <> 11 Then
        If Len(CPF) <> 9 Then
            MsgBox "N�mero de digitos invalido!Redigite", vbCritical, "Only Tech"
            FormatarCPF = "INV�LIDO"
            TextBox.SetFocus
            Exit Function
        End If
        AcrescentarD�gito = True
    End If
        
    For Passo = 11 To 12
        Soma = 0
        For Conta = 1 To Passo - 2
            Soma = Soma + Val(Mid(CPF, Conta, 1)) * (Passo - Conta)
        Next
        Resto = 11 - (Soma - (Int(Soma / 11) * 11))
        If Resto = 10 Or Resto = 11 Then Resto = 0
        
        If AcrescentarD�gito = True Then
            If Passo = 11 Then
                CPF = CPF & Resto
                D1 = Resto
            Else
                D2 = Resto
            End If
        Else
            If Resto <> Val(Mid(CPF, Passo - 1, 1)) Then
               AcrescentarD�gito = True
            End If
        End If
    Next
    
    If AcrescentarD�gito = True Then
        MsgBox "CPF Inv�lido!Redigite um V�lido!", vbCritical, "Only Tech"
        FormatarCPF = Format(CPF & D2, "000\.000\.000\-00")
        FormatarCPF = "INV�LIDO"
        TextBox.SetFocus
    Else
        TextBox.Text = Format(CPF, "000\.000\.000\-00")
    End If
    
End Function
