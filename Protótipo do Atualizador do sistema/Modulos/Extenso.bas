Attribute VB_Name = "Extenso"
'****************************************************************************************
' Nome.......: gsFormatarExtenso
' Descrição..: Lê o valor passado e retorna uma string com o valor em extenso
' Parâmetros.: rcValor - valor a ser formatado
' Retorno....: string sem aspas simples
'****************************************************************************************
Public Function gsFormatarExtenso(rcValor As Currency) As String
    
    Dim BUFFER As String, ExtParc As String
    Static Unidades(10) As String, Dezenas(10) As String
    Static Centenas(10) As String
    Static DezenaEspecial(10) As String, Quantia(4) As String
    Dim Indice As Integer, Tamanho As Integer
    Dim I As Integer, Passos As Integer
    Dim Temporario As Double, Resto As Double
    Dim Centavos As Variant
    
    Unidades(0) = "0": Unidades(1) = "Um": Unidades(2) = "Dois"
    Unidades(3) = "Três": Unidades(4) = "Quatro": Unidades(5) = "Cinco"
    Unidades(6) = "Seis": Unidades(7) = "Sete": Unidades(8) = "Oito"
    Unidades(9) = "Nove": Dezenas(0) = "Dez"
    Dezenas(1) = "Onze": Dezenas(2) = "Doze": Dezenas(3) = "Treze"
    Dezenas(4) = "Quatorze": Dezenas(5) = "Quinze": Dezenas(6) = "Dezesseis"
    Dezenas(7) = "Dezessete": Dezenas(8) = "Dezoito": Dezenas(9) = "Dezenove"
    DezenaEspecial(1) = "Dez": DezenaEspecial(2) = "Vinte"
    DezenaEspecial(3) = "Trinta": DezenaEspecial(4) = "Quarenta"
    DezenaEspecial(5) = "Cinquenta": DezenaEspecial(6) = "Sessenta"
    DezenaEspecial(7) = "Setenta": DezenaEspecial(8) = "Oitenta"
    DezenaEspecial(9) = "Noventa"

    Centenas(1) = "Cento": Centenas(2) = "Duzentos"
    Centenas(3) = "Trezentos": Centenas(4) = "Quatrocentos"
    Centenas(5) = "Quinhentos": Centenas(6) = "Seiscentos"
    Centenas(7) = "Setecentos": Centenas(8) = "Oitocentos"
    Centenas(9) = "Novecentos"
    Quantia(1) = " ": Quantia(2) = " Mil ": Quantia(3) = " Milhões "
    BUFFER = Format$(rcValor, "#########.00")
    Centavos = Right(BUFFER, 2)
    Tamanho = Len(BUFFER) - 3
    Passos = Tamanho
    Indice = 0
    Temporario = 0
    Resto = 0

    'Alteracao feita pelo Fernando - 20/10/00
    If rcValor = 0 Then             'Alterado
       gsFormatarExtenso = Empty       'Alterado
       Exit Function                   'Alterado
    End If                          'Alterado

    Do While (Tamanho > -1)
       Select Case (Tamanho Mod 3)
              Case 0
              If Mid$(BUFFER, 1, 1) = "," Then
                 If Right(BUFFER, 2) <> "00" And ExtParc <> "" Then
                 'Alteracao feito pelo Eduardo - 03/07/2000
                 If ExtParc = "Um" Then          'Alterado
                    ExtParc = ExtParc & " Real e "  'Alterado
                 Else                            'Alterado
                    ExtParc = ExtParc & " Reais e "
                 End If                          'Alterado
              ElseIf ExtParc <> "" Then
                  If ExtParc = "Um" Then          'Alterado
                     ExtParc = ExtParc & " Real  "   'Alterado
                   Else                            'Alterado
                     ExtParc = ExtParc & " Reais "
                   End If                          'Alterado
              End If
    
    Tamanho = 3
Else
If Tamanho > 0 Then
If Len(ExtParc) > 0 Then
ExtParc = ExtParc & Quantia(Tamanho / 3 + 1)
End If
If Mid$(BUFFER, 1, 1) <> "0" Then
I = Val(Mid$(BUFFER, 1, 1))
If (I = 1 And Mid$(BUFFER, 2, 2) = "00") Then
ExtParc = ExtParc & " Cem "
ElseIf (Mid$(BUFFER, 2, 2) <> "00") Then
ExtParc = ExtParc & Centenas(I) & " e "
ExtParc = ExtParc & Quantia(1)
Else
ExtParc = ExtParc & Centenas(I)
ExtParc = ExtParc & Quantia(1)
End If
End If
Else
I = Len(ExtParc) - 2
If Mid$(ExtParc, I, 1) = "d" Then
ExtParc = ExtParc
End If
If Centavos > 0 Then
'Alteracao feito pelo Eduardo - 13/07/2000
If Centavos = 1 Then            'Alteracao
ExtParc = ExtParc & " Centavo"  'Alteracao
Else                            'Alteracao
ExtParc = ExtParc & " Centavos"
End If                          'Alteracao
End If
End If
End If
Tamanho = Tamanho - 1
Case 1
If Mid$(BUFFER, 1, 1) <> "0" Then
I = Val(Mid$(BUFFER, 1, 1))
ExtParc = ExtParc & Unidades(I)
End If
If Mid$(BUFFER, 1, 1) = "0" And Passos = 1 Then
ExtParc = ExtParc
Else
If Tamanho = 1 And Len(ExtParc) = 1 Then
ExtParc = ExtParc & "Reais"
End If
End If
Tamanho = Tamanho - 1
Case 2
If Mid$(BUFFER, 1, 1) = "1" Then
I = Val(Mid$(BUFFER, 2, 1))
ExtParc = ExtParc & Dezenas(I)
BUFFER = Mid$(BUFFER, 2)
Tamanho = Tamanho - 2
Else
I = Val(Mid$(BUFFER, 1, 1))
ExtParc = ExtParc & DezenaEspecial(I)
If Mid$(BUFFER, 1, 1) <> "0" And Mid$(BUFFER, 2, 1) <> "0" Then
ExtParc = ExtParc & " e "
End If
Tamanho = Tamanho - 1
End If
End Select
BUFFER = Mid$(BUFFER, 2)
Loop
Mid$(ExtParc, 1, 1) = UCase$(Mid$(ExtParc, 1, 1))

'Retorna o valor por extenso
gsFormatarExtenso = ExtParc
End Function



