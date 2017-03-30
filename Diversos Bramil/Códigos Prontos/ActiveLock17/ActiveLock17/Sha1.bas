Attribute VB_Name = "SHA1"
Option Private Module
Option Explicit

' Author: Ian Lynagh
' Date  : 1998

Function BigAA1Add(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value1 = Space$(Abs(tempnum)) + value1
    ElseIf tempnum > 0 Then
        value2 = Space$(Abs(tempnum)) + value2
    End If

    tempnum = 0
    For loopit = Len(value1) To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(value1, loopit, 1)) + Val("&H" + Mid$(value2, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit
    If tempnum <> 0 Then
        valueans = Hex$(tempnum) + valueans
    End If

    BigAA1Add = valueans

End Function

Function BigAA1AND(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) And Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigAA1AND = valueans

End Function

Function BigAA1Mod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BigAA1Mod32Add = Right$(BigAA1Add(value1, value2), 8)

End Function

Function BigAA1NOT(ByVal value1 As String) As String
    Dim valueans As String
    Dim loopit As Integer

    value1 = Right$(value1, 8)
    value1 = String$(8 - Len(value1), "0") + value1
    For loopit = 1 To 8
        valueans = valueans + Hex$(15 Xor Val("&H" + Mid$(value1, loopit, 1)))
    Next loopit

    BigAA1NOT = valueans

End Function

Function BigAA1OR(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Or Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigAA1OR = valueans

End Function

Function BigAA1RotLeft(ByVal value1 As String, ByVal rots As Integer) As String
    Dim tempstr As String
    Dim loopit As Integer, loopinner As Integer, tempnum As Integer
    
    rots = rots Mod 32
    If rots = 0 Then
        BigAA1RotLeft = value1
        Exit Function
    End If
    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""
    
    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit

    tempstr = Mid$(value1, rots + 1) + Left$(value1, rots)
    
    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit
    
    BigAA1RotLeft = value1

End Function

Function BigAA1XOR(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Xor Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigAA1XOR = valueans

End Function

Function BigAA2Add(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String, temps1 As String, temps2 As String, tempstr As String
    Dim tempnum As Long

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value1 = String$(Abs(tempnum), "0") + value1
    ElseIf tempnum > 0 Then
        value2 = String$(tempnum, "0") + value2
    End If

    tempnum = 0
    While Len(value1) > 5
        temps1 = Right$(value1, 6)
        temps2 = Right$(value2, 6)
        tempnum = tempnum + Val("&H" + temps1 + "&") + Val("&H" + temps2 + "&")
        tempstr = Hex$(tempnum Mod 16777216)
        valueans = String$(6 - Len(tempstr), "0") + tempstr + valueans
        tempnum = Int(tempnum / 16777216)
        value1 = Left$(value1, Len(value1) - 6)
        value2 = Left$(value2, Len(value2) - 6)
    Wend
    tempnum = tempnum + Val("&H" + value1 + "&") + Val("&H" + value2 + "&")
    valueans = Hex$(tempnum Mod 16777216) + valueans

    BigAA2Add = valueans

End Function

Function BigAA2AND(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String, tempstr As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        value1 = Mid$(value1, tempnum + 1)
    End If

    While Len(value1) > 7
        tempstr = Hex$(Val("&H" + Left$(value1, 7) + "&") And Val("&H" + Left$(value2, 7) + "&"))
        valueans = valueans + String$(7 - Len(tempstr), "0") + tempstr
        value1 = Mid$(value1, 8)
        value2 = Mid$(value2, 8)
    Wend
    tempnum = Len(value1)
    tempstr = Hex$(Val("&H" + value1 + "&") And Val("&H" + value2 + "&"))
    valueans = valueans + String$(tempnum - Len(tempstr), "0") + tempstr

    BigAA2AND = valueans

End Function

Function BigAA2Mod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BigAA2Mod32Add = BigAA1Mod32Add(value1, value2)

End Function

Function BigAA2NOT(ByVal value1 As String) As String
    Dim valueans As String
    Dim loopit As Integer

    value1 = Right$(value1, 8)
    value1 = String$(8 - Len(value1), "0") + value1
    valueans = Hex$(65535 Xor Val("&H" + Right$(value1, 4) + "&"))
    valueans = String$(4 - Len(valueans), "0") + valueans
    valueans = Hex$(65535 Xor Val("&H" + Left$(value1, 4) + "&")) + valueans
    valueans = String$(8 - Len(valueans), "0") + valueans
    
    BigAA2NOT = valueans

End Function

Function BigAA2OR(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String, tempstr As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    While Len(value1) > 7
        tempstr = Hex$(Val("&H" + Left$(value1, 7) + "&") Or Val("&H" + Left$(value2, 7) + "&"))
        valueans = valueans + String$(7 - Len(tempstr), "0") + tempstr
        value1 = Mid$(value1, 8)
        value2 = Mid$(value2, 8)
    Wend
    tempnum = Len(value1)
    tempstr = Hex$(Val("&H" + value1 + "&") Or Val("&H" + value2 + "&"))
    valueans = valueans + String$(tempnum - Len(tempstr), "0") + tempstr

    BigAA2OR = valueans

End Function

Function BigAA2RotLeft(ByVal value1 As String, ByVal rots As Integer) As String
    BigAA2RotLeft = BigAA1RotLeft(value1, rots)

End Function

Function BigAA2XOR(ByVal value1 As String, ByVal value2 As String) As String
    Dim valueans As String, tempstr As String
    Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    While Len(value1) > 7
        tempstr = Hex$(Val("&H" + Left$(value1, 7) + "&") Xor Val("&H" + Left$(value2, 7) + "&"))
        valueans = valueans + String$(7 - Len(tempstr), "0") + tempstr
        value1 = Mid$(value1, 8)
        value2 = Mid$(value2, 8)
    Wend
    tempnum = Len(value1)
    tempstr = Hex$(Val("&H" + value1 + "&") Xor Val("&H" + value2 + "&"))
    valueans = valueans + String$(tempnum - Len(tempstr), "0") + tempstr
    
    BigAA2XOR = valueans

End Function

Function SHA1AA1Hash(ByVal hashthis As String) As String
    ReDim buf(0 To 4) As String
    ReDim in_(0 To 79) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, c As String, d As String, e As String
    Dim tempstr As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        tempstr = Chr$(tempnum Mod 256) + tempstr
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    hashthis = hashthis + tempstr
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    buf(4) = "c3d2e1f0"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
        e = buf(4)

        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 4 To 1 Step -1
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit

        For loopit = 16 To 79
            in_(loopit) = BigAA1RotLeft(BigAA1XOR(BigAA1XOR(BigAA1XOR(in_(loopit - 3), in_(loopit - 8)), in_(loopit - 14)), in_(loopit - 16)), 1)
        Next loopit

        For loopit = 0 To 19
            tempstr = BigAA1OR(BigAA1AND(b, c), BigAA1AND(BigAA1NOT(b), d))
            tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "5A827999"))))
            e = d
            d = c
            c = BigAA1RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 20 To 39
            tempstr = BigAA1XOR(BigAA1XOR(b, c), d)
            tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "6ED9EBA1"))))
            e = d
            d = c
            c = BigAA1RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 40 To 59
            tempstr = BigAA1OR(BigAA1OR(BigAA1AND(b, c), BigAA1AND(b, d)), BigAA1AND(c, d))
            tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "8F1BBCDC"))))
            e = d
            d = c
            c = BigAA1RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 60 To 79
            tempstr = BigAA1XOR(BigAA1XOR(b, c), d)
            tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "CA62C1D6"))))
            e = d
            d = c
            c = BigAA1RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit
        
        buf(0) = BigAA1Mod32Add(buf(0), a)
        buf(1) = BigAA1Mod32Add(buf(1), b)
        buf(2) = BigAA1Mod32Add(buf(2), c)
        buf(3) = BigAA1Mod32Add(buf(3), d)
        buf(4) = BigAA1Mod32Add(buf(4), e)
    
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 4
        For loopinner = 0 To 3
            hashthis = hashthis + Mid$(buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    SHA1AA1Hash = hashthis
    
End Function

Function SHA1AA2Hash(ByVal hashthis As String) As String
    ReDim buf(0 To 4) As String
    ReDim in_(0 To 79) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, c As String, d As String, e As String
    Dim tempstr As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        tempstr = Chr$(tempnum Mod 256) + tempstr
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    hashthis = hashthis + tempstr
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    buf(4) = "c3d2e1f0"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
        e = buf(4)

        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 4 To 1 Step -1
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit

        For loopit = 16 To 79
            in_(loopit) = BigAA2RotLeft(BigAA2XOR(BigAA2XOR(BigAA2XOR(in_(loopit - 3), in_(loopit - 8)), in_(loopit - 14)), in_(loopit - 16)), 1)
        Next loopit

        For loopit = 0 To 19
            tempstr = BigAA2OR(BigAA2AND(b, c), BigAA2AND(BigAA2NOT(b), d))
            tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "5A827999"))))
            e = d
            d = c
            c = BigAA2RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 20 To 39
            tempstr = BigAA2XOR(BigAA2XOR(b, c), d)
            tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "6ED9EBA1"))))
            e = d
            d = c
            c = BigAA2RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 40 To 59
            tempstr = BigAA2OR(BigAA2OR(BigAA2AND(b, c), BigAA2AND(b, d)), BigAA2AND(c, d))
            tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "8F1BBCDC"))))
            e = d
            d = c
            c = BigAA2RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 60 To 79
            tempstr = BigAA2XOR(BigAA2XOR(b, c), d)
            tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "CA62C1D6"))))
            e = d
            d = c
            c = BigAA2RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit
        
        buf(0) = BigAA2Mod32Add(buf(0), a)
        buf(1) = BigAA2Mod32Add(buf(1), b)
        buf(2) = BigAA2Mod32Add(buf(2), c)
        buf(3) = BigAA2Mod32Add(buf(3), d)
        buf(4) = BigAA2Mod32Add(buf(4), e)
    
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 4
        For loopinner = 0 To 3
            hashthis = hashthis + Mid$(buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    SHA1AA2Hash = hashthis
    
End Function

