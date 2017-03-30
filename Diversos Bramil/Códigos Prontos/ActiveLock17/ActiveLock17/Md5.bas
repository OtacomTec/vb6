Attribute VB_Name = "MD5"
Option Private Module
Option Explicit

' Author: Ian Lynagh
' Date  : 1998

Dim w1 As String, w2 As String, w3 As String, w4 As String

Function MD5AA1F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    MD5AA1F = BigAA1Mod32Add(BigAA1RotLeft(BigAA1Mod32Add(BigAA1Mod32Add(w, tempstr), BigAA1Mod32Add(in_, qdata)), rots), X)

End Function

Sub MD5AA1F1(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(z, BigAA1AND(X, BigAA1XOR(Y, z)))
    w = MD5AA1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AA1F2(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(Y, BigAA1AND(z, BigAA1XOR(X, Y)))
    w = MD5AA1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AA1F3(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(X, BigAA1XOR(Y, z))
    w = MD5AA1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AA1F4(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(Y, BigAA1OR(X, BigAA1NOT(z)))
    w = MD5AA1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Function MD5AA1Hash(ByVal hashthis As String) As String
    ReDim buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, c As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AA1F1 a, b, c, d, in_(0), "d76aa478", 7
        MD5AA1F1 d, a, b, c, in_(1), "e8c7b756", 12
        MD5AA1F1 c, d, a, b, in_(2), "242070db", 17
        MD5AA1F1 b, c, d, a, in_(3), "c1bdceee", 22
        MD5AA1F1 a, b, c, d, in_(4), "f57c0faf", 7
        MD5AA1F1 d, a, b, c, in_(5), "4787c62a", 12
        MD5AA1F1 c, d, a, b, in_(6), "a8304613", 17
        MD5AA1F1 b, c, d, a, in_(7), "fd469501", 22
        MD5AA1F1 a, b, c, d, in_(8), "698098d8", 7
        MD5AA1F1 d, a, b, c, in_(9), "8b44f7af", 12
        MD5AA1F1 c, d, a, b, in_(10), "ffff5bb1", 17
        MD5AA1F1 b, c, d, a, in_(11), "895cd7be", 22
        MD5AA1F1 a, b, c, d, in_(12), "6b901122", 7
        MD5AA1F1 d, a, b, c, in_(13), "fd987193", 12
        MD5AA1F1 c, d, a, b, in_(14), "a679438e", 17
        MD5AA1F1 b, c, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AA1F2 a, b, c, d, in_(1), "f61e2562", 5
        MD5AA1F2 d, a, b, c, in_(6), "c040b340", 9
        MD5AA1F2 c, d, a, b, in_(11), "265e5a51", 14
        MD5AA1F2 b, c, d, a, in_(0), "e9b6c7aa", 20
        MD5AA1F2 a, b, c, d, in_(5), "d62f105d", 5
        MD5AA1F2 d, a, b, c, in_(10), "02441453", 9
        MD5AA1F2 c, d, a, b, in_(15), "d8a1e681", 14
        MD5AA1F2 b, c, d, a, in_(4), "e7d3fbc8", 20
        MD5AA1F2 a, b, c, d, in_(9), "21e1cde6", 5
        MD5AA1F2 d, a, b, c, in_(14), "c33707d6", 9
        MD5AA1F2 c, d, a, b, in_(3), "f4d50d87", 14
        MD5AA1F2 b, c, d, a, in_(8), "455a14ed", 20
        MD5AA1F2 a, b, c, d, in_(13), "a9e3e905", 5
        MD5AA1F2 d, a, b, c, in_(2), "fcefa3f8", 9
        MD5AA1F2 c, d, a, b, in_(7), "676f02d9", 14
        MD5AA1F2 b, c, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AA1F3 a, b, c, d, in_(5), "fffa3942", 4
        MD5AA1F3 d, a, b, c, in_(8), "8771f681", 11
        MD5AA1F3 c, d, a, b, in_(11), "6d9d6122", 16
        MD5AA1F3 b, c, d, a, in_(14), "fde5380c", 23
        MD5AA1F3 a, b, c, d, in_(1), "a4beea44", 4
        MD5AA1F3 d, a, b, c, in_(4), "4bdecfa9", 11
        MD5AA1F3 c, d, a, b, in_(7), "f6bb4b60", 16
        MD5AA1F3 b, c, d, a, in_(10), "bebfbc70", 23
        MD5AA1F3 a, b, c, d, in_(13), "289b7ec6", 4
        MD5AA1F3 d, a, b, c, in_(0), "eaa127fa", 11
        MD5AA1F3 c, d, a, b, in_(3), "d4ef3085", 16
        MD5AA1F3 b, c, d, a, in_(6), "04881d05", 23
        MD5AA1F3 a, b, c, d, in_(9), "d9d4d039", 4
        MD5AA1F3 d, a, b, c, in_(12), "e6db99e5", 11
        MD5AA1F3 c, d, a, b, in_(15), "1fa27cf8", 16
        MD5AA1F3 b, c, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AA1F4 a, b, c, d, in_(0), "f4292244", 6
        MD5AA1F4 d, a, b, c, in_(7), "432aff97", 10
        MD5AA1F4 c, d, a, b, in_(14), "ab9423a7", 15
        MD5AA1F4 b, c, d, a, in_(5), "fc93a039", 21
        MD5AA1F4 a, b, c, d, in_(12), "655b59c3", 6
        MD5AA1F4 d, a, b, c, in_(3), "8f0ccc92", 10
        MD5AA1F4 c, d, a, b, in_(10), "ffeff47d", 15
        MD5AA1F4 b, c, d, a, in_(1), "85845dd1", 21
        MD5AA1F4 a, b, c, d, in_(8), "6fa87e4f", 6
        MD5AA1F4 d, a, b, c, in_(15), "fe2ce6e0", 10
        MD5AA1F4 c, d, a, b, in_(6), "a3014314", 15
        MD5AA1F4 b, c, d, a, in_(13), "4e0811a1", 21
        MD5AA1F4 a, b, c, d, in_(4), "f7537e82", 6
        MD5AA1F4 d, a, b, c, in_(11), "bd3af235", 10
        MD5AA1F4 c, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AA1F4 b, c, d, a, in_(9), "eb86d391", 21
    
        buf(0) = BigAA1Add(buf(0), a)
        buf(1) = BigAA1Add(buf(1), b)
        buf(2) = BigAA1Add(buf(2), c)
        buf(3) = BigAA1Add(buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AA1Hash = hashthis

End Function

Function MD5AA2F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    MD5AA2F = BigAA2Mod32Add(BigAA2RotLeft(BigAA2Mod32Add(BigAA2Mod32Add(w, tempstr), BigAA2Mod32Add(in_, qdata)), rots), X)

End Function

Sub MD5AA2F1(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(z, BigAA2AND(X, BigAA2XOR(Y, z)))
    w = MD5AA2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AA2F2(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(Y, BigAA2AND(z, BigAA2XOR(X, Y)))
    w = MD5AA2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AA2F3(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(X, BigAA2XOR(Y, z))
    w = MD5AA2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AA2F4(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(Y, BigAA2OR(X, BigAA2NOT(z)))
    w = MD5AA2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Function MD5AA2Hash(ByVal hashthis As String) As String
    ReDim buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, c As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AA2F1 a, b, c, d, in_(0), "d76aa478", 7
        MD5AA2F1 d, a, b, c, in_(1), "e8c7b756", 12
        MD5AA2F1 c, d, a, b, in_(2), "242070db", 17
        MD5AA2F1 b, c, d, a, in_(3), "c1bdceee", 22
        MD5AA2F1 a, b, c, d, in_(4), "f57c0faf", 7
        MD5AA2F1 d, a, b, c, in_(5), "4787c62a", 12
        MD5AA2F1 c, d, a, b, in_(6), "a8304613", 17
        MD5AA2F1 b, c, d, a, in_(7), "fd469501", 22
        MD5AA2F1 a, b, c, d, in_(8), "698098d8", 7
        MD5AA2F1 d, a, b, c, in_(9), "8b44f7af", 12
        MD5AA2F1 c, d, a, b, in_(10), "ffff5bb1", 17
        MD5AA2F1 b, c, d, a, in_(11), "895cd7be", 22
        MD5AA2F1 a, b, c, d, in_(12), "6b901122", 7
        MD5AA2F1 d, a, b, c, in_(13), "fd987193", 12
        MD5AA2F1 c, d, a, b, in_(14), "a679438e", 17
        MD5AA2F1 b, c, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AA2F2 a, b, c, d, in_(1), "f61e2562", 5
        MD5AA2F2 d, a, b, c, in_(6), "c040b340", 9
        MD5AA2F2 c, d, a, b, in_(11), "265e5a51", 14
        MD5AA2F2 b, c, d, a, in_(0), "e9b6c7aa", 20
        MD5AA2F2 a, b, c, d, in_(5), "d62f105d", 5
        MD5AA2F2 d, a, b, c, in_(10), "02441453", 9
        MD5AA2F2 c, d, a, b, in_(15), "d8a1e681", 14
        MD5AA2F2 b, c, d, a, in_(4), "e7d3fbc8", 20
        MD5AA2F2 a, b, c, d, in_(9), "21e1cde6", 5
        MD5AA2F2 d, a, b, c, in_(14), "c33707d6", 9
        MD5AA2F2 c, d, a, b, in_(3), "f4d50d87", 14
        MD5AA2F2 b, c, d, a, in_(8), "455a14ed", 20
        MD5AA2F2 a, b, c, d, in_(13), "a9e3e905", 5
        MD5AA2F2 d, a, b, c, in_(2), "fcefa3f8", 9
        MD5AA2F2 c, d, a, b, in_(7), "676f02d9", 14
        MD5AA2F2 b, c, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AA2F3 a, b, c, d, in_(5), "fffa3942", 4
        MD5AA2F3 d, a, b, c, in_(8), "8771f681", 11
        MD5AA2F3 c, d, a, b, in_(11), "6d9d6122", 16
        MD5AA2F3 b, c, d, a, in_(14), "fde5380c", 23
        MD5AA2F3 a, b, c, d, in_(1), "a4beea44", 4
        MD5AA2F3 d, a, b, c, in_(4), "4bdecfa9", 11
        MD5AA2F3 c, d, a, b, in_(7), "f6bb4b60", 16
        MD5AA2F3 b, c, d, a, in_(10), "bebfbc70", 23
        MD5AA2F3 a, b, c, d, in_(13), "289b7ec6", 4
        MD5AA2F3 d, a, b, c, in_(0), "eaa127fa", 11
        MD5AA2F3 c, d, a, b, in_(3), "d4ef3085", 16
        MD5AA2F3 b, c, d, a, in_(6), "04881d05", 23
        MD5AA2F3 a, b, c, d, in_(9), "d9d4d039", 4
        MD5AA2F3 d, a, b, c, in_(12), "e6db99e5", 11
        MD5AA2F3 c, d, a, b, in_(15), "1fa27cf8", 16
        MD5AA2F3 b, c, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AA2F4 a, b, c, d, in_(0), "f4292244", 6
        MD5AA2F4 d, a, b, c, in_(7), "432aff97", 10
        MD5AA2F4 c, d, a, b, in_(14), "ab9423a7", 15
        MD5AA2F4 b, c, d, a, in_(5), "fc93a039", 21
        MD5AA2F4 a, b, c, d, in_(12), "655b59c3", 6
        MD5AA2F4 d, a, b, c, in_(3), "8f0ccc92", 10
        MD5AA2F4 c, d, a, b, in_(10), "ffeff47d", 15
        MD5AA2F4 b, c, d, a, in_(1), "85845dd1", 21
        MD5AA2F4 a, b, c, d, in_(8), "6fa87e4f", 6
        MD5AA2F4 d, a, b, c, in_(15), "fe2ce6e0", 10
        MD5AA2F4 c, d, a, b, in_(6), "a3014314", 15
        MD5AA2F4 b, c, d, a, in_(13), "4e0811a1", 21
        MD5AA2F4 a, b, c, d, in_(4), "f7537e82", 6
        MD5AA2F4 d, a, b, c, in_(11), "bd3af235", 10
        MD5AA2F4 c, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AA2F4 b, c, d, a, in_(9), "eb86d391", 21
    
        buf(0) = BigAA2Add(buf(0), a)
        buf(1) = BigAA2Add(buf(1), b)
        buf(2) = BigAA2Add(buf(2), c)
        buf(3) = BigAA2Add(buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AA2Hash = hashthis

End Function

Function MD5AB1F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim valueans As String
    Dim loopit As Integer, tempnum As Integer

    w = String$(8 - Len(w), "0") + w
    tempstr = String$(8 - Len(tempstr), "0") + tempstr
    in_ = String$(8 - Len(in_), "0") + in_
    qdata = String$(8 - Len(qdata), "0") + qdata

    For loopit = 8 To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(w, loopit, 1)) + Val("&H" + Mid$(tempstr, loopit, 1)) + Val("&H" + Mid$(in_, loopit, 1)) + Val("&H" + Mid$(qdata, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit

    MD5AB1F = BigAA1Mod32Add(BigAA1RotLeft(valueans, rots), X)
                   
End Function

Sub MD5AB1F1(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer, tempnum As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(Y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(zn Xor (xn And (yn Xor zn)))
    Next loopit

    w = MD5AB1F(tempstr, w, X, Y, z, in_, qdata, rots)
    
End Sub

Sub MD5AB1F2(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer, tempnum As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(Y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(yn Xor (zn And (xn Xor yn)))
    Next loopit

    w = MD5AB1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AB1F3(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer, tempnum As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(Y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(zn Xor xn Xor yn)
    Next loopit

    w = MD5AB1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AB1F4(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer, tempnum As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(Y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(yn Xor (xn Or (15 Xor zn)))
    Next loopit

    w = MD5AB1F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Function MD5AB1Hash(ByVal hashthis As String) As String
    ReDim buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, c As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AB1F1 a, b, c, d, in_(0), "d76aa478", 7
        MD5AB1F1 d, a, b, c, in_(1), "e8c7b756", 12
        MD5AB1F1 c, d, a, b, in_(2), "242070db", 17
        MD5AB1F1 b, c, d, a, in_(3), "c1bdceee", 22
        MD5AB1F1 a, b, c, d, in_(4), "f57c0faf", 7
        MD5AB1F1 d, a, b, c, in_(5), "4787c62a", 12
        MD5AB1F1 c, d, a, b, in_(6), "a8304613", 17
        MD5AB1F1 b, c, d, a, in_(7), "fd469501", 22
        MD5AB1F1 a, b, c, d, in_(8), "698098d8", 7
        MD5AB1F1 d, a, b, c, in_(9), "8b44f7af", 12
        MD5AB1F1 c, d, a, b, in_(10), "ffff5bb1", 17
        MD5AB1F1 b, c, d, a, in_(11), "895cd7be", 22
        MD5AB1F1 a, b, c, d, in_(12), "6b901122", 7
        MD5AB1F1 d, a, b, c, in_(13), "fd987193", 12
        MD5AB1F1 c, d, a, b, in_(14), "a679438e", 17
        MD5AB1F1 b, c, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AB1F2 a, b, c, d, in_(1), "f61e2562", 5
        MD5AB1F2 d, a, b, c, in_(6), "c040b340", 9
        MD5AB1F2 c, d, a, b, in_(11), "265e5a51", 14
        MD5AB1F2 b, c, d, a, in_(0), "e9b6c7aa", 20
        MD5AB1F2 a, b, c, d, in_(5), "d62f105d", 5
        MD5AB1F2 d, a, b, c, in_(10), "02441453", 9
        MD5AB1F2 c, d, a, b, in_(15), "d8a1e681", 14
        MD5AB1F2 b, c, d, a, in_(4), "e7d3fbc8", 20
        MD5AB1F2 a, b, c, d, in_(9), "21e1cde6", 5
        MD5AB1F2 d, a, b, c, in_(14), "c33707d6", 9
        MD5AB1F2 c, d, a, b, in_(3), "f4d50d87", 14
        MD5AB1F2 b, c, d, a, in_(8), "455a14ed", 20
        MD5AB1F2 a, b, c, d, in_(13), "a9e3e905", 5
        MD5AB1F2 d, a, b, c, in_(2), "fcefa3f8", 9
        MD5AB1F2 c, d, a, b, in_(7), "676f02d9", 14
        MD5AB1F2 b, c, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AB1F3 a, b, c, d, in_(5), "fffa3942", 4
        MD5AB1F3 d, a, b, c, in_(8), "8771f681", 11
        MD5AB1F3 c, d, a, b, in_(11), "6d9d6122", 16
        MD5AB1F3 b, c, d, a, in_(14), "fde5380c", 23
        MD5AB1F3 a, b, c, d, in_(1), "a4beea44", 4
        MD5AB1F3 d, a, b, c, in_(4), "4bdecfa9", 11
        MD5AB1F3 c, d, a, b, in_(7), "f6bb4b60", 16
        MD5AB1F3 b, c, d, a, in_(10), "bebfbc70", 23
        MD5AB1F3 a, b, c, d, in_(13), "289b7ec6", 4
        MD5AB1F3 d, a, b, c, in_(0), "eaa127fa", 11
        MD5AB1F3 c, d, a, b, in_(3), "d4ef3085", 16
        MD5AB1F3 b, c, d, a, in_(6), "04881d05", 23
        MD5AB1F3 a, b, c, d, in_(9), "d9d4d039", 4
        MD5AB1F3 d, a, b, c, in_(12), "e6db99e5", 11
        MD5AB1F3 c, d, a, b, in_(15), "1fa27cf8", 16
        MD5AB1F3 b, c, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AB1F4 a, b, c, d, in_(0), "f4292244", 6
        MD5AB1F4 d, a, b, c, in_(7), "432aff97", 10
        MD5AB1F4 c, d, a, b, in_(14), "ab9423a7", 15
        MD5AB1F4 b, c, d, a, in_(5), "fc93a039", 21
        MD5AB1F4 a, b, c, d, in_(12), "655b59c3", 6
        MD5AB1F4 d, a, b, c, in_(3), "8f0ccc92", 10
        MD5AB1F4 c, d, a, b, in_(10), "ffeff47d", 15
        MD5AB1F4 b, c, d, a, in_(1), "85845dd1", 21
        MD5AB1F4 a, b, c, d, in_(8), "6fa87e4f", 6
        MD5AB1F4 d, a, b, c, in_(15), "fe2ce6e0", 10
        MD5AB1F4 c, d, a, b, in_(6), "a3014314", 15
        MD5AB1F4 b, c, d, a, in_(13), "4e0811a1", 21
        MD5AB1F4 a, b, c, d, in_(4), "f7537e82", 6
        MD5AB1F4 d, a, b, c, in_(11), "bd3af235", 10
        MD5AB1F4 c, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AB1F4 b, c, d, a, in_(9), "eb86d391", 21
    
        buf(0) = BigAA1Add(buf(0), a)
        buf(1) = BigAA1Add(buf(1), b)
        buf(2) = BigAA1Add(buf(2), c)
        buf(3) = BigAA1Add(buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AB1Hash = hashthis

End Function

Function MD5AB2F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim valueans As String, tempvalstr  As String
    Dim loopit As Integer, tempnum As Long
    Dim temps1 As String, temps2 As String, temps3 As String, temps4 As String

    w = String$(8 - Len(w), "0") + w
    tempstr = String$(8 - Len(tempstr), "0") + tempstr
    in_ = String$(8 - Len(in_), "0") + in_
    qdata = String$(8 - Len(qdata), "0") + qdata

    temps1 = Right$(w, 5)
    temps2 = Right$(tempstr, 5)
    temps3 = Right$(in_, 5)
    temps4 = Right$(qdata, 5)
    tempnum = Val("&H" + temps1 + "&") + Val("&H" + temps2 + "&") + Val("&H" + temps3 + "&") + Val("&H" + temps4 + "&")
    tempvalstr = Hex$(tempnum Mod 1048576)
    valueans = String$(5 - Len(tempvalstr), "0") + tempvalstr + valueans
    tempnum = Int(tempnum / 1048576)
    temps1 = Left$(w, 3)
    temps2 = Left$(tempstr, 3)
    temps3 = Left$(in_, 3)
    temps4 = Left$(qdata, 3)
    tempnum = tempnum + Val("&H" + temps1 + "&") + Val("&H" + temps2 + "&") + Val("&H" + temps3 + "&") + Val("&H" + temps4 + "&")
    valueans = Hex$(tempnum) + valueans

    MD5AB2F = BigAA2Mod32Add(BigAA2RotLeft(valueans, rots), X)
                   

End Function

Sub MD5AB2F1(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(Y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(zn Xor (xn And (yn Xor zn)))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(Y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(zn Xor (xn And (yn Xor zn))) + tempstr
    
    w = MD5AB2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AB2F2(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(Y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(yn Xor (zn And (xn Xor yn)))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(Y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(yn Xor (zn And (xn Xor yn))) + tempstr
    
    w = MD5AB2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AB2F3(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(Y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(xn Xor (yn Xor zn))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(Y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(xn Xor (yn Xor zn)) + tempstr
    
    w = MD5AB2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Sub MD5AB2F4(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    Y = String$(8 - Len(Y), "0") + Y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(Y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(yn Xor (xn Or (1048575 Xor zn)))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(Y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(yn Xor (xn Or (4095 Xor zn))) + tempstr
    
    w = MD5AB2F(tempstr, w, X, Y, z, in_, qdata, rots)

End Sub

Function MD5AB2Hash(ByVal hashthis As String) As String
    ReDim buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, c As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AB2F1 a, b, c, d, in_(0), "d76aa478", 7
        MD5AB2F1 d, a, b, c, in_(1), "e8c7b756", 12
        MD5AB2F1 c, d, a, b, in_(2), "242070db", 17
        MD5AB2F1 b, c, d, a, in_(3), "c1bdceee", 22
        MD5AB2F1 a, b, c, d, in_(4), "f57c0faf", 7
        MD5AB2F1 d, a, b, c, in_(5), "4787c62a", 12
        MD5AB2F1 c, d, a, b, in_(6), "a8304613", 17
        MD5AB2F1 b, c, d, a, in_(7), "fd469501", 22
        MD5AB2F1 a, b, c, d, in_(8), "698098d8", 7
        MD5AB2F1 d, a, b, c, in_(9), "8b44f7af", 12
        MD5AB2F1 c, d, a, b, in_(10), "ffff5bb1", 17
        MD5AB2F1 b, c, d, a, in_(11), "895cd7be", 22
        MD5AB2F1 a, b, c, d, in_(12), "6b901122", 7
        MD5AB2F1 d, a, b, c, in_(13), "fd987193", 12
        MD5AB2F1 c, d, a, b, in_(14), "a679438e", 17
        MD5AB2F1 b, c, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AB2F2 a, b, c, d, in_(1), "f61e2562", 5
        MD5AB2F2 d, a, b, c, in_(6), "c040b340", 9
        MD5AB2F2 c, d, a, b, in_(11), "265e5a51", 14
        MD5AB2F2 b, c, d, a, in_(0), "e9b6c7aa", 20
        MD5AB2F2 a, b, c, d, in_(5), "d62f105d", 5
        MD5AB2F2 d, a, b, c, in_(10), "02441453", 9
        MD5AB2F2 c, d, a, b, in_(15), "d8a1e681", 14
        MD5AB2F2 b, c, d, a, in_(4), "e7d3fbc8", 20
        MD5AB2F2 a, b, c, d, in_(9), "21e1cde6", 5
        MD5AB2F2 d, a, b, c, in_(14), "c33707d6", 9
        MD5AB2F2 c, d, a, b, in_(3), "f4d50d87", 14
        MD5AB2F2 b, c, d, a, in_(8), "455a14ed", 20
        MD5AB2F2 a, b, c, d, in_(13), "a9e3e905", 5
        MD5AB2F2 d, a, b, c, in_(2), "fcefa3f8", 9
        MD5AB2F2 c, d, a, b, in_(7), "676f02d9", 14
        MD5AB2F2 b, c, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AB2F3 a, b, c, d, in_(5), "fffa3942", 4
        MD5AB2F3 d, a, b, c, in_(8), "8771f681", 11
        MD5AB2F3 c, d, a, b, in_(11), "6d9d6122", 16
        MD5AB2F3 b, c, d, a, in_(14), "fde5380c", 23
        MD5AB2F3 a, b, c, d, in_(1), "a4beea44", 4
        MD5AB2F3 d, a, b, c, in_(4), "4bdecfa9", 11
        MD5AB2F3 c, d, a, b, in_(7), "f6bb4b60", 16
        MD5AB2F3 b, c, d, a, in_(10), "bebfbc70", 23
        MD5AB2F3 a, b, c, d, in_(13), "289b7ec6", 4
        MD5AB2F3 d, a, b, c, in_(0), "eaa127fa", 11
        MD5AB2F3 c, d, a, b, in_(3), "d4ef3085", 16
        MD5AB2F3 b, c, d, a, in_(6), "04881d05", 23
        MD5AB2F3 a, b, c, d, in_(9), "d9d4d039", 4
        MD5AB2F3 d, a, b, c, in_(12), "e6db99e5", 11
        MD5AB2F3 c, d, a, b, in_(15), "1fa27cf8", 16
        MD5AB2F3 b, c, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AB2F4 a, b, c, d, in_(0), "f4292244", 6
        MD5AB2F4 d, a, b, c, in_(7), "432aff97", 10
        MD5AB2F4 c, d, a, b, in_(14), "ab9423a7", 15
        MD5AB2F4 b, c, d, a, in_(5), "fc93a039", 21
        MD5AB2F4 a, b, c, d, in_(12), "655b59c3", 6
        MD5AB2F4 d, a, b, c, in_(3), "8f0ccc92", 10
        MD5AB2F4 c, d, a, b, in_(10), "ffeff47d", 15
        MD5AB2F4 b, c, d, a, in_(1), "85845dd1", 21
        MD5AB2F4 a, b, c, d, in_(8), "6fa87e4f", 6
        MD5AB2F4 d, a, b, c, in_(15), "fe2ce6e0", 10
        MD5AB2F4 c, d, a, b, in_(6), "a3014314", 15
        MD5AB2F4 b, c, d, a, in_(13), "4e0811a1", 21
        MD5AB2F4 a, b, c, d, in_(4), "f7537e82", 6
        MD5AB2F4 d, a, b, c, in_(11), "bd3af235", 10
        MD5AB2F4 c, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AB2F4 b, c, d, a, in_(9), "eb86d391", 21
    
        buf(0) = BigAA2Add(buf(0), a)
        buf(1) = BigAA2Add(buf(1), b)
        buf(2) = BigAA2Add(buf(2), c)
        buf(3) = BigAA2Add(buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AB2Hash = hashthis

End Function

