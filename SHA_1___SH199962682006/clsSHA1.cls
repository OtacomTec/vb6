VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSHA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       clsSHA1
' FILENAME:     clsSHA1.cls
' AUTHOR:       Philip Ciebiera
' CREATED:      06/08/2006
'
' DESCRIPTION:
' This class is used to generate a SHA1 'digest' or 'signature' of a string.
' The SHA1 algorithm is one of the industry standard methods for generating
' digital signatures. It is generically known as a digest, digital signature,
' one-way encryption, hash or checksum algorithm. A common use for SHA1 is
' for password encryption as it is one-way in nature, that does not mean that
' your passwords are not free from a dictionary attack. If you are using the
' routine for passwords, you can make it a little more secure by concatenating
' some known random characters to the password before you generate the signature
' and on subsequent tests, so even if a hacker knows you are using SHA1 for
' your passwords, the random characters will make it harder to dictionary attack.
'
' *** CAUTION ***
' See the comment attached to the SHA1 method below regarding use on systems
' with different character sets.
'
'*******************************************************************************
'
' NOTE:
' Most of the functions in this class were created by Phil Fresle for his
' SHA-256 implementation in VB. You can see his website at http://www.frez.co.uk.
' I've created the LRot function myself as well as the actual SHA1 implementation.
' So my thanks to him.
' -Phil
'
'*******************************************************************************
Option Explicit

Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long

Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE


'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Class_Initialize()
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
End Sub


'*******************************************************************************
' LShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' A left shift takes all the set binary bits and moves them left, in-filling
' with zeros in the vacated bits on the right. This function is equivalent to
' the << operator in Java and C++
'*******************************************************************************
Private Function LShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    '
    ' A shift of zero is no shift at all.
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
        
    ' A shift of 31 will result in the right most bit becoming the left most
    ' bit and all other bits being cleared
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
        
    ' A shift of less than zero or more than 31 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    ' If the left most bit that remains will end up in the negative bit
    ' position (&H80000000) we would end up with an overflow if we took the
    ' standard route. We need to strip the left most bit and add it back
    ' afterwards.
    If (lValue And m_l2Power(31 - iShiftBits)) Then
    
        ' (Value And OnBits(31 - (Shift + 1))) chops off the left most bits that
        ' we are shifting into, but also the left most bit we still want as this
        ' is going to end up in the negative bit marker position (&H80000000).
        ' After the multiplication/shift we Or the result with &H80000000 to
        ' turn the negative bit on.
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * _
            m_l2Power(iShiftBits)) Or &H80000000
    
    Else
    
        ' (Value And OnBits(31-Shift)) chops off the left most bits that we are
        ' shifting into so we do not get an overflow error when we do the
        ' multiplication/shift
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * _
            m_l2Power(iShiftBits))
        
    End If
End Function


'*******************************************************************************
' RShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' The right shift of an unsigned long integer involves shifting all the set bits
' to the right and in-filling on the left with zeros. This function is
' equivalent to the >>> operator in Java or the >> operator in C++ when used on
' an unsigned long.
'*******************************************************************************
Private Function RShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    '
    ' A shift of zero is no shift at all
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
        
    ' A shift of 31 will clear all bits and move the left most bit to the right
    ' most bit position
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
        
    ' A shift of less than zero or more than 31 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    ' We do not care about the top most bit or the final bit, the top most bit
    ' will be taken into account in the next stage, the final bit (whether it
    ' is an odd number or not) is being shifted into, so we do not give a jot
    ' about it
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    ' If the top most bit (&H80000000) was set we need to do things differently
    ' as in a normal VB signed long integer the top most bit is used to indicate
    ' the sign of the number, when it is set it is a negative number, so just
    ' deviding by a factor of 2 as above would not work.
    ' NOTE: (lValue And  &H80000000) is equivalent to (lValue < 0), you could
    ' get a very marginal speed improvement by changing the test to (lValue < 0)
    If (lValue And &H80000000) Then
        ' We take the value computed so far, and then add the left most negative
        ' bit after it has been shifted to the right the appropriate number of
        ' places
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function


'*******************************************************************************
' AddUnsigned (FUNCTION)
'
' PARAMETERS:
' (In) - lX - Long - First value
' (In) - lY - Long - Second value
'
' RETURN VALUE:
' Long - Result
'
' DESCRIPTION:
' Adds two potentially large unsigned numbers without overflowing
'*******************************************************************************
Private Function AddUnsigned(ByVal lX As Long, _
                             ByVal lY As Long) As Long
    Dim lX4     As Long
    Dim lY4     As Long
    Dim lX8     As Long
    Dim lY8     As Long
    Dim lResult As Long
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function


'*******************************************************************************
' LRot (FUNCTION)
'
' PARAMETERS:
' (In) - x - Long - Value to rotate
' (In) - y - Long - How many bits to rotate
'
' RETURN VALUE:
' Long - Result
'
' DESCRIPTION:
' Preforms a bitwise left rotation by y bits on number x
'*******************************************************************************
Private Function LRot(ByVal x As Long, ByVal n As Long) As Long
    LRot = LShift(x, n) Or RShift(x, (32 - n))
End Function


'*******************************************************************************
' ConvertToWordArray (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - String message
'
' RETURN VALUE:
' Long() - Converted message as long array
'
' DESCRIPTION:
' Takes the string message and puts it in a long array with padding according to
' the SHA1 rules (similar to MD5 routine).
'*******************************************************************************
Private Function ConvertToWordArray(sMessage As String) As Long()
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long
    
    Const MODULUS_BITS      As Long = 512
    Const CONGRUENT_BITS    As Long = 448
    
    lMessageLength = Len(sMessage)
    
    ' Get padded number of words. Message needs to be congruent to 448 bits,
    ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
    ' it must still have another 512 bits added. 512 bits = 64 bytes
    ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
    ' be a multiple of 16 (i.e. 16 * 4 (bytes) * 8 (bits))
    lNumberOfWords = (((lMessageLength + _
        ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ _
        (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * _
        (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    ' Combine each block of 4 bytes (ascii code of character) into one long
    ' value and store in the message. The high-order (most significant) bit of
    ' each byte is listed first. However, unlike MD5 we put the high-order
    ' (most significant) byte first in each word.
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        ' Each word is 4 bytes
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        
        ' NOTE: This is where we are using just the first byte of each unicode
        ' character, you may want to make the change here, or to the SHA256 method
        ' so it accepts a byte array.
        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
        
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
    ' bits stored in the last two words
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    ' Add a terminating 1 bit, all the rest of the bits to the end of the
    ' word array will default to zero
    lWordArray(lWordCount) = lWordArray(lWordCount) Or _
        LShift(&H80, lBytePosition)

    ' We put the length of the message in bits into the last two words, to get
    ' the length in bits we need to multiply by 8 (or left shift 3). This left
    ' shifted value is put in the last word. Any bits shifted off the left edge
    ' need to be put in the penultimate word, we can work out which bits by shifting
    ' right the length by 29 bits.
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function


'*******************************************************************************
' SHA1 (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - Message to digest
'
' RETURN VALUE:
' String - The digest
'
' DESCRIPTION:
' Takes a string and uses the SHA1 digest to produce a signature for it.
'
' NOTE: Due to the way in which the string is processed the routine assumes a
' single byte character set. VB passes unicode (2-byte) character strings, the
' ConvertToWordArray function works on the first byte for each character. This
' has been done this way for ease of use, to make the routine truely portable
' you could accept a byte array instead, it would then be up to the calling
' routine to make sure that the byte array is generated from their string in
' a manner consistent with the string type.
'*******************************************************************************
Public Function SHA1(sMessage As String) As String
    Dim HASH(4)         As Long
    Dim M()             As Long
    Dim W(79)           As Long
    Dim a, b, c, d, e   As Long
    Dim g, h, i, j      As Long
    Dim T1, T2          As Long
    
    ' Initial hash container values
    HASH(0) = &H67452301
    HASH(1) = &HEFCDAB89
    HASH(2) = &H98BADCFE
    HASH(3) = &H10325476
    HASH(4) = &HC3D2E1F0
    
    ' Preprocessing. Append padding bits and length and convert to words.
    M = ConvertToWordArray(sMessage)
    
    ' We process sixteen 32-bit words at a time, which is a 512-bit data block
    For i = 0 To UBound(M) Step 16
        ' Set inital values for the hash operators,
        ' this includes previous hash values.
        a = HASH(0)
        b = HASH(1)
        c = HASH(2)
        d = HASH(3)
        e = HASH(4)
        
        ' We grab the sixteen 32-bit words from our Message word array,
        ' this is the data we'll be working on.
        For g = 0 To 15
            W(g) = M(i + g)
        Next g
        
        ' These sixteen 32-bit words must now be extended through the
        ' initial hashing phase to eighty 32-bit words.
        For g = 16 To 79
            W(g) = LRot(W(g - 3) Xor W(g - 8) Xor W(g - 14) Xor W(g - 16), 1)
        Next g
        
        ' We now begin processing these eighty 32-bit words.
        For j = 0 To 79
            
            ' The processing below is as per SHA1's specification.
            If j <= 19 Then
                T1 = (b And c) Or ((Not b) And d)
                T2 = &H5A827999
            ElseIf j <= 39 Then
                T1 = b Xor c Xor d
                T2 = &H6ED9EBA1
            ElseIf j <= 59 Then
                T1 = (b And c) Or (b And d) Or (c And d)
                T2 = &H8F1BBCDC
            ElseIf j <= 79 Then
                T1 = b Xor c Xor d
                T2 = &HCA62C1D6
            End If
            
            ' For each word we process we run the below hashing function and
            ' set it equal to a, shifting the previous a's value down to b,
            ' so a becomes b, b becomes c, after a 30 Left Rotate of 30,
            ' c becomes d, d becomes e.
            h = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(LRot(a, 5), T1), e), T2), W(j))
            e = d
            d = c
            c = LRot(b, 30)
            b = a
            a = h
        Next j
            
        ' We now add the newley hashed values to the hash container.
        HASH(0) = AddUnsigned(a, HASH(0))
        HASH(1) = AddUnsigned(b, HASH(1))
        HASH(2) = AddUnsigned(c, HASH(2))
        HASH(3) = AddUnsigned(d, HASH(3))
        HASH(4) = AddUnsigned(e, HASH(4))
        
    Next i
    
    ' Output the 160-bit digest
    SHA1 = LCase(Right("00000000" & Hex(HASH(0)), 8) & _
        Right("00000000" & Hex(HASH(1)), 8) & _
        Right("00000000" & Hex(HASH(2)), 8) & _
        Right("00000000" & Hex(HASH(3)), 8) & _
        Right("00000000" & Hex(HASH(4)), 8))
End Function
