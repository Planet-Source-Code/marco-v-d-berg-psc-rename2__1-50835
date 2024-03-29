Attribute VB_Name = "Mod_Inflate64"
Option Explicit
'This code is created by Marco v/d Berg
'This inflate routine is the same as used by zip and could be used
'to make a ZIP decompressor

Private Type CodesType
    Lenght As Integer
    Code As Long
End Type

Private OutStream() As Byte
Private OutPos As Long
Private InStream() As Byte
Private Inpos As Long
Private ByteBuff As Long
Private BitNum As Integer
Private BitMask(16) As Long
Private BitVal(16) As Long
Private LC(31) As CodesType
Private DC(31) As CodesType
Private LitLen() As CodesType           'Literal/length tree
Private Dist() As CodesType             'Distance tree
Private LenOrder(18) As Integer
Private MinLLenght As Integer           'Minimum length used in literal/lenght codes
Private MaxLLenght As Integer           'Maximum length used in literal/lenght codes
Private MinDLenght As Integer           'Minimum length used in distance codes
Private MaxDLenght As Integer           'Maximum length used in distance codes

Public Function Inflate(ByteArray() As Byte, Optional UncompressedSize As Long = 1000, Optional ZIP64 As Boolean = False) As Long
    Dim IsLastBlock As Boolean
    Dim CompType As Integer
    Dim DoInflate As Boolean
    Dim Char As Long
    Dim Nubits As Integer
    Dim L1 As Long
    Dim L2 As Long
    Dim X As Long
'Copy local array to global array
    InStream = ByteArray
'    Erase ByteArray                                 'clear memory
'Init global variables
    Call Init_Inflate(UncompressedSize)
    Do
        IsLastBlock = (GetBits(1) = 1)
        CompType = GetBits(2)
        Select Case CompType
        Case 0              'file is stored
            'input position is already set 1 position further so whe need to line up next byte
            L1 = InStream(Inpos) + CLng(InStream(Inpos + 1)) * 256
            Inpos = Inpos + 2
            L2 = InStream(Inpos) + CLng(InStream(Inpos + 1)) * 256
            Inpos = Inpos + 2
            If L1 - (Not (L2) And &HFFFF&) Then Inflate = -2
            For X = 1 To L1
                Call PutByte(InStream(Inpos))
                Inpos = Inpos + 1
            Next
            DoInflate = False
        Case 1
            Call Create_Static_Tree
            DoInflate = True
        Case 2
            Call Create_Dynamic_Tree
            DoInflate = True
        Case 3
            Inflate = -1
            DoInflate = False
        End Select
        If DoInflate Then
            Do
                'read minimum nuber of bits to speed things up
                Char = Bit_Reverse(GetBits(MinLLenght), MinLLenght)
                Nubits = MinLLenght
                Do While LitLen(Char).Lenght <> Nubits
                    Char = Char + Char + GetBits(1)
                    Nubits = Nubits + 1
                Loop
                Char = LitLen(Char).Code
                If Char < 256 Then
                    Call PutByte(CByte(Char))
                ElseIf Char > 256 Then
                    Char = Char - 257
                    L1 = LC(Char).Code + GetBits(LC(Char).Lenght)
                    If L1 = 258 Then
                        If ZIP64 Then
                            L1 = GetBits(16) + 3
                        End If
                    End If
                    Char = Bit_Reverse(GetBits(MinDLenght), MinDLenght)
                    Nubits = MinDLenght
                    Do While Dist(Char).Lenght <> Nubits
                        Char = Char + Char + GetBits(1)
                        Nubits = Nubits + 1
                    Loop
                    Char = Dist(Char).Code
                    L2 = DC(Char).Code + GetBits(DC(Char).Lenght)
                    For X = 1 To L1
                        Char = OutStream(OutPos - L2)
                        Call PutByte(CByte(Char))
                    Next
                End If
            Loop While Char <> 256
        End If
    Loop While Not IsLastBlock
    If OutPos > 0 Then
        ReDim Preserve OutStream(OutPos - 1)
    Else
        Erase OutStream
    End If
'Clear memory
    Erase InStream
    Erase BitMask
    Erase BitVal
    Erase LC
    Erase DC
    Erase LitLen
    Erase Dist
    Erase LenOrder
    ByteArray = OutStream
End Function

Private Function Create_Static_Tree()
    Dim X As Long
    Dim Lenght(287) As Long
    For X = 0 To 143: Lenght(X) = 8: Next
    For X = 144 To 255: Lenght(X) = 9: Next
    For X = 256 To 279: Lenght(X) = 7: Next
    For X = 280 To 287: Lenght(X) = 8: Next
    If Create_Codes(LitLen, Lenght, 287, MaxLLenght, MinLLenght) <> 0 Then
        Create_Static_Tree = -1
        Exit Function
    End If
    For X = 0 To 31: Lenght(X) = 5: Next
    Create_Static_Tree = Create_Codes(Dist, Lenght, 31, MaxDLenght, MinDLenght)
End Function

Private Function Create_Dynamic_Tree() As Integer
    Dim Lenght() As Long
    Dim Bl_Tree() As CodesType
    Dim MinBL As Integer
    Dim MaxBL As Integer
    Dim NumLen As Long
    Dim Numdis As Long
    Dim NumCod As Long
    Dim Char As Integer
    Dim Nubits As Long
    Dim LN As Integer
    Dim Pos As Integer
    Dim X As Long
    NumLen = GetBits(5) + 257
    Numdis = GetBits(5) + 1
    NumCod = GetBits(4) + 4
    ReDim Lenght(18)
    For X = 0 To NumCod - 1
        Lenght(LenOrder(X)) = GetBits(3)
    Next
    For X = NumCod To 18
        Lenght(LenOrder(X)) = 0
    Next
    If Create_Codes(Bl_Tree, Lenght, 18, MaxBL, MinBL) <> 0 Then
        Create_Dynamic_Tree = -1
        Exit Function
    End If
    ReDim Lenght(NumLen + Numdis)
    Pos = 0
    Do While Pos < NumLen + Numdis
        Char = Bit_Reverse(GetBits(MinBL), MinBL)
        Nubits = MinBL
        Do While Bl_Tree(Char).Lenght <> Nubits
            Char = Char + Char + GetBits(1)
            Nubits = Nubits + 1
        Loop
        Char = Bl_Tree(Char).Code
        If Char < 16 Then
            Lenght(Pos) = Char
            Pos = Pos + 1
        Else
            If Char = 16 Then
                If Pos = 0 Then Create_Dynamic_Tree = -5: Exit Function 'no last lenght
                LN = Lenght(Pos - 1)
                Char = 3 + GetBits(2)
            ElseIf Char = 17 Then
                Char = 3 + GetBits(3)
                LN = 0
            Else
                Char = 11 + GetBits(7)
                LN = 0
            End If
            If Pos + Char > NumLen + Numdis Then
                Create_Dynamic_Tree = -6                    'to many lenghts
                Exit Function
            End If
            Do While Char > 0
                Char = Char - 1
                Lenght(Pos) = LN
                Pos = Pos + 1
            Loop
        End If
    Loop
    If Create_Codes(LitLen, Lenght, NumLen - 1, MaxLLenght, MinLLenght) <> 0 Then
        Create_Dynamic_Tree = -1
        Exit Function
    End If
    For X = 0 To Numdis
        Lenght(X) = Lenght(X + NumLen)
    Next
    Create_Dynamic_Tree = Create_Codes(Dist, Lenght, Numdis - 1, MaxDLenght, MinDLenght)
End Function


Private Sub Init_Inflate(UncompressedSize As Long)
    Dim Temp()
    Dim X As Long
    ReDim OutStream(UncompressedSize)
    Erase LitLen
    Erase Dist
    Erase DC
    Erase LC
    'Create the read order array
    Temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    For X = 0 To UBound(Temp): LenOrder(X) = Temp(X): Next
    'Create the Start lenghts array
    Temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For X = 0 To UBound(Temp): LC(X).Code = Temp(X): Next
    'Create the Extra lenght bits array
    Temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For X = 0 To UBound(Temp): LC(X).Lenght = Temp(X): Next
    'Create the distance code array
    Temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    For X = 0 To UBound(Temp): DC(X).Code = Temp(X): Next
    'Create the extra bits distance codes
    Temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For X = 0 To UBound(Temp): DC(X).Lenght = Temp(X): Next

    OutPos = 0
    Inpos = 0
    For X = 0 To 16
        BitMask(X) = 2 ^ X - 1
        BitVal(X) = 2 ^ X
    Next
    ByteBuff = 0
    BitNum = 0
End Sub

Private Function Create_Codes(tree() As CodesType, Lenghts() As Long, NumCodes As Long, MaxBits As Integer, Minbits As Integer) As Integer
    Dim bits(16) As Long
    Dim next_code(16) As Long
    Dim Code As Long
    Dim LN As Long
    Dim X As Long
'retrieve the bitlenght count and minimum and maximum bitlenghts
    Minbits = 16
    For X = 0 To NumCodes
        bits(Lenghts(X)) = bits(Lenghts(X)) + 1
        If Lenghts(X) > MaxBits Then MaxBits = Lenghts(X)
        If Lenghts(X) < Minbits And Lenghts(X) > 0 Then Minbits = Lenghts(X)
    Next
    LN = 1
    For X = 1 To MaxBits
        LN = LN + LN
        LN = LN - bits(X)
        If LN < 0 Then Create_Codes = LN: Exit Function 'Over subscribe, Return negative
    Next
    Create_Codes = LN

    ReDim tree(2 ^ MaxBits - 1) 'set the right dimensions
    Code = 0
    bits(0) = 0
    For X = 1 To MaxBits
        Code = (Code + bits(X - 1)) * 2
        next_code(X) = Code
    Next
    For X = 0 To NumCodes
        LN = Lenghts(X)
        If LN <> 0 Then
            tree(next_code(LN)).Lenght = LN
            tree(next_code(LN)).Code = X
            next_code(LN) = next_code(LN) + 1
        End If
    Next
End Function


Private Sub PutByte(Char As Byte)
    If OutPos > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 1000)
    OutStream(OutPos) = Char
    OutPos = OutPos + 1
End Sub

Private Function GetBits(Numbits As Integer) As Long
    If BitNum < Numbits Then
        Do
            ByteBuff = ByteBuff + (InStream(Inpos) * BitVal(BitNum))
            BitNum = BitNum + 8
            Inpos = Inpos + 1
        Loop While BitNum < Numbits
    End If
    GetBits = ByteBuff And BitMask(Numbits)
    ByteBuff = Fix(ByteBuff / BitVal(Numbits))
    BitNum = BitNum - Numbits
End Function

Private Function Bit_Reverse(ByVal Value As Long, ByVal Numbits As Long)
    Do While Numbits > 0
        Bit_Reverse = Bit_Reverse * 2 + (Value And 1)
        Numbits = Numbits - 1
        Value = Fix(Value / 2)
    Loop
End Function

